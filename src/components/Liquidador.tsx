import { useEffect, useMemo, useState } from 'react';
import {
    leerArchivo,
    resumenBalance,
    resumenImpuestos,
    type Categoria,
    type CuentaBalance,
    type GrupoImpuesto,
    type Linea,
    type Persona,
} from '@/lib/retencion';

type Separador = ',' | '.';
type TipoArchivo = 'impuestos' | 'balance';

const formato = new Intl.NumberFormat('es-CO', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
});

const redondear = (n: number) => Math.round(n * 100) / 100;

function paraCopiar(n: number, separador: Separador): string {
    const r = redondear(n);
    const s = Number.isInteger(r) ? String(r) : r.toFixed(2);
    return separador === ',' ? s.replace('.', ',') : s;
}

function leerPreferencia<T>(clave: string, porDefecto: T): T {
    try {
        const v = localStorage.getItem(clave);
        return v === null ? porDefecto : (JSON.parse(v) as T);
    } catch {
        return porDefecto;
    }
}

function guardarPreferencia(clave: string, valor: unknown) {
    try {
        localStorage.setItem(clave, JSON.stringify(valor));
    } catch {
        /* sin almacenamiento: no pasa nada */
    }
}

const nombresCategoria: Record<Categoria, string> = {
    general: 'Servicios / honorarios (PN o PJ)',
    arriendo: 'Arriendo',
    compras: 'Compras',
    excluir: 'No incluir',
};

interface EstadoCarga {
    archivo: string | null;
    error: string | null;
    cargando: boolean;
}

const cargaVacia: EstadoCarga = { archivo: null, error: null, cargando: false };

export default function Liquidador() {
    const [grupos, setGrupos] = useState<GrupoImpuesto[] | null>(null);
    const [cuentas, setCuentas] = useState<CuentaBalance[] | null>(null);
    const [carga, setCarga] = useState<Record<TipoArchivo, EstadoCarga>>({
        impuestos: cargaVacia,
        balance: cargaVacia,
    });
    const [error, setError] = useState<string | null>(null);
    const [categorias, setCategorias] = useState<Record<string, Categoria>>({});
    const [personas, setPersonas] = useState<Record<string, Persona>>({});
    const [copiado, setCopiado] = useState<string | null>(null);
    const [separador, setSeparador] = useState<Separador>(() =>
        leerPreferencia<Separador>('separador', ','),
    );
    const [tarifa, setTarifa] = useState<number>(() =>
        leerPreferencia('tarifaAutorretencion', 1.1),
    );

    useEffect(() => guardarPreferencia('separador', separador), [separador]);
    useEffect(() => guardarPreferencia('tarifaAutorretencion', tarifa), [tarifa]);

    const actualizarCarga = (tipo: TipoArchivo, cambios: Partial<EstadoCarga>) =>
        setCarga((c) => ({ ...c, [tipo]: { ...c[tipo], ...cambios } }));

    async function cargar(tipo: TipoArchivo, archivo: File | undefined) {
        if (!archivo) return;
        actualizarCarga(tipo, { cargando: true, error: null });
        try {
            const hojas = await leerArchivo(archivo);
            if (tipo === 'impuestos') {
                if (!hojas.impuestos) {
                    actualizarCarga(tipo, {
                        cargando: false,
                        error: hojas.balance
                            ? 'Este archivo es el balance de prueba. Súbelo en la casilla del balance.'
                            : 'No encontré el informe de impuestos detallados (la columna “Base compras”).',
                    });
                    return;
                }
                setGrupos(hojas.impuestos);
                setCategorias({});
                setPersonas({});
            } else {
                if (!hojas.balance) {
                    actualizarCarga(tipo, {
                        cargando: false,
                        error: hojas.impuestos
                            ? 'Este archivo es el informe de impuestos detallados. Súbelo en la otra casilla.'
                            : 'No encontré el balance de prueba (las columnas “Código cuenta contable” y “Movimiento crédito”).',
                    });
                    return;
                }
                setCuentas(hojas.balance);
            }
            actualizarCarga(tipo, { cargando: false, archivo: archivo.name });
        } catch (e) {
            console.error(e);
            actualizarCarga(tipo, {
                cargando: false,
                error: 'No pude leer el archivo. ¿Es un Excel (.xlsx o .xls)?',
            });
        }
    }

    function quitar(tipo: TipoArchivo) {
        if (tipo === 'impuestos') setGrupos(null);
        else setCuentas(null);
        actualizarCarga(tipo, cargaVacia);
    }

    const impuestos = useMemo(
        () => (grupos ? resumenImpuestos(grupos, categorias, personas) : null),
        [grupos, categorias, personas],
    );

    const balance = useMemo(
        () => (cuentas ? resumenBalance(cuentas, tarifa) : null),
        [cuentas, tarifa],
    );

    const subtotalInforme = useMemo(() => {
        if (!grupos) return 0;
        return grupos
            .filter(
                (g) =>
                    (categorias[g.nombre] ?? g.categoriaSugerida) !== 'excluir',
            )
            .reduce((s, g) => s + g.subtotalRetencion, 0);
    }, [grupos, categorias]);

    async function copiar(clave: string, contenido: string) {
        try {
            await navigator.clipboard.writeText(contenido);
            setCopiado(clave);
            setTimeout(() => setCopiado((c) => (c === clave ? null : c)), 1500);
        } catch {
            setError('El navegador no dejó copiar. Selecciona el número y usa Ctrl+C.');
        }
    }

    const n = (v: number) => paraCopiar(v, separador);
    const filaTsv = (l: Linea) =>
        l.base === null ? n(l.retencion) : `${n(l.base)}\t${n(l.retencion)}`;

    return (
        <div className="mx-auto max-w-5xl px-4 py-10 text-slate-900">
            <header className="mb-8">
                <h1 className="text-3xl font-bold">Liquidador de retención</h1>
                <p className="mt-2 text-slate-600">
                    Sube por separado el informe de <b>impuestos detallados</b> y
                    el <b>balance de prueba</b> de Siigo. Los archivos se leen en
                    tu computador; no se envían a ningún lado.
                </p>
            </header>

            <div className="grid gap-4 md:grid-cols-2">
                <ZonaCarga
                    titulo="Impuestos detallados"
                    descripcion="Saca PN, PJ, arriendo y compras."
                    estado={carga.impuestos}
                    onArchivo={(f) => cargar('impuestos', f)}
                    onQuitar={() => quitar('impuestos')}
                />
                <ZonaCarga
                    titulo="Balance de prueba"
                    descripcion="Saca la autorretención (clase 4 × tarifa) y las cuentas 2365."
                    estado={carga.balance}
                    onArchivo={(f) => cargar('balance', f)}
                    onQuitar={() => quitar('balance')}
                />
            </div>

            <div className="mt-4 flex flex-wrap items-center gap-x-6 gap-y-2 text-sm text-slate-700">
                <label className="flex items-center gap-2">
                    Decimales al copiar:
                    <select
                        value={separador}
                        onChange={(e) => setSeparador(e.target.value as Separador)}
                        className="rounded-md border border-slate-300 bg-white px-2 py-1"
                    >
                        <option value=",">coma (1234,56)</option>
                        <option value=".">punto (1234.56)</option>
                    </select>
                </label>
                <label className="flex items-center gap-2">
                    Tarifa autorretención:
                    <input
                        type="number"
                        step="0.01"
                        min="0"
                        value={tarifa}
                        onChange={(e) => setTarifa(Number(e.target.value) || 0)}
                        className="w-20 rounded-md border border-slate-300 bg-white px-2 py-1"
                    />
                    %
                </label>
            </div>

            {error && (
                <p className="mt-4 rounded-lg bg-red-50 px-4 py-3 text-red-700">
                    {error}
                </p>
            )}

            {impuestos && (
                <Tarjeta
                    titulo="Impuestos detallados"
                    accion={
                        <BotonCopiar
                            hecho={copiado === 'imp-todo'}
                            onClick={() =>
                                copiar(
                                    'imp-todo',
                                    impuestos.lineas.map(filaTsv).join('\n'),
                                )
                            }
                        >
                            Copiar tabla
                        </BotonCopiar>
                    }
                >
                    <TablaLineas
                        lineas={impuestos.lineas}
                        total={impuestos.total}
                        copiado={copiado}
                        copiar={copiar}
                        n={n}
                        filaTsv={filaTsv}
                        prefijo="imp"
                    />
                    <Cuadre
                        etiqueta="Contra los subtotales del informe"
                        diferencia={impuestos.total.retencion - subtotalInforme}
                    />
                </Tarjeta>
            )}

            {balance && (
                <Tarjeta
                    titulo="Balance"
                    accion={
                        <BotonCopiar
                            hecho={copiado === 'bal-todo'}
                            onClick={() =>
                                copiar(
                                    'bal-todo',
                                    balance.lineas
                                        .map((l) => n(l.retencion))
                                        .join('\n'),
                                )
                            }
                        >
                            Copiar retenciones
                        </BotonCopiar>
                    }
                >
                    <TablaLineas
                        lineas={balance.lineas}
                        copiado={copiado}
                        copiar={copiar}
                        n={n}
                        filaTsv={filaTsv}
                        prefijo="bal"
                    />
                    <p className="mt-4 text-sm text-slate-600">
                        Total retenciones en el balance (cuenta 2365):{' '}
                        <b>{formato.format(balance.totalRetenciones2365)}</b>
                    </p>
                </Tarjeta>
            )}

            {impuestos && balance && (
                <Cuadre
                    etiqueta="Cruce: retenciones de la 2365 contra impuestos detallados"
                    diferencia={
                        balance.otrasRetenciones - impuestos.total.retencion
                    }
                />
            )}

            {grupos && (
                <Tarjeta titulo="Detalle (para revisar)">
                    <p className="mb-4 text-sm text-slate-600">
                        Si algo quedó mal clasificado, cambia la categoría del
                        grupo o haz clic en PN/PJ del tercero. Los totales de
                        arriba se recalculan solos.
                    </p>
                    <div className="space-y-6">
                        {grupos.map((g) => {
                            const cat = categorias[g.nombre] ?? g.categoriaSugerida;
                            return (
                                <div key={g.nombre}>
                                    <div className="flex flex-wrap items-center justify-between gap-2">
                                        <h3 className="font-semibold">{g.nombre}</h3>
                                        <select
                                            value={cat}
                                            onChange={(e) =>
                                                setCategorias((c) => ({
                                                    ...c,
                                                    [g.nombre]: e.target
                                                        .value as Categoria,
                                                }))
                                            }
                                            className="rounded-md border border-slate-300 bg-white px-2 py-1 text-sm"
                                        >
                                            {Object.entries(nombresCategoria).map(
                                                ([valor, nombre]) => (
                                                    <option key={valor} value={valor}>
                                                        {nombre}
                                                    </option>
                                                ),
                                            )}
                                        </select>
                                    </div>
                                    <div className="mt-2 overflow-x-auto">
                                        <table
                                            className={`w-full text-sm ${cat === 'excluir' ? 'opacity-40' : ''}`}
                                        >
                                            <tbody>
                                                {g.movimientos.map((m, i) => {
                                                    const persona =
                                                        personas[m.identificacion] ??
                                                        m.persona;
                                                    return (
                                                        <tr
                                                            key={`${m.comprobante}-${i}`}
                                                            className="border-t border-slate-100"
                                                        >
                                                            <td className="py-1 pr-3 whitespace-nowrap text-slate-500">
                                                                {m.comprobante}
                                                            </td>
                                                            <td className="py-1 pr-3">
                                                                {m.tercero}
                                                                <span className="ml-2 text-slate-400">
                                                                    {m.identificacion}
                                                                </span>
                                                            </td>
                                                            <td className="py-1 pr-3">
                                                                <button
                                                                    type="button"
                                                                    title="Cambiar entre persona natural y jurídica"
                                                                    onClick={() =>
                                                                        setPersonas((p) => ({
                                                                            ...p,
                                                                            [m.identificacion]:
                                                                                persona === 'PN'
                                                                                    ? 'PJ'
                                                                                    : 'PN',
                                                                        }))
                                                                    }
                                                                    className={`rounded px-2 py-0.5 text-xs font-semibold ${
                                                                        persona === 'PN'
                                                                            ? 'bg-sky-100 text-sky-800'
                                                                            : 'bg-violet-100 text-violet-800'
                                                                    }`}
                                                                >
                                                                    {persona}
                                                                </button>
                                                            </td>
                                                            <td className="py-1 pr-3 text-right tabular-nums">
                                                                {formato.format(m.base)}
                                                            </td>
                                                            <td className="py-1 text-right tabular-nums">
                                                                {formato.format(m.retencion)}
                                                            </td>
                                                        </tr>
                                                    );
                                                })}
                                            </tbody>
                                        </table>
                                    </div>
                                </div>
                            );
                        })}
                    </div>
                </Tarjeta>
            )}
        </div>
    );
}

function Tarjeta({
    titulo,
    accion,
    children,
}: {
    titulo: string;
    accion?: React.ReactNode;
    children: React.ReactNode;
}) {
    return (
        <section className="mt-8 rounded-xl border border-slate-200 bg-white p-5 shadow-sm">
            <div className="mb-4 flex items-center justify-between gap-4">
                <h2 className="text-xl font-bold">{titulo}</h2>
                {accion}
            </div>
            {children}
        </section>
    );
}

function ZonaCarga({
    titulo,
    descripcion,
    estado,
    onArchivo,
    onQuitar,
}: {
    titulo: string;
    descripcion: string;
    estado: EstadoCarga;
    onArchivo: (archivo: File | undefined) => void;
    onQuitar: () => void;
}) {
    const [arrastrando, setArrastrando] = useState(false);
    const cargado = estado.archivo !== null;

    return (
        <div>
            <label
                onDragOver={(e) => {
                    e.preventDefault();
                    setArrastrando(true);
                }}
                onDragLeave={() => setArrastrando(false)}
                onDrop={(e) => {
                    e.preventDefault();
                    setArrastrando(false);
                    onArchivo(e.dataTransfer.files[0]);
                }}
                className={`flex h-full cursor-pointer flex-col items-center justify-center rounded-xl border-2 border-dashed px-6 py-8 text-center transition-colors ${
                    arrastrando
                        ? 'border-emerald-500 bg-emerald-50'
                        : cargado
                          ? 'border-emerald-300 bg-emerald-50/50'
                          : 'border-slate-300 bg-slate-50 hover:border-slate-400'
                }`}
            >
                <span className="text-lg font-semibold">{titulo}</span>
                <span className="mt-1 text-sm text-slate-500">{descripcion}</span>
                <span className="mt-3 text-sm font-medium text-slate-700">
                    {estado.cargando
                        ? 'Leyendo…'
                        : cargado
                          ? `✓ ${estado.archivo}`
                          : 'Arrastra el Excel aquí o haz clic para buscarlo'}
                </span>
                <input
                    type="file"
                    accept=".xlsx,.xls,.xlsm"
                    className="sr-only"
                    onChange={(e) => {
                        onArchivo(e.target.files?.[0]);
                        e.target.value = '';
                    }}
                />
            </label>
            {cargado && (
                <button
                    type="button"
                    onClick={onQuitar}
                    className="mt-1 text-sm text-slate-500 underline hover:text-slate-800"
                >
                    Quitar archivo
                </button>
            )}
            {estado.error && (
                <p className="mt-2 rounded-lg bg-red-50 px-3 py-2 text-sm text-red-700">
                    {estado.error}
                </p>
            )}
        </div>
    );
}

function BotonCopiar({
    hecho,
    onClick,
    children,
    pequeno = false,
}: {
    hecho: boolean;
    onClick: () => void;
    children: React.ReactNode;
    pequeno?: boolean;
}) {
    return (
        <button
            type="button"
            onClick={onClick}
            className={`shrink-0 rounded-md font-medium transition-colors ${
                pequeno ? 'px-2 py-0.5 text-xs' : 'px-3 py-1.5 text-sm'
            } ${
                hecho
                    ? 'bg-emerald-600 text-white'
                    : 'bg-slate-900 text-white hover:bg-slate-700'
            }`}
        >
            {hecho ? 'Copiado ✓' : children}
        </button>
    );
}

function TablaLineas({
    lineas,
    total,
    copiado,
    copiar,
    n,
    filaTsv,
    prefijo,
}: {
    lineas: Linea[];
    total?: Linea;
    copiado: string | null;
    copiar: (clave: string, contenido: string) => void;
    n: (v: number) => string;
    filaTsv: (l: Linea) => string;
    prefijo: string;
}) {
    const celda = (clave: string, valor: number | null) =>
        valor === null ? (
            <span className="text-slate-300">—</span>
        ) : (
            <button
                type="button"
                title="Clic para copiar"
                onClick={() => copiar(clave, n(valor))}
                className={`rounded px-1.5 py-0.5 tabular-nums transition-colors ${
                    copiado === clave
                        ? 'bg-emerald-100 text-emerald-800'
                        : 'hover:bg-slate-100'
                }`}
            >
                {copiado === clave ? 'Copiado ✓' : formato.format(valor)}
            </button>
        );

    const fila = (l: Linea, negrita = false) => (
        <tr
            key={l.clave}
            className={`border-t border-slate-100 ${negrita ? 'font-bold' : ''}`}
        >
            <td className="py-2 pr-4">{l.concepto}</td>
            <td className="py-2 pr-4 text-right">
                {celda(`${prefijo}-${l.clave}-base`, l.base)}
            </td>
            <td className="py-2 pr-4 text-right">
                {celda(`${prefijo}-${l.clave}-ret`, l.retencion)}
            </td>
            <td className="py-2 text-right">
                <BotonCopiar
                    pequeno
                    hecho={copiado === `${prefijo}-${l.clave}-fila`}
                    onClick={() => copiar(`${prefijo}-${l.clave}-fila`, filaTsv(l))}
                >
                    Copiar fila
                </BotonCopiar>
            </td>
        </tr>
    );

    return (
        <div className="overflow-x-auto">
            <table className="w-full">
                <thead>
                    <tr className="text-left text-sm text-slate-500">
                        <th className="pb-2 font-medium">Concepto</th>
                        <th className="pb-2 text-right font-medium">Base</th>
                        <th className="pb-2 text-right font-medium">Retención</th>
                        <th />
                    </tr>
                </thead>
                <tbody>
                    {lineas.map((l) => fila(l))}
                    {total && fila(total, true)}
                </tbody>
            </table>
        </div>
    );
}

function Cuadre({ etiqueta, diferencia }: { etiqueta: string; diferencia: number }) {
    const cuadra = Math.abs(diferencia) < 0.5;
    return (
        <p
            className={`mt-4 rounded-lg px-4 py-2 text-sm ${
                cuadra ? 'bg-emerald-50 text-emerald-800' : 'bg-red-50 text-red-700'
            }`}
        >
            {etiqueta}:{' '}
            {cuadra
                ? 'cuadra ✓'
                : `diferencia de ${formato.format(diferencia)} — revisa el detalle`}
        </p>
    );
}

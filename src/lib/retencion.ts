import * as XLSX from 'xlsx';

export type Fila = (string | number | null)[];

export type Categoria = 'general' | 'arriendo' | 'compras' | 'excluir';
export type Persona = 'PN' | 'PJ';

export interface Movimiento {
    comprobante: string;
    identificacion: string;
    tercero: string;
    fecha: string;
    base: number;
    retencion: number;
    persona: Persona;
}

export interface GrupoImpuesto {
    nombre: string;
    categoriaSugerida: Categoria;
    movimientos: Movimiento[];
    /** Subtotal que trae el informe en la fila "Impuesto: ...". */
    subtotalBase: number;
    subtotalRetencion: number;
}

export interface CuentaBalance {
    nivel: string;
    codigo: string;
    nombre: string;
    debito: number;
    credito: number;
}

export interface Hojas {
    impuestos: GrupoImpuesto[] | null;
    balance: CuentaBalance[] | null;
}

/** Quita tildes y todo lo que no sea letra/número para comparar encabezados. */
function normalizar(valor: unknown): string {
    return String(valor ?? '')
        .normalize('NFD')
        .replace(/[̀-ͯ]/g, '')
        .toLowerCase()
        .replace(/[^a-z0-9]/g, '');
}

function numero(valor: unknown): number {
    if (typeof valor === 'number') return valor;
    if (valor === null || valor === undefined || valor === '') return 0;
    // Texto con formato colombiano: 1.234.567,89
    const texto = String(valor).trim().replace(/\s/g, '');
    const limpio = /,\d{1,2}$/.test(texto)
        ? texto.replace(/\./g, '').replace(',', '.')
        : texto.replace(/,/g, '');
    const n = Number(limpio);
    return Number.isFinite(n) ? n : 0;
}

function texto(valor: unknown): string {
    return String(valor ?? '').trim();
}

function buscarEncabezado(
    filas: Fila[],
    requisito: (normalizados: string[]) => boolean,
): number {
    const limite = Math.min(filas.length, 40);
    for (let i = 0; i < limite; i++) {
        if (requisito((filas[i] ?? []).map(normalizar))) return i;
    }
    return -1;
}

function columna(
    encabezado: string[],
    prueba: (h: string) => boolean,
): number {
    return encabezado.findIndex(prueba);
}

/** NIT de empresa: 9 dígitos que empiezan por 8 o 9, o razón social. */
export function tipoPersona(identificacion: string, nombre: string): Persona {
    const digitos = identificacion.replace(/\D/g, '');
    if (digitos.length === 9 && /^[89]/.test(digitos)) return 'PJ';
    if (
        /\b(S\.?\s?A\.?\s?S\.?|S\.?\s?A\.?|LTDA\.?|CONSORCIO|UNION TEMPORAL|SOCIEDAD|E\.?S\.?P\.?|CORPORACION|FUNDACION|ASOCIACION|COOPERATIVA|S\.?\s?EN\s?C\.?)(\s|$)/i.test(
            nombre,
        )
    ) {
        return 'PJ';
    }
    return 'PN';
}

function sugerirCategoria(nombreGrupo: string): Categoria {
    const n = normalizar(nombreGrupo);
    if (/arriend|arrend/.test(n)) return 'arriendo';
    if (/compra/.test(n)) return 'compras';
    if (/\b(rete\s*)?(iva|ica)\b/i.test(nombreGrupo)) return 'excluir';
    const tarifa = nombreGrupo.match(/(\d+(?:[.,]\d+)?)\s*%?\s*$/);
    if (tarifa && Number(tarifa[1].replace(',', '.')) === 2.5) return 'compras';
    return 'general';
}

/** Informe "Impuestos detallados" de Siigo. */
function leerImpuestos(filas: Fila[]): GrupoImpuesto[] | null {
    const fila = buscarEncabezado(filas, (h) => h.includes('basecompras'));
    if (fila < 0) return null;
    const enc = (filas[fila] ?? []).map(normalizar);

    const cComprobante = 0;
    const cId = columna(enc, (h) => h.startsWith('identificaci'));
    const cTercero = columna(enc, (h) => h.includes('nombretercero'));
    const cFecha = columna(enc, (h) => h.startsWith('fecha'));
    const cBase = columna(enc, (h) => h === 'basecompras');
    const cValor = columna(enc, (h) => h.startsWith('valorimpuestocompras'));
    const cBaseDev = columna(
        enc,
        (h) => h.startsWith('basedevolu') && h.endsWith('compras'),
    );
    const cValorDev = columna(
        enc,
        (h) => h.startsWith('valorimpuestodevolu') && h.endsWith('compras'),
    );

    const grupos: GrupoImpuesto[] = [];
    let actual: GrupoImpuesto | null = null;

    for (const f of filas.slice(fila + 1)) {
        const primera = texto(f[cComprobante]);
        if (!primera) continue;
        if (/^procesado/i.test(primera)) break;

        const base = numero(f[cBase]) - (cBaseDev >= 0 ? numero(f[cBaseDev]) : 0);
        const retencion =
            numero(f[cValor]) - (cValorDev >= 0 ? numero(f[cValorDev]) : 0);

        if (/^tipo\s*:/i.test(primera)) {
            actual = null;
            continue;
        }
        const encabezadoGrupo = primera.match(/^impuesto\s*:\s*(.*)$/i);
        if (encabezadoGrupo) {
            actual = {
                nombre: encabezadoGrupo[1].trim(),
                categoriaSugerida: sugerirCategoria(encabezadoGrupo[1]),
                movimientos: [],
                subtotalBase: base,
                subtotalRetencion: retencion,
            };
            grupos.push(actual);
            continue;
        }
        if (!actual) continue;
        // Las filas que solo tienen ventas (retenciones que nos practicaron) no van.
        if (base === 0 && retencion === 0) continue;

        const identificacion = texto(f[cId]);
        const tercero = texto(f[cTercero]);
        actual.movimientos.push({
            comprobante: primera,
            identificacion,
            tercero,
            fecha: texto(f[cFecha]),
            base,
            retencion,
            persona: tipoPersona(identificacion, tercero),
        });
    }

    return grupos.filter(
        (g) => g.movimientos.length > 0 || g.subtotalBase !== 0,
    );
}

/** Balance de prueba de Siigo. */
function leerBalance(filas: Fila[]): CuentaBalance[] | null {
    const fila = buscarEncabezado(
        filas,
        (h) =>
            h.some((x) => x.startsWith('movimientocr')) &&
            h.some((x) => x.includes('cuentacontable')),
    );
    if (fila < 0) return null;
    const enc = (filas[fila] ?? []).map(normalizar);

    const cNivel = columna(enc, (h) => h === 'nivel');
    const cCodigo = columna(
        enc,
        (h) => h.includes('cuentacontable') && !h.startsWith('nombre'),
    );
    const cNombre = columna(enc, (h) => h.startsWith('nombrecuenta'));
    const cDebito = columna(enc, (h) => h.startsWith('movimientod'));
    const cCredito = columna(enc, (h) => h.startsWith('movimientocr'));

    const cuentas: CuentaBalance[] = [];
    for (const f of filas.slice(fila + 1)) {
        const codigo = texto(f[cCodigo]).replace(/\D/g, '');
        if (!codigo) continue;
        cuentas.push({
            nivel: normalizar(f[cNivel]),
            codigo,
            nombre: texto(f[cNombre]),
            debito: numero(f[cDebito]),
            credito: numero(f[cCredito]),
        });
    }
    return cuentas;
}

/** Busca en todas las hojas del archivo el informe de impuestos y el balance. */
export async function leerArchivo(archivo: File): Promise<Hojas> {
    const resultado: Hojas = { impuestos: null, balance: null };
    const libro = XLSX.read(await archivo.arrayBuffer(), { type: 'array' });

    for (const nombreHoja of libro.SheetNames) {
        const filas = XLSX.utils.sheet_to_json<Fila>(
            libro.Sheets[nombreHoja],
            { header: 1, raw: true, defval: null },
        );
        const impuestos = leerImpuestos(filas);
        if (impuestos) {
            resultado.impuestos ??= impuestos;
            continue;
        }
        resultado.balance ??= leerBalance(filas);
    }
    return resultado;
}

/* ---------- Cálculos ---------- */

export interface Linea {
    clave: string;
    concepto: string;
    base: number | null;
    retencion: number;
}

export function resumenImpuestos(
    grupos: GrupoImpuesto[],
    categorias: Record<string, Categoria>,
    personas: Record<string, Persona>,
): { lineas: Linea[]; total: Linea } {
    const acum = {
        PN: { base: 0, ret: 0 },
        PJ: { base: 0, ret: 0 },
        arriendo: { base: 0, ret: 0 },
        comprasPJ: { base: 0, ret: 0 },
        comprasPN: { base: 0, ret: 0 },
    };

    for (const g of grupos) {
        const cat = categorias[g.nombre] ?? g.categoriaSugerida;
        if (cat === 'excluir') continue;
        for (const m of g.movimientos) {
            const persona = personas[m.identificacion] ?? m.persona;
            const destino =
                cat === 'arriendo'
                    ? acum.arriendo
                    : cat === 'compras'
                      ? persona === 'PJ'
                          ? acum.comprasPJ
                          : acum.comprasPN
                      : acum[persona];
            destino.base += m.base;
            destino.ret += m.retencion;
        }
    }

    const lineas: Linea[] = [
        { clave: 'pn', concepto: 'PN', base: acum.PN.base, retencion: acum.PN.ret },
        { clave: 'pj', concepto: 'PJ', base: acum.PJ.base, retencion: acum.PJ.ret },
        {
            clave: 'arriendo',
            concepto: 'ARRIENDO',
            base: acum.arriendo.base,
            retencion: acum.arriendo.ret,
        },
        {
            clave: 'comprasPJ',
            concepto: 'COMPRA PJ',
            base: acum.comprasPJ.base,
            retencion: acum.comprasPJ.ret,
        },
    ];
    if (acum.comprasPN.base !== 0 || acum.comprasPN.ret !== 0) {
        lineas.push({
            clave: 'comprasPN',
            concepto: 'COMPRA PN',
            base: acum.comprasPN.base,
            retencion: acum.comprasPN.ret,
        });
    }

    const total: Linea = {
        clave: 'total',
        concepto: 'TOTAL',
        base: lineas.reduce((s, l) => s + (l.base ?? 0), 0),
        retencion: lineas.reduce((s, l) => s + l.retencion, 0),
    };
    return { lineas, total };
}

export interface ResumenBalance {
    lineas: Linea[];
    /** Retenciones del balance que también salen en impuestos detallados. */
    otrasRetenciones: number;
    totalRetenciones2365: number;
}

const esAuxiliar = (c: CuentaBalance) =>
    c.nivel === 'auxiliar' || c.codigo.length >= 8;

export function resumenBalance(
    cuentas: CuentaBalance[],
    tarifaAutorretencion: number,
): ResumenBalance {
    const auxiliares2365 = cuentas.filter(
        (c) => c.codigo.startsWith('2365') && esAuxiliar(c),
    );
    const suma = (lista: CuentaBalance[]) =>
        lista.reduce((s, c) => s + c.credito, 0);

    const salarios = auxiliares2365.filter((c) => c.codigo.startsWith('236505'));
    const independientes = auxiliares2365.filter((c) =>
        /independ/i.test(c.nombre),
    );
    const autorretencion = auxiliares2365.filter((c) =>
        c.codigo.startsWith('236575'),
    );
    const pagos = auxiliares2365.filter((c) =>
        /pago/i.test(normalizar(c.nombre)),
    );
    const apartadas = new Set(
        [...salarios, ...independientes, ...autorretencion, ...pagos].map(
            (c) => c.codigo,
        ),
    );
    const otras = auxiliares2365.filter((c) => !apartadas.has(c.codigo));

    const ingresos = cuentas.find((c) => c.codigo === '4');
    const baseAutorretencion = ingresos ? ingresos.credito - ingresos.debito : 0;

    const lineas: Linea[] = [
        {
            clave: 'autorretencionCalc',
            concepto: `AUTORRETENCIÓN (clase 4 crédito − débito × ${String(tarifaAutorretencion).replace('.', ',')}%)`,
            base: baseAutorretencion,
            retencion: (baseAutorretencion * tarifaAutorretencion) / 100,
        },
        { clave: 'salarios', concepto: 'SALARIOS', base: null, retencion: suma(salarios) },
        {
            clave: 'honorariosPN',
            concepto: 'HONORARIOS PN (independientes)',
            base: null,
            retencion: suma(independientes),
        },
        {
            clave: 'autorretencion',
            concepto: 'AUTORRETENCIÓN (balance)',
            base: null,
            retencion: suma(autorretencion),
        },
    ];

    const reteIva = cuentas.filter(
        (c) => c.codigo.startsWith('2367') && esAuxiliar(c),
    );
    if (reteIva.length > 0) {
        lineas.push({
            clave: 'reteIva',
            concepto: 'RTE IVA',
            base: null,
            retencion: suma(reteIva),
        });
    }

    return {
        lineas,
        otrasRetenciones: suma(otras),
        totalRetenciones2365: suma(auxiliares2365.filter((c) => !pagos.includes(c))),
    };
}

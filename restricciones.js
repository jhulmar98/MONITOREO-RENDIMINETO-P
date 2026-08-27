// =====================================================
// RESTRICCIONES.JS
// DASHBOARD DE RESTRICCIONES MÉDICAS
// =====================================================

import * as XLSX
from "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/+esm";
import {
    descargarExcelStorage,
    RUTAS_STORAGE
} from "../backend/api.js";

// =====================================================
// CONFIGURACIÓN
// =====================================================

const HOJA_PREFERIDA = "RESTRICCIONES";

const ALTO_GRAFICO_VISIBLE = 440;

const ITEMS_VISIBLES = 9;

const ALTO_ITEM_HORIZONTAL = 42;


// =====================================================
// MESES
// =====================================================

const MESES_ORDEN = [
    "ENERO",
    "FEBRERO",
    "MARZO",
    "ABRIL",
    "MAYO",
    "JUNIO",
    "JULIO",
    "AGOSTO",
    "SETIEMBRE",
    "OCTUBRE",
    "NOVIEMBRE",
    "DICIEMBRE"
];


// =====================================================
// ESTADO GLOBAL
// =====================================================

const estado = {

    data: [],

    dataFiltradaActual: [],

    filtroMedicoGrafico: "TOTAL",

    filtros: {

        meses: [],

        condicionLaboral: [],

        medico: [],

        cargo: [],

        zona: [],

        estado: []

    },

    charts: {

        meses: null,

        condicionLaboral: null,

        medicos: null,

        cargo: null,

        zona: null

    }

};


let filtroFlotanteActivo = null;


// =====================================================
// CONFIGURACIÓN FILTROS
// =====================================================

const CONFIG_FILTROS = {

    meses: {

        detailsId: "boxFiltroMes",

        panelId: "opcionesFiltroMes",

        textoId: "textoFiltroMes",

        textoTodos: "TODOS LOS MESES"

    },

    condicionLaboral: {

        detailsId: "boxFiltroCondicionLaboral",

        panelId: "opcionesFiltroCondicionLaboral",

        textoId: "textoFiltroCondicionLaboral",

        textoTodos: "TODAS LAS CONDICIONES"

    },

    medico: {

        detailsId: "boxFiltroMedico",

        panelId: "opcionesFiltroMedico",

        textoId: "textoFiltroMedico",

        textoTodos: "TODOS LOS MÉDICOS"

    },

    cargo: {

        detailsId: "boxFiltroCargo",

        panelId: "opcionesFiltroCargo",

        textoId: "textoFiltroCargo",

        textoTodos: "TODOS LOS CARGOS"

    },

    zona: {

        detailsId: "boxFiltroZona",

        panelId: "opcionesFiltroZona",

        textoId: "textoFiltroZona",

        textoTodos: "TODAS LAS ZONAS"

    },

    estado: {

        detailsId: "boxFiltroEstado",

        panelId: "opcionesFiltroEstado",

        textoId: "textoFiltroEstado",

        textoTodos: "TODOS LOS ESTADOS"

    }

};


// =====================================================
// INICIO
// =====================================================

window.addEventListener(
    "DOMContentLoaded",
    async () => {

        mostrarLoading(true);

        try {

            inyectarAjustesVisuales();

            reorganizarMedicosYAlertas();

            ocultarBloqueRankingSiExiste();

            prepararModalDetalle();

            crearModalMotivo();

            await cargarExcel();

            inicializarFiltros();

            inicializarBotonLimpiar();

            inicializarModal();

            inicializarFiltrosFlotantes();

            inicializarSelectorGraficoMedicos();

            renderDashboard();

        } catch (error) {

            console.error(
                "ERROR GENERAL RESTRICCIONES:",
                error
            );

            alert(
                "No se pudo cargar el Dashboard de Restricciones Médicas.\n\n" +
                "Verifica que restricciones.xlsx esté en la misma carpeta.\n\n" +
                "Revisa la consola del navegador."
            );

        } finally {

            mostrarLoading(false);

        }

    }
);


// =====================================================
// OCULTAR RANKING SI TODAVÍA EXISTE EN HTML
// =====================================================

function ocultarBloqueRankingSiExiste() {

    const panel =
        document.querySelector(
            ".panel-ranking"
        );

    if (!panel) {
        return;
    }

    const seccion =
        panel.closest(
            ".charts-grid"
        );

    if (seccion) {

        seccion.remove();

    } else {

        panel.remove();

    }

}


// =====================================================
// AJUSTES VISUALES
// =====================================================

function inyectarAjustesVisuales() {

    if (
        document.getElementById(
            "ajustesRestriccionesJS"
        )
    ) {
        return;
    }

    const style =
        document.createElement(
            "style"
        );

    style.id =
        "ajustesRestriccionesJS";

    style.textContent = `

        /* =================================================
           KPI - UNA SOLA FILA EN ESCRITORIO
        ================================================= */

        .kpi-grid--restricciones {

            grid-template-columns:
                repeat(5, minmax(0, 1fr)) !important;

        }

        .kpi-grid--restricciones .kpi-card {

            min-width: 0;

        }


        /* =================================================
           ALTURA GENERAL GRÁFICOS
        ================================================= */

        .chart-container,
        .chart-container--condicion,
        .chart-container--medicos,
        .chart-container--cargo,
        .chart-container--zona {

            height: ${ALTO_GRAFICO_VISIBLE}px !important;

            min-height: ${ALTO_GRAFICO_VISIBLE}px !important;

        }


        #graficoRestriccionesPorMes,
        #graficoCondicionLaboral,
        #graficoMedicosCertificadores,
        #graficoCargoActual,
        #graficoZonaActual {

            width: 100% !important;

            height: 100% !important;

        }


        /* =================================================
           FILA MÉDICO + ALERTAS
        ================================================= */

        .fila-medicos-alertas {

            display: grid !important;

            grid-template-columns:
                repeat(2, minmax(0, 1fr)) !important;

            gap: 22px !important;

            align-items: stretch !important;

        }

        .fila-medicos-alertas > .panel {

            width: 100%;

            min-width: 0;

        }


        /* =================================================
           MÉDICOS - TORTA
        ================================================= */

        .chart-container--medicos {

            height: ${ALTO_GRAFICO_VISIBLE}px !important;

            min-height: ${ALTO_GRAFICO_VISIBLE}px !important;

            overflow: hidden !important;

            display: flex;

            align-items: center;

            justify-content: center;

            padding: 22px !important;

        }

        .medicos-canvas-wrap {

            width: 100%;

            height: 100% !important;

            min-height: 0 !important;

            position: relative;

        }


        /* =================================================
           ALERTAS AL COSTADO
        ================================================= */

        .fila-medicos-alertas
        .panel-alertas-restricciones {

            height:
                calc(
                    ${ALTO_GRAFICO_VISIBLE}px + 100px
                );

            min-height:
                calc(
                    ${ALTO_GRAFICO_VISIBLE}px + 100px
                );

            display: flex;

            flex-direction: column;

            overflow: hidden;

        }


        .fila-medicos-alertas
        .panel-alertas-restricciones
        .panel-header {

            flex: 0 0 auto;

        }


        .fila-medicos-alertas
        .panel-alertas-restricciones
        .alerta-leyenda {

            flex: 0 0 auto;

        }


        .fila-medicos-alertas
        .tabla-alertas-scroll {

            flex: 1;

            min-height: 0;

            max-height: none !important;

            overflow: auto;

        }


        .fila-medicos-alertas
        .tabla-alertas-restricciones {

            min-width: 1180px;

        }


        /* =================================================
           MODAL DETALLE
        ================================================= */

        .detalle-table--restricciones {

            min-width: 1320px !important;

        }


        .btn-ver-motivo {

            min-width: 92px;

            padding: 8px 13px;

            display: inline-flex;

            align-items: center;

            justify-content: center;

            gap: 7px;

            color: #ffffff;

            font-size: .74rem;

            font-weight: 900;

            text-transform: uppercase;

            background:
                linear-gradient(
                    135deg,
                    #556b2f,
                    #3d4d1f
                );

            border: none;

            border-radius: 999px;

            cursor: pointer;

            box-shadow:
                0 6px 14px
                rgba(85,107,47,.18);

            transition:
                .22s ease;

        }


        .btn-ver-motivo:hover {

            transform:
                translateY(-2px);

            box-shadow:
                0 9px 20px
                rgba(85,107,47,.28);

        }


        /* =================================================
           MODAL MOTIVO
        ================================================= */

        .modal-motivo-restriccion {

            position: fixed;

            inset: 0;

            z-index: 1000000;

            padding: 20px;

            display: flex;

            align-items: center;

            justify-content: center;

            background:
                rgba(15,23,42,.58);

            backdrop-filter:
                blur(7px);

            opacity: 0;

            visibility: hidden;

            transition:
                .22s ease;

        }


        .modal-motivo-restriccion--show {

            opacity: 1;

            visibility: visible;

        }


        .modal-motivo-restriccion__box {

            width:
                min(
                    760px,
                    96vw
                );

            max-height:
                88vh;

            display: flex;

            flex-direction: column;

            background:
                #ffffff;

            border-radius:
                24px;

            box-shadow:
                0 30px 80px
                rgba(15,23,42,.34);

            overflow: hidden;

            transform:
                translateY(15px)
                scale(.98);

            transition:
                .22s ease;

        }


        .modal-motivo-restriccion--show
        .modal-motivo-restriccion__box {

            transform:
                translateY(0)
                scale(1);

        }


        .modal-motivo-restriccion__header {

            padding:
                20px
                22px;

            display: flex;

            align-items: flex-start;

            justify-content: space-between;

            gap: 15px;

            background:
                linear-gradient(
                    180deg,
                    rgba(85,107,47,.08),
                    #ffffff
                );

            border-bottom:
                1px solid
                #edf1e7;

        }


        .modal-motivo-restriccion__header h3 {

            margin:
                0
                0
                5px;

            color:
                #1e293b;

            font-size:
                1.05rem;

            font-weight:
                900;

            text-transform:
                uppercase;

        }


        .modal-motivo-restriccion__header p {

            margin: 0;

            color:
                #64748b;

            font-size:
                .82rem;

            font-weight:
                700;

        }


        .modal-motivo-restriccion__close {

            width: 40px;

            min-width: 40px;

            height: 40px;

            color:
                #475569;

            background:
                #f1f5f9;

            border: none;

            border-radius:
                12px;

            cursor: pointer;

            transition:
                .2s ease;

        }


        .modal-motivo-restriccion__close:hover {

            color:
                #dc2626;

            transform:
                rotate(90deg);

        }


        .modal-motivo-restriccion__body {

            padding: 22px;

            display: flex;

            flex-direction: column;

            gap: 18px;

            overflow-y: auto;

        }


        .motivo-bloque {

            padding: 18px;

            background:
                linear-gradient(
                    180deg,
                    #ffffff,
                    #f8faf5
                );

            border:
                1px solid
                #e3ead9;

            border-radius:
                17px;

        }


        .motivo-bloque h4 {

            margin:
                0
                0
                10px;

            color:
                #556b2f;

            font-size:
                .8rem;

            font-weight:
                900;

            text-transform:
                uppercase;

        }


        .motivo-bloque p {

            margin: 0;

            color:
                #334155;

            font-size:
                .88rem;

            font-weight:
                700;

            line-height:
                1.65;

            white-space:
                pre-wrap;

            overflow-wrap:
                anywhere;

        }


        /* =================================================
           RESPONSIVE
        ================================================= */

        @media (max-width: 1200px) {

            .kpi-grid--restricciones {

                grid-template-columns:
                    repeat(2, minmax(0, 1fr)) !important;

            }


            .fila-medicos-alertas {

                grid-template-columns:
                    1fr !important;

            }


            .fila-medicos-alertas
            .panel-alertas-restricciones {

                height: auto;

                min-height: 600px;

            }


            .fila-medicos-alertas
            .tabla-alertas-scroll {

                max-height: 500px !important;

            }

        }


        @media (max-width: 650px) {

            .kpi-grid--restricciones {

                grid-template-columns:
                    1fr !important;

            }

        }

    `;


    document.head.appendChild(
        style
    );

}


// =====================================================
// REORGANIZAR MÉDICOS + ALERTAS
// =====================================================

function reorganizarMedicosYAlertas() {

    const panelMedicos =
        document.querySelector(
            ".panel-medicos"
        );


    const panelAlertas =
        document.querySelector(
            ".panel-alertas-restricciones"
        );


    if (
        !panelMedicos ||
        !panelAlertas
    ) {

        return;

    }


    const seccionMedicos =
        panelMedicos.closest(
            ".charts-grid"
        );


    const seccionAlertas =
        panelAlertas.closest(
            ".charts-grid"
        );


    if (!seccionMedicos) {

        return;

    }


    seccionMedicos.classList.add(
        "fila-medicos-alertas"
    );


    seccionMedicos.classList.remove(
        "charts-grid--full"
    );


    panelMedicos.classList.remove(
        "panel-full"
    );


    panelAlertas.classList.remove(
        "panel-full"
    );


    seccionMedicos.appendChild(
        panelAlertas
    );


    if (
        seccionAlertas &&
        seccionAlertas !==
        seccionMedicos &&
        !seccionAlertas.querySelector(
            ".panel"
        )
    ) {

        seccionAlertas.remove();

    }

}


// =====================================================
// CARGAR EXCEL DESDE FIREBASE STORAGE
// =====================================================

async function cargarExcel() {

    console.log(
        "Cargando Excel desde Firebase Storage:",
        RUTAS_STORAGE.restriccionesMedicas
    );


    // =================================================
    // DESCARGAR ARCHIVO DESDE FIREBASE STORAGE
    // =================================================

    const buffer =
        await descargarExcelStorage(
            RUTAS_STORAGE.restriccionesMedicas
        );


    if (
        !buffer ||
        !buffer.byteLength
    ) {

        throw new Error(
            "El archivo restricciones.xlsx descargado desde Firebase está vacío."
        );

    }


    console.log(
        "Excel descargado correctamente:",
        buffer.byteLength,
        "bytes"
    );


    // =================================================
    // LEER ARCHIVO EXCEL
    // =================================================

    const workbook =
        XLSX.read(
            buffer,
            {
                type:
                    "array",

                cellDates:
                    true
            }
        );


    // =================================================
    // LOCALIZAR HOJA
    // =================================================

    let nombreHoja =
        workbook.SheetNames.find(
            nombre =>
                normalizarClave(
                    nombre
                ) ===
                normalizarClave(
                    HOJA_PREFERIDA
                )
        );


    if (!nombreHoja) {

        nombreHoja =
            workbook.SheetNames[0];

    }


    console.log(
        "Hoja utilizada:",
        nombreHoja
    );


    const hoja =
        workbook.Sheets[
            nombreHoja
        ];


    if (!hoja) {

        throw new Error(
            "No se encontró una hoja válida en restricciones.xlsx."
        );

    }


    // =================================================
    // CONVERTIR HOJA A MATRIZ
    // =================================================

    const matriz =
        XLSX.utils.sheet_to_json(
            hoja,
            {
                header:
                    1,

                defval:
                    "",

                raw:
                    true
            }
        );


    // =================================================
    // LOCALIZAR CABECERA
    // =================================================

    const indiceCabecera =
        localizarFilaCabeceraRestricciones(
            matriz
        );


    if (
        indiceCabecera <
        0
    ) {

        throw new Error(
            "No se encontró la fila de cabecera en restricciones.xlsx."
        );

    }


    const cabeceras =
        matriz[
            indiceCabecera
        ].map(
            valor =>
                String(
                    valor ??
                    ""
                ).trim()
        );


    console.log(
        "Cabeceras:",
        cabeceras
    );


    // =================================================
    // PROCESAR FILAS
    // =================================================

    const filas =
        [];


    for (
        let i =
            indiceCabecera + 1;

        i <
        matriz.length;

        i++
    ) {

        const fila =
            matriz[i] ||
            [];


        const objeto =
            {};


        cabeceras.forEach(
            (
                cabecera,
                columna
            ) => {

                if (!cabecera) {

                    return;

                }


                objeto[
                    cabecera
                ] =
                    fila[
                        columna
                    ] ??
                    "";

            }
        );


        const tieneDatos =
            Object.values(
                objeto
            ).some(
                valor =>
                    String(
                        valor ??
                        ""
                    ).trim() !==
                    ""
            );


        if (!tieneDatos) {

            continue;

        }


        filas.push({

            objeto,

            filaExcel:
                i + 1

        });

    }


    // =================================================
    // NORMALIZAR Y ELIMINAR DUPLICADOS
    // =================================================

    estado.data =
        eliminarDuplicados(
            filas
                .map(
                    item =>
                        normalizarRegistro(
                            item.objeto,
                            item.filaExcel
                        )
                )
                .filter(Boolean)
        );


    console.log(
        "Registros cargados desde Firebase:",
        estado.data.length
    );

}


// =====================================================
// LOCALIZAR CABECERA
// =====================================================

function localizarFilaCabeceraRestricciones(
    matriz
) {

    const limite =
        Math.min(
            matriz.length,
            30
        );


    for (
        let i = 0;

        i < limite;

        i++
    ) {

        const fila =
            matriz[i] ||
            [];


        const valores =
            fila.map(
                valor =>
                    normalizarCabecera(
                        valor
                    )
            );


        const tieneNombre =
            valores.includes(
                "APELLIDOS Y NOMBRE"
            ) ||
            valores.includes(
                "APELLIDOS Y NOMBRES"
            );


        const tieneCargo =
            valores.includes(
                "CARGO ACTUAL"
            );


        if (
            tieneNombre &&
            tieneCargo
        ) {

            return i;

        }

    }


    return -1;

}


// =====================================================
// NORMALIZAR REGISTRO
// =====================================================

function normalizarRegistro(
    row,
    filaExcel
) {

    const numero =
        limpiarNumero(
            obtenerCampo(
                row,
                [
                    "N°",
                    "Nº",
                    "NRO",
                    "NUMERO",
                    "NÚMERO"
                ]
            )
        );


    let dni =
        limpiarDocumento(
            obtenerCampo(
                row,
                [
                    "DNI",
                    "D.N.I.",
                    "N° DNI",
                    "Nº DNI",
                    "N DNI",
                    "NRO DNI",
                    "NÚMERO DNI",
                    "NUMERO DNI",

                    "NUMDOCUMENTO",
                    "NUM DOCUMENTO",
                    "NUM_DOCUMENTO",

                    "N° DOCUMENTO",
                    "Nº DOCUMENTO",
                    "NRO DOCUMENTO",

                    "NÚMERO DOCUMENTO",
                    "NUMERO DOCUMENTO",

                    "NÚMERO DE DOCUMENTO",
                    "NUMERO DE DOCUMENTO",

                    "DOCUMENTO",
                    "DOC. IDENTIDAD",
                    "DOCUMENTO IDENTIDAD",
                    "DOCUMENTO DE IDENTIDAD"
                ]
            )
        );


    if (!dni) {

        dni =
            buscarDocumentoAutomaticamente(
                row
            );

    }


    const nombreOriginal =
        normalizarTexto(
            obtenerCampo(
                row,
                [
                    "APELLIDOS Y NOMBRE",
                    "APELLIDOS Y NOMBRES",
                    "NOMBRE",
                    "NOMBRES"
                ]
            )
        );


    if (!nombreOriginal) {

        return null;

    }


    if (
        nombreOriginal ===
        "TOTAL" ||
        nombreOriginal ===
        "CANTIDAD"
    ) {

        return null;

    }


    const edad =
        extraerEdadDeNombre(
            nombreOriginal
        );


    const nombre =
        limpiarEdadDeNombre(
            nombreOriginal
        );


    if (!nombre) {

        return null;

    }


    const cargoActual =
        normalizarTexto(
            obtenerCampo(
                row,
                [
                    "CARGO ACTUAL",
                    "CARGO"
                ]
            )
        ) ||
        "NO ESPECIFICADO";


    const zonaActual =
        normalizarZona(
            obtenerCampo(
                row,
                [
                    "ZONA ACTUAL",
                    "ZONA"
                ]
            )
        );


    const contrato =
        normalizarTexto(
            obtenerCampo(
                row,
                [
                    "CONTRATO",
                    "CONTRATADO",
                    "TIPO CONTRATO"
                ]
            )
        ) ||
        "NO ESPECIFICADO";


    const condicionLaboral =
        normalizarTexto(
            obtenerCampo(
                row,
                [
                    "CONDICIÓN LABORAL",
                    "CONDICION LABORAL"
                ]
            )
        ) ||
        "NO ESPECIFICADO";


    const medico =
        normalizarTexto(
            obtenerCampo(
                row,
                [
                    "MEDICO QUE CERTIFICA",
                    "MÉDICO QUE CERTIFICA",
                    "MEDICO CERTIFICADOR",
                    "MÉDICO CERTIFICADOR"
                ]
            )
        ) ||
        "NO ESPECIFICADO";


    const restriccionMotivo =
        limpiarTexto(
            obtenerCampo(
                row,
                [
                    "RESTRICCION MOTIVO",
                    "RESTRICCIÓN MOTIVO",
                    "MOTIVO"
                ]
            )
        );


    const observacionesMedicas =
        limpiarTexto(
            obtenerCampo(
                row,
                [
                    "OBERVACIONES",
                    "OBSERVACIONES",

                    "OBSERVACIÓN MÉDICA",
                    "OBSERVACION MEDICA",

                    "OBSERVACIONES MÉDICAS",
                    "OBSERVACIONES MEDICAS"
                ]
            )
        );


    const restriccion =
        limpiarTexto(
            obtenerCampo(
                row,
                [
                    "RESTRICCIÓN",
                    "RESTRICCION"
                ]
            )
        );


    const observacion =
        limpiarTexto(
            obtenerCampo(
                row,
                [
                    "OBSERVACION",
                    "OBSERVACIÓN"
                ]
            )
        );


    const valorFechaInicio =
        obtenerCampo(
            row,
            [
                "F. INICIO",
                "F INICIO",
                "FECHA INICIO",
                "FECHA DE INICIO"
            ]
        );


    const valorFechaTermino =
        obtenerCampo(
            row,
            [
                "F. TERMINO",
                "F. TÉRMINO",

                "F TERMINO",
                "F TÉRMINO",

                "FECHA TERMINO",
                "FECHA TÉRMINO",

                "FECHA FIN",
                "FECHA DE FIN"
            ]
        );


    const fechaInicio =
        convertirFecha(
            valorFechaInicio
        );


    const fechaTermino =
        convertirFecha(
            valorFechaTermino
        );


    const anio =
        fechaInicio
            ? fechaInicio.getFullYear()
            : null;


    const mesNumero =
        fechaInicio
            ? fechaInicio.getMonth() + 1
            : null;


    const mes =
        fechaInicio
            ? MESES_ORDEN[
                fechaInicio.getMonth()
            ]
            : "INDEFINIDO";


    const calculoEstado =
        calcularEstadoRestriccion(
            fechaTermino,
            valorFechaTermino
        );


    return {

        filaExcel,

        numero,

        dni,

        nombre,

        nombreOriginal,

        edad,

        personaKey:
            dni ||
            normalizarClave(
                nombre
            ),

        cargoActual,

        zonaActual,

        contrato,

        condicionLaboral,

        medico,

        restriccionMotivo,

        observacionesMedicas,

        restriccion,

        observacion,

        fechaInicio,

        fechaTermino,

        fechaInicioTexto:
            fechaInicio
                ? formatearFecha(
                    fechaInicio
                )
                : "INDEFINIDO",

        fechaTerminoTexto:
            fechaTermino
                ? formatearFecha(
                    fechaTermino
                )
                : "INDEFINIDO",

        anio,

        mesNumero,

        mes,

        estado:
            calculoEstado.estado,

        diasRestantes:
            calculoEstado.diasRestantes

    };

}


// =====================================================
// BUSCAR DNI AUTOMÁTICAMENTE
// =====================================================

function buscarDocumentoAutomaticamente(
    row
) {

    const claves =
        Object.keys(
            row
        );


    const candidatos =
        claves.filter(
            clave => {

                const normalizada =
                    normalizarCabecera(
                        clave
                    );


                return (
                    normalizada.includes(
                        "DNI"
                    ) ||
                    normalizada.includes(
                        "DOCUMENTO"
                    )
                );

            }
        );


    for (
        const clave
        of candidatos
    ) {

        const documento =
            limpiarDocumento(
                row[
                    clave
                ]
            );


        if (
            documento.length >= 6 &&
            documento.length <= 15
        ) {

            return documento;

        }

    }


    return "";

}


// =====================================================
// ESTADO DE RESTRICCIÓN
// =====================================================

function calcularEstadoRestriccion(
    fechaTermino,
    valorOriginal
) {

    const texto =
        normalizarTexto(
            valorOriginal
        );


    if (
        texto.includes(
            "INDEFIN"
        )
    ) {

        return {

            estado:
                "INDEFINIDA",

            diasRestantes:
                null

        };

    }


    if (!fechaTermino) {

        return {

            estado:
                "INDEFINIDA",

            diasRestantes:
                null

        };

    }


    const hoy =
        normalizarDate(
            new Date()
        );


    const fin =
        normalizarDate(
            fechaTermino
        );


    const dias =
        Math.round(
            (
                fin.getTime() -
                hoy.getTime()
            ) /
            86400000
        );


    if (
        dias < 0
    ) {

        return {

            estado:
                "VENCIDA",

            diasRestantes:
                dias

        };

    }


    if (
        dias >= 0 &&
        dias <= 7
    ) {

        return {

            estado:
                "POR VENCER",

            diasRestantes:
                dias

        };

    }


    return {

        estado:
            "VIGENTE",

        diasRestantes:
            dias

    };

}


// =====================================================
// ELIMINAR DUPLICADOS
// =====================================================

function eliminarDuplicados(
    data
) {

    const mapa =
        new Map();


    data.forEach(
        item => {

            const clave = [

                item.personaKey,

                item.fechaInicioTexto,

                item.fechaTerminoTexto,

                normalizarClave(
                    item.restriccion
                ),

                normalizarClave(
                    item.medico
                ),

                normalizarClave(
                    item.cargoActual
                ),

                normalizarClave(
                    item.zonaActual
                )

            ].join("|");


            if (
                !mapa.has(
                    clave
                )
            ) {

                mapa.set(
                    clave,
                    item
                );

            }

        }
    );


    return [
        ...mapa.values()
    ];

}


// =====================================================
// FILTROS
// =====================================================

function inicializarFiltros() {

    cargarFiltroAnio();

    cargarListaNombres();

    cargarFiltrosMultiples();


    const filtroAnio =
        document.getElementById(
            "filtroAnio"
        );


    if (filtroAnio) {

        filtroAnio.addEventListener(
            "change",
            () => {

                limpiarFiltrosSinCambiarAnio();

                cargarFiltrosMultiples();

                cerrarPanelFlotanteFiltro();

                renderDashboard();

            }
        );

    }


    const inputNombre =
        document.getElementById(
            "filtroNombreGeneral"
        );


    if (inputNombre) {

        inputNombre.addEventListener(
            "input",
            () => {

                actualizarFiltrosDependientes(
                    "nombre"
                );

                cerrarPanelFlotanteFiltro();

                renderDashboard();

            }
        );

    }

}


// =====================================================
// AÑO
// =====================================================

function cargarFiltroAnio() {

    const select =
        document.getElementById(
            "filtroAnio"
        );


    if (!select) {

        return;

    }


    const anios = [
        ...new Set(
            estado.data
                .map(
                    item =>
                        item.anio
                )
                .filter(Boolean)
        )
    ].sort(
        (a, b) =>
            a - b
    );


    select.innerHTML =
        "";


    const todos =
        document.createElement(
            "option"
        );


    todos.value =
        "TODOS";


    todos.textContent =
        "TODOS LOS AÑOS";


    select.appendChild(
        todos
    );


    anios.forEach(
        anio => {

            const option =
                document.createElement(
                    "option"
                );


            option.value =
                String(
                    anio
                );


            option.textContent =
                String(
                    anio
                );


            select.appendChild(
                option
            );

        }
    );


    select.value =
        "TODOS";

}


// =====================================================
// FILTROS MÚLTIPLES
// =====================================================

function cargarFiltrosMultiples() {

    const data =
        obtenerDataPorAnio();


    cargarFiltroMes(
        data
    );


    cargarFiltroCampo({

        key:
            "condicionLaboral",

        data,

        campo:
            "condicionLaboral"

    });


    cargarFiltroCampo({

        key:
            "medico",

        data,

        campo:
            "medico"

    });


    cargarFiltroCampo({

        key:
            "cargo",

        data,

        campo:
            "cargoActual"

    });


    cargarFiltroCampo({

        key:
            "zona",

        data,

        campo:
            "zonaActual"

    });


    cargarFiltroCampo({

        key:
            "estado",

        data,

        campo:
            "estado",

        ordenPersonalizado: [

            "VIGENTE",
            "POR VENCER",
            "VENCIDA",
            "INDEFINIDA"

        ]

    });

}


// =====================================================
// FILTRO MES
// =====================================================

function cargarFiltroMes(
    data
) {

    renderFiltroCheckbox({

        key:
            "meses",

        opciones:
            obtenerOpcionesMes(
                data
            )

    });

}


function obtenerOpcionesMes(
    data
) {

    const opciones =
        [];


    if (
        data.some(
            item =>
                !item.fechaInicio
        )
    ) {

        opciones.push({

            value:
                "INDEFINIDO",

            label:
                "INDEFINIDO"

        });

    }


    const meses = [
        ...new Set(
            data
                .map(
                    item =>
                        item.mesNumero
                )
                .filter(Boolean)
        )
    ].sort(
        (a, b) =>
            a - b
    );


    meses.forEach(
        mes => {

            opciones.push({

                value:
                    String(
                        mes
                    ),

                label:
                    MESES_ORDEN[
                        mes - 1
                    ]

            });

        }
    );


    return opciones;

}


// =====================================================
// FILTRO CAMPO
// =====================================================

function cargarFiltroCampo({

    key,

    data,

    campo,

    conservarSeleccion = false,

    ordenPersonalizado = null

}) {

    let valores = [
        ...new Set(
            data
                .map(
                    item =>
                        item[
                            campo
                        ] ||
                        "NO ESPECIFICADO"
                )
                .filter(Boolean)
        )
    ];


    if (
        Array.isArray(
            ordenPersonalizado
        )
    ) {

        valores.sort(
            (a, b) => {

                const ia =
                    ordenPersonalizado.indexOf(
                        a
                    );


                const ib =
                    ordenPersonalizado.indexOf(
                        b
                    );


                if (
                    ia >= 0 &&
                    ib >= 0
                ) {

                    return ia - ib;

                }


                if (
                    ia >= 0
                ) {

                    return -1;

                }


                if (
                    ib >= 0
                ) {

                    return 1;

                }


                return a.localeCompare(
                    b,
                    "es"
                );

            }
        );

    } else {

        valores.sort(
            (a, b) =>
                a.localeCompare(
                    b,
                    "es"
                )
        );

    }


    if (
        conservarSeleccion
    ) {

        estado.filtros[
            key
        ] =
            estado.filtros[
                key
            ].filter(
                valor =>
                    valores.includes(
                        valor
                    )
            );

    }


    renderFiltroCheckbox({

        key,

        opciones:
            valores.map(
                valor => ({

                    value:
                        valor,

                    label:
                        valor

                })
            )

    });

}


// =====================================================
// RENDER CHECKBOX
// =====================================================

function renderFiltroCheckbox({

    key,

    opciones

}) {

    const config =
        CONFIG_FILTROS[
            key
        ];


    if (!config) {

        return;

    }


    const panel =
        document.getElementById(
            config.panelId
        );


    if (!panel) {

        return;

    }


    const seleccion =
        estado.filtros[
            key
        ] ||
        [];


    panel.innerHTML =
        "";


    panel.appendChild(
        crearOpcionCheckbox({

            key,

            value:
                "TODOS",

            label:
                config.textoTodos,

            checked:
                seleccion.length ===
                0,

            esTodos:
                true

        })
    );


    opciones.forEach(
        opcion => {

            panel.appendChild(
                crearOpcionCheckbox({

                    key,

                    value:
                        opcion.value,

                    label:
                        opcion.label,

                    checked:
                        seleccion.includes(
                            opcion.value
                        ),

                    esTodos:
                        false

                })
            );

        }
    );


    actualizarTextoFiltro(
        key,
        opciones
    );

}


// =====================================================
// CREAR CHECKBOX
// =====================================================

function crearOpcionCheckbox({

    key,

    value,

    label,

    checked,

    esTodos

}) {

    const item =
        document.createElement(
            "label"
        );


    item.className =
        esTodos
            ? "check-option check-option--all"
            : "check-option";


    const input =
        document.createElement(
            "input"
        );


    input.type =
        "checkbox";


    input.value =
        value;


    input.checked =
        checked;


    const span =
        document.createElement(
            "span"
        );


    span.textContent =
        label;


    item.append(
        input,
        span
    );


    input.addEventListener(
        "change",
        event => {

            event.stopPropagation();


            aplicarCambioFiltroCheckbox({

                key,

                value,

                checked:
                    input.checked

            });

        }
    );


    return item;

}


// =====================================================
// CAMBIO FILTRO
// =====================================================

function aplicarCambioFiltroCheckbox({

    key,

    value,

    checked

}) {

    if (
        value ===
        "TODOS"
    ) {

        estado.filtros[
            key
        ] =
            [];

    } else {

        let seleccion = [
            ...estado.filtros[
                key
            ]
        ];


        if (checked) {

            if (
                !seleccion.includes(
                    value
                )
            ) {

                seleccion.push(
                    value
                );

            }

        } else {

            seleccion =
                seleccion.filter(
                    item =>
                        item !==
                        value
                );

        }


        estado.filtros[
            key
        ] =
            seleccion;

    }


    actualizarFiltrosDependientes(
        key
    );


    renderDashboard();


    if (
        filtroFlotanteActivo &&
        filtroFlotanteActivo.key ===
        key
    ) {

        renderPanelFlotanteFiltro(
            key
        );


        requestAnimationFrame(
            posicionarPanelFlotanteFiltro
        );

    }

}


// =====================================================
// FILTROS DEPENDIENTES
// =====================================================

function actualizarFiltrosDependientes(
    filtroCambiado = null
) {

    const dataBase =
        obtenerDataPorAnio();


    if (
        filtroCambiado !==
        "meses"
    ) {

        const opciones =
            obtenerOpcionesMes(
                filtrarDataAvanzada(
                    dataBase,
                    {
                        ignorarMeses:
                            true
                    }
                )
            );


        estado.filtros.meses =
            estado.filtros.meses.filter(
                valor =>
                    opciones.some(
                        opcion =>
                            opcion.value ===
                            valor
                    )
            );


        renderFiltroCheckbox({

            key:
                "meses",

            opciones

        });

    }


    actualizarFiltroDependiente({

        filtroCambiado,

        key:
            "condicionLaboral",

        campo:
            "condicionLaboral",

        flag:
            "ignorarCondicionLaboral",

        dataBase

    });


    actualizarFiltroDependiente({

        filtroCambiado,

        key:
            "medico",

        campo:
            "medico",

        flag:
            "ignorarMedico",

        dataBase

    });


    actualizarFiltroDependiente({

        filtroCambiado,

        key:
            "cargo",

        campo:
            "cargoActual",

        flag:
            "ignorarCargo",

        dataBase

    });


    actualizarFiltroDependiente({

        filtroCambiado,

        key:
            "zona",

        campo:
            "zonaActual",

        flag:
            "ignorarZona",

        dataBase

    });


    actualizarFiltroDependiente({

        filtroCambiado,

        key:
            "estado",

        campo:
            "estado",

        flag:
            "ignorarEstado",

        dataBase,

        ordenPersonalizado: [

            "VIGENTE",
            "POR VENCER",
            "VENCIDA",
            "INDEFINIDA"

        ]

    });

}


function actualizarFiltroDependiente({

    filtroCambiado,

    key,

    campo,

    flag,

    dataBase,

    ordenPersonalizado = null

}) {

    if (
        filtroCambiado ===
        key
    ) {

        return;

    }


    cargarFiltroCampo({

        key,

        data:
            filtrarDataAvanzada(
                dataBase,
                {
                    [flag]:
                        true
                }
            ),

        campo,

        conservarSeleccion:
            true,

        ordenPersonalizado

    });

}


// =====================================================
// TEXTO FILTRO
// =====================================================

function actualizarTextoFiltro(
    key,
    opciones
) {

    const config =
        CONFIG_FILTROS[
            key
        ];


    const seleccion =
        estado.filtros[
            key
        ] ||
        [];


    if (
        !seleccion.length
    ) {

        setTexto(
            config.textoId,
            config.textoTodos
        );

        return;

    }


    if (
        seleccion.length ===
        1
    ) {

        const encontrado =
            opciones.find(
                opcion =>
                    opcion.value ===
                    seleccion[0]
            );


        setTexto(
            config.textoId,
            encontrado
                ? encontrado.label
                : seleccion[0]
        );

        return;

    }


    setTexto(
        config.textoId,
        `${seleccion.length} seleccionados`
    );

}


// =====================================================
// DATALIST NOMBRES
// =====================================================

function cargarListaNombres() {

    const datalist =
        document.getElementById(
            "listaNombresPersonal"
        );


    if (!datalist) {

        return;

    }


    datalist.innerHTML =
        obtenerPersonasUnicas(
            estado.data
        )
            .map(
                item => `

                    <option
                        value="${escapeHtml(item.nombre)}"
                        ${
                            item.dni
                                ? `label="DNI: ${escapeHtml(item.dni)}"`
                                : ""
                        }
                    ></option>

                `
            )
            .join("");

}


// =====================================================
// DATA POR AÑO
// =====================================================

function obtenerDataPorAnio() {

    const valor =
        document.getElementById(
            "filtroAnio"
        )?.value ||
        "TODOS";


    if (
        valor ===
        "TODOS" ||
        !valor
    ) {

        return [
            ...estado.data
        ];

    }


    const anio =
        Number(
            valor
        );


    return estado.data.filter(
        item =>
            item.anio ===
            anio ||
            !item.fechaInicio
    );

}


// =====================================================
// DATA FILTRADA
// =====================================================

function obtenerDataDinamicaFiltrada() {

    return filtrarDataAvanzada(
        obtenerDataPorAnio()
    );

}


function filtrarDataAvanzada(
    data,
    opciones = {}
) {

    const texto =
        normalizarTexto(
            document.getElementById(
                "filtroNombreGeneral"
            )?.value ||
            ""
        );


    return data.filter(
        item => {

            const valorMes =
                item.fechaInicio
                    ? String(
                        item.mesNumero
                    )
                    : "INDEFINIDO";


            const cumpleNombre =
                !texto ||
                item.nombre.includes(
                    texto
                ) ||
                item.dni.includes(
                    texto
                );


            const cumpleMes =
                opciones.ignorarMeses ||
                !estado.filtros.meses.length ||
                estado.filtros.meses.includes(
                    valorMes
                );


            const cumpleCondicion =
                opciones.ignorarCondicionLaboral ||
                !estado.filtros.condicionLaboral.length ||
                estado.filtros.condicionLaboral.includes(
                    item.condicionLaboral
                );


            const cumpleMedico =
                opciones.ignorarMedico ||
                !estado.filtros.medico.length ||
                estado.filtros.medico.includes(
                    item.medico
                );


            const cumpleCargo =
                opciones.ignorarCargo ||
                !estado.filtros.cargo.length ||
                estado.filtros.cargo.includes(
                    item.cargoActual
                );


            const cumpleZona =
                opciones.ignorarZona ||
                !estado.filtros.zona.length ||
                estado.filtros.zona.includes(
                    item.zonaActual
                );


            const cumpleEstado =
                opciones.ignorarEstado ||
                !estado.filtros.estado.length ||
                estado.filtros.estado.includes(
                    item.estado
                );


            return (
                cumpleNombre &&
                cumpleMes &&
                cumpleCondicion &&
                cumpleMedico &&
                cumpleCargo &&
                cumpleZona &&
                cumpleEstado
            );

        }
    );

}


// =====================================================
// RENDER DASHBOARD
// =====================================================

function renderDashboard() {

    const data =
        obtenerDataDinamicaFiltrada();


    estado.dataFiltradaActual =
        data;


    actualizarTextoAnio();


    actualizarKPIs(
        data
    );


    renderGraficoRestriccionesPorMes(
        data
    );


    renderGraficoCondicionLaboral(
        data
    );


    renderGraficoMedicos(
        data
    );


    renderGraficoCargoActual(
        data
    );


    renderGraficoZonaActual(
        data
    );


    renderTablaAlertas(
        data
    );

}


// =====================================================
// AÑO TEXTO
// =====================================================

function actualizarTextoAnio() {

    setTexto(
        "textoAnio",
        document.getElementById(
            "filtroAnio"
        )?.value ||
        "TODOS"
    );

}


// =====================================================
// KPI
// =====================================================

function actualizarKPIs(
    data
) {

    setTexto(
        "kpiTotalPersonas",
        contarPersonasUnicas(
            data
        ).toLocaleString(
            "es-PE"
        )
    );


    setTexto(
        "kpiVigentes",
        data.filter(
            item =>
                item.estado ===
                "VIGENTE"
        ).length.toLocaleString(
            "es-PE"
        )
    );


    setTexto(
        "kpiPorVencer",
        data.filter(
            item =>
                item.estado ===
                "POR VENCER"
        ).length.toLocaleString(
            "es-PE"
        )
    );


    setTexto(
        "kpiVencidas",
        data.filter(
            item =>
                item.estado ===
                "VENCIDA"
        ).length.toLocaleString(
            "es-PE"
        )
    );


    setTexto(
        "kpiIndefinidas",
        data.filter(
            item =>
                item.estado ===
                "INDEFINIDA"
        ).length.toLocaleString(
            "es-PE"
        )
    );

}


// =====================================================
// GRÁFICO MES
// =====================================================

function renderGraficoRestriccionesPorMes(
    data
) {

    destruirChart(
        "meses"
    );


    const canvas =
        document.getElementById(
            "graficoRestriccionesPorMes"
        );


    if (!canvas) {

        return;

    }


    ajustarGraficoNormal(
        canvas
    );


    const labels =
        [];


    const valores =
        [];


    const detalles =
        [];


    const indefinidos =
        data.filter(
            item =>
                !item.fechaInicio
        );


    if (
        indefinidos.length
    ) {

        labels.push(
            "INDEFINIDO"
        );


        valores.push(
            indefinidos.length
        );


        detalles.push(
            indefinidos
        );

    }


    const meses = [
        ...new Set(
            data
                .map(
                    item =>
                        item.mesNumero
                )
                .filter(Boolean)
        )
    ].sort(
        (a, b) =>
            a - b
    );


    meses.forEach(
        mes => {

            const detalle =
                data.filter(
                    item =>
                        item.mesNumero ===
                        mes
                );


            labels.push(
                MESES_ORDEN[
                    mes - 1
                ]
            );


            valores.push(
                detalle.length
            );


            detalles.push(
                detalle
            );

        }
    );


    estado.charts.meses =
        crearGraficoBarras({

            canvas,

            labels,

            valores,

            label:
                "Restricciones médicas",

            horizontal:
                false,

            onClickItem:
                index => {

                    abrirModalDetalleGeneral({

                        titulo:
                            "RESTRICCIONES MÉDICAS POR MES",

                        subtitulo:
                            labels[index],

                        data:
                            detalles[index] ||
                            []

                    });

                }

        });

}


// =====================================================
// CONDICIÓN LABORAL
// =====================================================

function renderGraficoCondicionLaboral(
    data
) {

    destruirChart(
        "condicionLaboral"
    );


    const canvas =
        document.getElementById(
            "graficoCondicionLaboral"
        );


    if (!canvas) {

        return;

    }


    const conteo =
        contarPorCampo(
            data,
            "condicionLaboral"
        );


    prepararGraficoHorizontal(
        canvas,
        conteo.length,
        ".chart-container--condicion",
        ".condicion-canvas-wrap"
    );


    const labels =
        conteo.map(
            item =>
                item.label
        );


    const valores =
        conteo.map(
            item =>
                item.valor
        );


    estado.charts.condicionLaboral =
        crearGraficoBarras({

            canvas,

            labels,

            valores,

            label:
                "Restricciones médicas",

            horizontal:
                true,

            onClickItem:
                index => {

                    const valor =
                        labels[index];


                    abrirModalDetalleGeneral({

                        titulo:
                            "RESTRICCIONES POR CONDICIÓN LABORAL",

                        subtitulo:
                            valor,

                        data:
                            data.filter(
                                item =>
                                    item.condicionLaboral ===
                                    valor
                            )

                    });

                }

        });

}


// =====================================================
// SELECTOR MÉDICOS
// =====================================================

function inicializarSelectorGraficoMedicos() {

    const acciones =
        document.querySelector(
            ".panel-medicos .panel-actions"
        );


    if (
        !acciones ||
        document.getElementById(
            "filtroGraficoMedicos"
        )
    ) {

        return;

    }


    const caja =
        document.createElement(
            "div"
        );


    caja.className =
        "mini-filtro";


    caja.innerHTML = `

        <span>
            Vista
        </span>

        <select
            id="filtroGraficoMedicos"
        >

            <option value="TOTAL">
                TOTAL
            </option>

            <option value="INDEFINIDO">
                INDEFINIDO
            </option>

            ${MESES_ORDEN.map(
                (mes, index) => `

                    <option
                        value="${index + 1}"
                    >
                        ${mes}
                    </option>

                `
            ).join("")}

        </select>

    `;


    acciones.prepend(
        caja
    );


    document.getElementById(
        "filtroGraficoMedicos"
    ).addEventListener(
        "change",
        event => {

            estado.filtroMedicoGrafico =
                event.target.value;


            renderGraficoMedicos(
                estado.dataFiltradaActual
            );

        }
    );

}


// =====================================================
// MÉDICOS - TORTA
// =====================================================

function renderGraficoMedicos(
    data
) {

    destruirChart(
        "medicos"
    );


    const canvas =
        document.getElementById(
            "graficoMedicosCertificadores"
        );


    if (!canvas) {

        return;

    }


    ajustarGraficoNormal(
        canvas
    );


    let dataGrafico = [
        ...data
    ];


    const vista =
        estado.filtroMedicoGrafico;


    if (
        vista ===
        "INDEFINIDO"
    ) {

        dataGrafico =
            data.filter(
                item =>
                    !item.fechaInicio
            );

    } else if (
        vista !==
        "TOTAL"
    ) {

        const mes =
            Number(
                vista
            );


        dataGrafico =
            data.filter(
                item =>
                    item.mesNumero ===
                    mes
            );

    }


    const conteo =
        contarPorCampo(
            dataGrafico,
            "medico"
        );


    const labels =
        conteo.map(
            item =>
                item.label
        );


    const valores =
        conteo.map(
            item =>
                item.valor
        );


    setTexto(
        "totalMedicosCertificadores",
        new Set(
            dataGrafico
                .map(
                    item =>
                        item.medico
                )
                .filter(
                    item =>
                        item &&
                        item !==
                        "NO ESPECIFICADO"
                )
        ).size.toLocaleString(
            "es-PE"
        )
    );


    estado.charts.medicos =
        crearGraficoTorta({

            canvas,

            labels,

            valores,

            label:
                "Restricciones certificadas",

            onClickItem:
                index => {

                    const medico =
                        labels[index];


                    abrirModalDetalleGeneral({

                        titulo:
                            "RESTRICCIONES POR MÉDICO CERTIFICADOR",

                        subtitulo:
                            medico,

                        data:
                            dataGrafico.filter(
                                item =>
                                    item.medico ===
                                    medico
                            )

                    });

                }

        });

}


// =====================================================
// CARGO
// =====================================================

function renderGraficoCargoActual(
    data
) {

    destruirChart(
        "cargo"
    );


    const canvas =
        document.getElementById(
            "graficoCargoActual"
        );


    if (!canvas) {

        return;

    }


    const conteo =
        contarPorCampo(
            data,
            "cargoActual"
        );


    prepararGraficoHorizontal(
        canvas,
        conteo.length,
        ".chart-container--cargo",
        ".cargo-canvas-wrap"
    );


    const labels =
        conteo.map(
            item =>
                item.label
        );


    const valores =
        conteo.map(
            item =>
                item.valor
        );


    estado.charts.cargo =
        crearGraficoBarras({

            canvas,

            labels,

            valores,

            label:
                "Restricciones médicas",

            horizontal:
                true,

            onClickItem:
                index => {

                    const cargo =
                        labels[index];


                    abrirModalDetalleGeneral({

                        titulo:
                            "RESTRICCIONES POR CARGO ACTUAL",

                        subtitulo:
                            cargo,

                        data:
                            data.filter(
                                item =>
                                    item.cargoActual ===
                                    cargo
                            )

                    });

                }

        });

}


// =====================================================
// ZONA
// =====================================================

function renderGraficoZonaActual(
    data
) {

    destruirChart(
        "zona"
    );


    const canvas =
        document.getElementById(
            "graficoZonaActual"
        );


    if (!canvas) {

        return;

    }


    const conteo =
        contarPorCampo(
            data,
            "zonaActual"
        );


    prepararGraficoHorizontal(
        canvas,
        conteo.length,
        ".chart-container--zona",
        ".zona-canvas-wrap"
    );


    const labels =
        conteo.map(
            item =>
                item.label
        );


    const valores =
        conteo.map(
            item =>
                item.valor
        );


    estado.charts.zona =
        crearGraficoBarras({

            canvas,

            labels,

            valores,

            label:
                "Restricciones médicas",

            horizontal:
                true,

            onClickItem:
                index => {

                    const zona =
                        labels[index];


                    abrirModalDetalleGeneral({

                        titulo:
                            "RESTRICCIONES POR ZONA ACTUAL",

                        subtitulo:
                            zona,

                        data:
                            data.filter(
                                item =>
                                    item.zonaActual ===
                                    zona
                            )

                    });

                }

        });

}


// =====================================================
// ALERTAS
// =====================================================

function renderTablaAlertas(
    data
) {

    const tbody =
        document.getElementById(
            "tablaAlertasRestricciones"
        );


    if (!tbody) {

        return;

    }


    const registros =
        ordenarAlertas(
            data
        );


    setTexto(
        "totalAlertasRestricciones",
        registros.length.toLocaleString(
            "es-PE"
        )
    );


    if (
        !registros.length
    ) {

        tbody.innerHTML = `

            <tr>

                <td
                    colspan="9"
                    class="empty-cell"
                >
                    No se encontraron restricciones.
                </td>

            </tr>

        `;

        return;

    }


    tbody.innerHTML =
        registros.map(
            (
                item,
                index
            ) => `

                <tr
                    class="fila-alerta-restriccion"
                    data-indice="${index}"
                    title="Haz clic para ver el detalle"
                >

                    <td>
                        ${index + 1}
                    </td>

                    <td>
                        ${escapeHtml(item.dni || "-")}
                    </td>

                    <td>

                        <strong>
                            ${escapeHtml(item.nombre)}
                        </strong>

                    </td>

                    <td>
                        ${escapeHtml(item.cargoActual)}
                    </td>

                    <td>
                        ${escapeHtml(item.zonaActual)}
                    </td>

                    <td>
                        ${escapeHtml(item.fechaInicioTexto)}
                    </td>

                    <td>
                        ${escapeHtml(item.fechaTerminoTexto)}
                    </td>

                    <td>

                        ${
                            item.diasRestantes ===
                            null

                                ? "-"

                                : escapeHtml(
                                    formatearDiasRestantes(
                                        item.diasRestantes
                                    )
                                )
                        }

                    </td>

                    <td>

                        <span
                            class="
                                badge-restriccion-estado
                                ${obtenerClaseEstado(item.estado)}
                            "
                        >

                            ${escapeHtml(item.estado)}

                        </span>

                    </td>

                </tr>

            `
        ).join("");


    tbody
        .querySelectorAll(
            ".fila-alerta-restriccion"
        )
        .forEach(
            fila => {

                fila.addEventListener(
                    "click",
                    () => {

                        const item =
                            registros[
                                Number(
                                    fila.dataset.indice
                                )
                            ];


                        if (!item) {

                            return;

                        }


                        abrirModalDetalleGeneral({

                            titulo:
                                "DETALLE DE RESTRICCIÓN MÉDICA",

                            subtitulo:
                                item.nombre,

                            data:
                                [
                                    item
                                ]

                        });

                    }
                );

            }
        );

}


// =====================================================
// ORDEN ALERTAS
// =====================================================

function ordenarAlertas(
    data
) {

    const prioridad = {

        "POR VENCER":
            1,

        "VENCIDA":
            2,

        "VIGENTE":
            3,

        "INDEFINIDA":
            4

    };


    return [
        ...data
    ].sort(
        (a, b) => {

            const pa =
                prioridad[
                    a.estado
                ] ||
                99;


            const pb =
                prioridad[
                    b.estado
                ] ||
                99;


            if (
                pa !==
                pb
            ) {

                return pa - pb;

            }


            if (
                a.diasRestantes !==
                null &&
                b.diasRestantes !==
                null
            ) {

                return (
                    a.diasRestantes -
                    b.diasRestantes
                );

            }


            return a.nombre.localeCompare(
                b.nombre,
                "es"
            );

        }
    );

}


// =====================================================
// PREPARAR MODAL DETALLE
// =====================================================

function prepararModalDetalle() {

    const tabla =
        document.querySelector(
            ".detalle-table--restricciones"
        );


    if (!tabla) {

        return;

    }


    const thead =
        tabla.querySelector(
            "thead"
        );


    if (thead) {

        thead.innerHTML = `

            <tr>

                <th>
                    DNI
                </th>

                <th>
                    Personal
                </th>

                <th>
                    Cargo actual
                </th>

                <th>
                    Zona actual
                </th>

                <th>
                    Contrato
                </th>

                <th>
                    Fecha inicio
                </th>

                <th>
                    Fecha término
                </th>

                <th>
                    Médico certificador
                </th>

                <th>
                    Motivo
                </th>

            </tr>

        `;

    }

}


// =====================================================
// MODAL GENERAL
// =====================================================

function inicializarModal() {

    const cerrar =
        document.getElementById(
            "btnCerrarModalDetalleGeneral"
        );


    if (cerrar) {

        cerrar.addEventListener(
            "click",
            cerrarModalDetalle
        );

    }


    const modal =
        document.getElementById(
            "modalDetalleGeneral"
        );


    if (modal) {

        modal.addEventListener(
            "click",
            event => {

                if (
                    event.target ===
                    modal
                ) {

                    cerrarModalDetalle();

                }

            }
        );

    }


    window.addEventListener(
        "keydown",
        event => {

            if (
                event.key ===
                "Escape"
            ) {

                cerrarModalMotivo();

                cerrarModalDetalle();

            }

        }
    );

}


// =====================================================
// ABRIR MODAL
// =====================================================

function abrirModalDetalleGeneral({

    titulo,

    subtitulo,

    data

}) {

    const modal =
        document.getElementById(
            "modalDetalleGeneral"
        );


    if (!modal) {

        return;

    }


    setTexto(
        "modalDetalleTitulo",
        titulo ||
        "DETALLE DE RESTRICCIONES MÉDICAS"
    );


    setTexto(
        "modalDetalleSubtitulo",
        subtitulo ||
        "Relación de registros"
    );


    setTexto(
        "modalResumenRegistros",
        data.length.toLocaleString(
            "es-PE"
        )
    );


    setTexto(
        "modalResumenPersonas",
        contarPersonasUnicas(
            data
        ).toLocaleString(
            "es-PE"
        )
    );


    setTexto(
        "modalResumenVigentes",
        data.filter(
            item =>
                item.estado ===
                "VIGENTE" ||
                item.estado ===
                "POR VENCER"
        ).length.toLocaleString(
            "es-PE"
        )
    );


    renderTablaDetalleGeneral(
        data
    );


    modal.classList.add(
        "modal-dashboard--show"
    );

}


// =====================================================
// TABLA DETALLE
// =====================================================

function renderTablaDetalleGeneral(
    data
) {

    const tbody =
        document.getElementById(
            "tablaDetalleGeneral"
        );


    if (!tbody) {

        return;

    }


    setTexto(
        "contadorTablaDetalleGeneral",
        `${data.length.toLocaleString("es-PE")} registros`
    );


    if (
        !data.length
    ) {

        tbody.innerHTML = `

            <tr>

                <td
                    colspan="9"
                    class="empty-cell"
                >
                    Sin registros.
                </td>

            </tr>

        `;

        return;

    }


    tbody.innerHTML =
        data.map(
            (
                item,
                index
            ) => `

                <tr>

                    <td>
                        ${escapeHtml(item.dni || "-")}
                    </td>

                    <td>

                        <strong>
                            ${escapeHtml(item.nombre)}
                        </strong>

                    </td>

                    <td>
                        ${escapeHtml(item.cargoActual)}
                    </td>

                    <td>
                        ${escapeHtml(item.zonaActual)}
                    </td>

                    <td>
                        ${escapeHtml(item.contrato)}
                    </td>

                    <td>
                        ${escapeHtml(item.fechaInicioTexto)}
                    </td>

                    <td>
                        ${escapeHtml(item.fechaTerminoTexto)}
                    </td>

                    <td>
                        ${escapeHtml(item.medico)}
                    </td>

                    <td>

                        <button
                            type="button"
                            class="btn-ver-motivo"
                            data-indice="${index}"
                        >

                            <i
                                class="fa-solid fa-eye"
                            ></i>

                            Motivo

                        </button>

                    </td>

                </tr>

            `
        ).join("");


    tbody
        .querySelectorAll(
            ".btn-ver-motivo"
        )
        .forEach(
            boton => {

                boton.addEventListener(
                    "click",
                    event => {

                        event.stopPropagation();


                        const item =
                            data[
                                Number(
                                    boton.dataset.indice
                                )
                            ];


                        if (!item) {

                            return;

                        }


                        abrirModalMotivo(
                            item
                        );

                    }
                );

            }
        );

}


// =====================================================
// MODAL MOTIVO
// =====================================================

function crearModalMotivo() {

    if (
        document.getElementById(
            "modalMotivoRestriccion"
        )
    ) {

        return;

    }


    const modal =
        document.createElement(
            "div"
        );


    modal.id =
        "modalMotivoRestriccion";


    modal.className =
        "modal-motivo-restriccion";


    modal.innerHTML = `

        <div
            class="modal-motivo-restriccion__box"
        >

            <div
                class="modal-motivo-restriccion__header"
            >

                <div>

                    <h3>
                        Motivo de la restricción
                    </h3>

                    <p
                        id="motivoNombrePersona"
                    >
                        -
                    </p>

                </div>


                <button
                    type="button"
                    id="btnCerrarModalMotivo"
                    class="modal-motivo-restriccion__close"
                    title="Cerrar"
                >

                    <i
                        class="fa-solid fa-xmark"
                    ></i>

                </button>

            </div>


            <div
                class="modal-motivo-restriccion__body"
            >

                <section
                    class="motivo-bloque"
                >

                    <h4>
                        Observaciones médicas
                    </h4>

                    <p
                        id="textoObservacionesMedicas"
                    >
                        -
                    </p>

                </section>


                <section
                    class="motivo-bloque"
                >

                    <h4>
                        Restricciones
                    </h4>

                    <p
                        id="textoRestriccionesMedicas"
                    >
                        -
                    </p>

                </section>

            </div>

        </div>

    `;


    document.body.appendChild(
        modal
    );


    document.getElementById(
        "btnCerrarModalMotivo"
    ).addEventListener(
        "click",
        cerrarModalMotivo
    );


    modal.addEventListener(
        "click",
        event => {

            if (
                event.target ===
                modal
            ) {

                cerrarModalMotivo();

            }

        }
    );

}


// =====================================================
// ABRIR MOTIVO
// =====================================================

function abrirModalMotivo(
    item
) {

    const modal =
        document.getElementById(
            "modalMotivoRestriccion"
        );


    if (!modal) {

        return;

    }


    setTexto(
        "motivoNombrePersona",
        item.dni
            ? `${item.nombre} | DNI: ${item.dni}`
            : item.nombre
    );


    const observaciones =
        item.observacionesMedicas ||
        item.restriccionMotivo ||
        item.observacion ||
        "SIN OBSERVACIONES REGISTRADAS";


    const restricciones =
        item.restriccion ||
        "SIN RESTRICCIONES ESPECIFICADAS";


    setTexto(
        "textoObservacionesMedicas",
        observaciones
    );


    setTexto(
        "textoRestriccionesMedicas",
        restricciones
    );


    modal.classList.add(
        "modal-motivo-restriccion--show"
    );

}


// =====================================================
// CERRAR MODALES
// =====================================================

function cerrarModalMotivo() {

    const modal =
        document.getElementById(
            "modalMotivoRestriccion"
        );


    if (modal) {

        modal.classList.remove(
            "modal-motivo-restriccion--show"
        );

    }

}


function cerrarModalDetalle() {

    const modal =
        document.getElementById(
            "modalDetalleGeneral"
        );


    if (modal) {

        modal.classList.remove(
            "modal-dashboard--show"
        );

    }

}


// =====================================================
// GRÁFICO BARRAS
// =====================================================

function crearGraficoBarras({

    canvas,

    labels,

    valores,

    label,

    horizontal = false,

    onClickItem = null

}) {

    return new Chart(
        canvas,
        {

            type:
                "bar",

            data: {

                labels,

                datasets: [
                    {

                        label,

                        data:
                            valores,

                        backgroundColor:
                            generarColoresPorValor(
                                valores
                            ),

                        borderRadius:
                            10,

                        borderSkipped:
                            false,

                        maxBarThickness:
                            horizontal
                                ? 25
                                : 42,

                        barPercentage:
                            0.88,

                        categoryPercentage:
                            0.88

                    }
                ]

            },

            options: {

                responsive:
                    true,

                maintainAspectRatio:
                    false,

                animation: {

                    duration:
                        550

                },

                indexAxis:
                    horizontal
                        ? "y"
                        : "x",

                interaction: {

                    mode:
                        "nearest",

                    intersect:
                        false

                },

                onClick:
                    (
                        event,
                        elements
                    ) => {

                        if (
                            !elements.length ||
                            !onClickItem
                        ) {

                            return;

                        }


                        onClickItem(
                            elements[0].index
                        );

                    },

                layout: {

                    padding: {

                        right:
                            horizontal
                                ? 65
                                : 12,

                        left:
                            horizontal
                                ? 12
                                : 4,

                        top:
                            18,

                        bottom:
                            6

                    }

                },

                plugins: {

                    legend: {

                        display:
                            false

                    },

                    tooltip: {

                        backgroundColor:
                            "rgba(15,23,42,.96)",

                        padding:
                            13,

                        displayColors:
                            false,

                        callbacks: {

                            label:
                                context =>
                                    `${label}: ${context.raw}`

                        }

                    },

                    datalabels: {

                        display:
                            true,

                        color:
                            "#334155",

                        anchor:
                            "end",

                        align:
                            horizontal
                                ? "right"
                                : "top",

                        offset:
                            horizontal
                                ? 5
                                : 4,

                        clip:
                            false,

                        clamp:
                            false,

                        font: {

                            weight:
                                "bold",

                            size:
                                10

                        },

                        formatter:
                            value =>
                                value > 0
                                    ? value
                                    : ""

                    }

                },

                scales: {

                    x: {

                        beginAtZero:
                            true,

                        ticks: {

                            precision:
                                0,

                            color:
                                "#475569",

                            font: {

                                size:
                                    10,

                                weight:
                                    "600"

                            }

                        },

                        grid: {

                            color:
                                "rgba(0,0,0,.05)"

                        }

                    },

                    y: {

                        beginAtZero:
                            true,

                        ticks: {

                            autoSkip:
                                false,

                            color:
                                "#475569",

                            font: {

                                size:
                                    10,

                                weight:
                                    "700"

                            },

                            callback:
                                function(value) {

                                    const texto =
                                        String(
                                            this.getLabelForValue(
                                                value
                                            )
                                        );


                                    return texto.length >
                                        55
                                        ? texto.substring(
                                            0,
                                            55
                                        ) +
                                        "..."
                                        : texto;

                                }

                        },

                        grid: {

                            color:
                                "rgba(0,0,0,.05)"

                        }

                    }

                }

            },

            plugins:
                obtenerPlugins()

        }
    );

}


// =====================================================
// GRÁFICO TORTA
// =====================================================

function crearGraficoTorta({

    canvas,

    labels,

    valores,

    label,

    onClickItem = null

}) {

    return new Chart(
        canvas,
        {

            type:
                "pie",

            data: {

                labels,

                datasets: [
                    {

                        label,

                        data:
                            valores,

                        backgroundColor:
                            generarColoresPorValor(
                                valores
                            ),

                        borderColor:
                            "#ffffff",

                        borderWidth:
                            3,

                        hoverOffset:
                            10

                    }
                ]

            },

            options: {

                responsive:
                    true,

                maintainAspectRatio:
                    false,

                animation: {

                    duration:
                        550

                },

                onClick:
                    (
                        event,
                        elements
                    ) => {

                        if (
                            !elements.length ||
                            !onClickItem
                        ) {

                            return;

                        }


                        onClickItem(
                            elements[0].index
                        );

                    },

                layout: {

                    padding: {

                        top:
                            12,

                        right:
                            20,

                        bottom:
                            6,

                        left:
                            20

                    }

                },

                plugins: {

                    legend: {

                        display:
                            true,

                        position:
                            "right",

                        labels: {

                            color:
                                "#475569",

                            boxWidth:
                                13,

                            boxHeight:
                                13,

                            padding:
                                13,

                            font: {

                                size:
                                    10,

                                weight:
                                    "700"

                            }

                        }

                    },

                    tooltip: {

                        backgroundColor:
                            "rgba(15,23,42,.96)",

                        padding:
                            13,

                        callbacks: {

                            label:
                                context => {

                                    const total =
                                        context.dataset.data.reduce(
                                            (
                                                suma,
                                                valor
                                            ) =>
                                                suma +
                                                Number(
                                                    valor ||
                                                    0
                                                ),
                                            0
                                        );


                                    const porcentaje =
                                        total
                                            ? (
                                                context.raw /
                                                total *
                                                100
                                            ).toFixed(
                                                1
                                            )
                                            : "0.0";


                                    return (
                                        `${context.label}: ` +
                                        `${context.raw} ` +
                                        `(${porcentaje}%)`
                                    );

                                }

                        }

                    },

                    datalabels: {

                        display:
                            context => {

                                const total =
                                    context.dataset.data.reduce(
                                        (
                                            suma,
                                            valor
                                        ) =>
                                            suma +
                                            Number(
                                                valor ||
                                                0
                                            ),
                                        0
                                    );


                                if (!total) {

                                    return false;

                                }


                                const valor =
                                    context.dataset.data[
                                        context.dataIndex
                                    ];


                                return (
                                    valor /
                                    total *
                                    100
                                ) >=
                                4;

                            },

                        color:
                            "#ffffff",

                        font: {

                            size:
                                10,

                            weight:
                                "900"

                        },

                        formatter:
                            (
                                valor,
                                context
                            ) => {

                                const total =
                                    context.dataset.data.reduce(
                                        (
                                            suma,
                                            item
                                        ) =>
                                            suma +
                                            Number(
                                                item ||
                                                0
                                            ),
                                        0
                                    );


                                if (!total) {

                                    return "";

                                }


                                return (
                                    valor /
                                    total *
                                    100
                                ).toFixed(
                                    1
                                ) +
                                "%";

                            }

                    }

                }

            },

            plugins:
                obtenerPlugins()

        }
    );

}


// =====================================================
// ALTURA GRÁFICO NORMAL
// =====================================================

function ajustarGraficoNormal(
    canvas
) {

    const contenedor =
        canvas.closest(
            ".chart-container"
        );


    if (contenedor) {

        contenedor.style.height =
            `${ALTO_GRAFICO_VISIBLE}px`;


        contenedor.style.minHeight =
            `${ALTO_GRAFICO_VISIBLE}px`;

    }

}


// =====================================================
// ALTURA HORIZONTAL + SCROLL
// =====================================================

function prepararGraficoHorizontal(
    canvas,
    cantidad,
    selectorContenedor,
    selectorWrapper
) {

    const contenedor =
        canvas.closest(
            selectorContenedor
        );


    const wrapper =
        canvas.closest(
            selectorWrapper
        );


    const alturaVisible =
        ALTO_GRAFICO_VISIBLE;


    const alturaContenido =
        cantidad >
        ITEMS_VISIBLES
            ? Math.max(
                alturaVisible,
                cantidad *
                ALTO_ITEM_HORIZONTAL +
                70
            )
            : alturaVisible;


    if (contenedor) {

        contenedor.style.height =
            `${alturaVisible}px`;


        contenedor.style.minHeight =
            `${alturaVisible}px`;


        contenedor.style.overflowY =
            cantidad >
            ITEMS_VISIBLES
                ? "auto"
                : "hidden";


        contenedor.style.overflowX =
            "hidden";


        contenedor.style.scrollbarGutter =
            "stable";

    }


    if (wrapper) {

        wrapper.style.width =
            "100%";


        wrapper.style.position =
            "relative";


        wrapper.style.height =
            `${alturaContenido}px`;


        wrapper.style.minHeight =
            `${alturaContenido}px`;

    }


    canvas.style.width =
        "100%";


    canvas.style.height =
        "100%";

}


// =====================================================
// CONTAR POR CAMPO
// =====================================================

function contarPorCampo(
    data,
    campo
) {

    const conteo =
        {};


    data.forEach(
        item => {

            const valor =
                item[
                    campo
                ] ||
                "NO ESPECIFICADO";


            conteo[
                valor
            ] =
                (
                    conteo[
                        valor
                    ] ||
                    0
                ) +
                1;

        }
    );


    return Object.entries(
        conteo
    )
        .map(
            (
                [
                    label,
                    valor
                ]
            ) => ({

                label,

                valor

            })
        )
        .sort(
            (a, b) => {

                if (
                    b.valor !==
                    a.valor
                ) {

                    return (
                        b.valor -
                        a.valor
                    );

                }


                return a.label.localeCompare(
                    b.label,
                    "es"
                );

            }
        );

}


// =====================================================
// PERSONAS ÚNICAS
// =====================================================

function obtenerPersonasUnicas(
    data
) {

    const mapa =
        new Map();


    data.forEach(
        item => {

            if (
                !mapa.has(
                    item.personaKey
                )
            ) {

                mapa.set(
                    item.personaKey,
                    {

                        nombre:
                            item.nombre,

                        dni:
                            item.dni

                    }
                );

            }

        }
    );


    return [
        ...mapa.values()
    ].sort(
        (a, b) =>
            a.nombre.localeCompare(
                b.nombre,
                "es"
            )
    );

}


function contarPersonasUnicas(
    data
) {

    return new Set(
        data
            .map(
                item =>
                    item.personaKey
            )
            .filter(Boolean)
    ).size;

}


// =====================================================
// FILTROS FLOTANTES
// =====================================================

function inicializarFiltrosFlotantes() {

    crearPanelFlotanteFiltro();


    document
        .querySelectorAll(
            ".filtro-check summary"
        )
        .forEach(
            summary => {

                summary.addEventListener(
                    "click",
                    event => {

                        event.preventDefault();

                        event.stopPropagation();


                        const details =
                            summary.closest(
                                ".filtro-check"
                            );


                        if (!details) {

                            return;

                        }


                        const key =
                            obtenerKeyFiltroDesdeDetails(
                                details.id
                            );


                        if (!key) {

                            return;

                        }


                        if (
                            filtroFlotanteActivo &&
                            filtroFlotanteActivo.key ===
                            key
                        ) {

                            cerrarPanelFlotanteFiltro();

                            return;

                        }


                        abrirPanelFlotanteFiltro(
                            details,
                            key
                        );

                    }
                );

            }
        );


    document.addEventListener(
        "click",
        event => {

            if (
                event.target.closest(
                    ".filtro-check"
                ) ||
                event.target.closest(
                    "#panelFiltroFlotante"
                )
            ) {

                return;

            }


            cerrarPanelFlotanteFiltro();

        }
    );


    window.addEventListener(
        "resize",
        reposicionarPanelFlotanteFiltro
    );


    window.addEventListener(
        "scroll",
        reposicionarPanelFlotanteFiltro,
        true
    );

}


// =====================================================
// CREAR PANEL FLOTANTE
// =====================================================

function crearPanelFlotanteFiltro() {

    if (
        document.getElementById(
            "panelFiltroFlotante"
        )
    ) {

        return;

    }


    const panel =
        document.createElement(
            "div"
        );


    panel.id =
        "panelFiltroFlotante";


    panel.className =
        "panel-filtro-flotante";


    panel.style.display =
        "none";


    document.body.appendChild(
        panel
    );

}


function obtenerKeyFiltroDesdeDetails(
    id
) {

    for (
        const [
            key,
            config
        ]
        of Object.entries(
            CONFIG_FILTROS
        )
    ) {

        if (
            config.detailsId ===
            id
        ) {

            return key;

        }

    }


    return null;

}


function abrirPanelFlotanteFiltro(
    details,
    key
) {

    const panel =
        document.getElementById(
            "panelFiltroFlotante"
        );


    if (!panel) {

        return;

    }


    document
        .querySelectorAll(
            ".filtro-check"
        )
        .forEach(
            item =>
                item.classList.remove(
                    "filtro-check--activo"
                )
        );


    filtroFlotanteActivo = {

        details,

        key

    };


    details.classList.add(
        "filtro-check--activo"
    );


    renderPanelFlotanteFiltro(
        key
    );


    panel.style.display =
        "block";


    posicionarPanelFlotanteFiltro();

}


function renderPanelFlotanteFiltro(
    key
) {

    const panel =
        document.getElementById(
            "panelFiltroFlotante"
        );


    const config =
        CONFIG_FILTROS[
            key
        ];


    const original =
        document.getElementById(
            config?.panelId
        );


    if (
        !panel ||
        !original
    ) {

        return;

    }


    const seleccion =
        estado.filtros[
            key
        ] ||
        [];


    panel.innerHTML =
        "";


    [
        ...original.querySelectorAll(
            ".check-option"
        )
    ].forEach(
        opcion => {

            const originalInput =
                opcion.querySelector(
                    "input"
                );


            const originalSpan =
                opcion.querySelector(
                    "span"
                );


            if (
                !originalInput ||
                !originalSpan
            ) {

                return;

            }


            const value =
                originalInput.value;


            const esTodos =
                value ===
                "TODOS";


            const item =
                document.createElement(
                    "label"
                );


            item.className =
                esTodos
                    ? "check-option check-option--all"
                    : "check-option";


            const input =
                document.createElement(
                    "input"
                );


            input.type =
                "checkbox";


            input.value =
                value;


            input.checked =
                esTodos
                    ? seleccion.length ===
                        0
                    : seleccion.includes(
                        value
                    );


            const span =
                document.createElement(
                    "span"
                );


            span.textContent =
                originalSpan.textContent;


            item.append(
                input,
                span
            );


            input.addEventListener(
                "click",
                event =>
                    event.stopPropagation()
            );


            input.addEventListener(
                "change",
                event => {

                    event.stopPropagation();


                    aplicarCambioFiltroCheckbox({

                        key,

                        value,

                        checked:
                            input.checked

                    });

                }
            );


            panel.appendChild(
                item
            );

        }
    );

}


function posicionarPanelFlotanteFiltro() {

    if (
        !filtroFlotanteActivo
    ) {

        return;

    }


    const panel =
        document.getElementById(
            "panelFiltroFlotante"
        );


    const details =
        filtroFlotanteActivo.details;


    if (
        !panel ||
        !details
    ) {

        return;

    }


    const rect =
        details.getBoundingClientRect();


    const margen =
        12;


    const separacion =
        10;


    let ancho =
        details.classList.contains(
            "filtro-check--wide"
        )
            ? 440
            : 330;


    ancho =
        Math.min(
            ancho,
            window.innerWidth -
            margen *
            2
        );


    let left =
        rect.left;


    if (
        left +
        ancho >
        window.innerWidth -
        margen
    ) {

        left =
            window.innerWidth -
            ancho -
            margen;

    }


    left =
        Math.max(
            margen,
            left
        );


    const abajo =
        window.innerHeight -
        rect.bottom -
        separacion -
        margen;


    const arriba =
        rect.top -
        separacion -
        margen;


    const abrirArriba =
        abajo <
        280 &&
        arriba >
        abajo;


    const disponible =
        abrirArriba
            ? arriba
            : abajo;


    panel.style.position =
        "fixed";


    panel.style.left =
        `${left}px`;


    panel.style.width =
        `${ancho}px`;


    panel.style.maxHeight =
        `${Math.max(
            180,
            Math.min(
                460,
                disponible
            )
        )}px`;


    panel.style.overflowY =
        "auto";


    panel.style.overflowX =
        "hidden";


    panel.style.zIndex =
        "999999";


    panel.style.background =
        "#ffffff";


    if (
        abrirArriba
    ) {

        panel.style.top =
            "auto";


        panel.style.bottom =
            `${
                window.innerHeight -
                rect.top +
                separacion
            }px`;

    } else {

        panel.style.bottom =
            "auto";


        panel.style.top =
            `${
                rect.bottom +
                separacion
            }px`;

    }

}


function reposicionarPanelFlotanteFiltro() {

    if (
        filtroFlotanteActivo
    ) {

        posicionarPanelFlotanteFiltro();

    }

}


function cerrarPanelFlotanteFiltro() {

    const panel =
        document.getElementById(
            "panelFiltroFlotante"
        );


    if (panel) {

        panel.style.display =
            "none";


        panel.innerHTML =
            "";

    }


    document
        .querySelectorAll(
            ".filtro-check"
        )
        .forEach(
            item =>
                item.classList.remove(
                    "filtro-check--activo"
                )
        );


    filtroFlotanteActivo =
        null;

}


// =====================================================
// LIMPIAR
// =====================================================

function inicializarBotonLimpiar() {

    const boton =
        document.getElementById(
            "btnLimpiarTodo"
        );


    if (!boton) {

        return;

    }


    boton.addEventListener(
        "click",
        () => {

            limpiarFiltrosSinCambiarAnio();


            cargarFiltrosMultiples();


            cerrarPanelFlotanteFiltro();


            estado.filtroMedicoGrafico =
                "TOTAL";


            const selector =
                document.getElementById(
                    "filtroGraficoMedicos"
                );


            if (selector) {

                selector.value =
                    "TOTAL";

            }


            renderDashboard();

        }
    );

}


function limpiarFiltrosSinCambiarAnio() {

    estado.filtros = {

        meses: [],

        condicionLaboral: [],

        medico: [],

        cargo: [],

        zona: [],

        estado: []

    };


    const input =
        document.getElementById(
            "filtroNombreGeneral"
        );


    if (input) {

        input.value =
            "";

    }

}


// =====================================================
// OBTENER CAMPO
// =====================================================

function obtenerCampo(
    objeto,
    alternativas
) {

    const claves =
        Object.keys(
            objeto
        );


    for (
        const alternativa
        of alternativas
    ) {

        const buscada =
            normalizarCabecera(
                alternativa
            );


        const clave =
            claves.find(
                item =>
                    normalizarCabecera(
                        item
                    ) ===
                    buscada
            );


        if (
            clave ===
            undefined
        ) {

            continue;

        }


        const valor =
            objeto[
                clave
            ];


        if (
            valor !==
            undefined &&
            valor !==
            null &&
            String(
                valor
            ).trim() !==
            ""
        ) {

            return valor;

        }

    }


    return "";

}


// =====================================================
// NORMALIZACIONES
// =====================================================

function normalizarCabecera(
    valor
) {

    return String(
        valor ??
        ""
    )
        .trim()
        .replace(
            /\s+/g,
            " "
        )
        .normalize(
            "NFD"
        )
        .replace(
            /[\u0300-\u036f]/g,
            ""
        )
        .toUpperCase();

}


function normalizarTexto(
    valor
) {

    return String(
        valor ??
        ""
    )
        .trim()
        .replace(
            /\s+/g,
            " "
        )
        .toUpperCase();

}


function limpiarTexto(
    valor
) {

    return String(
        valor ??
        ""
    )
        .trim()
        .replace(
            /\r\n/g,
            "\n"
        )
        .replace(
            /\r/g,
            "\n"
        );

}


function normalizarClave(
    valor
) {

    return normalizarTexto(
        valor
    )
        .normalize(
            "NFD"
        )
        .replace(
            /[\u0300-\u036f]/g,
            ""
        );

}


// =====================================================
// DOCUMENTO
// =====================================================

function limpiarDocumento(
    valor
) {

    if (
        valor ===
        null ||
        valor ===
        undefined
    ) {

        return "";

    }


    if (
        typeof valor ===
        "number" &&
        Number.isFinite(
            valor
        )
    ) {

        return String(
            Math.trunc(
                valor
            )
        );

    }


    return String(
        valor
    )
        .trim()
        .replace(
            /\.0$/,
            ""
        )
        .replace(
            /\D/g,
            ""
        );

}


// =====================================================
// NÚMERO
// =====================================================

function limpiarNumero(
    valor
) {

    if (
        valor ===
        null ||
        valor ===
        undefined ||
        valor ===
        ""
    ) {

        return null;

    }


    const numero =
        Number(
            valor
        );


    return Number.isFinite(
        numero
    )
        ? numero
        : null;

}


// =====================================================
// ZONA
// =====================================================

function normalizarZona(
    valor
) {

    const texto =
        normalizarTexto(
            valor
        );


    if (!texto) {

        return "NO ESPECIFICADO";

    }


    if (
        /^\d+$/.test(
            texto
        )
    ) {

        return `ZONA ${texto}`;

    }


    const match =
        texto.match(
            /^ZONA\s*(\d+)$/
        );


    if (match) {

        return `ZONA ${match[1]}`;

    }


    return texto;

}


// =====================================================
// EDAD
// =====================================================

function extraerEdadDeNombre(
    nombre
) {

    const match =
        String(
            nombre ??
            ""
        ).match(
            /\((\d{1,3})\)\s*$/
        );


    return match
        ? Number(
            match[1]
        )
        : null;

}


function limpiarEdadDeNombre(
    nombre
) {

    return normalizarTexto(
        String(
            nombre ??
            ""
        ).replace(
            /\s*\(\d{1,3}\)\s*$/,
            ""
        )
    );

}


// =====================================================
// FECHAS
// =====================================================

function convertirFecha(
    valor
) {

    if (
        valor ===
        null ||
        valor ===
        undefined ||
        valor ===
        ""
    ) {

        return null;

    }


    if (
        valor instanceof
        Date &&
        !Number.isNaN(
            valor.getTime()
        )
    ) {

        return normalizarDate(
            valor
        );

    }


    if (
        typeof valor ===
        "number"
    ) {

        const info =
            XLSX.SSF.parse_date_code(
                valor
            );


        if (
            info &&
            info.y &&
            info.m &&
            info.d
        ) {

            return new Date(
                info.y,
                info.m - 1,
                info.d
            );

        }

    }


    const texto =
        String(
            valor
        ).trim();


    if (
        !texto ||
        normalizarTexto(
            texto
        ).includes(
            "INDEFIN"
        )
    ) {

        return null;

    }


    let match =
        texto.match(
            /(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/
        );


    if (match) {

        let anio =
            Number(
                match[3]
            );


        if (
            anio <
            100
        ) {

            anio +=
                2000;

        }


        const fecha =
            new Date(
                anio,
                Number(
                    match[2]
                ) - 1,
                Number(
                    match[1]
                )
            );


        if (
            !Number.isNaN(
                fecha.getTime()
            )
        ) {

            return normalizarDate(
                fecha
            );

        }

    }


    match =
        texto.match(
            /(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/
        );


    if (match) {

        const fecha =
            new Date(
                Number(
                    match[1]
                ),
                Number(
                    match[2]
                ) - 1,
                Number(
                    match[3]
                )
            );


        if (
            !Number.isNaN(
                fecha.getTime()
            )
        ) {

            return normalizarDate(
                fecha
            );

        }

    }


    return null;

}


function normalizarDate(
    fecha
) {

    return new Date(
        fecha.getFullYear(),
        fecha.getMonth(),
        fecha.getDate()
    );

}


function formatearFecha(
    fecha
) {

    if (
        !(
            fecha instanceof
            Date
        ) ||
        Number.isNaN(
            fecha.getTime()
        )
    ) {

        return "";

    }


    return fecha.toLocaleDateString(
        "es-PE",
        {

            day:
                "2-digit",

            month:
                "2-digit",

            year:
                "numeric"

        }
    );

}


// =====================================================
// DÍAS RESTANTES
// =====================================================

function formatearDiasRestantes(
    dias
) {

    if (
        dias ===
        0
    ) {

        return "VENCE HOY";

    }


    if (
        dias ===
        1
    ) {

        return "1 DÍA";

    }


    if (
        dias >
        1
    ) {

        return `${dias} DÍAS`;

    }


    if (
        dias ===
        -1
    ) {

        return "VENCIDA HACE 1 DÍA";

    }


    return (
        `VENCIDA HACE ` +
        `${Math.abs(dias)} DÍAS`
    );

}


// =====================================================
// CLASE ESTADO
// =====================================================

function obtenerClaseEstado(
    valor
) {

    switch (
        valor
    ) {

        case "VIGENTE":

            return "badge-restriccion-estado--vigente";


        case "POR VENCER":

            return "badge-restriccion-estado--por-vencer";


        case "VENCIDA":

            return "badge-restriccion-estado--vencida";


        case "INDEFINIDA":

            return "badge-restriccion-estado--indefinida";


        default:

            return "";

    }

}


// =====================================================
// COLORES
// =====================================================

function generarColoresPorValor(
    valores
) {

    if (
        !Array.isArray(
            valores
        ) ||
        !valores.length
    ) {

        return [];

    }


    const max =
        Math.max(
            ...valores
        );


    return valores.map(
        valor => {

            const ratio =
                max > 0
                    ? valor / max
                    : 0;


            if (
                ratio >= 0.90
            ) {
                return "#E78B8B";
            }


            if (
                ratio >= 0.80
            ) {
                return "#F0A36B";
            }


            if (
                ratio >= 0.70
            ) {
                return "#F6BF5B";
            }


            if (
                ratio >= 0.60
            ) {
                return "#E4D86F";
            }


            if (
                ratio >= 0.50
            ) {
                return "#C9D97B";
            }


            if (
                ratio >= 0.40
            ) {
                return "#A9D18E";
            }


            if (
                ratio >= 0.30
            ) {
                return "#7FC8A9";
            }


            if (
                ratio >= 0.20
            ) {
                return "#66BFBF";
            }


            return "#8EC5E8";

        }
    );

}


// =====================================================
// PLUGINS
// =====================================================

function obtenerPlugins() {

    return window.ChartDataLabels
        ? [
            ChartDataLabels
        ]
        : [];

}


// =====================================================
// DESTRUIR CHART
// =====================================================

function destruirChart(
    key
) {

    if (
        estado.charts[
            key
        ]
    ) {

        estado.charts[
            key
        ].destroy();


        estado.charts[
            key
        ] =
            null;

    }

}


// =====================================================
// DOM
// =====================================================

function setTexto(
    id,
    valor
) {

    const elemento =
        document.getElementById(
            id
        );


    if (elemento) {

        elemento.textContent =
            valor;

    }

}


function mostrarLoading(
    mostrar
) {

    const overlay =
        document.getElementById(
            "loadingSistema"
        );


    if (!overlay) {

        return;

    }


    overlay.style.display =
        mostrar
            ? "flex"
            : "none";

}


// =====================================================
// ESCAPE HTML
// =====================================================

function escapeHtml(
    valor
) {

    return String(
        valor ??
        ""
    )
        .replaceAll(
            "&",
            "&amp;"
        )
        .replaceAll(
            "<",
            "&lt;"
        )
        .replaceAll(
            ">",
            "&gt;"
        )
        .replaceAll(
            '"',
            "&quot;"
        )
        .replaceAll(
            "'",
            "&#039;"
        );

}
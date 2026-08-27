
// =====================================================
// CONSULTA.JS
// CONSULTA INTEGRAL DE PERSONAL
// LISTO PARA PRODUCCIÓN
//
// FUNCIONES:
// - Personal desde Firestore.
// - Descansos médicos desde DM.xlsx.
// - Papeletas desde papeletas.xlsx.
// - Búsqueda por DNI o nombre.
// - Imagen según sexo.
// - Dashboard desplegable por tarjeta.
// - Gráficos interactivos.
// - Modal de detalle al hacer clic.
// - Colores institucionales y paleta dinámica.
// =====================================================

import * as XLSX from "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/+esm";

import {
    collection,
    getDocs
} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js";

import {
    db,
    obtenerArchivoDescansosMedicos,
    descargarExcelStorage
} from "../backend/api.js";

// =====================================================
// CONFIGURACIÓN
// =====================================================

const CONFIG = {
    coleccionPersonal: "personal",

    rutaPapeletas:
        "RENDIMIENTO PERSONAL/PAPELETAS/papeletas.xlsx",

    limiteSugerencias: 12,

    imagenHombre:
        "../imagenes/hombre.png",

    imagenMujer:
        "../imagenes/mujer.png"
};

// =====================================================
// MESES
// =====================================================

const MESES = [
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

const NUMERO_MES = {
    ENERO: 1,
    FEBRERO: 2,
    MARZO: 3,
    ABRIL: 4,
    MAYO: 5,
    JUNIO: 6,
    JULIO: 7,
    AGOSTO: 8,
    SETIEMBRE: 9,
    SEPTIEMBRE: 9,
    OCTUBRE: 10,
    NOVIEMBRE: 11,
    DICIEMBRE: 12
};

// =====================================================
// COLORES
// =====================================================

const COLORES = {
    primary: "#556B2F",
    primaryDark: "#3D4D1F",

    rojoSuave: "#E78B8B",
    naranjaSuave: "#F0A36B",
    amarilloSuave: "#F6BF5B",
    limaSuave: "#E4D86F",
    verdeLima: "#C9D97B",
    verdeClaro: "#A9D18E",
    verdeAgua: "#7FC8A9",
    turquesa: "#66BFBF",
    azulClaro: "#8EC5E8",

    blanco: "#FFFFFF",
    texto: "#374151",
    textoSuave: "#475569",
    grid: "rgba(0,0,0,0.07)"
};

// =====================================================
// ESTADO GLOBAL
// =====================================================

const estado = {
    personal: [],

    descansosMedicos: [],

    papeletas: [],

    personaSeleccionada: null,

    descansosSeleccionados: [],

    papeletasSeleccionadas: [],

    dashboardActivo: null,

    eventosInicializados: false,

    charts: {
        descansosMes: null,
        descansosTipo: null,
        descansosDiagnostico: null,

        papeletasMes: null,
        papeletasTipo: null,
        papeletasMotivo: null
    }
};

// =====================================================
// DOM
// =====================================================

const elementos = {
    loading:
        document.getElementById("loadingSistema"),

    buscador:
        document.getElementById("buscadorPersonal"),

    btnLimpiar:
        document.getElementById("btnLimpiarBusqueda"),

    listaSugerencias:
        document.getElementById("listaSugerenciasPersonal"),

    contenidoSugerencias:
        document.getElementById("contenidoSugerenciasPersonal"),

    mensajeSinResultados:
        document.getElementById("mensajeSinResultados"),

    estadoInicial:
        document.getElementById("estadoInicialConsulta"),

    estadoError:
        document.getElementById("estadoErrorConsulta"),

    mensajeError:
        document.getElementById("mensajeErrorConsulta"),

    btnReintentar:
        document.getElementById("btnReintentarConsulta"),

    fichaPersonal:
        document.getElementById("fichaPersonal"),

    mensajeSinDescansos:
        document.getElementById("mensajeSinDescansos"),

    mensajeSinPapeletas:
        document.getElementById("mensajeSinPapeletas"),

    contenedorGraficosDescansos:
        document.getElementById("contenedorGraficosDescansos"),

    contenedorGraficosPapeletas:
        document.getElementById("contenedorGraficosPapeletas"),

    btnDashboardDescansos:
        document.getElementById("btnDashboardDescansos"),

    btnCerrarDashboardDescansos:
        document.getElementById("btnCerrarDashboardDescansos"),

    dashboardDescansos:
        document.getElementById("dashboardDescansos"),

    btnDashboardPapeletas:
        document.getElementById("btnDashboardPapeletas"),

    btnCerrarDashboardPapeletas:
        document.getElementById("btnCerrarDashboardPapeletas"),

    dashboardPapeletas:
        document.getElementById("dashboardPapeletas"),

    modalDetalle:
        document.getElementById("modalDetalleGrafico"),

    btnCerrarModalDetalle:
        document.getElementById("btnCerrarModalDetalle"),

    cabeceraModalDetalle:
        document.getElementById("cabeceraModalDetalle"),

    cuerpoModalDetalle:
        document.getElementById("cuerpoModalDetalle")
};

// =====================================================
// INICIO
// =====================================================

window.addEventListener(
    "DOMContentLoaded",
    inicializarSistema
);

async function inicializarSistema() {
    mostrarLoading(true);
    ocultarError();

    try {
        inicializarEventos();
        registrarPluginChart();

        await cargarTodasLasFuentes();

        console.info("✅ Consulta integral inicializada");
        console.info("Personal:", estado.personal.length);
        console.info("Descansos:", estado.descansosMedicos.length);
        console.info("Papeletas:", estado.papeletas.length);

    } catch (error) {
        console.error(
            "❌ Error al inicializar consulta:",
            error
        );

        mostrarError(
            "No se pudo cargar la información. Revisa Firestore, Firebase Storage y las rutas configuradas."
        );

    } finally {
        mostrarLoading(false);
    }
}

// =====================================================
// CHART.JS
// =====================================================

function registrarPluginChart() {
    if (typeof window.Chart === "undefined") {
        throw new Error(
            "Chart.js no está disponible."
        );
    }

    if (window.ChartDataLabels) {
        Chart.register(
            window.ChartDataLabels
        );
    }
}

function obtenerPluginsChart() {
    return window.ChartDataLabels
        ? [window.ChartDataLabels]
        : [];
}

// =====================================================
// CARGA DE FUENTES
// =====================================================

async function cargarTodasLasFuentes() {
    const resultados = await Promise.allSettled([
        cargarPersonalFirestore(),
        cargarDescansosMedicos(),
        cargarPapeletas()
    ]);

    const [
        resultadoPersonal,
        resultadoDescansos,
        resultadoPapeletas
    ] = resultados;

    if (resultadoPersonal.status === "rejected") {
        throw resultadoPersonal.reason;
    }

    if (resultadoDescansos.status === "rejected") {
        console.warn(
            "No se pudo cargar DM.xlsx:",
            resultadoDescansos.reason
        );

        estado.descansosMedicos = [];
    }

    if (resultadoPapeletas.status === "rejected") {
        console.warn(
            "No se pudo cargar papeletas.xlsx:",
            resultadoPapeletas.reason
        );

        estado.papeletas = [];
    }
}

// =====================================================
// PERSONAL DESDE FIRESTORE
// =====================================================

async function cargarPersonalFirestore() {
    const referencia = collection(
        db,
        CONFIG.coleccionPersonal
    );

    const snapshot = await getDocs(
        referencia
    );

    const mapaPersonal = new Map();

    snapshot.forEach(documento => {
        const datos = documento.data() || {};

        const dni = limpiarDNI(
            datos.dni ||
            documento.id
        );

        if (!dni) {
            return;
        }

        const sexo = normalizarSexo(
            datos.sexo
        );

        const persona = {
            id: documento.id,

            dni,

            apellidosNombres:
                normalizarTextoVisible(
                    datos.apellidos_nombres ||
                    datos.apellidosNombres ||
                    datos.nombre ||
                    datos.nombres ||
                    ""
                ),

            area:
                normalizarTextoVisible(
                    datos.area || ""
                ),

            condicionLaboral:
                normalizarTextoVisible(
                    datos.condicion_laboral ||
                    datos.condicionLaboral ||
                    ""
                ),

            funcion:
                normalizarTextoVisible(
                    datos.funcion ||
                    datos.cargo ||
                    ""
                ),

            sexo,

            promedioNota:
                convertirNota(
                    datos.promedio_nota
                ),

            estado:
                normalizarTextoVisible(
                    datos.estado ||
                    "ACTIVO"
                ),

            imagen:
                obtenerImagenSegunSexo(
                    sexo
                ),

            busqueda:
                ""
        };

        persona.busqueda =
            construirTextoBusqueda(
                persona
            );

        const anterior =
            mapaPersonal.get(
                dni
            );

        mapaPersonal.set(
            dni,
            anterior
                ? combinarPersonas(
                    anterior,
                    persona
                )
                : persona
        );
    });

    estado.personal = [
        ...mapaPersonal.values()
    ].sort(
        (a, b) =>
            a.apellidosNombres.localeCompare(
                b.apellidosNombres,
                "es"
            )
    );

    if (!estado.personal.length) {
        throw new Error(
            "La colección personal no contiene registros."
        );
    }
}

function construirTextoBusqueda(persona) {
    return normalizarBusqueda(
        [
            persona.dni,
            persona.apellidosNombres,
            persona.area,
            persona.funcion,
            persona.condicionLaboral
        ].join(" ")
    );
}

function combinarPersonas(
    anterior,
    nuevo
) {
    const combinado = {
        ...anterior,

        apellidosNombres:
            nuevo.apellidosNombres ||
            anterior.apellidosNombres,

        area:
            nuevo.area ||
            anterior.area,

        condicionLaboral:
            nuevo.condicionLaboral ||
            anterior.condicionLaboral,

        funcion:
            nuevo.funcion ||
            anterior.funcion,

        sexo:
            nuevo.sexo ||
            anterior.sexo,

        promedioNota:
            nuevo.promedioNota !== null
                ? nuevo.promedioNota
                : anterior.promedioNota,

        estado:
            nuevo.estado ||
            anterior.estado
    };

    combinado.imagen =
        obtenerImagenSegunSexo(
            combinado.sexo
        );

    combinado.busqueda =
        construirTextoBusqueda(
            combinado
        );

    return combinado;
}

function obtenerImagenSegunSexo(sexo) {
    return normalizarSexo(sexo) === "FEMENINO"
        ? CONFIG.imagenMujer
        : CONFIG.imagenHombre;
}

// =====================================================
// DESCANSOS MÉDICOS
// =====================================================

async function cargarDescansosMedicos() {
    const buffer =
        await obtenerArchivoDescansosMedicos();

    validarBufferExcel(
        buffer,
        "DM.xlsx"
    );

    const workbook = XLSX.read(
        buffer,
        {
            type: "array",
            cellDates: true
        }
    );

    const registros = [];

    workbook.SheetNames.forEach(
        nombreHoja => {
            const hoja =
                workbook.Sheets[nombreHoja];

            if (!hoja) {
                return;
            }

            const filas =
                XLSX.utils.sheet_to_json(
                    hoja,
                    {
                        defval: "",
                        raw: true
                    }
                );

            filas.forEach(
                (fila, indice) => {
                    const registro =
                        normalizarDescansoMedico(
                            fila,
                            indice + 2,
                            nombreHoja
                        );

                    if (
                        registro.dni &&
                        registro.fechaInicio
                    ) {
                        registros.push(
                            registro
                        );
                    }
                }
            );
        }
    );

    estado.descansosMedicos =
        eliminarDuplicadosDescansos(
            registros
        );
}

function normalizarDescansoMedico(
    fila,
    filaExcel,
    hoja
) {
    const dni = limpiarDNI(
        obtenerCampoFlexible(
            fila,
            [
                "numdocumento",
                "NUMDOCUMENTO",
                "NUM_DOCUMENTO",
                "DNI",
                "DOCUMENTO"
            ]
        )
    );

    const nombre =
        normalizarTextoVisible(
            obtenerCampoFlexible(
                fila,
                [
                    "des_nomapell",
                    "DES_NOMAPELL",
                    "APELLIDOS Y NOMBRES",
                    "NOMBRE"
                ]
            )
        );

    const fechaInicio =
        convertirFecha(
            obtenerCampoFlexible(
                fila,
                [
                    "fec_inicio",
                    "FEC_INICIO",
                    "FECHA INICIO",
                    "FECHA DE INICIO",
                    "INICIO"
                ]
            )
        );

    const fechaFin =
        convertirFecha(
            obtenerCampoFlexible(
                fila,
                [
                    "fec_termino",
                    "FEC_TERMINO",
                    "FECHA FIN",
                    "FECHA DE FIN",
                    "FIN"
                ]
            )
        );

    const numeroDias =
        convertirNumero(
            obtenerCampoFlexible(
                fila,
                [
                    "cant. Días",
                    "cant. días",
                    "CANT. DÍAS",
                    "CANT. DIAS",
                    "CANT DIAS",
                    "CANT_DIAS",
                    "DIAS",
                    "DÍAS"
                ]
            )
        );

    const tipoDocumento =
        normalizarTextoVisible(
            obtenerCampoFlexible(
                fila,
                [
                    "txttipodocumento",
                    "TXTTIPODOCUMENTO",
                    "TIPO DOCUMENTO",
                    "TIPO DE DOCUMENTO"
                ]
            )
        );

    const contingencia =
        normalizarTextoVisible(
            obtenerCampoFlexible(
                fila,
                [
                    "txtcontingencia",
                    "TXTCONTINGENCIA",
                    "CONTINGENCIA"
                ]
            )
        );

    const diagnostico =
        normalizarTextoVisible(
            obtenerCampoFlexible(
                fila,
                [
                    "txtotrosdiagnosticos",
                    "TXTOTROSDIAGNOSTICOS",
                    "OTROS DIAGNOSTICOS",
                    "OTROS DIAGNÓSTICOS",
                    "DIAGNOSTICO",
                    "DIAGNÓSTICO"
                ]
            )
        );

    const establecimiento =
        normalizarTextoVisible(
            obtenerCampoFlexible(
                fila,
                [
                    "txtestablecimientosalud",
                    "TXTESTABLECIMIENTOSALUD",
                    "ESTABLECIMIENTO SALUD",
                    "ESTABLECIMIENTO DE SALUD",
                    "HOSPITAL"
                ]
            )
        );

    const mesNumero =
        fechaInicio
            ? fechaInicio.getMonth() + 1
            : null;

    return {
        dni,

        nombre,

        fechaInicio,

        fechaFin,

        fechaInicioTexto:
            formatearFecha(
                fechaInicio
            ),

        fechaFinTexto:
            formatearFecha(
                fechaFin
            ),

        numeroDias,

        mesNumero,

        mes:
            mesNumero
                ? MESES[mesNumero - 1]
                : "",

        tipoDocumento:
            tipoDocumento ||
            "NO ESPECIFICADO",

        contingencia:
            contingencia ||
            "NO ESPECIFICADO",

        diagnostico:
            diagnostico ||
            "NO ESPECIFICADO",

        establecimiento:
            establecimiento ||
            "NO ESPECIFICADO",

        filaExcel,

        hoja
    };
}

function eliminarDuplicadosDescansos(
    registros
) {
    const mapa = new Map();

    registros.forEach(
        registro => {
            const clave = [
                registro.dni,
                registro.fechaInicioTexto,
                registro.fechaFinTexto,
                registro.numeroDias,
                registro.tipoDocumento,
                registro.contingencia,
                registro.diagnostico,
                registro.establecimiento
            ].join("|");

            if (!mapa.has(clave)) {
                mapa.set(
                    clave,
                    registro
                );
            }
        }
    );

    return [...mapa.values()];
}

// =====================================================
// PAPELETAS
// =====================================================

async function cargarPapeletas() {
    const buffer =
        await descargarExcelStorage(
            CONFIG.rutaPapeletas
        );

    validarBufferExcel(
        buffer,
        "papeletas.xlsx"
    );

    const workbook = XLSX.read(
        buffer,
        {
            type: "array",
            cellDates: true
        }
    );

    const registros = [];

    workbook.SheetNames.forEach(
        nombreHoja => {
            const hoja =
                workbook.Sheets[nombreHoja];

            if (!hoja) {
                return;
            }

            const filas =
                XLSX.utils.sheet_to_json(
                    hoja,
                    {
                        defval: "",
                        raw: true
                    }
                );

            filas.forEach(
                (fila, indice) => {
                    const registro =
                        normalizarPapeleta(
                            fila,
                            indice + 2,
                            nombreHoja
                        );

                    if (registro.dni) {
                        registros.push(
                            registro
                        );
                    }
                }
            );
        }
    );

    estado.papeletas =
        eliminarDuplicadosPapeletas(
            registros
        );
}

function normalizarPapeleta(
    fila,
    filaExcel,
    hoja
) {
    const dni = limpiarDNI(
        obtenerCampoFlexible(
            fila,
            [
                "DNI",
                "N° DNI",
                "Nº DNI",
                "NUMERO DNI",
                "NUMDOCUMENTO",
                "NUM DOCUMENTO",
                "DOCUMENTO"
            ]
        )
    );

    const nombre =
        normalizarTextoVisible(
            obtenerCampoFlexible(
                fila,
                [
                    "APELLIDOS Y NOMBRES",
                    "NOMBRES Y APELLIDOS",
                    "APELLIDOS_NOMBRES",
                    "NOMBRES",
                    "NOMBRE",
                    "PERSONAL"
                ]
            )
        );

    const mesExcel =
        normalizarTextoVisible(
            obtenerCampoFlexible(
                fila,
                ["MES"]
            )
        );

    const unidadOrganica =
        normalizarTextoVisible(
            obtenerCampoFlexible(
                fila,
                [
                    "UNIDAD ORGANICA",
                    "UNIDAD ORGÁNICA",
                    "UNIDAD_ORGANICA",
                    "UNIDAD",
                    "SUB GERENCIA",
                    "SUBGERENCIA",
                    "AREA",
                    "ÁREA"
                ]
            )
        );

    const tipoPapeleta =
        normalizarTextoVisible(
            obtenerCampoFlexible(
                fila,
                [
                    "TIPO PAPELETA",
                    "TIPO DE PAPELETA",
                    "TIPO_PAPELETA",
                    "TIPO"
                ]
            )
        );

    const motivoPapeleta =
        normalizarTextoVisible(
            obtenerCampoFlexible(
                fila,
                [
                    "MOTIVO PAPELETA",
                    "MOTIVO DE PAPELETA",
                    "MOTIVO_PAPELETA",
                    "MOTIVO"
                ]
            )
        );

    const fechaInicio =
        convertirFecha(
            obtenerCampoFlexible(
                fila,
                [
                    "FECHA INICIO",
                    "FECHA DE INICIO",
                    "FECHA_INICIO",
                    "FEC_INICIO",
                    "INICIO"
                ]
            )
        );

    const fechaFin =
        convertirFecha(
            obtenerCampoFlexible(
                fila,
                [
                    "FECHA FIN",
                    "FECHA DE FIN",
                    "FECHA_FIN",
                    "FEC_FIN",
                    "FIN"
                ]
            )
        );

    const observaciones =
        normalizarTextoVisible(
            obtenerCampoFlexible(
                fila,
                [
                    "OBSERVACIONES",
                    "OBSERVACION",
                    "OBSERVACIÓN",
                    "ONSERVACION",
                    "COMENTARIOS",
                    "COMENTARIO"
                ]
            )
        );

    const mesNumero =
        fechaInicio
            ? fechaInicio.getMonth() + 1
            : obtenerNumeroMes(
                mesExcel
            );

    return {
        dni,

        nombre,

        mes:
            mesNumero
                ? MESES[mesNumero - 1]
                : (
                    mesExcel ||
                    "NO ESPECIFICADO"
                ),

        mesNumero,

        unidadOrganica:
            unidadOrganica ||
            "NO ESPECIFICADO",

        tipoPapeleta:
            tipoPapeleta ||
            "NO ESPECIFICADO",

        motivoPapeleta:
            motivoPapeleta ||
            "NO ESPECIFICADO",

        fechaInicio,

        fechaFin,

        fechaInicioTexto:
            formatearFecha(
                fechaInicio
            ),

        fechaFinTexto:
            formatearFecha(
                fechaFin
            ),

        observaciones:
            observaciones ||
            "NO ESPECIFICADO",

        filaExcel,

        hoja
    };
}

function eliminarDuplicadosPapeletas(
    registros
) {
    const mapa = new Map();

    registros.forEach(
        registro => {
            const clave = [
                registro.dni,
                registro.mes,
                registro.tipoPapeleta,
                registro.motivoPapeleta,
                registro.fechaInicioTexto,
                registro.fechaFinTexto,
                registro.observaciones,
                registro.unidadOrganica
            ].join("|");

            if (!mapa.has(clave)) {
                mapa.set(
                    clave,
                    registro
                );
            }
        }
    );

    return [...mapa.values()];
}

// =====================================================
// EVENTOS
// =====================================================

function inicializarEventos() {
    if (estado.eventosInicializados) {
        return;
    }

    estado.eventosInicializados = true;

    elementos.buscador?.addEventListener(
        "input",
        manejarBusqueda
    );

    elementos.buscador?.addEventListener(
        "focus",
        () => {
            const texto =
                elementos.buscador.value.trim();

            if (texto) {
                mostrarSugerencias(
                    buscarPersonal(
                        texto
                    )
                );
            }
        }
    );

    elementos.buscador?.addEventListener(
        "keydown",
        evento => {
            if (evento.key === "Escape") {
                cerrarSugerencias();
            }

            if (evento.key === "Enter") {
                const primera =
                    elementos.contenidoSugerencias
                        ?.querySelector(
                            ".sugerencia-personal"
                        );

                if (primera) {
                    evento.preventDefault();
                    primera.click();
                }
            }
        }
    );

    elementos.btnLimpiar?.addEventListener(
        "click",
        limpiarConsulta
    );

    elementos.btnReintentar?.addEventListener(
        "click",
        reintentarCarga
    );

    elementos.btnDashboardDescansos
        ?.addEventListener(
            "click",
            () =>
                alternarDashboard(
                    "descansos"
                )
        );

    elementos.btnCerrarDashboardDescansos
        ?.addEventListener(
            "click",
            () =>
                cerrarDashboard(
                    "descansos"
                )
        );

    elementos.btnDashboardPapeletas
        ?.addEventListener(
            "click",
            () =>
                alternarDashboard(
                    "papeletas"
                )
        );

    elementos.btnCerrarDashboardPapeletas
        ?.addEventListener(
            "click",
            () =>
                cerrarDashboard(
                    "papeletas"
                )
        );

    elementos.btnCerrarModalDetalle
        ?.addEventListener(
            "click",
            cerrarModalDetalle
        );

    elementos.modalDetalle?.addEventListener(
        "click",
        evento => {
            if (
                evento.target ===
                elementos.modalDetalle
            ) {
                cerrarModalDetalle();
            }
        }
    );

    document.addEventListener(
        "click",
        evento => {
            if (
                !evento.target.closest(
                    ".buscador-personal__campo-wrap"
                )
            ) {
                cerrarSugerencias();
            }
        }
    );

    window.addEventListener(
        "keydown",
        evento => {
            if (evento.key === "Escape") {
                cerrarSugerencias();
                cerrarModalDetalle();
            }
        }
    );

    window.addEventListener(
        "resize",
        debounce(
            () => {
                Object.values(
                    estado.charts
                ).forEach(
                    chart =>
                        chart?.resize()
                );
            },
            150
        )
    );
}

async function reintentarCarga() {
    mostrarLoading(true);
    ocultarError();

    try {
        await cargarTodasLasFuentes();
        limpiarConsulta();

    } catch (error) {
        console.error(
            "Error al reintentar:",
            error
        );

        mostrarError(
            "No se pudo recargar la información."
        );

    } finally {
        mostrarLoading(false);
    }
}

// =====================================================
// BÚSQUEDA
// =====================================================

function manejarBusqueda() {
    const texto =
        elementos.buscador?.value || "";

    if (!texto.trim()) {
        cerrarSugerencias();
        return;
    }

    mostrarSugerencias(
        buscarPersonal(
            texto
        )
    );
}

function buscarPersonal(texto) {
    const consulta =
        normalizarBusqueda(
            texto
        );

    if (!consulta) {
        return [];
    }

    const palabras =
        consulta
            .split(" ")
            .filter(Boolean);

    return estado.personal
        .filter(
            persona =>
                palabras.every(
                    palabra =>
                        persona.busqueda.includes(
                            palabra
                        )
                )
        )
        .sort(
            (a, b) => {
                const dniA =
                    a.dni.startsWith(
                        consulta
                    );

                const dniB =
                    b.dni.startsWith(
                        consulta
                    );

                if (dniA !== dniB) {
                    return dniA ? -1 : 1;
                }

                const nombreA =
                    normalizarBusqueda(
                        a.apellidosNombres
                    ).startsWith(
                        consulta
                    );

                const nombreB =
                    normalizarBusqueda(
                        b.apellidosNombres
                    ).startsWith(
                        consulta
                    );

                if (nombreA !== nombreB) {
                    return nombreA
                        ? -1
                        : 1;
                }

                return a.apellidosNombres
                    .localeCompare(
                        b.apellidosNombres,
                        "es"
                    );
            }
        )
        .slice(
            0,
            CONFIG.limiteSugerencias
        );
}

function mostrarSugerencias(resultados) {
    if (
        !elementos.listaSugerencias ||
        !elementos.contenidoSugerencias
    ) {
        return;
    }

    elementos.listaSugerencias.hidden =
        false;

    if (!resultados.length) {
        elementos.contenidoSugerencias.innerHTML =
            "";

        if (elementos.mensajeSinResultados) {
            elementos.mensajeSinResultados.hidden =
                false;
        }

        return;
    }

    if (elementos.mensajeSinResultados) {
        elementos.mensajeSinResultados.hidden =
            true;
    }

    elementos.contenidoSugerencias.innerHTML =
        resultados.map(
            persona => `
                <button
                    type="button"
                    class="sugerencia-personal"
                    data-dni="${escapeHtml(persona.dni)}"
                >
                    <div class="sugerencia-personal__avatar">
                        <img
                            src="${escapeHtml(persona.imagen)}"
                            alt="Imagen de ${escapeHtml(persona.apellidosNombres)}"
                        >
                    </div>

                    <div class="sugerencia-personal__datos">
                        <strong>
                            ${escapeHtml(persona.apellidosNombres || "SIN NOMBRE")}
                        </strong>

                        <span>
                            DNI: ${escapeHtml(persona.dni)}
                            ·
                            ${escapeHtml(persona.funcion || "SIN FUNCIÓN")}
                        </span>
                    </div>

                    <i
                        class="fa-solid fa-chevron-right"
                        aria-hidden="true"
                    ></i>
                </button>
            `
        ).join("");

    elementos.contenidoSugerencias
        .querySelectorAll(
            ".sugerencia-personal"
        )
        .forEach(
            boton => {
                boton.addEventListener(
                    "click",
                    () =>
                        seleccionarPersona(
                            boton.dataset.dni
                        )
                );
            }
        );
}

function cerrarSugerencias() {
    if (elementos.listaSugerencias) {
        elementos.listaSugerencias.hidden =
            true;
    }
}

// =====================================================
// SELECCIÓN DE PERSONA
// =====================================================

function seleccionarPersona(dni) {
    const persona =
        estado.personal.find(
            item =>
                compararDNI(
                    item.dni,
                    dni
                )
        );

    if (!persona) {
        mostrarError(
            "No se encontró el personal seleccionado."
        );

        return;
    }

    estado.personaSeleccionada =
        persona;

    estado.descansosSeleccionados =
        estado.descansosMedicos
            .filter(
                registro =>
                    compararDNI(
                        registro.dni,
                        persona.dni
                    )
            )
            .sort(
                ordenarDescansosRecientes
            );

    estado.papeletasSeleccionadas =
        estado.papeletas
            .filter(
                registro =>
                    compararDNI(
                        registro.dni,
                        persona.dni
                    )
            )
            .sort(
                ordenarPapeletasRecientes
            );

    if (elementos.buscador) {
        elementos.buscador.value =
            persona.apellidosNombres;
    }

    cerrarSugerencias();
    ocultarError();
    cerrarTodosLosDashboards();

    renderDashboardPersonal();
}

function renderDashboardPersonal() {
    const persona =
        estado.personaSeleccionada;

    if (!persona) {
        return;
    }

    if (elementos.estadoInicial) {
        elementos.estadoInicial.hidden =
            true;
    }

    if (elementos.fichaPersonal) {
        elementos.fichaPersonal.hidden =
            false;
    }

    renderDatosPersonales(
        persona
    );

    renderCapacitacion(
        persona
    );

    renderIndicadores();

    prepararDashboardDescansos();

    prepararDashboardPapeletas();
}

// =====================================================
// DATOS DEL PERFIL
// =====================================================

function renderDatosPersonales(persona) {
    setTexto(
        "nombrePersonal",
        persona.apellidosNombres ||
        "SIN NOMBRE"
    );

    setTexto(
        "dniPersonal",
        persona.dni || "-"
    );

    setTexto(
        "apellidosNombresPersonal",
        persona.apellidosNombres || "-"
    );

    setTexto(
        "areaPersonal",
        persona.area ||
        "NO ESPECIFICADO"
    );

    setTexto(
        "condicionLaboralPersonal",
        persona.condicionLaboral ||
        "NO ESPECIFICADO"
    );

    setTexto(
        "funcionPersonal",
        persona.funcion ||
        "NO ESPECIFICADO"
    );

    setTexto(
        "sexoPersonal",
        persona.sexo ||
        "NO ESPECIFICADO"
    );

    const imagen =
        document.getElementById(
            "fotoPersonal"
        );

    if (imagen) {
        imagen.onerror = () => {
            imagen.onerror = null;

            imagen.src =
                CONFIG.imagenHombre;
        };

        imagen.src =
            obtenerImagenSegunSexo(
                persona.sexo
            );

        imagen.alt =
            `Imagen de ${persona.apellidosNombres}`;
    }

    const badge =
        document.getElementById(
            "estadoPersonalBadge"
        );

    if (badge) {
        const textoEstado =
            persona.estado ||
            "ACTIVO";

        badge.innerHTML = `
            <span></span>
            ${escapeHtml(textoEstado)}
        `;

        badge.classList.toggle(
            "perfil-card__estado--inactivo",
            textoEstado !== "ACTIVO"
        );
    }
}

// =====================================================
// CAPACITACIÓN
// =====================================================

function renderCapacitacion(persona) {
    const nota =
        persona.promedioNota;

    const promedio =
        document.getElementById(
            "promedioNotaPersonal"
        );

    const escala =
        document.getElementById(
            "escalaPromedioPersonal"
        );

    const badge =
        document.getElementById(
            "estadoCapacitacionPersonal"
        );

    const mensaje =
        document.getElementById(
            "mensajeCapacitacionPersonal"
        );

    if (
        nota === null ||
        Number.isNaN(nota)
    ) {
        if (promedio) {
            promedio.textContent = "—";
        }

        if (escala) {
            escala.textContent = "";
        }

        if (badge) {
            badge.textContent =
                "SIN CAPACITACIÓN";

            badge.className =
                "capacitacion-card__badge " +
                "capacitacion-card__badge--sin-nota";
        }

        if (mensaje) {
            mensaje.textContent =
                "Aún no recibe capacitación.";
        }

        return;
    }

    if (promedio) {
        promedio.textContent =
            formatearNota(
                nota
            );
    }

    if (escala) {
        escala.textContent =
            "/100";
    }

    const clasificacion =
        clasificarNota(
            nota
        );

    if (badge) {
        badge.textContent =
            clasificacion.texto;

        badge.className =
            "capacitacion-card__badge " +
            clasificacion.clase;
    }

    if (mensaje) {
        mensaje.textContent =
            `Promedio registrado: ${formatearNota(nota)} de 100.`;
    }
}

function clasificarNota(nota) {
    if (nota >= 90) {
        return {
            texto: "EXCELENTE",
            clase:
                "capacitacion-card__badge--excelente"
        };
    }

    if (nota >= 75) {
        return {
            texto: "BUENO",
            clase:
                "capacitacion-card__badge--bueno"
        };
    }

    if (nota >= 60) {
        return {
            texto: "REGULAR",
            clase:
                "capacitacion-card__badge--regular"
        };
    }

    return {
        texto: "REQUIERE MEJORA",
        clase:
            "capacitacion-card__badge--bajo"
    };
}

// =====================================================
// KPIS
// =====================================================

function renderIndicadores() {
    const totalDescansos =
        estado.descansosSeleccionados.length;

    const diasAcumulados =
        estado.descansosSeleccionados.reduce(
            (total, registro) =>
                total +
                Number(
                    registro.numeroDias || 0
                ),
            0
        );

    const totalPapeletas =
        estado.papeletasSeleccionadas.length;

    setTexto(
        "totalDescansosMedicos",
        totalDescansos.toLocaleString(
            "es-PE"
        )
    );

    setTexto(
        "totalDiasDescanso",
        diasAcumulados.toLocaleString(
            "es-PE"
        )
    );

    setTexto(
        "totalPapeletas",
        totalPapeletas.toLocaleString(
            "es-PE"
        )
    );
}

// =====================================================
// DASHBOARDS DESPLEGABLES
// =====================================================

function alternarDashboard(tipo) {
    if (!estado.personaSeleccionada) {
        return;
    }

    const estaAbierto =
        estado.dashboardActivo === tipo;

    if (estaAbierto) {
        cerrarDashboard(tipo);
        return;
    }

    cerrarTodosLosDashboards();

    if (tipo === "descansos") {
        abrirDashboardDescansos();
    }

    if (tipo === "papeletas") {
        abrirDashboardPapeletas();
    }
}

function abrirDashboardDescansos() {
    if (!elementos.dashboardDescansos) {
        return;
    }

    estado.dashboardActivo =
        "descansos";

    elementos.dashboardDescansos.hidden =
        false;

    actualizarEstadoBotonDashboard(
        elementos.btnDashboardDescansos,
        true
    );

    requestAnimationFrame(
        () => {
            renderDashboardDescansos();

            elementos.dashboardDescansos.scrollIntoView({
                behavior: "smooth",
                block: "start"
            });
        }
    );
}

function abrirDashboardPapeletas() {
    if (!elementos.dashboardPapeletas) {
        return;
    }

    estado.dashboardActivo =
        "papeletas";

    elementos.dashboardPapeletas.hidden =
        false;

    actualizarEstadoBotonDashboard(
        elementos.btnDashboardPapeletas,
        true
    );

    requestAnimationFrame(
        () => {
            renderDashboardPapeletas();

            elementos.dashboardPapeletas.scrollIntoView({
                behavior: "smooth",
                block: "start"
            });
        }
    );
}

function cerrarDashboard(tipo) {
    if (tipo === "descansos") {
        if (elementos.dashboardDescansos) {
            elementos.dashboardDescansos.hidden =
                true;
        }

        actualizarEstadoBotonDashboard(
            elementos.btnDashboardDescansos,
            false
        );

        destruirChart(
            "descansosMes"
        );

        destruirChart(
            "descansosTipo"
        );

        destruirChart(
            "descansosDiagnostico"
        );
    }

    if (tipo === "papeletas") {
        if (elementos.dashboardPapeletas) {
            elementos.dashboardPapeletas.hidden =
                true;
        }

        actualizarEstadoBotonDashboard(
            elementos.btnDashboardPapeletas,
            false
        );

        destruirChart(
            "papeletasMes"
        );

        destruirChart(
            "papeletasTipo"
        );

        destruirChart(
            "papeletasMotivo"
        );
    }

    if (estado.dashboardActivo === tipo) {
        estado.dashboardActivo =
            null;
    }
}

function cerrarTodosLosDashboards() {
    cerrarDashboard(
        "descansos"
    );

    cerrarDashboard(
        "papeletas"
    );

    estado.dashboardActivo =
        null;
}

function actualizarEstadoBotonDashboard(
    boton,
    abierto
) {
    if (!boton) {
        return;
    }

    boton.setAttribute(
        "aria-expanded",
        String(abierto)
    );

    boton.classList.toggle(
        "is-active",
        abierto
    );

    const flecha =
        boton.querySelector(
            ".indicador-card__flecha"
        );

    if (flecha) {
        flecha.style.transform =
            abierto
                ? "rotate(180deg)"
                : "rotate(0deg)";
    }
}

// =====================================================
// PREPARAR DASHBOARDS
// =====================================================

function prepararDashboardDescansos() {
    const registros =
        estado.descansosSeleccionados;

    setTexto(
        "contadorDescansosDashboard",
        `${registros.length.toLocaleString("es-PE")} registros`
    );

    if (elementos.mensajeSinDescansos) {
        elementos.mensajeSinDescansos.hidden =
            registros.length > 0;
    }

    if (elementos.contenedorGraficosDescansos) {
        elementos.contenedorGraficosDescansos.hidden =
            registros.length === 0;
    }
}

function prepararDashboardPapeletas() {
    const registros =
        estado.papeletasSeleccionadas;

    setTexto(
        "contadorPapeletasDashboard",
        `${registros.length.toLocaleString("es-PE")} registros`
    );

    if (elementos.mensajeSinPapeletas) {
        elementos.mensajeSinPapeletas.hidden =
            registros.length > 0;
    }

    if (elementos.contenedorGraficosPapeletas) {
        elementos.contenedorGraficosPapeletas.hidden =
            registros.length === 0;
    }
}

// =====================================================
// DASHBOARD DESCANSOS
// =====================================================

function renderDashboardDescansos() {
    const registros =
        estado.descansosSeleccionados;

    prepararDashboardDescansos();

    destruirChart(
        "descansosMes"
    );

    destruirChart(
        "descansosTipo"
    );

    destruirChart(
        "descansosDiagnostico"
    );

    if (!registros.length) {
        return;
    }

    renderGraficoDescansosPorMes(
        registros
    );

    renderGraficoDescansosTipoDocumento(
        registros
    );

    renderGraficoDescansosDiagnostico(
        registros
    );
}

function renderGraficoDescansosPorMes(registros) {
    const canvas =
        document.getElementById(
            "graficoDescansosPorMes"
        );

    if (!canvas) {
        return;
    }

    const valores =
        MESES.map(
            (_, index) =>
                registros.filter(
                    registro =>
                        registro.mesNumero ===
                        index + 1
                ).length
        );

    estado.charts.descansosMes =
        new Chart(
            canvas,
            {
                type: "line",

                data: {
                    labels: MESES,

                    datasets: [{
                        label:
                            "Descansos médicos por mes",

                        data:
                            valores,

                        borderColor:
                            COLORES.primary,

                        backgroundColor:
                            "rgba(85,107,47,0.12)",

                        pointBackgroundColor:
                            COLORES.primary,

                        pointBorderColor:
                            COLORES.blanco,

                        pointBorderWidth:
                            2,

                        pointRadius:
                            4,

                        pointHoverRadius:
                            6,

                        borderWidth:
                            2.5,

                        fill:
                            true,

                        tension:
                            0.35
                    }]
                },

                options: {
                    ...opcionesVerticales(),

                    interaction: {
                        mode: "nearest",
                        intersect: false
                    },

                    onHover(
                        evento
                    ) {
                        const items =
                            this.getElementsAtEventForMode(
                                evento,
                                "nearest",
                                {
                                    intersect: false
                                },
                                true
                            );

                        cambiarCursorGrafico(
                            evento,
                            items.length > 0
                        );
                    },

                    onClick(
                        evento
                    ) {
                        const items =
                            this.getElementsAtEventForMode(
                                evento,
                                "nearest",
                                {
                                    intersect: false
                                },
                                true
                            );

                        if (!items.length) {
                            return;
                        }

                        const index =
                            items[0].index;

                        const mesNumero =
                            index + 1;

                        abrirModalDescansos({
                            titulo:
                                "DESCANSOS MÉDICOS POR MES",

                            subtitulo:
                                MESES[index],

                            registros:
                                registros.filter(
                                    registro =>
                                        registro.mesNumero ===
                                        mesNumero
                                )
                        });
                    }
                },

                plugins:
                    obtenerPluginsChart()
            }
        );
}

function renderGraficoDescansosTipoDocumento(
    registros
) {
    const canvas =
        document.getElementById(
            "graficoDescansosTipoDocumento"
        );

    if (!canvas) {
        return;
    }

    const conteo =
        contarPorCampo(
            registros,
            "tipoDocumento"
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

    estado.charts.descansosTipo =
        new Chart(
            canvas,
            {
                type: "doughnut",

                data: {
                    labels,

                    datasets: [{
                        label:
                            "Descansos médicos",

                        data:
                            valores,

                        backgroundColor:
                            generarColoresPorValor(
                                valores
                            ),

                        borderColor:
                            COLORES.blanco,

                        borderWidth:
                            3,

                        hoverOffset:
                            8
                    }]
                },

                options:
                    opcionesTorta({
                        etiqueta:
                            "Descansos médicos",

                        onClickItem:
                            index => {
                                const valor =
                                    labels[index];

                                abrirModalDescansos({
                                    titulo:
                                        "DESCANSOS POR TIPO DE DOCUMENTO",

                                    subtitulo:
                                        valor,

                                    registros:
                                        registros.filter(
                                            registro =>
                                                registro.tipoDocumento ===
                                                valor
                                        )
                                });
                            }
                    }),

                plugins:
                    obtenerPluginsChart()
            }
        );
}

function renderGraficoDescansosDiagnostico(
    registros
) {
    const canvas =
        document.getElementById(
            "graficoDescansosDiagnostico"
        );

    if (!canvas) {
        return;
    }

    const conteo =
        contarPorCampo(
            registros,
            "diagnostico"
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

    estado.charts.descansosDiagnostico =
        new Chart(
            canvas,
            {
                type: "bar",

                data: {
                    labels,

                    datasets: [{
                        label:
                            "Descansos médicos",

                        data:
                            valores,

                        backgroundColor:
                            generarColoresPorValor(
                                valores
                            ),

                        borderRadius:
                            7,

                        borderSkipped:
                            false,

                        minBarLength:
                            10,

                        maxBarThickness:
                            24
                    }]
                },

                options: {
                    ...opcionesHorizontales(),

                    interaction: {
                        mode: "index",
                        intersect: false,
                        axis: "y"
                    },

                    onHover(
                        evento
                    ) {
                        const items =
                            this.getElementsAtEventForMode(
                                evento,
                                "index",
                                {
                                    intersect: false,
                                    axis: "y"
                                },
                                true
                            );

                        cambiarCursorGrafico(
                            evento,
                            items.length > 0
                        );
                    },

                    onClick(
                        evento
                    ) {
                        const items =
                            this.getElementsAtEventForMode(
                                evento,
                                "index",
                                {
                                    intersect: false,
                                    axis: "y"
                                },
                                true
                            );

                        if (!items.length) {
                            return;
                        }

                        const diagnostico =
                            labels[
                                items[0].index
                            ];

                        abrirModalDescansos({
                            titulo:
                                "DESCANSOS POR DIAGNÓSTICO",

                            subtitulo:
                                diagnostico,

                            registros:
                                registros.filter(
                                    registro =>
                                        registro.diagnostico ===
                                        diagnostico
                                )
                        });
                    }
                },

                plugins:
                    obtenerPluginsChart()
            }
        );
}

// =====================================================
// DASHBOARD PAPELETAS
// =====================================================

function renderDashboardPapeletas() {
    const registros =
        estado.papeletasSeleccionadas;

    prepararDashboardPapeletas();

    destruirChart(
        "papeletasMes"
    );

    destruirChart(
        "papeletasTipo"
    );

    destruirChart(
        "papeletasMotivo"
    );

    if (!registros.length) {
        return;
    }

    renderGraficoPapeletasPorMes(
        registros
    );

    renderGraficoPapeletasTipo(
        registros
    );

    renderGraficoPapeletasMotivo(
        registros
    );
}

function renderGraficoPapeletasPorMes(registros) {
    const canvas =
        document.getElementById(
            "graficoPapeletasPorMes"
        );

    if (!canvas) {
        return;
    }

    const valores =
        MESES.map(
            (_, index) =>
                registros.filter(
                    registro =>
                        registro.mesNumero ===
                        index + 1
                ).length
        );

    estado.charts.papeletasMes =
        new Chart(
            canvas,
            {
                type: "line",

                data: {
                    labels: MESES,

                    datasets: [{
                        label:
                            "Papeletas por mes",

                        data:
                            valores,

                        borderColor:
                            COLORES.primary,

                        backgroundColor:
                            "rgba(85,107,47,0.12)",

                        pointBackgroundColor:
                            COLORES.primary,

                        pointBorderColor:
                            COLORES.blanco,

                        pointBorderWidth:
                            2,

                        pointRadius:
                            4,

                        pointHoverRadius:
                            6,

                        borderWidth:
                            2.5,

                        fill:
                            true,

                        tension:
                            0.35
                    }]
                },

                options: {
                    ...opcionesVerticales(),

                    interaction: {
                        mode: "nearest",
                        intersect: false
                    },

                    onHover(
                        evento
                    ) {
                        const items =
                            this.getElementsAtEventForMode(
                                evento,
                                "nearest",
                                {
                                    intersect: false
                                },
                                true
                            );

                        cambiarCursorGrafico(
                            evento,
                            items.length > 0
                        );
                    },

                    onClick(
                        evento
                    ) {
                        const items =
                            this.getElementsAtEventForMode(
                                evento,
                                "nearest",
                                {
                                    intersect: false
                                },
                                true
                            );

                        if (!items.length) {
                            return;
                        }

                        const index =
                            items[0].index;

                        const mesNumero =
                            index + 1;

                        abrirModalPapeletas({
                            titulo:
                                "PAPELETAS POR MES",

                            subtitulo:
                                MESES[index],

                            registros:
                                registros.filter(
                                    registro =>
                                        registro.mesNumero ===
                                        mesNumero
                                )
                        });
                    }
                },

                plugins:
                    obtenerPluginsChart()
            }
        );
}

function renderGraficoPapeletasTipo(registros) {
    const canvas =
        document.getElementById(
            "graficoPapeletasTipo"
        );

    if (!canvas) {
        return;
    }

    const conteo =
        contarPorCampo(
            registros,
            "tipoPapeleta"
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

    estado.charts.papeletasTipo =
        new Chart(
            canvas,
            {
                type: "doughnut",

                data: {
                    labels,

                    datasets: [{
                        label:
                            "Papeletas",

                        data:
                            valores,

                        backgroundColor:
                            generarColoresPorValor(
                                valores
                            ),

                        borderColor:
                            COLORES.blanco,

                        borderWidth:
                            3,

                        hoverOffset:
                            8
                    }]
                },

                options:
                    opcionesTorta({
                        etiqueta:
                            "Papeletas",

                        onClickItem:
                            index => {
                                const tipo =
                                    labels[index];

                                abrirModalPapeletas({
                                    titulo:
                                        "PAPELETAS POR TIPO",

                                    subtitulo:
                                        tipo,

                                    registros:
                                        registros.filter(
                                            registro =>
                                                registro.tipoPapeleta ===
                                                tipo
                                        )
                                });
                            }
                    }),

                plugins:
                    obtenerPluginsChart()
            }
        );
}

function renderGraficoPapeletasMotivo(registros) {
    const canvas =
        document.getElementById(
            "graficoPapeletasMotivo"
        );

    if (!canvas) {
        return;
    }

    const conteo =
        contarPorCampo(
            registros,
            "motivoPapeleta"
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

    estado.charts.papeletasMotivo =
        new Chart(
            canvas,
            {
                type: "bar",

                data: {
                    labels,

                    datasets: [{
                        label:
                            "Papeletas",

                        data:
                            valores,

                        backgroundColor:
                            generarColoresPorValor(
                                valores
                            ),

                        borderRadius:
                            7,

                        borderSkipped:
                            false,

                        minBarLength:
                            10,

                        maxBarThickness:
                            24
                    }]
                },

                options: {
                    ...opcionesHorizontales(),

                    interaction: {
                        mode: "index",
                        intersect: false,
                        axis: "y"
                    },

                    onHover(
                        evento
                    ) {
                        const items =
                            this.getElementsAtEventForMode(
                                evento,
                                "index",
                                {
                                    intersect: false,
                                    axis: "y"
                                },
                                true
                            );

                        cambiarCursorGrafico(
                            evento,
                            items.length > 0
                        );
                    },

                    onClick(
                        evento
                    ) {
                        const items =
                            this.getElementsAtEventForMode(
                                evento,
                                "index",
                                {
                                    intersect: false,
                                    axis: "y"
                                },
                                true
                            );

                        if (!items.length) {
                            return;
                        }

                        const motivo =
                            labels[
                                items[0].index
                            ];

                        abrirModalPapeletas({
                            titulo:
                                "PAPELETAS POR MOTIVO",

                            subtitulo:
                                motivo,

                            registros:
                                registros.filter(
                                    registro =>
                                        registro.motivoPapeleta ===
                                        motivo
                                )
                        });
                    }
                },

                plugins:
                    obtenerPluginsChart()
            }
        );
}

// =====================================================
// OPCIONES CHART.JS
// =====================================================

function opcionesVerticales() {
    return {
        responsive: true,

        maintainAspectRatio: false,

        animation: false,

        normalized: true,

        layout: {
            padding: {
                top: 28,
                right: 15,
                bottom: 4,
                left: 5
            }
        },

        plugins: {
            legend: {
                display: false
            },

            tooltip: {
                callbacks: {
                    label:
                        contexto =>
                            ` ${Number(contexto.raw || 0).toLocaleString("es-PE")} registro(s)`
                }
            },

            datalabels: {
                color:
                    COLORES.texto,

                anchor:
                    "end",

                align:
                    "top",

                offset:
                    3,

                clamp:
                    true,
                clip: false,

                font: {
                    weight:
                        "bold",

                    size:
                        11
                },

                formatter:
                    valor =>
                        valor > 0
                            ? valor
                            : ""
            }
        },

        scales: {
            x: {
                ticks: {
                    color:
                        COLORES.textoSuave,

                    maxRotation:
                        45,

                    minRotation:
                        0
                },

                grid: {
                    color:
                        "rgba(0,0,0,0.05)"
                }
            },

            y: {
                beginAtZero: true,

                grace: "15%",

                ticks: {
                    color: COLORES.textoSuave,
                    precision: 0,
                    stepSize: 1
                },

                grid: {
                    color: COLORES.grid
                }
            }
        }
    };
}

function opcionesHorizontales() {
    return {
        responsive: true,

        maintainAspectRatio: false,

        animation: false,

        normalized: true,

        indexAxis:
            "y",

        layout: {
            padding: {
                right: 42,
                left: 5,
                top: 4,
                bottom: 4
            }
        },

        plugins: {
            legend: {
                display: false
            },

            tooltip: {
                callbacks: {
                    title:
                        elementos =>
                            elementos[0]?.label ||
                            "",

                    label:
                        contexto =>
                            ` ${Number(contexto.raw || 0).toLocaleString("es-PE")} registro(s)`
                }
            },

            datalabels: {
                color:
                    COLORES.texto,

                anchor:
                    "end",

                align:
                    "right",

                offset:
                    6,

                clamp:
                    true,

                clip:
                    false,

                font: {
                    weight:
                        "bold",

                    size:
                        11
                },

                formatter:
                    valor =>
                        valor > 0
                            ? valor
                            : ""
            }
        },

        scales: {
            x: {
                beginAtZero:
                    true,

                grace:
                    "8%",

                ticks: {
                    color:
                        COLORES.textoSuave,

                    precision:
                        0
                },

                grid: {
                    color:
                        COLORES.grid
                }
            },

            y: {
                ticks: {
                    color:
                        COLORES.texto,

                    autoSkip:
                        false,

                    font: {
                        size: 11
                    },

                    callback(valor) {
                        const etiqueta =
                            this.getLabelForValue(
                                valor
                            );

                        return truncarTexto(
                            etiqueta,
                            48
                        );
                    }
                },

                grid: {
                    display:
                        false
                }
            }
        }
    };
}

function opcionesTorta({
    etiqueta,
    onClickItem
}) {
    return {
        responsive: true,

        maintainAspectRatio: false,

        cutout: "58%",

        animation: false,

        layout: {
            padding: {
                top: 5,
                right: 10,
                bottom: 5,
                left: 10
            }
        },

        onHover(evento) {
            const items =
                this.getElementsAtEventForMode(
                    evento,
                    "nearest",
                    {
                        intersect: true
                    },
                    true
                );

            cambiarCursorGrafico(
                evento,
                items.length > 0
            );
        },

        onClick(evento) {
            const items =
                this.getElementsAtEventForMode(
                    evento,
                    "nearest",
                    {
                        intersect: true
                    },
                    true
                );

            if (!items.length) {
                return;
            }

            onClickItem?.(
                items[0].index
            );
        },

        plugins: {
            legend: {
                display: true,

                position: "bottom",

                align: "center",

                onClick(
                    evento,
                    legendItem,
                    legend
                ) {
                    const chart =
                        legend.chart;

                    const index =
                        legendItem.index;

                    if (
                        index === undefined ||
                        index === null
                    ) {
                        return;
                    }

                    chart.toggleDataVisibility(
                        index
                    );

                    chart.update();
                },

                labels: {
                    color:
                        COLORES.textoSuave,

                    usePointStyle: true,

                    pointStyle:
                        "rectRounded",

                    boxWidth:
                        12,

                    boxHeight:
                        12,

                    padding:
                        12,

                    font: {
                        size:
                            10,

                        weight:
                            "700"
                    },

                    generateLabels(chart) {
                        const labels =
                            chart.data.labels || [];

                        const dataset =
                            chart.data.datasets?.[0];

                        const valores =
                            dataset?.data || [];

                        const colores =
                            dataset?.backgroundColor || [];

                        const total =
                            valores.reduce(
                                (
                                    suma,
                                    valor
                                ) =>
                                    suma +
                                    Number(
                                        valor || 0
                                    ),
                                0
                            );

                        return labels.map(
                            (
                                label,
                                index
                            ) => {
                                const valor =
                                    Number(
                                        valores[index] || 0
                                    );

                                const porcentaje =
                                    total > 0
                                        ? (
                                            (
                                                valor /
                                                total
                                            ) *
                                            100
                                        ).toFixed(1)
                                        : "0.0";

                                const color =
                                    Array.isArray(
                                        colores
                                    )
                                        ? colores[index]
                                        : colores;

                                return {
                                    text:
                                        `${truncarTexto(
                                            label,
                                            26
                                        )} · ` +
                                        `${valor.toLocaleString(
                                            "es-PE"
                                        )} ` +
                                        `(${porcentaje}%)`,

                                    fillStyle:
                                        color,

                                    strokeStyle:
                                        color,

                                    lineWidth:
                                        0,

                                    hidden:
                                        !chart.getDataVisibility(
                                            index
                                        ),

                                    index,

                                    datasetIndex:
                                        0,

                                    pointStyle:
                                        "rectRounded"
                                };
                            }
                        );
                    }
                }
            },

            tooltip: {
                displayColors:
                    true,

                callbacks: {
                    title(contextos) {
                        return (
                            contextos[0]?.label ||
                            etiqueta ||
                            ""
                        );
                    },

                    label(contexto) {
                        const valores =
                            contexto.dataset.data || [];

                        const total =
                            valores.reduce(
                                (
                                    suma,
                                    valor
                                ) =>
                                    suma +
                                    Number(
                                        valor || 0
                                    ),
                                0
                            );

                        const valor =
                            Number(
                                contexto.raw || 0
                            );

                        const porcentaje =
                            total > 0
                                ? (
                                    (
                                        valor /
                                        total
                                    ) *
                                    100
                                ).toFixed(1)
                                : "0.0";

                        return (
                            ` ${valor.toLocaleString(
                                "es-PE"
                            )} registro(s) ` +
                            `(${porcentaje}%)`
                        );
                    }
                }
            },

            datalabels: {
                display(contexto) {
                    const datos =
                        contexto.dataset.data || [];

                    const total =
                        datos.reduce(
                            (
                                suma,
                                valor
                            ) =>
                                suma +
                                Number(
                                    valor || 0
                                ),
                            0
                        );

                    const valor =
                        Number(
                            datos[
                                contexto.dataIndex
                            ] || 0
                        );

                    const porcentaje =
                        total > 0
                            ? (
                                valor /
                                total
                            ) *
                            100
                            : 0;

                    return porcentaje >= 6;
                },

                color:
                    COLORES.blanco,

                anchor:
                    "center",

                align:
                    "center",

                clamp:
                    true,

                clip:
                    true,

                font: {
                    size:
                        11,

                    weight:
                        "bold"
                },

                formatter(
                    valor,
                    contexto
                ) {
                    const datos =
                        contexto.dataset.data || [];

                    const total =
                        datos.reduce(
                            (
                                suma,
                                numero
                            ) =>
                                suma +
                                Number(
                                    numero || 0
                                ),
                            0
                        );

                    if (!total) {
                        return "";
                    }

                    const porcentaje =
                        (
                            (
                                Number(
                                    valor || 0
                                ) /
                                total
                            ) *
                            100
                        ).toFixed(1);

                    return `${porcentaje}%`;
                }
            }
        }
    };
}
// =====================================================
// COLORES DINÁMICOS
// =====================================================

function generarColoresPorValor(valores) {
    if (
        !Array.isArray(valores) ||
        !valores.length
    ) {
        return [];
    }

    const maximo =
        Math.max(
            ...valores
        );

    return valores.map(
        valor => {
            const proporcion =
                maximo > 0
                    ? valor / maximo
                    : 0;

            if (proporcion >= 0.90) {
                return COLORES.rojoSuave;
            }

            if (proporcion >= 0.80) {
                return COLORES.naranjaSuave;
            }

            if (proporcion >= 0.70) {
                return COLORES.amarilloSuave;
            }

            if (proporcion >= 0.60) {
                return COLORES.limaSuave;
            }

            if (proporcion >= 0.50) {
                return COLORES.verdeLima;
            }

            if (proporcion >= 0.40) {
                return COLORES.verdeClaro;
            }

            if (proporcion >= 0.30) {
                return COLORES.verdeAgua;
            }

            if (proporcion >= 0.20) {
                return COLORES.turquesa;
            }

            return COLORES.azulClaro;
        }
    );
}

function cambiarCursorGrafico(
    evento,
    activo
) {
    const target =
        evento?.native?.target;

    if (target) {
        target.style.cursor =
            activo
                ? "pointer"
                : "default";
    }
}

// =====================================================
// MODAL DESCANSOS
// =====================================================

function abrirModalDescansos({
    titulo,
    subtitulo,
    registros
}) {
    const ordenados =
        [...registros].sort(
            ordenarDescansosRecientes
        );

    setTexto(
        "modalDetalleEyebrow",
        "Detalle de descansos médicos"
    );

    setTexto(
        "modalDetalleTitulo",
        titulo
    );

    setTexto(
        "modalDetalleSubtitulo",
        `${subtitulo} | ${ordenados.length} registros`
    );

    setTexto(
        "modalTotalRegistros",
        ordenados.length.toLocaleString(
            "es-PE"
        )
    );

    const totalDias =
        ordenados.reduce(
            (total, registro) =>
                total +
                Number(
                    registro.numeroDias || 0
                ),
            0
        );

    setTexto(
        "modalResumenSecundarioTitulo",
        "Días acumulados"
    );

    setTexto(
        "modalResumenSecundarioValor",
        totalDias.toLocaleString(
            "es-PE"
        )
    );

    setTexto(
        "modalContadorDetalle",
        `${ordenados.length.toLocaleString("es-PE")} registros`
    );

    if (elementos.cabeceraModalDetalle) {
        elementos.cabeceraModalDetalle.innerHTML = `
            <th>N.°</th>
            <th>Inicio</th>
            <th>Fin</th>
            <th>Días</th>
            <th>Tipo de documento</th>
            <th>Contingencia</th>
            <th>Diagnóstico</th>
            <th>Establecimiento</th>
        `;
    }

    if (elementos.cuerpoModalDetalle) {
        if (!ordenados.length) {
            elementos.cuerpoModalDetalle.innerHTML = `
                <tr>
                    <td
                        colspan="8"
                        class="tabla-consulta__vacio"
                    >
                        No se encontraron registros.
                    </td>
                </tr>
            `;

        } else {
            elementos.cuerpoModalDetalle.innerHTML =
                ordenados.map(
                    (registro, index) => `
                        <tr>
                            <td>${index + 1}</td>

                            <td>
                                ${escapeHtml(registro.fechaInicioTexto || "-")}
                            </td>

                            <td>
                                ${escapeHtml(registro.fechaFinTexto || "-")}
                            </td>

                            <td>
                                ${Number(registro.numeroDias || 0)}
                            </td>

                            <td>
                                ${escapeHtml(registro.tipoDocumento)}
                            </td>

                            <td>
                                ${escapeHtml(registro.contingencia)}
                            </td>

                            <td>
                                ${escapeHtml(registro.diagnostico)}
                            </td>

                            <td>
                                ${escapeHtml(registro.establecimiento)}
                            </td>
                        </tr>
                    `
                ).join("");
        }
    }

    mostrarModalDetalle();
}

// =====================================================
// MODAL PAPELETAS
// =====================================================

function abrirModalPapeletas({
    titulo,
    subtitulo,
    registros
}) {
    const ordenados =
        [...registros].sort(
            ordenarPapeletasRecientes
        );

    setTexto(
        "modalDetalleEyebrow",
        "Detalle de papeletas"
    );

    setTexto(
        "modalDetalleTitulo",
        titulo
    );

    setTexto(
        "modalDetalleSubtitulo",
        `${subtitulo} | ${ordenados.length} registros`
    );

    setTexto(
        "modalTotalRegistros",
        ordenados.length.toLocaleString(
            "es-PE"
        )
    );

    setTexto(
        "modalResumenSecundarioTitulo",
        "Meses con registros"
    );

    const mesesUnicos =
        new Set(
            ordenados
                .map(
                    registro =>
                        registro.mesNumero
                )
                .filter(Boolean)
        ).size;

    setTexto(
        "modalResumenSecundarioValor",
        mesesUnicos.toLocaleString(
            "es-PE"
        )
    );

    setTexto(
        "modalContadorDetalle",
        `${ordenados.length.toLocaleString("es-PE")} registros`
    );

    if (elementos.cabeceraModalDetalle) {
        elementos.cabeceraModalDetalle.innerHTML = `
            <th>N.°</th>
            <th>Mes</th>
            <th>Tipo de papeleta</th>
            <th>Motivo de papeleta</th>
            <th>Fecha inicio</th>
            <th>Fecha fin</th>
            <th>Observaciones</th>
            <th>Unidad orgánica</th>
        `;
    }

    if (elementos.cuerpoModalDetalle) {
        if (!ordenados.length) {
            elementos.cuerpoModalDetalle.innerHTML = `
                <tr>
                    <td
                        colspan="8"
                        class="tabla-consulta__vacio"
                    >
                        No se encontraron registros.
                    </td>
                </tr>
            `;

        } else {
            elementos.cuerpoModalDetalle.innerHTML =
                ordenados.map(
                    (registro, index) => `
                        <tr>
                            <td>${index + 1}</td>

                            <td>
                                ${escapeHtml(registro.mes)}
                            </td>

                            <td>
                                ${escapeHtml(registro.tipoPapeleta)}
                            </td>

                            <td>
                                ${escapeHtml(registro.motivoPapeleta)}
                            </td>

                            <td>
                                ${escapeHtml(registro.fechaInicioTexto || "-")}
                            </td>

                            <td>
                                ${escapeHtml(registro.fechaFinTexto || "-")}
                            </td>

                            <td>
                                ${escapeHtml(registro.observaciones)}
                            </td>

                            <td>
                                ${escapeHtml(registro.unidadOrganica)}
                            </td>
                        </tr>
                    `
                ).join("");
        }
    }

    mostrarModalDetalle();
}

function mostrarModalDetalle() {
    if (!elementos.modalDetalle) {
        return;
    }

    elementos.modalDetalle.classList.add(
        "modal-consulta--show"
    );

    elementos.modalDetalle.setAttribute(
        "aria-hidden",
        "false"
    );

    document.body.classList.add(
        "modal-abierto"
    );
}

function cerrarModalDetalle() {
    if (!elementos.modalDetalle) {
        return;
    }

    elementos.modalDetalle.classList.remove(
        "modal-consulta--show"
    );

    elementos.modalDetalle.setAttribute(
        "aria-hidden",
        "true"
    );

    document.body.classList.remove(
        "modal-abierto"
    );
}

// =====================================================
// CONTEOS
// =====================================================

function contarPorCampo(
    registros,
    campo
) {
    const mapa = new Map();

    registros.forEach(
        registro => {
            const valor =
                registro[campo] ||
                "NO ESPECIFICADO";

            mapa.set(
                valor,
                (
                    mapa.get(valor) ||
                    0
                ) + 1
            );
        }
    );

    return [...mapa.entries()]
        .map(
            ([label, valor]) => ({
                label,
                valor
            })
        )
        .sort(
            (a, b) => {
                if (b.valor !== a.valor) {
                    return b.valor - a.valor;
                }

                return a.label.localeCompare(
                    b.label,
                    "es"
                );
            }
        );
}

// =====================================================
// CHART HELPERS
// =====================================================

function destruirChart(clave) {
    if (estado.charts[clave]) {
        estado.charts[clave].destroy();

        estado.charts[clave] =
            null;
    }
}

// =====================================================
// ORDEN
// =====================================================

function ordenarDescansosRecientes(
    a,
    b
) {
    const fechaA =
        a.fechaInicio instanceof Date
            ? a.fechaInicio.getTime()
            : 0;

    const fechaB =
        b.fechaInicio instanceof Date
            ? b.fechaInicio.getTime()
            : 0;

    return fechaB - fechaA;
}

function ordenarPapeletasRecientes(
    a,
    b
) {
    const fechaA =
        a.fechaInicio instanceof Date
            ? a.fechaInicio.getTime()
            : 0;

    const fechaB =
        b.fechaInicio instanceof Date
            ? b.fechaInicio.getTime()
            : 0;

    return fechaB - fechaA;
}

// =====================================================
// LIMPIAR CONSULTA
// =====================================================

function limpiarConsulta() {
    estado.personaSeleccionada =
        null;

    estado.descansosSeleccionados =
        [];

    estado.papeletasSeleccionadas =
        [];

    cerrarTodosLosDashboards();

    Object.keys(
        estado.charts
    ).forEach(
        destruirChart
    );

    if (elementos.buscador) {
        elementos.buscador.value =
            "";

        elementos.buscador.focus();
    }

    if (elementos.fichaPersonal) {
        elementos.fichaPersonal.hidden =
            true;
    }

    if (elementos.estadoInicial) {
        elementos.estadoInicial.hidden =
            false;
    }

    cerrarSugerencias();
    cerrarModalDetalle();
}

// =====================================================
// EXCEL
// =====================================================

function validarBufferExcel(
    buffer,
    nombreArchivo
) {
    if (
        !(buffer instanceof ArrayBuffer) ||
        buffer.byteLength === 0
    ) {
        throw new Error(
            `El archivo ${nombreArchivo} está vacío o no es válido.`
        );
    }
}

function obtenerCampoFlexible(
    objeto,
    alternativas
) {
    if (
        !objeto ||
        typeof objeto !== "object"
    ) {
        return "";
    }

    const mapa = new Map();

    Object.keys(objeto).forEach(
        clave => {
            mapa.set(
                normalizarCabecera(
                    clave
                ),
                clave
            );
        }
    );

    for (
        const alternativa of alternativas
    ) {
        const claveReal =
            mapa.get(
                normalizarCabecera(
                    alternativa
                )
            );

        if (claveReal !== undefined) {
            const valor =
                objeto[claveReal];

            if (
                valor !== undefined &&
                valor !== null &&
                String(valor).trim() !== ""
            ) {
                return valor;
            }
        }
    }

    return "";
}

function normalizarCabecera(valor) {
    return String(valor ?? "")
        .trim()
        .normalize("NFD")
        .replace(
            /[\u0300-\u036f]/g,
            ""
        )
        .replace(
            /[_-]+/g,
            " "
        )
        .replace(
            /[.]/g,
            ""
        )
        .replace(
            /\s+/g,
            " "
        )
        .toUpperCase();
}

// =====================================================
// DNI Y TEXTOS
// =====================================================

function limpiarDNI(valor) {
    const documento =
        String(valor ?? "")
            .trim()
            .replace(
                /\.0$/,
                ""
            )
            .replace(
                /\D/g,
                ""
            );

    if (!documento) {
        return "";
    }

    return documento.padStart(
        8,
        "0"
    );
}

function compararDNI(
    dniA,
    dniB
) {
    const a =
        limpiarDNI(
            dniA
        );

    const b =
        limpiarDNI(
            dniB
        );

    return Boolean(
        a &&
        b &&
        a === b
    );
}

function normalizarTextoVisible(valor) {
    return String(valor ?? "")
        .trim()
        .replace(
            /\s+/g,
            " "
        )
        .toUpperCase();
}

function normalizarBusqueda(valor) {
    return String(valor ?? "")
        .trim()
        .normalize("NFD")
        .replace(
            /[\u0300-\u036f]/g,
            ""
        )
        .replace(
            /\s+/g,
            " "
        )
        .toUpperCase();
}

function normalizarSexo(valor) {
    const sexo =
        normalizarBusqueda(
            valor
        );

    const equivalencias = {
        M: "MASCULINO",
        MASCULINO: "MASCULINO",
        HOMBRE: "MASCULINO",
        VARON: "MASCULINO",

        F: "FEMENINO",
        FEMENINO: "FEMENINO",
        MUJER: "FEMENINO"
    };

    return equivalencias[sexo] ||
        sexo ||
        "";
}

// =====================================================
// MESES
// =====================================================

function obtenerNumeroMes(valor) {
    const mes =
        normalizarBusqueda(
            valor
        );

    return NUMERO_MES[mes] ||
        null;
}

// =====================================================
// FECHAS
// =====================================================

function convertirFecha(valor) {
    if (
        valor === null ||
        valor === undefined ||
        valor === ""
    ) {
        return null;
    }

    if (
        valor instanceof Date &&
        !Number.isNaN(
            valor.getTime()
        )
    ) {
        return normalizarFecha(
            valor
        );
    }

    if (
        typeof valor === "number"
    ) {
        const partes =
            XLSX.SSF.parse_date_code(
                valor
            );

        if (!partes) {
            return null;
        }

        return new Date(
            partes.y,
            partes.m - 1,
            partes.d
        );
    }

    const texto =
        String(valor)
            .trim()
            .split(" ")[0]
            .replace(
                /-/g,
                "/"
            );

    const partes =
        texto.split("/");

    if (partes.length !== 3) {
        return null;
    }

    let dia;
    let mes;
    let anio;

    if (partes[0].length === 4) {
        anio =
            Number(partes[0]);

        mes =
            Number(partes[1]);

        dia =
            Number(partes[2]);

    } else {
        dia =
            Number(partes[0]);

        mes =
            Number(partes[1]);

        anio =
            Number(partes[2]);
    }

    if (anio < 100) {
        anio +=
            anio >= 50
                ? 1900
                : 2000;
    }

    if (
        !dia ||
        !mes ||
        !anio
    ) {
        return null;
    }

    const fecha =
        new Date(
            anio,
            mes - 1,
            dia
        );

    if (
        fecha.getFullYear() !== anio ||
        fecha.getMonth() !== mes - 1 ||
        fecha.getDate() !== dia
    ) {
        return null;
    }

    return fecha;
}

function normalizarFecha(fecha) {
    return new Date(
        fecha.getFullYear(),
        fecha.getMonth(),
        fecha.getDate()
    );
}

function formatearFecha(fecha) {
    if (
        !(fecha instanceof Date) ||
        Number.isNaN(
            fecha.getTime()
        )
    ) {
        return "";
    }

    return fecha.toLocaleDateString(
        "es-PE",
        {
            day: "2-digit",
            month: "2-digit",
            year: "numeric"
        }
    );
}

// =====================================================
// NÚMEROS
// =====================================================

function convertirNumero(valor) {
    if (
        valor === null ||
        valor === undefined ||
        valor === ""
    ) {
        return 0;
    }

    const numero =
        Number(
            String(valor)
                .trim()
                .replace(
                    ",",
                    "."
                )
        );

    return Number.isFinite(numero)
        ? numero
        : 0;
}

function convertirNota(valor) {
    if (
        valor === null ||
        valor === undefined ||
        String(valor).trim() === ""
    ) {
        return null;
    }

    const numero =
        Number(
            String(valor)
                .trim()
                .replace(
                    ",",
                    "."
                )
        );

    return Number.isFinite(numero)
        ? numero
        : null;
}

function formatearNota(nota) {
    return Number(nota)
        .toLocaleString(
            "es-PE",
            {
                minimumFractionDigits:
                    Number.isInteger(
                        Number(nota)
                    )
                        ? 0
                        : 1,

                maximumFractionDigits:
                    2
            }
        );
}

// =====================================================
// UI
// =====================================================

function mostrarLoading(mostrar) {
    if (!elementos.loading) {
        return;
    }

    elementos.loading.style.display =
        mostrar
            ? "flex"
            : "none";
}

function mostrarError(mensaje) {
    if (elementos.estadoInicial) {
        elementos.estadoInicial.hidden =
            true;
    }

    if (elementos.fichaPersonal) {
        elementos.fichaPersonal.hidden =
            true;
    }

    if (elementos.estadoError) {
        elementos.estadoError.hidden =
            false;
    }

    if (elementos.mensajeError) {
        elementos.mensajeError.textContent =
            mensaje;
    }
}

function ocultarError() {
    if (elementos.estadoError) {
        elementos.estadoError.hidden =
            true;
    }
}

// =====================================================
// UTILIDADES
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

function truncarTexto(
    texto,
    limite
) {
    const valor =
        String(texto ?? "");

    if (valor.length <= limite) {
        return valor;
    }

    return `${valor.slice(0, limite - 3)}...`;
}

function escapeHtml(valor) {
    return String(valor ?? "")
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

function debounce(
    funcion,
    espera = 200
) {
    let temporizador;

    return (...argumentos) => {
        clearTimeout(
            temporizador
        );

        temporizador =
            setTimeout(
                () =>
                    funcion(
                        ...argumentos
                    ),
                espera
            );
    };
}

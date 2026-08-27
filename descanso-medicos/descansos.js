// =====================================================
// DESCANSOS.JS — DASHBOARD DM CORREGIDO
// Base principal de evaluación: numdocumento
// Evita duplicados por DNI + fechas + documento + diagnóstico
// Filtros múltiples con panel flotante estable
// 2 primeros gráficos: total del año
// Resto de gráficos: dinámicos según filtros
// Incluye matriz horizontal de alertas 25 a 30 días
// =====================================================

import * as XLSX from "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/+esm";

import {
    obtenerArchivoDescansosMedicos,
    descargarExcelStorage
} from "../backend/api.js";
// =====================================================
// ESTADO GLOBAL
// =====================================================
const estado = {
    data: [],
    dataFiltradaActual: [],
    busquedaPersonaRanking: "",
    filtros: {
        meses: [],
        tipoDocumento: [],
        contingencia: [],
        diagnostico: [],
        hospital: []
    },
    charts: {
        meses: null,
        condicionMes: null,
        lineasMesesPorDia: null,
        tipoDocumento: null,
        contingencia: null,
        topPersonas: null,
        cargo: null,
        hospitales: null,
        reincidentesPosteriorDM: null
    },

    asistencia: {
        archivosPorMes: new Map(),
        mesesCargados: new Set(),
        mesesNoEncontrados: new Set(),
        resultados: []
    }
};

let filtroFlotanteActivo = null;

// =====================================================
// CONSTANTES
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

const MESES_CORTOS = [
    "ENE",
    "FEB",
    "MAR",
    "ABR",
    "MAY",
    "JUN",
    "JUL",
    "AGO",
    "SET",
    "OCT",
    "NOV",
    "DIC"
];
// =====================================================

// ROL DE SERVICIO EN FIREBASE STORAGE

// =====================================================

const RUTA_ROL_SERVICIO =

    "RENDIMIENTO PERSONAL/ROL DE SERVICIO";

const ARCHIVOS_ASISTENCIA = {
    1: ["enero.xlsx"],
    2: ["febrero.xlsx"],
    3: ["marzo.xlsx"],
    4: ["abril.xlsx"],
    5: ["mayo.xlsx"],
    6: ["junio.xlsx"],
    7: ["julio.xlsx"],
    8: ["agosto.xlsx"],
    9: ["setiembre.xlsx", "septiembre.xlsx"],
    10: ["octubre.xlsx"],
    11: ["noviembre.xlsx"],
    12: ["diciembre.xlsx"]
};


const COLORES_LINEAS = [
    "#556b2f",
    "#2563eb",
    "#dc2626",
    "#f59e0b",
    "#7c3aed",
    "#0891b2",
    "#16a34a",
    "#be123c",
    "#9333ea",
    "#0f766e",
    "#ea580c",
    "#475569"
];

const CONFIG_FILTROS = {
    meses: {
        panelId: "opcionesFiltroMes",
        textoId: "textoFiltroMes",
        textoTodos: "TODOS LOS MESES",
        campo: "mesNumero"
    },
    tipoDocumento: {
        panelId: "opcionesFiltroTipoDocumento",
        textoId: "textoFiltroTipoDocumento",
        textoTodos: "TODOS LOS TIPOS",
        campo: "tipoDocumento"
    },
    contingencia: {
        panelId: "opcionesFiltroContingencia",
        textoId: "textoFiltroContingencia",
        textoTodos: "TODAS LAS CONTINGENCIAS",
        campo: "contingencia"
    },
    diagnostico: {
        panelId: "opcionesFiltroDiagnostico",
        textoId: "textoFiltroDiagnostico",
        textoTodos: "TODOS LOS DIAGNÓSTICOS",
        campo: "diagnostico"
    },
    hospital: {
        panelId: "opcionesFiltroHospital",
        textoId: "textoFiltroHospital",
        textoTodos: "TODOS LOS HOSPITALES",
        campo: "establecimientoSalud"
    }
};

// =====================================================
// INICIO
// =====================================================
window.addEventListener("DOMContentLoaded", async () => {

    mostrarLoading(true);

    try {

        await cargarExcel();
        await cargarArchivosAsistenciaDisponibles();

        inicializarFiltros();

        inicializarBuscadorRanking();

        inicializarBotonLimpiarTodo();

        inicializarModal();

        inicializarDesplegablesFiltrosSinScroll();

        renderDashboard();

    } catch (error) {

        console.error("Error general:", error);
            alert(
                "No se pudo cargar la información desde Firebase Storage.\n\n" +
                "Verifica que exista el archivo:\n" +
                "RENDIMIENTO PERSONAL/DESCANSO MEDICOS/DM.xlsx\n\n" +
                "También revisa las reglas de acceso de Firebase Storage."
            );

    } finally {

        mostrarLoading(false);

    }

});

// =====================================================
// CARGAR EXCEL
// =====================================================
// =====================================================
// CARGAR DM.XLSX DESDE FIREBASE STORAGE
// =====================================================
async function cargarExcel() {

    console.log(
        "Descargando DM.xlsx desde Firebase Storage..."
    );

    const buffer =
        await obtenerArchivoDescansosMedicos();

    if (
        !(buffer instanceof ArrayBuffer) ||
        buffer.byteLength === 0
    ) {

        throw new Error(
            "El archivo DM.xlsx descargado está vacío o no es válido."
        );

    }

    const workbook = XLSX.read(buffer, {
        type: "array",
        cellDates: true
    });

    if (!workbook.SheetNames.length) {

        throw new Error(
            "El archivo DM.xlsx no contiene ninguna hoja."
        );

    }

    const nombreHoja =
        workbook.SheetNames[0];

    const hoja =
        workbook.Sheets[nombreHoja];

    if (!hoja) {

        throw new Error(
            "No se pudo leer la primera hoja de DM.xlsx."
        );

    }

    const filas = XLSX.utils.sheet_to_json(hoja, {
        defval: "",
        raw: true
    });

    const dataNormalizada = filas
        .map((row, index) =>
            normalizarRegistro(
                row,
                index + 2
            )
        )
        .filter(item =>
            item.nombre &&
            item.fechaInicio &&
            item.numeroDias > 0 &&
            item.numdocumento
        );

    estado.data =
        eliminarDuplicadosPorDNI(
            dataNormalizada
        );

    console.log(
        "DM.xlsx cargado desde Firebase Storage."
    );

    console.log(
        "Hoja procesada:",
        nombreHoja
    );

    console.log(
        "Filas encontradas:",
        filas.length
    );

    console.log(
        "Registros normalizados:",
        dataNormalizada.length
    );

    console.log(
        "Registros sin duplicados:",
        estado.data.length
    );

    console.log(
        "Primer registro:",
        estado.data[0]
    );

}


// =====================================================
// NORMALIZAR REGISTRO
// =====================================================
function normalizarRegistro(row, filaExcel) {

    const nombre = normalizarTexto(
        obtenerCampo(row, [
            "des_nomapell",
            "DES_NOMAPELL",
            "Des_NomApell",
            "APELLIDOS Y NOMBRES",
            "NOMBRES"
        ])
    );

    const numdocumento = limpiarDocumento(
        obtenerCampo(row, [
            "numdocumento",
            "NUMDOCUMENTO",
            "NUM_DOCUMENTO",
            "DNI",
            "dni",
            "Documento",
            "DOCUMENTO"
        ])
    );

    const fechaInicio = convertirFecha(
        obtenerCampo(row, [
            "fec_inicio",
            "FEC_INICIO",
            "fec inicio",
            "FECHA DE INICIO",
            "FECHA INICIO"
        ])
    );

    const fechaFin = convertirFecha(
        obtenerCampo(row, [
            "fec_termino",
            "FEC_TERMINO",
            "fec termino",
            "FECHA DE FIN",
            "FECHA FIN"
        ])
    );

    const numeroDias = Number(
        obtenerCampo(row, [
            "cant. Días",
            "cant. días",
            "CANT. DÍAS",
            "CANT. DIAS",
            "cant. Dias",
            "cant_dias",
            "CANT_DIAS",
            "DIAS",
            "DÍAS"
        ])
    ) || 0;

    const nomCondicion = normalizarTexto(
        obtenerCampo(row, [
            "nom_condicion",
            "NOM_CONDICION",
            "nom condicion",
            "CONDICION",
            "CONDICIÓN"
        ])
    );

    const cargo = normalizarTexto(
        obtenerCampo(row, [
            "des_car",
            "DES_CAR",
            "des car",
            "CARGO"
        ])
    );

    const tipoDocumento = normalizarTexto(
        obtenerCampo(row, [
            "txttipodocumento",
            "TXTTIPODOCUMENTO",
            "txt tipo documento",
            "TIPO DOCUMENTO",
            "TIPO DE DOCUMENTO"
        ])
    );

    const contingencia = normalizarTexto(
        obtenerCampo(row, [
            "txtcontingencia",
            "TXTCONTINGENCIA",
            "txt contingencia",
            "CONTINGENCIA"
        ])
    );

    const diagnostico = normalizarTexto(
        obtenerCampo(row, [
            "txtotrosdiagnosticos",
            "TXTOTROSDIAGNOSTICOS",
            "txt otros diagnosticos",
            "OTROS DIAGNOSTICOS",
            "OTROS DIAGNÓSTICOS",
            "DIAGNOSTICO",
            "DIAGNÓSTICO"
        ])
    );

    const establecimientoSalud = normalizarTexto(
        obtenerCampo(row, [
            "txtestablecimientosalud",
            "TXTESTABLECIMIENTOSALUD",
            "txt establecimiento salud",
            "ESTABLECIMIENTO SALUD",
            "ESTABLECIMIENTO DE SALUD",
            "HOSPITAL"
        ])
    );

    const anio = fechaInicio
        ? fechaInicio.getFullYear()
        : null;

    const mesNumero = fechaInicio
        ? fechaInicio.getMonth() + 1
        : null;

    const mes = fechaInicio
        ? MESES_ORDEN[fechaInicio.getMonth()]
        : "";

    const diaMes = fechaInicio
        ? fechaInicio.getDate()
        : null;

    return {
        filaExcel,
        nombre,
        numdocumento,
        personaKey: numdocumento,
        etiquetaPersona: `${nombre} | DNI: ${numdocumento}`,
        fechaInicio,
        fechaFin,
        fechaInicioTexto: formatearFecha(fechaInicio),
        fechaFinTexto: formatearFecha(fechaFin),
        numeroDias,
        anio,
        mes,
        mesNumero,
        diaMes,
        nomCondicion,
        cargo,
        tipoDocumento,
        contingencia,
        diagnostico,
        establecimientoSalud
    };

}

// =====================================================
// ELIMINAR DUPLICADOS REALES
// =====================================================
function eliminarDuplicadosPorDNI(data) {

    const mapa = new Map();

    data.forEach(item => {

        const key = [
            item.numdocumento,
            item.fechaInicioTexto,
            item.fechaFinTexto,
            item.numeroDias,
            item.tipoDocumento,
            item.contingencia,
            item.diagnostico,
            item.establecimientoSalud
        ].join("|");

        if (!mapa.has(key)) {
            mapa.set(key, item);
        }

    });

    return [...mapa.values()];

}

// =====================================================
// FILTROS
// =====================================================
function inicializarFiltros() {

    const filtroAnio = document.getElementById("filtroAnio");

    if (!filtroAnio) return;

    const anios = [...new Set(
        estado.data
            .map(item => item.anio)
            .filter(Boolean)
    )].sort((a, b) => a - b);

    filtroAnio.innerHTML = "";

    if (!anios.length) {

        const option = document.createElement("option");
        option.value = "";
        option.textContent = "SIN DATOS";
        filtroAnio.appendChild(option);

        return;

    }

    anios.forEach(anio => {

        const option = document.createElement("option");
        option.value = String(anio);
        option.textContent = String(anio);

        filtroAnio.appendChild(option);

    });

    filtroAnio.value = String(anios[anios.length - 1]);

    cargarFiltrosPorAnio();

    filtroAnio.addEventListener("change", () => {

        limpiarFiltrosSinCambiarAnio();

        cargarFiltrosPorAnio();

        cerrarPanelFlotanteFiltro();

        renderDashboard();

    });

    const inputNombre = document.getElementById("filtroNombreGeneral");

    if (inputNombre) {

        inputNombre.addEventListener("input", () => {

            limpiarBuscadorRanking();

            actualizarFiltrosDependientes("nombre");

            cerrarPanelFlotanteFiltro();

            renderDashboard();

        });

    }

}

function cargarFiltrosPorAnio() {

    const dataAnio = obtenerDataAnioActual();

    cargarListaNombres(dataAnio);

    cargarFiltroMeses(dataAnio);

    cargarFiltroCampo({
        key: "tipoDocumento",
        data: dataAnio,
        campo: "tipoDocumento"
    });

    cargarFiltroCampo({
        key: "contingencia",
        data: dataAnio,
        campo: "contingencia"
    });

    cargarFiltroCampo({
        key: "diagnostico",
        data: dataAnio,
        campo: "diagnostico"
    });

    cargarFiltroCampo({
        key: "hospital",
        data: dataAnio,
        campo: "establecimientoSalud"
    });

    actualizarTextoTodosFiltros();

}

function cargarFiltroMeses(dataAnio) {

    const mesesDisponibles = obtenerMesesDisponibles(dataAnio);

    const opciones = mesesDisponibles.map(mesNumero => ({
        value: String(mesNumero),
        label: MESES_ORDEN[mesNumero - 1]
    }));

    renderFiltroCheckbox({
        key: "meses",
        opciones
    });

}

function cargarFiltroCampo({
    key,
    data,
    campo,
    conservarSeleccion = false
}) {

    const valores = [...new Set(
        data
            .map(item => item[campo] || "NO ESPECIFICADO")
            .filter(Boolean)
    )].sort((a, b) => a.localeCompare(b, "es"));

    if (conservarSeleccion) {

        estado.filtros[key] = estado.filtros[key].filter(valor =>
            valores.includes(valor)
        );

    } else {

        estado.filtros[key] = [];

    }

    const opciones = valores.map(valor => ({
        value: valor,
        label: valor
    }));

    renderFiltroCheckbox({
        key,
        opciones
    });

}

function renderFiltroCheckbox({
    key,
    opciones
}) {

    const config = CONFIG_FILTROS[key];

    if (!config) return;

    const panel = document.getElementById(config.panelId);

    if (!panel) return;

    const seleccion = estado.filtros[key] || [];

    panel.innerHTML = "";

    const opcionTodos = crearOpcionCheckbox({
        key,
        value: "TODOS",
        label: config.textoTodos,
        checked: seleccion.length === 0,
        esTodos: true
    });

    panel.appendChild(opcionTodos);

    opciones.forEach(opcion => {

        const item = crearOpcionCheckbox({
            key,
            value: opcion.value,
            label: opcion.label,
            checked: seleccion.includes(opcion.value),
            esTodos: false
        });

        panel.appendChild(item);

    });

    actualizarTextoFiltro(key, opciones);

    if (filtroFlotanteActivo && filtroFlotanteActivo.key === key) {

        renderPanelFlotanteFiltro(key);

        requestAnimationFrame(() => {
            posicionarPanelFlotanteFiltro();
        });

    }

}

function crearOpcionCheckbox({
    key,
    value,
    label,
    checked,
    esTodos
}) {

    const item = document.createElement("label");

    item.className = esTodos
        ? "check-option check-option--all"
        : "check-option";

    const input = document.createElement("input");

    input.type = "checkbox";
    input.value = value;
    input.checked = checked;

    const span = document.createElement("span");

    span.textContent = label;

    item.appendChild(input);
    item.appendChild(span);

    input.addEventListener("change", event => {

        event.stopPropagation();

        aplicarCambioFiltroCheckbox({
            key,
            value,
            checked: input.checked
        });

    });

    return item;

}

function aplicarCambioFiltroCheckbox({
    key,
    value,
    checked
}) {

    if (value === "TODOS") {

        estado.filtros[key] = [];

    } else {

        let seleccion = [...estado.filtros[key]];

        if (checked) {

            if (!seleccion.includes(value)) {
                seleccion.push(value);
            }

        } else {

            seleccion = seleccion.filter(item => item !== value);

        }

        estado.filtros[key] = seleccion;

    }

    limpiarBuscadorRanking();

    actualizarFiltrosDependientes(key);

    renderDashboard();

    if (filtroFlotanteActivo && filtroFlotanteActivo.key === key) {

        renderPanelFlotanteFiltro(key);

        requestAnimationFrame(() => {
            posicionarPanelFlotanteFiltro();
        });

    }

}

function actualizarFiltrosDependientes(filtroCambiado = null) {

    const dataBase = obtenerDataPorAnioNombreMes();

    if (filtroCambiado !== "meses") {

        const mesesDisponibles = obtenerMesesDisponibles(
            filtrarDataAvanzada(dataBase, {
                ignorarMeses: true
            })
        );

        const opcionesMes = mesesDisponibles.map(mesNumero => ({
            value: String(mesNumero),
            label: MESES_ORDEN[mesNumero - 1]
        }));

        estado.filtros.meses = estado.filtros.meses.filter(mes =>
            opcionesMes.some(opcion => opcion.value === mes)
        );

        renderFiltroCheckbox({
            key: "meses",
            opciones: opcionesMes
        });

    }

    if (filtroCambiado !== "tipoDocumento") {

        cargarFiltroCampo({
            key: "tipoDocumento",
            data: filtrarDataAvanzada(dataBase, {
                ignorarTipo: true
            }),
            campo: "tipoDocumento",
            conservarSeleccion: true
        });

    }

    if (filtroCambiado !== "contingencia") {

        cargarFiltroCampo({
            key: "contingencia",
            data: filtrarDataAvanzada(dataBase, {
                ignorarContingencia: true
            }),
            campo: "contingencia",
            conservarSeleccion: true
        });

    }

    if (filtroCambiado !== "diagnostico") {

        cargarFiltroCampo({
            key: "diagnostico",
            data: filtrarDataAvanzada(dataBase, {
                ignorarDiagnostico: true
            }),
            campo: "diagnostico",
            conservarSeleccion: true
        });

    }

    if (filtroCambiado !== "hospital") {

        cargarFiltroCampo({
            key: "hospital",
            data: filtrarDataAvanzada(dataBase, {
                ignorarHospital: true
            }),
            campo: "establecimientoSalud",
            conservarSeleccion: true
        });

    }

}

function actualizarTextoFiltro(key, opciones) {

    const config = CONFIG_FILTROS[key];

    if (!config) return;

    const seleccion = estado.filtros[key] || [];

    if (!seleccion.length) {

        setTexto(config.textoId, config.textoTodos);

        return;

    }

    if (seleccion.length === 1) {

        const encontrado = opciones.find(opcion =>
            opcion.value === seleccion[0]
        );

        setTexto(
            config.textoId,
            encontrado ? encontrado.label : seleccion[0]
        );

        return;

    }

    setTexto(config.textoId, `${seleccion.length} seleccionados`);

}

function actualizarTextoTodosFiltros() {

    setTexto("textoFiltroMes", "TODOS LOS MESES");
    setTexto("textoFiltroTipoDocumento", "TODOS LOS TIPOS");
    setTexto("textoFiltroContingencia", "TODAS LAS CONTINGENCIAS");
    setTexto("textoFiltroDiagnostico", "TODOS LOS DIAGNÓSTICOS");
    setTexto("textoFiltroHospital", "TODOS LOS HOSPITALES");

}

function cargarListaNombres(data) {

    const datalist =
        document.getElementById(
            "listaNombresPersonal"
        );

    if (!datalist) return;

    const personas =
        obtenerPersonasUnicas(data);

    datalist.innerHTML =
        personas.map(
            item => `
                <option
                    value="${escapeHtml(item.nombre)}"
                    label="DNI: ${escapeHtml(item.dni)}"
                ></option>
            `
        ).join("");

}

// =====================================================
// PANEL FLOTANTE DE FILTROS
// =====================================================
function inicializarDesplegablesFiltrosSinScroll() {

    crearPanelFlotanteFiltro();

    document.querySelectorAll(".filtro-check summary").forEach(summary => {

        summary.addEventListener("click", event => {

            event.preventDefault();
            event.stopPropagation();

            const details = summary.closest(".filtro-check");

            if (!details) return;

            const key = obtenerKeyFiltroDesdeDetails(details.id);

            if (!key) return;

            if (
                filtroFlotanteActivo &&
                filtroFlotanteActivo.key === key &&
                filtroFlotanteActivo.details === details
            ) {

                cerrarPanelFlotanteFiltro();

                return;

            }

            abrirPanelFlotanteFiltro(details, key);

        });

    });

    document.addEventListener("click", event => {

        if (
            event.target.closest(".filtro-check") ||
            event.target.closest("#panelFiltroFlotante")
        ) {
            return;
        }

        cerrarPanelFlotanteFiltro();

    });

    window.addEventListener("resize", () => {
        reposicionarPanelFlotanteFiltro();
    });

    window.addEventListener("scroll", () => {
        reposicionarPanelFlotanteFiltro();
    }, true);

    const contenedorScroll = document.querySelector(".top-header__right--simple");

    if (contenedorScroll) {

        contenedorScroll.addEventListener("scroll", () => {
            reposicionarPanelFlotanteFiltro();
        });

    }

}

function crearPanelFlotanteFiltro() {

    if (document.getElementById("panelFiltroFlotante")) return;

    const panel = document.createElement("div");

    panel.id = "panelFiltroFlotante";
    panel.className = "panel-filtro-flotante";
    panel.style.display = "none";

    document.body.appendChild(panel);

}

function obtenerKeyFiltroDesdeDetails(id) {

    const mapa = {
        boxFiltroMes: "meses",
        boxFiltroTipoDocumento: "tipoDocumento",
        boxFiltroContingencia: "contingencia",
        boxFiltroDiagnostico: "diagnostico",
        boxFiltroHospital: "hospital"
    };

    return mapa[id] || null;

}

function abrirPanelFlotanteFiltro(details, key) {

    const panel = document.getElementById("panelFiltroFlotante");

    if (!panel) return;

    filtroFlotanteActivo = {
        details,
        key
    };

    document.querySelectorAll(".filtro-check").forEach(item => {
        item.removeAttribute("open");
        item.classList.remove("filtro-check--activo");
    });

    details.classList.add("filtro-check--activo");

    renderPanelFlotanteFiltro(key);

    panel.style.display = "block";

    posicionarPanelFlotanteFiltro();

}

function renderPanelFlotanteFiltro(key) {

    const panel = document.getElementById("panelFiltroFlotante");

    if (!panel) return;

    const config = CONFIG_FILTROS[key];

    if (!config) return;

    const panelOriginal = document.getElementById(config.panelId);

    if (!panelOriginal) return;

    const seleccion = estado.filtros[key] || [];

    const opcionesOriginales = [...panelOriginal.querySelectorAll(".check-option")].map(label => {

        const input = label.querySelector("input");
        const span = label.querySelector("span");

        return {
            value: input ? input.value : "",
            label: span ? span.textContent : "",
            esTodos: input ? input.value === "TODOS" : false
        };

    });

    panel.innerHTML = "";

    opcionesOriginales.forEach(opcion => {

        const item = document.createElement("label");

        item.className = opcion.esTodos
            ? "check-option check-option--all"
            : "check-option";

        const input = document.createElement("input");

        input.type = "checkbox";
        input.value = opcion.value;

        input.checked = opcion.esTodos
            ? seleccion.length === 0
            : seleccion.includes(opcion.value);

        const span = document.createElement("span");

        span.textContent = opcion.label;

        item.appendChild(input);
        item.appendChild(span);

        input.addEventListener("click", event => {
            event.stopPropagation();
        });

        input.addEventListener("change", event => {

            event.stopPropagation();

            aplicarCambioFiltroCheckbox({
                key,
                value: opcion.value,
                checked: input.checked
            });

        });

        panel.appendChild(item);

    });

}
function posicionarPanelFlotanteFiltro() {

    if (!filtroFlotanteActivo) return;

    const panel = document.getElementById("panelFiltroFlotante");

    if (!panel) return;

    const details = filtroFlotanteActivo.details;

    if (!details) return;

    const rect = details.getBoundingClientRect();

    const margen = 12;
    const separacion = 10;

    let ancho = details.classList.contains("filtro-check--wide")
        ? 440
        : 330;

    ancho = Math.min(
        ancho,
        window.innerWidth - margen * 2
    );

    let left = rect.left;

    if (left + ancho > window.innerWidth - margen) {
        left = window.innerWidth - ancho - margen;
    }

    if (left < margen) {
        left = margen;
    }

    /*
     * Altura disponible desde la parte inferior
     * del botón hasta el borde inferior de la pantalla.
     */
    const espacioAbajo =
        window.innerHeight -
        rect.bottom -
        separacion -
        margen;

    /*
     * Altura disponible encima del botón.
     */
    const espacioArriba =
        rect.top -
        separacion -
        margen;

    /*
     * Por defecto se abre hacia abajo.
     * Si no existe suficiente espacio, se abre hacia arriba.
     */
    const abrirHaciaArriba =
        espacioAbajo < 280 &&
        espacioArriba > espacioAbajo;

    /*
     * La altura máxima nunca será mayor de 460 px
     * ni mayor que el espacio disponible.
     */
    const alturaDisponible = abrirHaciaArriba
        ? espacioArriba
        : espacioAbajo;

    const alturaMaxima = Math.max(
        180,
        Math.min(460, alturaDisponible)
    );

    panel.style.position = "fixed";
    panel.style.left = `${left}px`;
    panel.style.width = `${ancho}px`;
    panel.style.maxHeight = `${alturaMaxima}px`;

    /*
     * Esta es la corrección principal.
     */
    panel.style.overflowY = "auto";
    panel.style.overflowX = "hidden";
    panel.style.overscrollBehavior = "contain";

    panel.style.zIndex = "999999";
    panel.style.background = "#ffffff";
    panel.style.opacity = "1";

    if (abrirHaciaArriba) {

        panel.style.top = "auto";

        panel.style.bottom =
            `${window.innerHeight - rect.top + separacion}px`;

    } else {

        panel.style.bottom = "auto";

        panel.style.top =
            `${rect.bottom + separacion}px`;

    }

}


function reposicionarPanelFlotanteFiltro() {

    if (!filtroFlotanteActivo) return;

    posicionarPanelFlotanteFiltro();

}

function cerrarPanelFlotanteFiltro() {

    const panel = document.getElementById("panelFiltroFlotante");

    if (panel) {
        panel.style.display = "none";
        panel.innerHTML = "";
    }

    document.querySelectorAll(".filtro-check").forEach(item => {
        item.removeAttribute("open");
        item.classList.remove("filtro-check--activo");
    });

    filtroFlotanteActivo = null;

}

function cerrarTodosLosFiltrosAbiertos() {

    cerrarPanelFlotanteFiltro();

}

// =====================================================
// LIMPIAR TODO
// =====================================================
function inicializarBotonLimpiarTodo() {

    const btn = document.getElementById("btnLimpiarTodo");

    if (!btn) return;

    btn.addEventListener("click", () => {
        limpiarTodosLosFiltros();
    });

}

function limpiarTodosLosFiltros() {

    limpiarFiltrosSinCambiarAnio();

    cargarFiltrosPorAnio();

    cerrarPanelFlotanteFiltro();

    renderDashboard();

}

function limpiarFiltrosSinCambiarAnio() {

    estado.filtros = {
        meses: [],
        tipoDocumento: [],
        contingencia: [],
        diagnostico: [],
        hospital: []
    };

    estado.busquedaPersonaRanking = "";

    const inputNombre = document.getElementById("filtroNombreGeneral");

    if (inputNombre) {
        inputNombre.value = "";
    }

    const buscadorRanking = document.getElementById("buscarPersonaDescanso");

    if (buscadorRanking) {
        buscadorRanking.value = "";
    }

}

// =====================================================
// BUSCADOR RANKING PERSONA
// =====================================================
function inicializarBuscadorRanking() {

    const input = document.getElementById("buscarPersonaDescanso");

    if (!input) return;

    input.addEventListener("input", () => {

        estado.busquedaPersonaRanking = normalizarTexto(input.value);

        renderGraficoTopPersonas(estado.dataFiltradaActual);

    });

}

function limpiarBuscadorRanking() {

    estado.busquedaPersonaRanking = "";

    const input = document.getElementById("buscarPersonaDescanso");

    if (input) {
        input.value = "";
    }

}

// =====================================================
// OBTENER DATA SEGÚN FILTROS
// =====================================================
function obtenerDataAnioActual() {

    const anioSeleccionado = Number(getValorSelect("filtroAnio"));

    return estado.data.filter(item =>
        item.anio === anioSeleccionado
    );

}

function obtenerDataPorAnioNombreMes() {

    const textoBuscado = getValorInput("filtroNombreGeneral");
    const meses = estado.filtros.meses;

    let data = obtenerDataAnioActual();

    if (textoBuscado) {

        data = data.filter(item =>
            item.nombre.includes(textoBuscado) ||
            item.numdocumento.includes(textoBuscado)
        );

    }

    if (meses.length) {

        data = data.filter(item =>
            meses.includes(String(item.mesNumero))
        );

    }

    return data;

}

function obtenerDataDinamicaFiltrada() {

    const dataBase = obtenerDataPorAnioNombreMes();

    return filtrarDataAvanzada(dataBase);

}

function filtrarDataAvanzada(data, opciones = {}) {

    const tipos = estado.filtros.tipoDocumento;
    const contingencias = estado.filtros.contingencia;
    const diagnosticos = estado.filtros.diagnostico;
    const hospitales = estado.filtros.hospital;
    const meses = estado.filtros.meses;

    return data.filter(item => {

        const itemMes = String(item.mesNumero);
        const itemTipo = item.tipoDocumento || "NO ESPECIFICADO";
        const itemContingencia = item.contingencia || "NO ESPECIFICADO";
        const itemDiagnostico = item.diagnostico || "NO ESPECIFICADO";
        const itemHospital = item.establecimientoSalud || "NO ESPECIFICADO";

        const cumpleMes =
            opciones.ignorarMeses ||
            !meses.length ||
            meses.includes(itemMes);

        const cumpleTipo =
            opciones.ignorarTipo ||
            !tipos.length ||
            tipos.includes(itemTipo);

        const cumpleContingencia =
            opciones.ignorarContingencia ||
            !contingencias.length ||
            contingencias.includes(itemContingencia);

        const cumpleDiagnostico =
            opciones.ignorarDiagnostico ||
            !diagnosticos.length ||
            diagnosticos.includes(itemDiagnostico);

        const cumpleHospital =
            opciones.ignorarHospital ||
            !hospitales.length ||
            hospitales.includes(itemHospital);

        return cumpleMes &&
            cumpleTipo &&
            cumpleContingencia &&
            cumpleDiagnostico &&
            cumpleHospital;

    });

}

function getValorSelect(id) {

    return document.getElementById(id)?.value || "";

}

function getValorInput(id) {

    return normalizarTexto(document.getElementById(id)?.value || "");

}

// =====================================================
// RENDER GENERAL
// =====================================================
function renderDashboard() {

    const anioSeleccionado = Number(getValorSelect("filtroAnio"));

    const dataAnio = obtenerDataAnioActual();

    const dataDinamica = obtenerDataDinamicaFiltrada();

    estado.dataFiltradaActual = dataDinamica;

    setTexto("textoAnio", anioSeleccionado || "-");
    setTexto("textoAnioCondicion", anioSeleccionado || "-");

    actualizarKPIs(dataDinamica);

    renderGraficoDescansosPorMes(dataAnio);
    renderGraficoCondicionPorMes(dataAnio);

    renderGraficoLineasMesesPorDia(dataDinamica);
    renderGraficoTipoDocumento(dataDinamica);
    renderGraficoContingencia(dataDinamica);
    renderGraficoTopPersonas(dataDinamica);
    renderGraficoCargo(dataDinamica);
    renderMatrizAlertaDescansos(dataDinamica);

    renderDescansosAntesAusencia(dataDinamica);

    renderGraficoReincidentesPosteriorDM(dataDinamica);

    renderGraficoHospitales(dataDinamica);
}

// =====================================================
// KPI
// =====================================================
function actualizarKPIs(data) {

    setTexto(
        "kpiPersonasUnicas",
        contarDNIsUnicos(data).toLocaleString("es-PE")
    );

    setTexto(
        "kpiTotalDescansos",
        data.length.toLocaleString("es-PE")
    );

}

// =====================================================
// GRÁFICOS
// =====================================================
function renderGraficoDescansosPorMes(data) {

    destruirChart("meses");

    const canvas = document.getElementById("graficoDescansosPorMes");

    if (!canvas) return;

    const mesesDisponibles = obtenerMesesDisponibles(data);

    const labels = mesesDisponibles.map(
        mesNumero => MESES_ORDEN[mesNumero - 1]
    );

    const valores = mesesDisponibles.map(mesNumero =>
        data.filter(item => item.mesNumero === mesNumero).length
    );

    estado.charts.meses = crearGraficoBarras({
        canvas,
        labels,
        valores,
        label: "Descansos médicos",
        horizontal: false,
        onClickItem: index => {

            const mesNumero = mesesDisponibles[index];

            const detalle = data.filter(item =>
                item.mesNumero === mesNumero
            );

            abrirModalDetalleGeneral({
                titulo: "DESCANSOS MÉDICOS POR MES",
                subtitulo: MESES_ORDEN[mesNumero - 1],
                data: detalle,
                mostrarSubtipos: false
            });

        }
    });

}

function renderGraficoCondicionPorMes(data) {

    destruirChart("condicionMes");

    const canvas = document.getElementById("graficoCondicionPorMes");

    if (!canvas) return;

    const mesesDisponibles = obtenerMesesDisponibles(data);

    const labels = mesesDisponibles.map(
        mesNumero => MESES_ORDEN[mesNumero - 1]
    );

    const condiciones = [...new Set(
        data
            .map(item => item.nomCondicion || "NO ESPECIFICADO")
            .filter(Boolean)
    )].sort((a, b) => a.localeCompare(b, "es"));

    const datasets = condiciones.map((condicion, index) => {

        const valores = mesesDisponibles.map(mesNumero =>
            data.filter(item =>
                item.mesNumero === mesNumero &&
                (item.nomCondicion || "NO ESPECIFICADO") === condicion
            ).length
        );

        return {
            label: condicion,
            data: valores,
            borderWidth: 3,
            tension: 0.35,
            fill: false,
            pointRadius: 4,
            pointHoverRadius: 6,
            pointBorderWidth: 2,
            borderColor: COLORES_LINEAS[index % COLORES_LINEAS.length],
            backgroundColor: COLORES_LINEAS[index % COLORES_LINEAS.length],
            pointBackgroundColor: COLORES_LINEAS[index % COLORES_LINEAS.length],
            pointBorderColor: "#ffffff"
        };

    });

    estado.charts.condicionMes = new Chart(canvas, {

        type: "line",

        data: {
            labels,
            datasets
        },

        options: {

            responsive: true,
            maintainAspectRatio: false,

            animation: {
                duration: 650
            },

            interaction: {
                mode: "nearest",
                intersect: false
            },

            hover: {
                mode: "nearest",
                intersect: false
            },

            onClick: (event, elements) => {

                if (!elements.length) return;

                const element = elements[0];

                const condicion = datasets[element.datasetIndex].label;
                const mesNumero = mesesDisponibles[element.index];

                const detalle = data.filter(item =>
                    item.mesNumero === mesNumero &&
                    (item.nomCondicion || "NO ESPECIFICADO") === condicion
                );

                abrirModalDetalleGeneral({
                    titulo: "CONDICIÓN LABORAL POR MES",
                    subtitulo: `${condicion} | ${MESES_ORDEN[mesNumero - 1]}`,
                    data: detalle,
                    mostrarSubtipos: false
                });

            },

            layout: {
                padding: {
                    right: 18,
                    left: 8,
                    top: 18,
                    bottom: 8
                }
            },

            plugins: {

                legend: {
                    display: true,
                    position: "bottom",
                    labels: {
                        color: "#475569",
                        font: {
                            size: 11,
                            weight: "700"
                        },
                        boxWidth: 12,
                        padding: 14
                    }
                },

                tooltip: {
                    enabled: true,
                    intersect: false,
                    mode: "nearest",
                    backgroundColor: "rgba(15,23,42,0.96)",
                    padding: 14,
                    titleColor: "#ffffff",
                    bodyColor: "#ffffff",
                    displayColors: true,
                    callbacks: {
                        title: items => items.length ? items[0].label : "",
                        label: ctx => `${ctx.dataset.label}: ${ctx.raw}`
                    }
                },

                datalabels: {
                    display: true,
                    color: "#334155",
                    anchor: "end",
                    align: "top",
                    offset: 4,
                    clamp: true,
                    clip: false,
                    font: {
                        weight: "bold",
                        size: 10
                    },
                    formatter: value => value > 0 ? value : ""
                }

            },

            scales: {

                x: {
                    ticks: {
                        color: "#475569",
                        font: {
                            size: 11,
                            weight: "700"
                        }
                    },
                    grid: {
                        color: "rgba(0,0,0,0.04)",
                        drawBorder: false
                    }
                },

                y: {
                    beginAtZero: true,
                    ticks: {
                        color: "#475569",
                        precision: 0,
                        font: {
                            size: 11,
                            weight: "700"
                        }
                    },
                    grid: {
                        color: "rgba(0,0,0,0.06)",
                        drawBorder: false
                    }
                }

            }

        },

        plugins: obtenerPlugins()

    });

}

function renderGraficoLineasMesesPorDia(data) {

    destruirChart("lineasMesesPorDia");

    const canvas = document.getElementById("graficoLineasMesesPorDia");

    if (!canvas) return;

    const mesesDisponibles = obtenerMesesDisponibles(data);

    const labels = Array.from(
        { length: 31 },
        (_, index) => String(index + 1)
    );

    const datasets = mesesDisponibles.map((mesNumero, index) => {

        const valores = labels.map(diaTexto => {

            const dia = Number(diaTexto);

            return data.filter(item =>
                item.mesNumero === mesNumero &&
                item.diaMes === dia
            ).length;

        });

        return {
            label: MESES_ORDEN[mesNumero - 1],
            data: valores,
            borderWidth: 3,
            tension: 0.35,
            fill: false,
            pointRadius: 3,
            pointHoverRadius: 6,
            pointBorderWidth: 2,
            borderColor: COLORES_LINEAS[index % COLORES_LINEAS.length],
            backgroundColor: COLORES_LINEAS[index % COLORES_LINEAS.length],
            pointBackgroundColor: COLORES_LINEAS[index % COLORES_LINEAS.length],
            pointBorderColor: "#ffffff",
            mesNumero
        };

    });

    estado.charts.lineasMesesPorDia = new Chart(canvas, {

        type: "line",

        data: {
            labels,
            datasets
        },

        options: {

            responsive: true,
            maintainAspectRatio: false,

            interaction: {
                mode: "nearest",
                intersect: false
            },

            hover: {
                mode: "nearest",
                intersect: false
            },

            onClick: (event, elements) => {

                if (!elements.length) return;

                const element = elements[0];

                const dataset = datasets[element.datasetIndex];

                const dia = Number(labels[element.index]);

                const detalle = data.filter(item =>
                    item.mesNumero === dataset.mesNumero &&
                    item.diaMes === dia
                );

                abrirModalDetalleGeneral({
                    titulo: "EVOLUCIÓN DIARIA DE DESCANSOS MÉDICOS",
                    subtitulo: `${dataset.label} | Día ${dia}`,
                    data: detalle,
                    mostrarSubtipos: false
                });

            },

            layout: {
                padding: {
                    top: 20,
                    right: 22,
                    bottom: 12,
                    left: 10
                }
            },

            plugins: {

                legend: {
                    display: true,
                    position: "bottom",
                    labels: {
                        color: "#475569",
                        font: {
                            size: 11,
                            weight: "700"
                        },
                        boxWidth: 12,
                        padding: 14
                    }
                },

                tooltip: {
                    enabled: true,
                    intersect: false,
                    mode: "nearest",
                    backgroundColor: "rgba(15,23,42,0.96)",
                    padding: 14,
                    titleColor: "#ffffff",
                    bodyColor: "#ffffff",
                    displayColors: true,
                    callbacks: {
                        title: items => items.length ? `Día ${items[0].label}` : "",
                        label: ctx => `${ctx.dataset.label}: ${ctx.raw} descansos médicos`
                    }
                },

                datalabels: {
                    display: true,
                    color: "#334155",
                    anchor: "end",
                    align: "top",
                    offset: 4,
                    font: {
                        weight: "bold",
                        size: 9
                    },
                    formatter: value => value > 0 ? value : ""
                }

            },

            scales: {

                x: {
                    ticks: {
                        color: "#475569",
                        font: {
                            size: 10,
                            weight: "700"
                        }
                    },
                    grid: {
                        color: "rgba(0,0,0,0.04)",
                        drawBorder: false
                    }
                },

                y: {
                    beginAtZero: true,
                    ticks: {
                        color: "#475569",
                        precision: 0,
                        font: {
                            size: 11,
                            weight: "700"
                        }
                    },
                    grid: {
                        color: "rgba(0,0,0,0.06)",
                        drawBorder: false
                    }
                }

            }

        },

        plugins: obtenerPlugins()

    });

}

function renderGraficoTipoDocumento(data) {

    destruirChart("tipoDocumento");

    const canvas = document.getElementById("graficoTipoDocumento");

    if (!canvas) return;

    const conteo = contarPorCampo(data, "tipoDocumento");

    const labels = conteo.map(item => item.label);
    const valores = conteo.map(item => item.valor);

    renderTotalesTipoDocumento(conteo);

    estado.charts.tipoDocumento = crearGraficoTorta({
        canvas,
        labels,
        valores,
        label: "Tipo de documento",
        onClickItem: index => {

            const tipo = labels[index];

            const detalle = data.filter(item =>
                (item.tipoDocumento || "NO ESPECIFICADO") === tipo
            );

            abrirModalDetalleGeneral({
                titulo: "DESCANSOS MÉDICOS POR TIPO DE DOCUMENTO",
                subtitulo: tipo,
                data: detalle,
                mostrarSubtipos: false
            });

        }
    });

}

function renderTotalesTipoDocumento(ordenado) {

    const contenedor = document.getElementById("totalTipoDocumento");

    if (!contenedor) return;

    if (!ordenado.length) {

        contenedor.innerHTML = `
            <div class="tipo-total-item">
                <strong>0</strong>
            </div>
        `;

        return;

    }

    contenedor.innerHTML = ordenado.map(item => `
        <div class="tipo-total-item">
            <strong>${Number(item.valor || 0).toLocaleString("es-PE")}</strong>
        </div>
    `).join("");

}

function renderGraficoContingencia(data) {

    destruirChart("contingencia");

    const canvas = document.getElementById("graficoContingencia");

    if (!canvas) return;

    const conteo = contarPorCampo(data, "contingencia");

    const labels = conteo.map(item => item.label);
    const valores = conteo.map(item => item.valor);

    estado.charts.contingencia = crearGraficoBarras({
        canvas,
        labels,
        valores,
        label: "Descansos médicos",
        horizontal: true,
        onClickItem: index => {

            const contingencia = labels[index];

            const detalle = data.filter(item =>
                (item.contingencia || "NO ESPECIFICADO") === contingencia
            );

            abrirModalDetalleGeneral({
                titulo: "DESCANSOS MÉDICOS POR CONTINGENCIA",
                subtitulo: contingencia,
                data: detalle,
                mostrarSubtipos: true
            });

        }
    });

}

function renderGraficoTopPersonas(data) {

    destruirChart("topPersonas");

    const canvas = document.getElementById("graficoTopPersonas");

    if (!canvas) return;

    const texto = estado.busquedaPersonaRanking;

    const dataFiltradaNombre = texto
        ? data.filter(item =>
            item.nombre.includes(texto) ||
            item.numdocumento.includes(texto)
        )
        : data;

    const conteo = contarDescansosPorPersona(dataFiltradaNombre);

    const labels = conteo.map(item => item.label);
    const valores = conteo.map(item => item.valor);

    ajustarAlturaRanking(canvas, labels.length);

    estado.charts.topPersonas = crearGraficoBarras({
        canvas,
        labels,
        valores,
        label: "Descansos médicos",
        horizontal: true,
        ranking: true,
        onClickItem: index => {

            const dni = conteo[index].dni;

            const detalle = dataFiltradaNombre.filter(item =>
                item.numdocumento === dni
            );

            abrirModalDetalleGeneral({
                titulo: "PERSONAS CON MÁS DESCANSOS MÉDICOS SOLICITADOS",
                subtitulo: conteo[index].label,
                data: detalle,
                mostrarSubtipos: false
            });

        }
    });

}

function renderGraficoCargo(data) {

    destruirChart("cargo");

    const canvas = document.getElementById("graficoCargo");

    if (!canvas) return;

    const conteo = contarPorCampo(data, "cargo");

    const labels = conteo.map(item => item.label);
    const valores = conteo.map(item => item.valor);

    estado.charts.cargo = crearGraficoBarras({
        canvas,
        labels,
        valores,
        label: "Descansos médicos",
        horizontal: true,
        onClickItem: index => {

            const cargo = labels[index];

            const detalle = data.filter(item =>
                (item.cargo || "NO ESPECIFICADO") === cargo
            );

            abrirModalDetalleGeneral({
                titulo: "DESCANSOS MÉDICOS POR CARGO",
                subtitulo: cargo,
                data: detalle,
                mostrarSubtipos: false
            });

        }
    });

}

// =====================================================
// MATRIZ HORIZONTAL DE ALERTAS 25 A 30 DÍAS
// =====================================================
function renderMatrizAlertaDescansos(data) {

    const tbody = document.getElementById("tablaMatrizAlertaDescansos");

    if (!tbody) return;

    const matriz = construirMatrizAlertaDescansos(data, 25, 30);

    if (!matriz.length) {

        tbody.innerHTML = `
            <tr>
                <td colspan="14" class="empty-cell">
                    No se encontraron personas con 25 a 30 días acumulados en el mes según los filtros aplicados.
                </td>
            </tr>
        `;

        return;

    }

    tbody.innerHTML = matriz.map(item => {

        const celdasMeses = Array.from({ length: 12 }, (_, index) => {

            const mesNumero = index + 1;

            const celda = item.meses[mesNumero];

            if (!celda) {

                return `
                    <td class="matriz-mes-cell matriz-mes-cell--vacia">
                        -
                    </td>
                `;

            }

            return `
                <td
                    class="matriz-mes-cell ${escapeHtml(celda.clase)}"
                    data-dni="${escapeHtml(item.dni)}"
                    data-mes="${mesNumero}"
                    title="${escapeHtml(item.nombre)} | ${MESES_ORDEN[index]} | ${celda.dias} días | ${escapeHtml(celda.estado)}"
                >
                    <button
                        type="button"
                        class="matriz-mes-btn"
                        data-dni="${escapeHtml(item.dni)}"
                        data-mes="${mesNumero}"
                    >
                        <strong>${celda.dias}</strong>
                        <span>${escapeHtml(celda.estadoCorto)}</span>
                    </button>
                </td>
            `;

        }).join("");

        return `
            <tr>
                <td class="matriz-persona-cell">
                    <strong>${escapeHtml(item.nombre)}</strong>
                    <span>DNI: ${escapeHtml(item.dni)}</span>
                </td>

                ${celdasMeses}

                <td class="matriz-estado-cell">
                    <span class="matriz-badge ${escapeHtml(item.claseEstado)}">
                        ${escapeHtml(item.estadoGeneral)}
                    </span>
                    <small>
                        ${escapeHtml(item.resumen)}
                    </small>
                </td>
            </tr>
        `;

    }).join("");

    tbody.querySelectorAll(".matriz-mes-btn").forEach(btn => {

        btn.addEventListener("click", () => {

            const dni = btn.dataset.dni;
            const mesNumero = Number(btn.dataset.mes);

            const persona = matriz.find(item => item.dni === dni);

            if (!persona) return;

            const detalle = data.filter(item =>
                item.numdocumento === dni &&
                item.mesNumero === mesNumero
            );

            abrirModalDetalleGeneral({
                titulo: "ALERTA DE DESCANSO MÉDICO PROLONGADO",
                subtitulo: `${persona.nombre} | ${MESES_ORDEN[mesNumero - 1]} | ${persona.meses[mesNumero].dias} días acumulados`,
                data: detalle,
                mostrarSubtipos: false
            });

        });

    });

}

function construirMatrizAlertaDescansos(data, minimo = 25, maximo = 30) {

    const mapaPersonas = new Map();

    data.forEach(item => {

        if (!item.numdocumento || !item.mesNumero) return;

        const dni = item.numdocumento;

        if (!mapaPersonas.has(dni)) {

            mapaPersonas.set(dni, {
                dni,
                nombre: item.nombre,
                meses: {}
            });

        }

        const persona = mapaPersonas.get(dni);

        if (!persona.meses[item.mesNumero]) {

            persona.meses[item.mesNumero] = {
                dias: 0,
                cantidadDescansos: 0,
                registros: []
            };

        }

        persona.meses[item.mesNumero].dias += Number(item.numeroDias || 0);
        persona.meses[item.mesNumero].cantidadDescansos += 1;
        persona.meses[item.mesNumero].registros.push(item);

    });

    const resultado = [];

    mapaPersonas.forEach(persona => {

        const mesesAlertados = Object.entries(persona.meses)
            .map(([mesNumero, info]) => ({
                mesNumero: Number(mesNumero),
                dias: Number(info.dias || 0),
                cantidadDescansos: Number(info.cantidadDescansos || 0)
            }))
            .filter(item =>
                item.dias >= minimo &&
                item.dias <= maximo
            )
            .sort((a, b) => a.mesNumero - b.mesNumero);

        if (!mesesAlertados.length) return;

        const secuencias = obtenerSecuenciasConsecutivas(mesesAlertados.map(item => item.mesNumero));

        const mesesEnSecuenciaDos = new Set();
        const mesesEnSecuenciaTres = new Set();

        secuencias.forEach(secuencia => {

            if (secuencia.length >= 3) {

                secuencia.forEach(mes => {
                    mesesEnSecuenciaTres.add(mes);
                });

            } else if (secuencia.length === 2) {

                secuencia.forEach(mes => {
                    mesesEnSecuenciaDos.add(mes);
                });

            }

        });

        const mesesResultado = {};

        mesesAlertados.forEach(item => {

            let estado = "Mensual";
            let estadoCorto = "Mensual";
            let clase = "matriz-mes-cell--mensual";

            if (mesesEnSecuenciaTres.has(item.mesNumero)) {

                estado = "Crítico: 3 o más meses seguidos";
                estadoCorto = "Crítico";
                clase = "matriz-mes-cell--critico";

            } else if (mesesEnSecuenciaDos.has(item.mesNumero)) {

                estado = "Seguido: 2 meses seguidos";
                estadoCorto = "Seguido";
                clase = "matriz-mes-cell--seguido";

            } else if (mesesAlertados.length >= 2) {

                estado = "No seguido";
                estadoCorto = "No seguido";
                clase = "matriz-mes-cell--no-seguido";

            }

            mesesResultado[item.mesNumero] = {
                dias: item.dias,
                cantidadDescansos: item.cantidadDescansos,
                estado,
                estadoCorto,
                clase
            };

        });

        const mayorSecuencia = secuencias.reduce(
            (max, secuencia) => Math.max(max, secuencia.length),
            0
        );

        let estadoGeneral = "Mensual";
        let claseEstado = "matriz-badge--mensual";

        if (mayorSecuencia >= 3) {

            estadoGeneral = "Crítico";
            claseEstado = "matriz-badge--critico";

        } else if (mayorSecuencia === 2) {

            estadoGeneral = "Seguido";
            claseEstado = "matriz-badge--seguido";

        } else if (mesesAlertados.length >= 2) {

            estadoGeneral = "No seguido";
            claseEstado = "matriz-badge--no-seguido";

        }

        const mesesTexto = mesesAlertados
            .map(item => `${MESES_CORTOS[item.mesNumero - 1]}: ${item.dias}`)
            .join(" | ");

        resultado.push({
            dni: persona.dni,
            nombre: persona.nombre,
            meses: mesesResultado,
            mesesAlertados,
            mayorSecuencia,
            estadoGeneral,
            claseEstado,
            resumen: mesesTexto
        });

    });

    return resultado.sort((a, b) => {

        const prioridad = {
            "Crítico": 1,
            "Seguido": 2,
            "No seguido": 3,
            "Mensual": 4
        };

        const prioridadA = prioridad[a.estadoGeneral] || 99;
        const prioridadB = prioridad[b.estadoGeneral] || 99;

        if (prioridadA !== prioridadB) {
            return prioridadA - prioridadB;
        }

        if (b.mayorSecuencia !== a.mayorSecuencia) {
            return b.mayorSecuencia - a.mayorSecuencia;
        }

        if (b.mesesAlertados.length !== a.mesesAlertados.length) {
            return b.mesesAlertados.length - a.mesesAlertados.length;
        }

        return a.nombre.localeCompare(b.nombre, "es");

    });

}

function obtenerSecuenciasConsecutivas(meses) {

    if (!meses.length) return [];

    const ordenados = [...meses].sort((a, b) => a - b);

    const secuencias = [];

    let actual = [ordenados[0]];

    for (let i = 1; i < ordenados.length; i++) {

        if (ordenados[i] === ordenados[i - 1] + 1) {

            actual.push(ordenados[i]);

        } else {

            if (actual.length >= 2) {
                secuencias.push(actual);
            }

            actual = [ordenados[i]];

        }

    }

    if (actual.length >= 2) {
        secuencias.push(actual);
    }

    return secuencias;

}

function renderGraficoHospitales(data) {

    destruirChart("hospitales");

    const canvas = document.getElementById("graficoHospitales");

    if (!canvas) return;

    const conteo = contarPorCampo(data, "establecimientoSalud");

    const labels = conteo.map(item => item.label);
    const valores = conteo.map(item => item.valor);

    ajustarAlturaGraficoHorizontal(
        canvas,
        labels.length,
        "chart-container--hospital"
    );

    estado.charts.hospitales = crearGraficoBarras({
        canvas,
        labels,
        valores,
        label: "Descansos médicos",
        horizontal: true,
        onClickItem: index => {

            const hospital = labels[index];

            const detalle = data.filter(item =>
                (item.establecimientoSalud || "NO ESPECIFICADO") === hospital
            );

            abrirModalDetalleGeneral({
                titulo: "HOSPITALES / ESTABLECIMIENTOS MÁS FRECUENTES",
                subtitulo: hospital,
                data: detalle,
                mostrarSubtipos: false
            });

        }
    });

}

// =====================================================
// MODAL DETALLE GENERAL
// =====================================================
function abrirModalDetalleGeneral({
    titulo,
    subtitulo,
    data,
    mostrarSubtipos = false
}) {

    const modal = document.getElementById("modalDetalleGeneral");

    if (!modal) return;

    const dataOrdenada = ordenarCronologico(data);

    setTexto("modalDetalleTitulo", titulo || "DETALLE DE DESCANSOS MÉDICOS");
    setTexto("modalDetalleSubtitulo", subtitulo || "Registros ordenados cronológicamente");

    renderResumenGeneral(dataOrdenada);

    renderTablaDetalleGeneral(dataOrdenada);

    if (mostrarSubtipos) {
        renderSubtiposContingencia(dataOrdenada);
    } else {
        ocultarSubtiposContingencia();
    }

    modal.classList.add("modal-dashboard--show");

}

function renderResumenGeneral(data) {

    const personas = contarDNIsUnicos(data);

    const totalDias = data.reduce(
        (acc, item) => acc + Number(item.numeroDias || 0),
        0
    );

    setTexto("modalResumenRegistros", data.length.toLocaleString("es-PE"));
    setTexto("modalResumenPersonas", personas.toLocaleString("es-PE"));
    setTexto("modalResumenDias", totalDias.toLocaleString("es-PE"));

}
function renderTablaDetalleGeneral(data) {

    const tbody = document.getElementById(
        "tablaDetalleGeneral"
    );

    if (!tbody) return;

    setTexto(
        "contadorTablaDetalleGeneral",
        `${data.length.toLocaleString("es-PE")} registros`
    );

    setTexto(
        "tituloTablaDetalleGeneral",
        "Relación cronológica"
    );

    if (!data.length) {

        tbody.innerHTML = `
            <tr>
                <td colspan="10" class="empty-cell">
                    No se encontraron registros.
                </td>
            </tr>
        `;

        return;

    }

    tbody.innerHTML = data.map((item, index) => `
        <tr>
            <td>
                ${index + 1}
            </td>

            <td>
                ${escapeHtml(item.nombre || "-")}
            </td>

            <td>
                ${escapeHtml(item.fechaInicioTexto || "-")}
            </td>

            <td>
                ${escapeHtml(item.fechaFinTexto || "-")}
            </td>

            <td>
                ${Number(item.numeroDias || 0)}
            </td>

            <td>
                ${escapeHtml(item.tipoDocumento || "-")}
            </td>

            <td>
                ${escapeHtml(item.contingencia || "-")}
            </td>

            <td>
                ${escapeHtml(item.diagnostico || "-")}
            </td>

            <td>
                ${escapeHtml(item.establecimientoSalud || "-")}
            </td>

            <td>
                ${
                    item.marcaPosterior
                        ? `
                            <span
                                class="badge-dia-ausencia"
                                title="Marca registrada el ${escapeHtml(item.fechaPosteriorTexto || "")}"
                            >
                                ${escapeHtml(item.marcaPosterior)}
                            </span>
                        `
                        : "-"
                }
            </td>
        </tr>
    `).join("");

}

function renderSubtiposContingencia(data) {

    const bloque = document.getElementById("bloqueSubtiposContingencia");
    const lista = document.getElementById("listaSubtiposContingencia");

    if (!bloque || !lista) return;

    bloque.style.display = "block";

    const conteo = {};

    data.forEach(item => {

        const contingencia = item.contingencia || "NO ESPECIFICADO";
        const diagnostico = item.diagnostico || "NO ESPECIFICADO";
        const key = `${contingencia} | ${diagnostico}`;

        conteo[key] = (conteo[key] || 0) + 1;

    });

    const ordenado = Object.entries(conteo)
        .map(([label, valor]) => ({ label, valor }))
        .sort((a, b) => b.valor - a.valor);

    if (!ordenado.length) {

        lista.innerHTML = `
            <div class="subtipo-card">
                <span>SIN DATOS</span>
                <strong>0</strong>
            </div>
        `;

        return;

    }

    lista.innerHTML = ordenado.map(item => `
        <button class="subtipo-card" data-subtipo="${escapeHtml(item.label)}">
            <span>${escapeHtml(item.label)}</span>
            <strong>${item.valor}</strong>
        </button>
    `).join("");

    lista.querySelectorAll(".subtipo-card").forEach(btn => {

        btn.addEventListener("click", () => {

            const subtipo = btn.dataset.subtipo;

            const detalle = data.filter(item => {

                const contingencia = item.contingencia || "NO ESPECIFICADO";
                const diagnostico = item.diagnostico || "NO ESPECIFICADO";
                const key = `${contingencia} | ${diagnostico}`;

                return key === subtipo;

            });

            setTexto("tituloTablaDetalleGeneral", subtipo);

            renderTablaDetalleGeneral(ordenarCronologico(detalle));

        });

    });

}

function ocultarSubtiposContingencia() {

    const bloque = document.getElementById("bloqueSubtiposContingencia");
    const lista = document.getElementById("listaSubtiposContingencia");

    if (bloque) {
        bloque.style.display = "none";
    }

    if (lista) {
        lista.innerHTML = "";
    }

    setTexto("tituloTablaDetalleGeneral", "Relación cronológica");

}

// =====================================================
// MODAL
// =====================================================
function inicializarModal() {

    const btnCerrar = document.getElementById("btnCerrarModalDetalleGeneral");

    if (btnCerrar) {

        btnCerrar.addEventListener("click", () => {
            cerrarModal("modalDetalleGeneral");
        });

    }

    const modal = document.getElementById("modalDetalleGeneral");

    if (modal) {

        modal.addEventListener("click", event => {

            if (event.target.classList.contains("modal-dashboard")) {
                cerrarModal("modalDetalleGeneral");
            }

        });

    }

    window.addEventListener("keydown", event => {

        if (event.key === "Escape") {
            cerrarModal("modalDetalleGeneral");
        }

    });

}

function cerrarModal(id) {

    const modal = document.getElementById(id);

    if (!modal) return;

    modal.classList.remove("modal-dashboard--show");

}

// =====================================================
// FACTORÍAS DE GRÁFICOS
// =====================================================
function crearGraficoBarras({
    canvas,
    labels,
    valores,
    label,
    horizontal = false,
    ranking = false,
    onClickItem = null
}) {

    return new Chart(canvas, {

        type: "bar",

        data: {

            labels,

            datasets: [{

                label,
                data: valores,
                backgroundColor: generarColoresPorValor(valores),
                borderRadius: ranking ? 5 : 10,
                borderSkipped: false,
                maxBarThickness: ranking ? 20 : horizontal ? 26 : 42,
                barPercentage: ranking ? 0.72 : 0.9,
                categoryPercentage: ranking ? 0.72 : 0.9

            }]

        },

        options: {

            responsive: true,
            maintainAspectRatio: false,
            animation: {
                duration: 600
            },
            indexAxis: horizontal ? "y" : "x",
            interaction: {
                mode: "nearest",
                intersect: false
            },
            hover: {
                mode: "nearest",
                intersect: false
            },

            onClick: (event, elements) => {

                if (!elements.length || !onClickItem) return;

                const index = elements[0].index;

                onClickItem(index, labels[index], valores[index]);

            },

            layout: {
                padding: {
                    right: horizontal ? 60 : 8,
                    left: horizontal ? 16 : 0,
                    top: 10,
                    bottom: 5
                }
            },

            plugins: {

                legend: {
                    display: false
                },

                tooltip: {
                    enabled: true,
                    intersect: false,
                    mode: "nearest",
                    backgroundColor: "rgba(15,23,42,0.96)",
                    padding: 14,
                    titleColor: "#ffffff",
                    bodyColor: "#ffffff",
                    displayColors: false,
                    callbacks: {
                        title: items => items.length ? items[0].label : "",
                        label: ctx => `${label}: ${ctx.raw}`
                    }
                },

                datalabels: {
                    display: true,
                    color: "#334155",
                    anchor: "end",
                    align: horizontal ? "right" : "top",
                    offset: horizontal ? 6 : 4,
                    clamp: false,
                    clip: false,
                    font: {
                        weight: "bold",
                        size: ranking ? 10 : horizontal ? 10 : 12
                    },
                    formatter: value => value
                }

            },

            scales: {

                x: {
                    beginAtZero: true,
                    ticks: {
                        color: "#475569",
                        precision: 0,
                        font: {
                            size: ranking ? 10 : horizontal ? 11 : 12,
                            weight: "600"
                        }
                    },
                    grid: {
                        color: "rgba(0,0,0,0.05)",
                        drawBorder: false
                    }
                },

                y: {
                    beginAtZero: true,
                    ticks: {
                        color: "#475569",
                        precision: 0,
                        autoSkip: false,
                        padding: 8,
                        font: {
                            size: ranking ? 10 : horizontal ? 11 : 12,
                            weight: "700"
                        },
                        callback: function(value) {

                            const labelValue = this.getLabelForValue(value);

                            if (ranking && labelValue.length > 40) {
                                return labelValue.substring(0, 40) + "...";
                            }

                            if (!ranking && horizontal && labelValue.length > 50) {
                                return labelValue.substring(0, 50) + "...";
                            }

                            return labelValue;

                        }
                    },
                    grid: {
                        color: "rgba(0,0,0,0.06)",
                        drawBorder: false
                    }
                }

            }

        },

        plugins: obtenerPlugins()

    });

}

function crearGraficoTorta({
    canvas,
    labels,
    valores,
    label,
    onClickItem = null
}) {

    return new Chart(canvas, {

        type: "pie",

        data: {

            labels,

            datasets: [{

                label,
                data: valores,
                backgroundColor: generarColoresPorValor(valores),
                borderWidth: 2,
                borderColor: "#ffffff"

            }]

        },

        options: {

            responsive: true,
            maintainAspectRatio: false,

            onClick: (event, elements) => {

                if (!elements.length || !onClickItem) return;

                const index = elements[0].index;

                onClickItem(index, labels[index], valores[index]);

            },

            layout: {
                padding: {
                    top: 5,
                    bottom: 5
                }
            },

            plugins: {

                legend: {
                    position: "bottom",
                    labels: {
                        color: "#475569",
                        font: {
                            size: 12,
                            weight: "600"
                        },
                        boxWidth: 14,
                        padding: 16
                    }
                },

                tooltip: {
                    backgroundColor: "rgba(15,23,42,0.96)",
                    padding: 14,
                    titleColor: "#ffffff",
                    bodyColor: "#ffffff",
                    displayColors: true,
                    callbacks: {
                        label: ctx => {

                            const total = ctx.dataset.data.reduce(
                                (acc, value) => acc + value,
                                0
                            );

                            const porcentaje = total > 0
                                ? ((ctx.raw / total) * 100).toFixed(1)
                                : "0.0";

                            return `${ctx.label}: ${ctx.raw} (${porcentaje}%)`;

                        }
                    }
                },

                datalabels: {
                    display: true,
                    color: "#ffffff",
                    font: {
                        weight: "bold",
                        size: 12
                    },
                    formatter: (value, ctx) => {

                        const total = ctx.dataset.data.reduce(
                            (acc, item) => acc + item,
                            0
                        );

                        if (!total) return "";

                        const porcentaje = (value / total) * 100;

                        return porcentaje >= 5
                            ? `${porcentaje.toFixed(1)}%`
                            : "";

                    }
                }

            }

        },

        plugins: obtenerPlugins()

    });

}

// =====================================================
// ALTURAS
// =====================================================
function ajustarAlturaRanking(canvas, cantidadRegistros) {

    const registrosVisibles = 15;
    const altoPorPersona = 38;

    const alturaVisible = registrosVisibles * altoPorPersona + 90;

    const alturaTotal = Math.max(
        alturaVisible,
        cantidadRegistros * altoPorPersona + 80
    );

    const contenedor = canvas.closest(".chart-container--ranking");
    const wrapper = canvas.closest(".ranking-canvas-wrap");

    if (contenedor) {

        contenedor.style.height = `${alturaVisible}px`;
        contenedor.style.minHeight = `${alturaVisible}px`;
        contenedor.style.overflowY = "auto";
        contenedor.style.overflowX = "hidden";

    }

    if (wrapper) {

        wrapper.style.height = `${alturaTotal}px`;
        wrapper.style.minHeight = `${alturaTotal}px`;

    }

    canvas.style.height = "100%";
    canvas.style.minHeight = "100%";

}

function ajustarAlturaGraficoHorizontal(canvas, cantidadRegistros, claseContenedor) {

    const registrosVisibles = 14;
    const altoPorItem = 42;

    const alturaVisible = registrosVisibles * altoPorItem + 100;

    const alturaTotal = Math.max(
        alturaVisible,
        cantidadRegistros * altoPorItem + 90
    );

    const contenedor = canvas.closest(`.${claseContenedor}`);
    const wrapper = canvas.closest(".hospital-canvas-wrap");

    if (contenedor) {

        contenedor.style.height = `${alturaVisible}px`;
        contenedor.style.minHeight = `${alturaVisible}px`;
        contenedor.style.overflowY = "auto";
        contenedor.style.overflowX = "hidden";

    }

    if (wrapper) {

        wrapper.style.height = `${alturaTotal}px`;
        wrapper.style.minHeight = `${alturaTotal}px`;

    }

    canvas.style.height = "100%";
    canvas.style.minHeight = "100%";

}

// =====================================================
// HELPERS DATOS
// =====================================================
function contarPorCampo(data, campo) {

    const conteo = {};

    data.forEach(item => {

        const key = item[campo] || "NO ESPECIFICADO";

        conteo[key] = (conteo[key] || 0) + 1;

    });

    return Object.entries(conteo)
        .map(([label, valor]) => ({ label, valor }))
        .sort((a, b) => b.valor - a.valor);

}

function contarDescansosPorPersona(data) {

    const mapa = new Map();

    data.forEach(item => {

        const dni = item.numdocumento;

        if (!mapa.has(dni)) {

            mapa.set(dni, {
                dni,
                nombre: item.nombre,
                label: item.nombre,
                valor: 0
            });

        }

        mapa.get(dni).valor += 1;

    });

    return [...mapa.values()]
        .sort((a, b) => b.valor - a.valor);

}

function obtenerPersonasUnicas(data) {

    const mapa = new Map();

    data.forEach(item => {

        if (!mapa.has(item.numdocumento)) {

            mapa.set(item.numdocumento, {
                dni: item.numdocumento,
                nombre: item.nombre
            });

        }

    });

    return [...mapa.values()]
        .sort((a, b) => a.nombre.localeCompare(b.nombre, "es"));

}

function contarDNIsUnicos(data) {

    return new Set(
        data
            .map(item => item.numdocumento)
            .filter(Boolean)
    ).size;

}

function obtenerMesesDisponibles(data) {

    return [...new Set(
        data
            .map(item => item.mesNumero)
            .filter(Boolean)
    )].sort((a, b) => a - b);

}

function ordenarCronologico(data) {

    return [...data].sort((a, b) => {

        const fechaA = a.fechaInicio instanceof Date
            ? a.fechaInicio.getTime()
            : 0;

        const fechaB = b.fechaInicio instanceof Date
            ? b.fechaInicio.getTime()
            : 0;

        if (fechaA !== fechaB) return fechaA - fechaB;

        return a.numdocumento.localeCompare(b.numdocumento, "es");

    });

}

// =====================================================
// HELPERS GENERALES
// =====================================================
function obtenerCampo(obj, fields) {

    for (const field of fields) {

        if (
            obj[field] !== undefined &&
            obj[field] !== null &&
            String(obj[field]).trim() !== ""
        ) {
            return obj[field];
        }

    }

    return "";

}

function normalizarTexto(texto) {

    return String(texto || "")
        .trim()
        .replace(/\s+/g, " ")
        .toUpperCase();

}

function limpiarDocumento(valor) {

    return String(valor || "")
        .trim()
        .replace(/\D/g, "");

}

function convertirFecha(valor) {

    if (!valor) return null;

    if (valor instanceof Date && !Number.isNaN(valor.getTime())) {
        return normalizarDate(valor);
    }

    if (typeof valor === "number") {

        const fecha = new Date(
            Math.round((valor - 25569) * 86400 * 1000)
        );

        return normalizarDate(fecha);

    }

    if (typeof valor === "string") {

        const limpio = valor
            .trim()
            .split(" ")[0]
            .replaceAll("-", "/");

        const partes = limpio.split("/");

        if (partes.length !== 3) return null;

        const dia = Number(partes[0]);
        const mes = Number(partes[1]);

        let anio = Number(partes[2]);

        if (!dia || !mes || !anio) return null;

        if (anio < 100) {
            anio = 2000 + anio;
        }

        return new Date(anio, mes - 1, dia);

    }

    return null;

}

function normalizarDate(fecha) {

    return new Date(
        fecha.getFullYear(),
        fecha.getMonth(),
        fecha.getDate()
    );

}

function formatearFecha(fecha) {

    if (
        !(fecha instanceof Date) ||
        Number.isNaN(fecha.getTime())
    ) {
        return "";
    }

    return fecha.toLocaleDateString("es-PE", {
        day: "2-digit",
        month: "2-digit",
        year: "numeric"
    });

}

function setTexto(id, valor) {

    const el = document.getElementById(id);

    if (el) {
        el.textContent = valor;
    }

}

function mostrarLoading(show) {

    const overlay = document.getElementById("loadingSistema");

    if (!overlay) return;

    overlay.style.display = show ? "flex" : "none";

}

function destruirChart(key) {

    if (estado.charts[key]) {

        estado.charts[key].destroy();

        estado.charts[key] = null;

    }

}

function obtenerPlugins() {

    return window.ChartDataLabels
        ? [ChartDataLabels]
        : [];

}

function generarColoresPorValor(valores) {

    if (!Array.isArray(valores) || !valores.length) {
        return [];
    }

    const max = Math.max(...valores);

    return valores.map(valor => {

        const ratio = max > 0 ? valor / max : 0;

        if (ratio >= 0.90) return "#E78B8B";
        if (ratio >= 0.80) return "#F0A36B";
        if (ratio >= 0.70) return "#F6BF5B";
        if (ratio >= 0.60) return "#E4D86F";
        if (ratio >= 0.50) return "#C9D97B";
        if (ratio >= 0.40) return "#A9D18E";
        if (ratio >= 0.30) return "#7FC8A9";
        if (ratio >= 0.20) return "#66BFBF";

        return "#8EC5E8";

    });

}

function escapeHtml(texto) {

    return String(texto || "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");

}
// =====================================================
// CARGAR AUTOMÁTICAMENTE LOS ROLES HASTA EL MES ACTUAL
// =====================================================
async function cargarArchivosAsistenciaDisponibles() {

    const fechaActual = new Date();

    const mesActual =
        fechaActual.getMonth() + 1;

    const tareas = [];

    console.log(
        `Cargando roles de servicio desde enero hasta ${MESES_ORDEN[mesActual - 1]}`
    );

    for (
        let mesNumero = 1;
        mesNumero <= mesActual;
        mesNumero++
    ) {

        tareas.push(
            cargarArchivoAsistenciaMes(
                mesNumero
            )
        );

    }

    await Promise.all(
        tareas
    );

    console.log(
        "✅ Meses de asistencia cargados:",
        [...estado.asistencia.mesesCargados]
            .sort((a, b) => a - b)
            .map(mesNumero =>
                MESES_ORDEN[mesNumero - 1]
            )
    );

    console.log(
        "⚠️ Meses no encontrados:",
        [...estado.asistencia.mesesNoEncontrados]
            .sort((a, b) => a - b)
            .map(mesNumero =>
                MESES_ORDEN[mesNumero - 1]
            )
    );

}


// =====================================================
// CARGAR ROL DE SERVICIO MENSUAL DESDE FIREBASE STORAGE
// =====================================================
async function cargarArchivoAsistenciaMes(mesNumero) {

    if (
        estado.asistencia.mesesCargados.has(
            mesNumero
        )
    ) {

        return estado.asistencia.archivosPorMes.get(
            mesNumero
        );

    }

    const nombresArchivo =
        ARCHIVOS_ASISTENCIA[mesNumero] || [];

    for (const nombreArchivo of nombresArchivo) {

        const rutaStorage =
            `${RUTA_ROL_SERVICIO}/${nombreArchivo}`;

        try {

            console.log(
                `Descargando rol de servicio: ${rutaStorage}`
            );

            const buffer =
                await descargarExcelStorage(
                    rutaStorage
                );

            if (
                !(buffer instanceof ArrayBuffer) ||
                buffer.byteLength === 0
            ) {

                console.warn(
                    `El archivo está vacío o no es válido: ${rutaStorage}`
                );

                continue;

            }

            const registros =
                procesarArchivoAsistencia(
                    buffer,
                    mesNumero,
                    rutaStorage
                );

            estado.asistencia.archivosPorMes.set(
                mesNumero,
                registros
            );

            estado.asistencia.mesesCargados.add(
                mesNumero
            );

            estado.asistencia.mesesNoEncontrados.delete(
                mesNumero
            );

            console.log(
                `✅ Archivo cargado: ${rutaStorage}`
            );

            console.log(
                `Personas encontradas: ${registros.size}`
            );

            return registros;

        } catch (error) {

            console.warn(
                `No se pudo cargar ${rutaStorage}:`,
                error
            );

        }

    }

    estado.asistencia.archivosPorMes.set(
        mesNumero,
        new Map()
    );

    estado.asistencia.mesesNoEncontrados.add(
        mesNumero
    );

    console.warn(
        `No se encontró el archivo del mes: ${MESES_ORDEN[mesNumero - 1]}`
    );

    return new Map();

}

function procesarArchivoAsistencia(
    buffer,
    mesNumero,
    rutaArchivo
) {

    const workbook = XLSX.read(buffer, {
        type: "array",
        cellDates: false
    });

    const mapaPersonas = new Map();

    workbook.SheetNames.forEach(nombreHoja => {

        const hoja = workbook.Sheets[nombreHoja];

        const filas = XLSX.utils.sheet_to_json(hoja, {
            header: 1,
            defval: "",
            raw: false
        });

        if (!filas.length) return;

        const indiceCabecera =
            localizarFilaCabeceraAsistencia(filas);

        if (indiceCabecera < 0) {
            return;
        }

        const cabeceraNormalizada =
            filas[indiceCabecera].map(valor =>
                normalizarCabeceraAsistencia(valor)
            );

        const indiceDNI = buscarIndiceCabecera(
            cabeceraNormalizada,
            [
                "DNI",
                "NUMDOCUMENTO",
                "NUM DOCUMENTO",
                "DOCUMENTO"
            ]
        );

        const indiceNombre = buscarIndiceCabecera(
            cabeceraNormalizada,
            [
                "NOMBRE",
                "NOMBRES",
                "APELLIDOS Y NOMBRES",
                "PERSONAL",
                "DES NOMAPELL"
            ]
        );

        if (indiceDNI < 0) return;

        const columnasDias =
            obtenerColumnasDiasAsistencia(
                filas[indiceCabecera]
            );

        for (
            let filaIndex = indiceCabecera + 1;
            filaIndex < filas.length;
            filaIndex++
        ) {

            const fila = filas[filaIndex];

            const dni = limpiarDocumento(
                fila[indiceDNI]
            );

            if (!dni) continue;

            const nombre = indiceNombre >= 0
                ? normalizarTexto(fila[indiceNombre])
                : "";

            const dias = {};

            columnasDias.forEach(({ indice, dia }) => {

                dias[dia] = normalizarMarcaAsistencia(
                    fila[indice]
                );

            });

            const dniClave =
                normalizarDNIAsistencia(dni);

            if (!mapaPersonas.has(dniClave)) {

                mapaPersonas.set(dniClave, {
                    dni,
                    nombre,
                    mesNumero,
                    dias,
                    hoja: nombreHoja,
                    filaExcel: filaIndex + 1,
                    rutaArchivo
                });

                continue;

            }

            const existente =
                mapaPersonas.get(dniClave);

            Object.entries(dias).forEach(
                ([dia, valor]) => {

                    if (
                        valor &&
                        !existente.dias[dia]
                    ) {
                        existente.dias[dia] = valor;
                    }

                }
            );

            if (!existente.nombre && nombre) {
                existente.nombre = nombre;
            }

        }

    });

    return mapaPersonas;

}

function localizarFilaCabeceraAsistencia(filas) {

    const limite = Math.min(
        filas.length,
        30
    );

    for (let index = 0; index < limite; index++) {

        const fila = filas[index] || [];

        const normalizada = fila.map(valor =>
            normalizarCabeceraAsistencia(valor)
        );

        const tieneDNI = normalizada.some(valor =>
            [
                "DNI",
                "NUMDOCUMENTO",
                "NUM DOCUMENTO",
                "DOCUMENTO"
            ].includes(valor)
        );

        const cantidadDias = fila.filter(valor => {

            const dia =
                obtenerNumeroDiaCabecera(valor);

            return dia >= 1 && dia <= 31;

        }).length;

        if (
            tieneDNI &&
            cantidadDias >= 5
        ) {
            return index;
        }

    }

    return -1;

}

function obtenerColumnasDiasAsistencia(
    cabeceraOriginal
) {

    const columnas = [];

    cabeceraOriginal.forEach(
        (valor, indice) => {

            const dia =
                obtenerNumeroDiaCabecera(valor);

            if (
                dia >= 1 &&
                dia <= 31
            ) {

                columnas.push({
                    indice,
                    dia
                });

            }

        }
    );

    return columnas;

}

function obtenerNumeroDiaCabecera(valor) {

    const texto = String(valor ?? "")
        .trim()
        .replace(/\.0$/, "");

    if (!/^\d{1,2}$/.test(texto)) {
        return 0;
    }

    const numero = Number(texto);

    return numero >= 1 && numero <= 31
        ? numero
        : 0;

}

function buscarIndiceCabecera(
    cabecera,
    alternativas
) {

    for (const alternativa of alternativas) {

        const buscado =
            normalizarCabeceraAsistencia(
                alternativa
            );

        const indice =
            cabecera.indexOf(buscado);

        if (indice >= 0) {
            return indice;
        }

    }

    return -1;

}

function normalizarCabeceraAsistencia(valor) {

    return String(valor ?? "")
        .trim()
        .replace(/[_-]+/g, " ")
        .replace(/\s+/g, " ")
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .toUpperCase();

}

function normalizarMarcaAsistencia(valor) {

    return String(valor ?? "")
        .trim()
        .replace(/\s+/g, "")
        .toUpperCase();

}

function normalizarDNIAsistencia(valor) {

    const documento = limpiarDocumento(valor);

    if (!documento) return "";

    return documento.replace(/^0+/, "") || "0";

}

// =====================================================
// DESCANSO MÉDICO ANTES DE UNA AUSENCIA
// =====================================================
function renderDescansosAntesAusencia(data) {

    const tbody = document.getElementById(
        "tablaDescansosAntesAusencia"
    );

    if (!tbody) return;

    const resultados =
        detectarDescansosAntesAusencia(data);

    estado.asistencia.resultados = resultados;

    setTexto(
        "totalDescansosAntesAusencia",
        resultados.length.toLocaleString("es-PE")
    );

    actualizarEstadoArchivosAsistencia(data);

    if (!resultados.length) {

        tbody.innerHTML = `
            <tr>
                <td colspan="6" class="empty-cell">
                    No se encontraron personas que registren una marca
                    distinta de T1, T2 o T3 inmediatamente después de
                    finalizar su descanso médico.
                </td>
            </tr>
        `;

        return;

    }

    tbody.innerHTML = resultados.map(
        (resultado, index) => `
            <tr
                class="fila-ausencia-dm"
                data-indice="${index}"
                title="Haz clic para ver el detalle"
            >
                <td>
                    ${escapeHtml(resultado.dni)}
                </td>

                <td>
                    <strong>
                        ${escapeHtml(resultado.nombre)}
                    </strong>
                </td>

                <td>
                    ${escapeHtml(resultado.fechaInicioTexto)}
                </td>

                <td>
                    ${escapeHtml(resultado.fechaFinTexto)}
                </td>

                <td>
                    <span
                        class="badge-dia-ausencia"
                        title="Marca registrada: ${escapeHtml(resultado.marcaPosterior)}"
                    >
                        ${escapeHtml(resultado.fechaPosteriorTexto)}
                        ·
                        ${escapeHtml(resultado.marcaPosterior)}
                    </span>
                </td>

                <td>
                    ${escapeHtml(resultado.mesPosterior)}
                </td>
            </tr>
        `
    ).join("");

    tbody
        .querySelectorAll(".fila-ausencia-dm")
        .forEach(fila => {

            fila.addEventListener(
                "click",
                () => {

                    const indice = Number(
                        fila.dataset.indice
                    );

                    const resultado =
                        resultados[indice];

                    if (!resultado) return;

                    abrirModalDetalleGeneral({

                        titulo:
                            "DESCANSO MÉDICO ANTES DE UNA MARCA DIFERENTE",

                        subtitulo:
                            `${resultado.nombre} | ` +
                            `DM hasta ${resultado.fechaFinTexto} | ` +
                            `${resultado.fechaPosteriorTexto}: ` +
                            `${resultado.marcaPosterior}`,

                        data: [
                            resultado.registroDM
                        ],

                        mostrarSubtipos: false
                    });

                }
            );

        });

}
function detectarDescansosAntesAusencia(data) {

    const resultados = [];

    const clavesAgregadas = new Set();

    // Únicas marcas consideradas turnos normales
    const TURNOS_VALIDOS = new Set([
        "T1",
        "T2",
        "T3"
    ]);

    data.forEach(descanso => {

        if (
            !descanso.numdocumento ||
            !(descanso.fechaFin instanceof Date) ||
            Number.isNaN(descanso.fechaFin.getTime())
        ) {
            return;
        }

        // Se analiza el día inmediatamente posterior
        // al término del descanso médico
        const fechaPosterior = sumarDiasFecha(
            descanso.fechaFin,
            1
        );

        const mesPosteriorNumero =
            fechaPosterior.getMonth() + 1;

        const diaPosterior =
            fechaPosterior.getDate();

        const archivoMes =
            estado.asistencia.archivosPorMes.get(
                mesPosteriorNumero
            );

        if (
            !(archivoMes instanceof Map) ||
            !archivoMes.size
        ) {
            return;
        }

        const dniClave =
            normalizarDNIAsistencia(
                descanso.numdocumento
            );

        const personaAsistencia =
            archivoMes.get(dniClave);

        if (!personaAsistencia) {
            return;
        }

        const marcaPosterior =
            normalizarMarcaAsistencia(
                personaAsistencia.dias[diaPosterior]
            );

        // No considerar celdas vacías
        if (!marcaPosterior) {
            return;
        }

        // No considerar los turnos normales
        if (TURNOS_VALIDOS.has(marcaPosterior)) {
            return;
        }

        const clave = [
            descanso.numdocumento,
            descanso.fechaInicioTexto,
            descanso.fechaFinTexto,
            formatearFecha(fechaPosterior),
            marcaPosterior
        ].join("|");

        if (clavesAgregadas.has(clave)) {
            return;
        }

        clavesAgregadas.add(clave);

        resultados.push({

            dni: descanso.numdocumento,

            nombre:
                descanso.nombre ||
                personaAsistencia.nombre ||
                "NO ESPECIFICADO",

            fechaInicioTexto:
                descanso.fechaInicioTexto,

            fechaFinTexto:
                descanso.fechaFinTexto,

            fechaAusenciaTexto:
                formatearFecha(fechaPosterior),

            fechaPosteriorTexto:
                formatearFecha(fechaPosterior),

            marcaPosterior,

            mesAusencia:
                MESES_ORDEN[
                    mesPosteriorNumero - 1
                ],

            mesPosterior:
                MESES_ORDEN[
                    mesPosteriorNumero - 1
                ],

            registroDM:
                descanso
        });

    });

    return resultados.sort((a, b) => {

        const fechaA =
            convertirFecha(a.fechaPosteriorTexto)?.getTime() || 0;

        const fechaB =
            convertirFecha(b.fechaPosteriorTexto)?.getTime() || 0;

        if (fechaA !== fechaB) {
            return fechaA - fechaB;
        }

        return a.nombre.localeCompare(
            b.nombre,
            "es"
        );

    });

}

// =====================================================
// GRÁFICO: PERSONAL CON DOS O MÁS CASOS POSTERIORES AL DM
// =====================================================
function renderGraficoReincidentesPosteriorDM(data) {

    destruirChart("reincidentesPosteriorDM");

    const canvas = document.getElementById(
        "graficoReincidentesPosteriorDM"
    );

    if (!canvas) return;

    const resultados =
        detectarDescansosAntesAusencia(data);

    const reincidentes =
        agruparReincidentesPosteriorDM(resultados);

    setTexto(
        "totalReincidentesPosteriorDM",
        reincidentes.length.toLocaleString("es-PE")
    );

    const labels = reincidentes.map(item =>
        item.nombre
    );

    const valores = reincidentes.map(item =>
        item.cantidad
    );

    ajustarAlturaReincidentesPosteriorDM(
        canvas,
        reincidentes.length
    );

    estado.charts.reincidentesPosteriorDM =
        crearGraficoBarras({

            canvas,

            labels,

            valores,

            label:
                "Registros diferentes a T1, T2 o T3",

            horizontal:
                true,

            ranking:
                true,

            onClickItem: index => {

                const persona =
                    reincidentes[index];

                if (!persona) return;

                const registrosDM = persona.casos.map(caso => ({
                    ...caso.registroDM,

                    marcaPosterior:
                        caso.marcaPosterior,

                    fechaPosteriorTexto:
                        caso.fechaPosteriorTexto
                }));

                abrirModalDetalleGeneral({

                    titulo:
                        "PERSONAL CON DOS O MÁS DESCANSOS MÉDICOS Y REGISTRO DIFERENTE A TURNO",

                    subtitulo:
                        `${persona.nombre} | DNI: ${persona.dni} | ` +
                        `${persona.cantidad} casos detectados`,

                    data:
                        registrosDM,

                    mostrarSubtipos:
                        false

                });

            }
        });

}


// =====================================================
// AGRUPAR CASOS POSTERIORES AL DM POR DNI
// =====================================================
function agruparReincidentesPosteriorDM(resultados) {

    const mapaPersonas = new Map();

    resultados.forEach(resultado => {

        const dni = resultado.dni;

        if (!dni) return;

        if (!mapaPersonas.has(dni)) {

            mapaPersonas.set(dni, {

                dni,

                nombre:
                    resultado.nombre ||
                    "NO ESPECIFICADO",

                cantidad:
                    0,

                casos:
                    []

            });

        }

        const persona =
            mapaPersonas.get(dni);

        persona.cantidad += 1;

        persona.casos.push(resultado);

    });

    return [...mapaPersonas.values()]
        .filter(persona =>
            persona.cantidad >= 2
        )
        .sort((a, b) => {

            if (b.cantidad !== a.cantidad) {
                return b.cantidad - a.cantidad;
            }

            return a.nombre.localeCompare(
                b.nombre,
                "es"
            );

        });

}


// =====================================================
// ALTURA DEL GRÁFICO DE REINCIDENTES
// =====================================================
function ajustarAlturaReincidentesPosteriorDM(
    canvas,
    cantidadRegistros
) {

    const registrosVisibles = 12;
    const altoPorPersona = 44;

    const alturaVisible =
        registrosVisibles * altoPorPersona + 100;

    const alturaTotal = Math.max(
        alturaVisible,
        cantidadRegistros * altoPorPersona + 90
    );

    const contenedor = canvas.closest(
        ".chart-container--reincidencia-dm"
    );

    const wrapper = canvas.closest(
        ".reincidencia-dm-canvas-wrap"
    );

    if (contenedor) {

        contenedor.style.height =
            `${alturaVisible}px`;

        contenedor.style.minHeight =
            `${alturaVisible}px`;

        contenedor.style.overflowY =
            "auto";

        contenedor.style.overflowX =
            "hidden";

    }

    if (wrapper) {

        wrapper.style.height =
            `${alturaTotal}px`;

        wrapper.style.minHeight =
            `${alturaTotal}px`;

    }

    canvas.style.height = "100%";
    canvas.style.minHeight = "100%";

}
function sumarDiasFecha(fecha, cantidad) {

    const nuevaFecha = new Date(
        fecha.getFullYear(),
        fecha.getMonth(),
        fecha.getDate()
    );

    nuevaFecha.setDate(
        nuevaFecha.getDate() + cantidad
    );

    return nuevaFecha;

}

function actualizarEstadoArchivosAsistencia(data) {

    const contenedor =
        document.getElementById(
            "estadoArchivoAsistencia"
        );

    if (!contenedor) return;

    const mesesNecesarios = [
        ...new Set(
            data
                .filter(item =>
                    item.fechaFin instanceof Date
                )
                .map(item =>
                    sumarDiasFecha(
                        item.fechaFin,
                        1
                    ).getMonth() + 1
                )
        )
    ].sort((a, b) => a - b);

    if (!mesesNecesarios.length) {

        contenedor.innerHTML = `
            <i class="fa-solid fa-circle-info"></i>

            <span>
                No existen descansos médicos para analizar.
            </span>
        `;

        return;

    }

    const faltantes =
        mesesNecesarios.filter(
            mesNumero =>
                !estado.asistencia.mesesCargados.has(
                    mesNumero
                )
        );

    if (!faltantes.length) {

        const nombres =
            mesesNecesarios
                .map(mesNumero =>
                    `${MESES_ORDEN[
                        mesNumero - 1
                    ].toLowerCase()}.xlsx`
                )
                .join(", ");

        contenedor.innerHTML = `
            <i class="fa-solid fa-file-circle-check"></i>

            <span>
                Archivos analizados:
                ${escapeHtml(nombres)}
            </span>
        `;

        return;

    }

    const nombresFaltantes =
        faltantes
            .map(mesNumero =>
                `${MESES_ORDEN[
                    mesNumero - 1
                ].toLowerCase()}.xlsx`
            )
            .join(", ");

    contenedor.innerHTML = `
        <i class="fa-solid fa-triangle-exclamation"></i>

        <span>
            No se encontraron los archivos:
            ${escapeHtml(nombresFaltantes)}
        </span>
    `;

}
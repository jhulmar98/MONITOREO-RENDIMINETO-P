// =====================================================
// 📊 DASHBOARD.JS
// SISTEMA DE MONITOREO DE PAPELETAS
// LISTO PARA PRODUCCIÓN
//
// FUNCIONES:
// - Lectura de papeletas.xlsx
// - Fechas normalizadas en formato DD/MM/AAAA
// - Filtros dependientes con selección múltiple
// - Búsqueda por nombre o DNI con sugerencias
// - Gráficos interactivos
// - Toda la fila de los gráficos horizontales es clicable
// - Modal con tabla tipo Excel
// - Registros ordenados cronológicamente
// =====================================================
import {
    descargarExcelStorage
} from "../backend/api.js";

// =====================================================
// ⚙️ CONFIGURACIÓN
// =====================================================
const CONFIG = {

    rutaStorage:
        "RENDIMIENTO PERSONAL/PAPELETAS/papeletas.xlsx",

    reincidenciaCritica:
        3,

    maximoSugerencias:
        12,

    maximoElementosGrafico:
        150,

    alturaMinimaGrafico:
        380,

    alturaPorFilaGrafico:
        38,

    columnas: {

        mes: [
            "MES"
        ],

        nombre: [
            "APELLIDOS Y NOMBRES",
            "NOMBRES Y APELLIDOS",
            "NOMBRE"
        ],

        dni: [
            "DNI",
            "N° DNI",
            "Nº DNI",
            "NUMERO DNI",
            "DOCUMENTO"
        ],

        unidad: [
            "UNIDAD ORGANICA",
            "UNIDAD ORGÁNICA",
            "UNIDAD",
            "SUB GERENCIA",
            "SUBGERENCIA"
        ],

        tipo: [
            "TIPO PAPELETA",
            "TIPO DE PAPELETA",
            "TIPO"
        ],

        motivo: [
            "MOTIVO PAPELETA",
            "MOTIVO DE PAPELETA",
            "MOTIVO"
        ],

        fechaInicio: [
            "FECHA INICIO",
            "FECHA DE INICIO"
        ],

        fechaFin: [
            "FECHA FIN",
            "FECHA DE FIN"
        ],

        observacion: [
            "OBSERVACIONES",
            "OBSERVACION",
            "OBSERVACIÓN",
            "ONSERVACION"
        ]

    }

};

// =====================================================
// 📅 MESES
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
// 🌍 ESTADO GLOBAL
// =====================================================
const estado = {

    registrosOriginales: [],

    registrosFiltrados: [],

    registrosModal: [],

    personalDisponible: [],

    personaSeleccionada: null,

    filtros: {

        mes:
            new Set(),

        tipo:
            new Set(),

        motivo:
            new Set(),

        observacion:
            new Set(),

        nombre:
            ""

    },

    opciones: {

        mes: [],

        tipo: [],

        motivo: [],

        observacion: []

    },

    charts: {

        meses:
            null,

        dias:
            null,

        personas:
            null,

        motivos:
            null,

        tipos:
            null,

        observaciones:
            null

    }

};


// =====================================================
// 🎛️ CONFIGURACIÓN DE MULTISELECTS
// =====================================================
const MULTISELECTS = {

    mes: {

        boton:
            "btnFiltroMes",

        menu:
            "menuFiltroMes",

        texto:
            "textoFiltroMes",

        opciones:
            "opcionesMes",

        todos:
            "seleccionarTodosMes",

        buscador:
            "buscarMes",

        textoTodos:
            "Todos los meses",

        plural:
            "meses seleccionados"

    },

    tipo: {

        boton:
            "btnFiltroTipo",

        menu:
            "menuFiltroTipo",

        texto:
            "textoFiltroTipo",

        opciones:
            "opcionesTipo",

        todos:
            "seleccionarTodosTipo",

        buscador:
            "buscarTipo",

        textoTodos:
            "Todos los tipos",

        plural:
            "tipos seleccionados"

    },

    motivo: {

        boton:
            "btnFiltroMotivo",

        menu:
            "menuFiltroMotivo",

        texto:
            "textoFiltroMotivo",

        opciones:
            "opcionesMotivo",

        todos:
            "seleccionarTodosMotivo",

        buscador:
            "buscarMotivo",

        textoTodos:
            "Todos los motivos",

        plural:
            "motivos seleccionados"

    },

    observacion: {

        boton:
            "btnFiltroObservacion",

        menu:
            "menuFiltroObservacion",

        texto:
            "textoFiltroObservacion",

        opciones:
            "opcionesObservacion",

        todos:
            "seleccionarTodosObservacion",

        buscador:
            "buscarObservacion",

        textoTodos:
            "Todas las observaciones",

        plural:
            "observaciones seleccionadas"

    }

};


// =====================================================
// 🚀 INICIO
// =====================================================
window.addEventListener(
    "DOMContentLoaded",
    iniciarSistema
);

async function iniciarSistema() {

    try {

        mostrarLoading(
            "Inicializando sistema...",
            "Preparando filtros y gráficos"
        );

        registrarPluginChart();

        configurarMultiselects();

        configurarEventosGenerales();

        configurarBuscadorPersonal();

        configurarModal();

        
        await cargarExcel();
        construirPersonalDisponible();

        recalcularSistema();

    } catch (error) {

        console.error(
            "❌ Error al iniciar el sistema:",
            error
        );

        mostrarError(
            "No se pudo cargar el archivo de papeletas."
        );

    } finally {

        ocultarLoading();

    }

}

// =====================================================
// 📥 CARGAR PAPELETAS DESDE FIREBASE STORAGE
// =====================================================
async function cargarExcel() {

    if (typeof XLSX === "undefined") {

        throw new Error(
            "La librería XLSX no está disponible."
        );

    }

    mostrarLoading(
        "Cargando papeletas...",
        "Descargando la base desde Firebase Storage"
    );

    console.log(
        "Descargando archivo:",
        CONFIG.rutaStorage
    );

    const buffer = await descargarExcelStorage(
        CONFIG.rutaStorage
    );

    if (
        !(buffer instanceof ArrayBuffer) ||
        buffer.byteLength === 0
    ) {

        throw new Error(
            `El archivo "${CONFIG.rutaStorage}" está vacío o no es válido.`
        );

    }

    const libro = XLSX.read(
        buffer,
        {
            type: "array",
            cellDates: true,
            raw: true
        }
    );

    if (!libro.SheetNames.length) {

        throw new Error(
            "El archivo papeletas.xlsx no contiene hojas."
        );

    }

    const nombreHoja = libro.SheetNames[0];

    const hoja = libro.Sheets[nombreHoja];

    if (!hoja) {

        throw new Error(
            `No se pudo leer la hoja "${nombreHoja}".`
        );

    }

    const filas = XLSX.utils.sheet_to_json(
        hoja,
        {
            defval: "",
            raw: true,
            dateNF: "dd/mm/yyyy"
        }
    );

    estado.registrosOriginales = filas
        .map(normalizarRegistro)
        .filter(registro => {

            return (
                registro.nombre ||
                registro.dni ||
                registro.tipo ||
                registro.motivo
            );

        });

    estado.registrosFiltrados = [
        ...estado.registrosOriginales
    ];

    console.log(
        "✅ papeletas.xlsx cargado desde Firebase Storage"
    );

    console.log(
        "Ruta:",
        CONFIG.rutaStorage
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
        "Registros válidos:",
        estado.registrosOriginales.length
    );

    if (!estado.registrosOriginales.length) {

        throw new Error(
            "El archivo fue descargado, pero no se encontraron registros válidos."
        );

    }

}


// =====================================================
// 🧹 NORMALIZAR REGISTRO
// =====================================================
function normalizarRegistro(
    filaOriginal
) {

    const fila =
        normalizarClavesObjeto(
            filaOriginal
        );

    const fechaInicio =
        normalizarFecha(
            obtenerCampo(
                fila,
                CONFIG.columnas.fechaInicio
            )
        );

    const fechaFin =
        normalizarFecha(
            obtenerCampo(
                fila,
                CONFIG.columnas.fechaFin
            )
        );

    return {

        mes:
            normalizarMes(
                obtenerCampo(
                    fila,
                    CONFIG.columnas.mes
                )
            ),

        nombre:
            limpiarTexto(
                obtenerCampo(
                    fila,
                    CONFIG.columnas.nombre
                )
            ),

        dni:
            limpiarDni(
                obtenerCampo(
                    fila,
                    CONFIG.columnas.dni
                )
            ),

        unidad:
            limpiarTexto(
                obtenerCampo(
                    fila,
                    CONFIG.columnas.unidad
                )
            ),

        tipo:
            limpiarTexto(
                obtenerCampo(
                    fila,
                    CONFIG.columnas.tipo
                )
            ),

        motivo:
            limpiarTexto(
                obtenerCampo(
                    fila,
                    CONFIG.columnas.motivo
                )
            ),

        fechaInicio,

        fechaFin,

        observacion:
            limpiarTexto(
                obtenerCampo(
                    fila,
                    CONFIG.columnas.observacion
                )
            )

    };

}


// =====================================================
// 🔑 NORMALIZAR ENCABEZADOS
// =====================================================
function normalizarClavesObjeto(
    objeto
) {

    const resultado = {};

    Object.entries(
        objeto || {}
    ).forEach(
        ([clave, valor]) => {

            resultado[
                normalizarClave(clave)
            ] = valor;

        }
    );

    return resultado;

}

function normalizarClave(
    valor
) {

    return String(
        valor ?? ""
    )
        .normalize("NFD")
        .replace(
            /[\u0300-\u036f]/g,
            ""
        )
        .replace(
            /\s+/g,
            " "
        )
        .trim()
        .toUpperCase();

}

function obtenerCampo(
    objeto,
    columnas
) {

    for (
        const columna of columnas
    ) {

        const clave =
            normalizarClave(
                columna
            );

        if (
            Object.prototype.hasOwnProperty.call(
                objeto,
                clave
            )
        ) {

            const valor =
                objeto[clave];

            if (
                valor !== undefined &&
                valor !== null
            ) {

                return valor;

            }

        }

    }

    return "";

}


// =====================================================
// 🎛️ CONFIGURAR MULTISELECTS
// =====================================================
function configurarMultiselects() {

    Object.entries(
        MULTISELECTS
    ).forEach(
        ([filtro, config]) => {

            const boton =
                document.getElementById(
                    config.boton
                );

            const menu =
                document.getElementById(
                    config.menu
                );

            const seleccionarTodos =
                document.getElementById(
                    config.todos
                );

            const buscador =
                document.getElementById(
                    config.buscador
                );

            if (
                !boton ||
                !menu
            ) {

                return;

            }

            boton.addEventListener(
                "click",
                evento => {

                    evento.stopPropagation();

                    cerrarSugerenciasPersonal();

                    const estabaAbierto =
                        menu.classList.contains(
                            "is-open"
                        );

                    cerrarMultiselects();

                    if (
                        !estabaAbierto
                    ) {

                        menu.classList.add(
                            "is-open"
                        );

                        boton.classList.add(
                            "is-active"
                        );

                        boton.setAttribute(
                            "aria-expanded",
                            "true"
                        );

                        setTimeout(
                            () =>
                                buscador?.focus(),
                            40
                        );

                    }

                }
            );

            menu.addEventListener(
                "click",
                evento =>
                    evento.stopPropagation()
            );

            seleccionarTodos?.addEventListener(
                "change",
                () => {

                    if (
                        seleccionarTodos.checked
                    ) {

                        estado.filtros[
                            filtro
                        ] = new Set(
                            estado.opciones[
                                filtro
                            ]
                        );

                    } else {

                        estado.filtros[
                            filtro
                        ].clear();

                    }

                    gestionarCambioFiltro(
                        filtro
                    );

                }
            );

            buscador?.addEventListener(
                "input",
                () => {

                    filtrarOpcionesVisuales(
                        filtro,
                        buscador.value
                    );

                }
            );

        }
    );

    document.addEventListener(
        "click",
        () => {

            cerrarMultiselects();

            cerrarSugerenciasPersonal();

        }
    );

    document.addEventListener(
        "keydown",
        evento => {

            if (
                evento.key === "Escape"
            ) {

                cerrarMultiselects();

                cerrarSugerenciasPersonal();

                cerrarModalDetalle();

            }

        }
    );

}


// =====================================================
// 🎚️ EVENTOS GENERALES
// =====================================================
function configurarEventosGenerales() {

    document
        .getElementById(
            "btnLimpiarFiltros"
        )
        ?.addEventListener(
            "click",
            limpiarTodosLosFiltros
        );

}


// =====================================================
// 👤 CONSTRUIR LISTA DE PERSONAL
// =====================================================
function construirPersonalDisponible() {

    const mapa =
        new Map();

    estado.registrosOriginales.forEach(
        registro => {

            const clave =
                registro.dni ||
                normalizarBusqueda(
                    registro.nombre
                );

            if (
                !clave
            ) {

                return;

            }

            if (
                !mapa.has(clave)
            ) {

                mapa.set(
                    clave,
                    {
                        nombre:
                            registro.nombre ||
                            "Sin nombre registrado",

                        dni:
                            registro.dni || "",

                        cantidad:
                            0
                    }
                );

            }

            mapa.get(
                clave
            ).cantidad++;

        }
    );

    estado.personalDisponible =
        [...mapa.values()]
            .sort(
                (a, b) =>
                    ordenarTexto(
                        a.nombre,
                        b.nombre
                    )
            );

}


// =====================================================
// 🔎 BUSCADOR DE PERSONAL
// =====================================================
function configurarBuscadorPersonal() {

    const contenedor =
        document.getElementById(
            "buscadorPersonal"
        );

    const input =
        document.getElementById(
            "filtroNombre"
        );

    const btnLimpiar =
        document.getElementById(
            "btnLimpiarNombre"
        );

    if (
        !contenedor ||
        !input
    ) {

        return;

    }

    contenedor.addEventListener(
        "click",
        evento =>
            evento.stopPropagation()
    );

    input.addEventListener(
        "focus",
        () => {

            if (
                input.value.trim()
            ) {

                renderizarSugerenciasPersonal(
                    input.value
                );

            }

        }
    );

    input.addEventListener(
        "input",
        debounce(
            () => {

                estado.personaSeleccionada =
                    null;

                estado.filtros.nombre =
                    normalizarBusqueda(
                        input.value
                    );

                renderizarSugerenciasPersonal(
                    input.value
                );

                aplicarFiltros();

            },
            120
        )
    );

    input.addEventListener(
        "keydown",
        evento => {

            if (
                evento.key !== "ArrowDown"
            ) {

                return;

            }

            const primera =
                document.querySelector(
                    ".sugerencia-personal"
                );

            if (
                primera
            ) {

                evento.preventDefault();

                primera.focus();

            }

        }
    );

    btnLimpiar?.addEventListener(
        "click",
        () => {

            input.value =
                "";

            estado.filtros.nombre =
                "";

            estado.personaSeleccionada =
                null;

            cerrarSugerenciasPersonal();

            aplicarFiltros();

            input.focus();

        }
    );

}


// =====================================================
// 👥 RENDERIZAR SUGERENCIAS
// =====================================================
function renderizarSugerenciasPersonal(
    texto
) {

    const panel =
        document.getElementById(
            "sugerenciasPersonal"
        );

    const lista =
        document.getElementById(
            "listaSugerenciasPersonal"
        );

    const vacio =
        document.getElementById(
            "sugerenciasPersonalVacio"
        );

    const contador =
        document.getElementById(
            "contadorSugerenciasPersonal"
        );

    const input =
        document.getElementById(
            "filtroNombre"
        );

    if (
        !panel ||
        !lista ||
        !vacio ||
        !input
    ) {

        return;

    }

    const busqueda =
        normalizarBusqueda(
            texto
        );

    if (
        !busqueda
    ) {

        cerrarSugerenciasPersonal();

        return;

    }

    cerrarMultiselects();

    const coincidencias =
        estado.personalDisponible
            .filter(persona => {

                const contenido =
                    normalizarBusqueda(
                        `${persona.nombre} ${persona.dni}`
                    );

                return contenido.includes(
                    busqueda
                );

            })
            .slice(
                0,
                CONFIG.maximoSugerencias
            );

    lista.innerHTML =
        "";

    panel.hidden =
        false;

    input.setAttribute(
        "aria-expanded",
        "true"
    );

    if (
        contador
    ) {

        contador.textContent =
            `${coincidencias.length} resultado${coincidencias.length === 1 ? "" : "s"}`;

    }

    if (
        !coincidencias.length
    ) {

        lista.hidden =
            true;

        vacio.hidden =
            false;

        return;

    }

    lista.hidden =
        false;

    vacio.hidden =
        true;

    const fragmento =
        document.createDocumentFragment();

    coincidencias.forEach(
        persona => {

            fragmento.appendChild(
                crearSugerenciaPersonal(
                    persona
                )
            );

        }
    );

    lista.appendChild(
        fragmento
    );

}


// =====================================================
// 🧩 CREAR SUGERENCIA
// =====================================================
function crearSugerenciaPersonal(
    persona
) {

    const template =
        document.getElementById(
            "templateSugerenciaPersonal"
        );

    let boton;

    if (
        template?.content?.firstElementChild
    ) {

        boton =
            template.content
                .firstElementChild
                .cloneNode(
                    true
                );

    } else {

        boton =
            document.createElement(
                "button"
            );

        boton.type =
            "button";

        boton.className =
            "sugerencia-personal";

        boton.innerHTML = `

            <div class="sugerencia-personal__avatar">
                <i class="fa-solid fa-user"></i>
            </div>

            <div class="sugerencia-personal__contenido">

                <strong class="sugerencia-personal__nombre"></strong>

                <span class="sugerencia-personal__dni"></span>

            </div>

            <i class="fa-solid fa-chevron-right sugerencia-personal__arrow"></i>

        `;

    }

    boton.dataset.nombre =
        persona.nombre;

    boton.dataset.dni =
        persona.dni;

    asignarTextoSelector(
        boton,
        ".sugerencia-personal__nombre",
        persona.nombre
    );

    asignarTextoSelector(
        boton,
        ".sugerencia-personal__dni",
        persona.dni
            ? `DNI: ${persona.dni}`
            : "DNI no registrado"
    );

    boton.addEventListener(
        "click",
        () => {

            seleccionarPersona(
                persona
            );

        }
    );

    boton.addEventListener(
        "keydown",
        evento => {

            const opciones =
                [
                    ...document.querySelectorAll(
                        ".sugerencia-personal"
                    )
                ];

            const indice =
                opciones.indexOf(
                    boton
                );

            if (
                evento.key === "ArrowDown"
            ) {

                evento.preventDefault();

                opciones[
                    Math.min(
                        indice + 1,
                        opciones.length - 1
                    )
                ]?.focus();

            }

            if (
                evento.key === "ArrowUp"
            ) {

                evento.preventDefault();

                if (
                    indice <= 0
                ) {

                    document
                        .getElementById(
                            "filtroNombre"
                        )
                        ?.focus();

                } else {

                    opciones[
                        indice - 1
                    ]?.focus();

                }

            }

            if (
                evento.key === "Enter"
            ) {

                evento.preventDefault();

                seleccionarPersona(
                    persona
                );

            }

        }
    );

    return boton;

}


// =====================================================
// ✅ SELECCIONAR PERSONA
// =====================================================
function seleccionarPersona(
    persona
) {

    const input =
        document.getElementById(
            "filtroNombre"
        );

    estado.personaSeleccionada = {

        nombre:
            persona.nombre,

        dni:
            persona.dni

    };

    estado.filtros.nombre =
        normalizarBusqueda(
            persona.dni ||
            persona.nombre
        );

    if (
        input
    ) {

        input.value =
            persona.dni
                ? `${persona.nombre} — ${persona.dni}`
                : persona.nombre;

    }

    cerrarSugerenciasPersonal();

    aplicarFiltros();

}


// =====================================================
// ❌ CERRAR SUGERENCIAS
// =====================================================
function cerrarSugerenciasPersonal() {

    const panel =
        document.getElementById(
            "sugerenciasPersonal"
        );

    const input =
        document.getElementById(
            "filtroNombre"
        );

    if (
        panel
    ) {

        panel.hidden =
            true;

    }

    input?.setAttribute(
        "aria-expanded",
        "false"
    );

}


// =====================================================
// 🔄 RECALCULAR SISTEMA
// =====================================================
function recalcularSistema() {

    actualizarOpcionesDependientes();

    renderizarMultiselects();

    aplicarFiltros();

}


// =====================================================
// 🔗 FILTROS DEPENDIENTES
// MES → TIPO → MOTIVO → OBSERVACIÓN
// =====================================================
function actualizarOpcionesDependientes() {

    estado.opciones.mes =
        obtenerValoresUnicos(
            estado.registrosOriginales,
            "mes",
            ordenarMeses
        );

    depurarFiltro(
        "mes"
    );


    const registrosTipo =
        filtrarPorCampos(
            estado.registrosOriginales,
            [
                "mes"
            ]
        );

    estado.opciones.tipo =
        obtenerValoresUnicos(
            registrosTipo,
            "tipo"
        );

    depurarFiltro(
        "tipo"
    );


    const registrosMotivo =
        filtrarPorCampos(
            estado.registrosOriginales,
            [
                "mes",
                "tipo"
            ]
        );

    estado.opciones.motivo =
        obtenerValoresUnicos(
            registrosMotivo,
            "motivo"
        );

    depurarFiltro(
        "motivo"
    );


    const registrosObservacion =
        filtrarPorCampos(
            estado.registrosOriginales,
            [
                "mes",
                "tipo",
                "motivo"
            ]
        );

    estado.opciones.observacion =
        obtenerValoresUnicos(
            registrosObservacion,
            "observacion"
        );

    depurarFiltro(
        "observacion"
    );

}


// =====================================================
// 🧹 DEPURAR FILTRO
// =====================================================
function depurarFiltro(
    filtro
) {

    const disponibles =
        new Set(
            estado.opciones[
                filtro
            ]
        );

    estado.filtros[
        filtro
    ] = new Set(

        [...estado.filtros[filtro]]
            .filter(valor =>
                disponibles.has(
                    valor
                )
            )

    );

}


// =====================================================
// 🔁 CAMBIO DE FILTRO
// =====================================================
function gestionarCambioFiltro(
    filtro
) {

    limpiarFiltrosPosteriores(
        filtro
    );

    actualizarOpcionesDependientes();

    renderizarMultiselects();

    aplicarFiltros();

}


// =====================================================
// 🧹 LIMPIAR FILTROS POSTERIORES
// =====================================================
function limpiarFiltrosPosteriores(
    filtro
) {

    const orden = [

        "mes",
        "tipo",
        "motivo",
        "observacion"

    ];

    const indice =
        orden.indexOf(
            filtro
        );

    if (
        indice < 0
    ) {

        return;

    }

    orden
        .slice(
            indice + 1
        )
        .forEach(
            filtroPosterior => {

                estado.filtros[
                    filtroPosterior
                ].clear();

            }
        );

}


// =====================================================
// 🧩 RENDERIZAR MULTISELECTS
// =====================================================
function renderizarMultiselects() {

    Object.keys(
        MULTISELECTS
    ).forEach(
        filtro => {

            renderizarMultiselect(
                filtro
            );

        }
    );

}

function renderizarMultiselect(
    filtro
) {

    const config =
        MULTISELECTS[
            filtro
        ];

    const contenedor =
        document.getElementById(
            config.opciones
        );

    const seleccionarTodos =
        document.getElementById(
            config.todos
        );

    if (
        !contenedor
    ) {

        return;

    }

    const opciones =
        estado.opciones[
            filtro
        ];

    const seleccionados =
        estado.filtros[
            filtro
        ];

    contenedor.innerHTML =
        "";

    if (
        !opciones.length
    ) {

        contenedor.innerHTML = `

            <div class="multiselect__empty">

                <i class="fa-solid fa-circle-info"></i>

                <span>
                    No hay opciones disponibles
                </span>

            </div>

        `;

        if (
            seleccionarTodos
        ) {

            seleccionarTodos.checked =
                false;

            seleccionarTodos.indeterminate =
                false;

            seleccionarTodos.disabled =
                true;

        }

        actualizarTextoMultiselect(
            filtro
        );

        return;

    }

    const fragmento =
        document.createDocumentFragment();

    opciones.forEach(
        (opcion, indice) => {

            const etiqueta =
                document.createElement(
                    "label"
                );

            etiqueta.className =
                "multiselect__option";

            etiqueta.dataset.search =
                normalizarBusqueda(
                    opcion
                );

            const id =
                `filtro-${filtro}-${indice}`;

            etiqueta.innerHTML = `

                <input
                    type="checkbox"
                    id="${id}"
                    ${seleccionados.has(opcion)
                        ? "checked"
                        : ""}
                >

                <span class="checkbox-custom"></span>

                <span class="multiselect__option-text">
                    ${escapeHtml(opcion)}
                </span>

            `;

            const checkbox =
                etiqueta.querySelector(
                    "input"
                );

            checkbox.addEventListener(
                "change",
                () => {

                    if (
                        checkbox.checked
                    ) {

                        estado.filtros[
                            filtro
                        ].add(
                            opcion
                        );

                    } else {

                        estado.filtros[
                            filtro
                        ].delete(
                            opcion
                        );

                    }

                    gestionarCambioFiltro(
                        filtro
                    );

                }
            );

            fragmento.appendChild(
                etiqueta
            );

        }
    );

    contenedor.appendChild(
        fragmento
    );

    if (
        seleccionarTodos
    ) {

        seleccionarTodos.disabled =
            false;

        seleccionarTodos.checked =
            seleccionados.size > 0 &&
            seleccionados.size ===
                opciones.length;

        seleccionarTodos.indeterminate =
            seleccionados.size > 0 &&
            seleccionados.size <
                opciones.length;

    }

    actualizarTextoMultiselect(
        filtro
    );

}


// =====================================================
// 📝 TEXTO DEL MULTISELECT
// =====================================================
function actualizarTextoMultiselect(
    filtro
) {

    const config =
        MULTISELECTS[
            filtro
        ];

    const elemento =
        document.getElementById(
            config.texto
        );

    if (
        !elemento
    ) {

        return;

    }

    const seleccionados =
        [...estado.filtros[filtro]];

    const total =
        estado.opciones[
            filtro
        ].length;

    if (
        !seleccionados.length ||
        seleccionados.length === total
    ) {

        elemento.textContent =
            config.textoTodos;

        return;

    }

    if (
        seleccionados.length === 1
    ) {

        elemento.textContent =
            seleccionados[0];

        return;

    }

    elemento.textContent =
        `${seleccionados.length} ${config.plural}`;

}


// =====================================================
// 🔍 FILTRAR OPCIONES VISUALES
// =====================================================
function filtrarOpcionesVisuales(
    filtro,
    texto
) {

    const config =
        MULTISELECTS[
            filtro
        ];

    const contenedor =
        document.getElementById(
            config.opciones
        );

    if (
        !contenedor
    ) {

        return;

    }

    const busqueda =
        normalizarBusqueda(
            texto
        );

    contenedor
        .querySelectorAll(
            ".multiselect__option"
        )
        .forEach(
            opcion => {

                const contenido =
                    opcion.dataset.search || "";

                opcion.style.display =
                    contenido.includes(
                        busqueda
                    )
                        ? "flex"
                        : "none";

            }
        );

}


// =====================================================
// ❌ CERRAR MULTISELECTS
// =====================================================
function cerrarMultiselects() {

    Object.values(
        MULTISELECTS
    ).forEach(
        config => {

            document
                .getElementById(
                    config.menu
                )
                ?.classList.remove(
                    "is-open"
                );

            const boton =
                document.getElementById(
                    config.boton
                );

            boton?.classList.remove(
                "is-active"
            );

            boton?.setAttribute(
                "aria-expanded",
                "false"
            );

        }
    );

}


// =====================================================
// 🎯 APLICAR FILTROS
// =====================================================
function aplicarFiltros() {

    const resultado =
        estado.registrosOriginales.filter(
            registro => {

                return (

                    cumpleSet(
                        registro.mes,
                        estado.filtros.mes
                    ) &&

                    cumpleSet(
                        registro.tipo,
                        estado.filtros.tipo
                    ) &&

                    cumpleSet(
                        registro.motivo,
                        estado.filtros.motivo
                    ) &&

                    cumpleSet(
                        registro.observacion,
                        estado.filtros.observacion
                    ) &&

                    cumpleBusquedaPersonal(
                        registro
                    )

                );

            }
        );

    estado.registrosFiltrados =
        resultado;

    actualizarDashboard(
        resultado
    );

}

function cumpleSet(
    valor,
    seleccionados
) {

    if (
        !seleccionados ||
        seleccionados.size === 0
    ) {

        return true;

    }

    return seleccionados.has(
        valor
    );

}

function cumpleBusquedaPersonal(
    registro
) {

    if (
        estado.personaSeleccionada
    ) {

        const persona =
            estado.personaSeleccionada;

        if (
            persona.dni
        ) {

            return registro.dni ===
                persona.dni;

        }

        return normalizarBusqueda(
            registro.nombre
        ) === normalizarBusqueda(
            persona.nombre
        );

    }

    const busqueda =
        estado.filtros.nombre;

    if (
        !busqueda
    ) {

        return true;

    }

    const contenido =
        normalizarBusqueda(
            `${registro.nombre} ${registro.dni}`
        );

    return contenido.includes(
        busqueda
    );

}


// =====================================================
// 🧮 FILTRAR PARA OPCIONES DEPENDIENTES
// =====================================================
function filtrarPorCampos(
    registros,
    campos
) {

    return registros.filter(
        registro => {

            return campos.every(
                campo => {

                    return cumpleSet(
                        registro[campo],
                        estado.filtros[campo]
                    );

                }
            );

        }
    );

}


// =====================================================
// 🖥️ ACTUALIZAR DASHBOARD
// =====================================================
function actualizarDashboard(
    data
) {

    actualizarKPIs(
        data
    );

    renderGraficoMeses(
        data
    );

    renderGraficoDias(
        data
    );

    renderGraficoPersonas(
        data
    );

    renderGraficoMotivos(
        data
    );

    renderGraficoTipos(
        data
    );

    renderGraficoObservaciones(
        data
    );

    actualizarResumenFiltros(
        data
    );

}


// =====================================================
// 📊 KPI
// =====================================================
function actualizarKPIs(
    data
) {

    setTexto(
        "kpiTotal",
        formatearNumero(
            data.length
        )
    );

    const personas =
        new Set();

    const conteoPersonas = {};

    data.forEach(
        registro => {

            const clave =
                registro.dni ||
                normalizarBusqueda(
                    registro.nombre
                );

            if (
                !clave
            ) {

                return;

            }

            personas.add(
                clave
            );

            conteoPersonas[
                clave
            ] = (
                conteoPersonas[
                    clave
                ] || 0
            ) + 1;

        }
    );

    setTexto(
        "kpiPersonas",
        formatearNumero(
            personas.size
        )
    );

    const motivoMayor =
        obtenerMayor(
            contarPorCampo(
                data,
                "motivo"
            )
        );

    setTexto(
        "kpiUnidad",
        motivoMayor.clave
            ? `${motivoMayor.clave} (${motivoMayor.valor})`
            : "-"
    );

    const reincidentes =
        Object.values(
            conteoPersonas
        ).filter(
            cantidad =>
                cantidad >=
                CONFIG.reincidenciaCritica
        ).length;

    setTexto(
        "kpiReincidentes",
        formatearNumero(
            reincidentes
        )
    );

}


// =====================================================
// 📈 GRÁFICO POR MESES
// =====================================================
function renderGraficoMeses(
    data
) {

    destruirChart(
        "meses"
    );

    const canvas =
        document.getElementById(
            "graficoMeses"
        );

    if (
        !canvas
    ) {

        return;

    }

    const conteo =
        Object.fromEntries(
            MESES.map(
                mes => [
                    mes,
                    0
                ]
            )
        );

    data.forEach(
        registro => {

            if (
                Object.prototype.hasOwnProperty.call(
                    conteo,
                    registro.mes
                )
            ) {

                conteo[
                    registro.mes
                ]++;

            }

        }
    );

    const labels =
        [...MESES];

    const valores =
        labels.map(
            mes =>
                conteo[mes]
        );

    estado.charts.meses =
        new Chart(
            canvas,
            {

                type:
                    "line",

                data: {

                    labels,

                    datasets: [{

                        label:
                            "Papeletas por mes",

                        data:
                            valores,

                        borderColor:
                            "#556B2F",

                        backgroundColor:
                            "rgba(85,107,47,0.12)",

                        pointBackgroundColor:
                            "#556B2F",

                        pointBorderColor:
                            "#FFFFFF",

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

                        mode:
                            "nearest",

                        intersect:
                            false

                    },

                    onHover(
                        evento
                    ) {

                        const elementos =
                            this.getElementsAtEventForMode(
                                evento,
                                "nearest",
                                {
                                    intersect:
                                        false
                                },
                                true
                            );

                        cambiarCursorGrafico(
                            evento,
                            elementos.length > 0
                        );

                    },

                    onClick(
                        evento
                    ) {

                        const elementos =
                            this.getElementsAtEventForMode(
                                evento,
                                "nearest",
                                {
                                    intersect:
                                        false
                                },
                                true
                            );

                        if (
                            !elementos.length
                        ) {

                            return;

                        }

                        const indice =
                            elementos[0].index;

                        const mes =
                            labels[indice];

                        abrirModalDetalle({

                            categoria:
                                "Comportamiento mensual",

                            titulo:
                                `Papeletas de ${mes}`,

                            seleccion:
                                mes,

                            icono:
                                "fa-solid fa-chart-line",

                            registros:
                                data.filter(
                                    registro =>
                                        registro.mes ===
                                        mes
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
// 📊 GRÁFICO POR DÍA
// =====================================================
function renderGraficoDias(
    data
) {

    destruirChart(
        "dias"
    );

    const canvas =
        document.getElementById(
            "graficoDias"
        );

    if (
        !canvas
    ) {

        return;

    }

    const conteo = {};

    data.forEach(
        registro => {

            if (
                !registro.fechaInicio
            ) {

                return;

            }

            conteo[
                registro.fechaInicio
            ] = (
                conteo[
                    registro.fechaInicio
                ] || 0
            ) + 1;

        }
    );

    const entradas =
        Object.entries(
            conteo
        )
            .sort(
                ([fechaA], [fechaB]) =>
                    fechaATimestamp(fechaA) -
                    fechaATimestamp(fechaB)
            );

    const labels =
        entradas.map(
            ([fecha]) =>
                fecha
        );

    const valores =
        entradas.map(
            ([, cantidad]) =>
                cantidad
        );

    estado.charts.dias =
        new Chart(
            canvas,
            {

                type:
                    "bar",

                data: {

                    labels,

                    datasets: [{

                        label:
                            "Papeletas por día",

                        data:
                            valores,

                        backgroundColor:
                            generarColoresPorValor(
                                valores
                            ),

                        borderRadius:
                            8,

                        borderSkipped:
                            false,

                        minBarLength:
                            6,

                        maxBarThickness:
                            36

                    }]

                },

                options: {

                    ...opcionesVerticales(),

                    interaction: {

                        mode:
                            "index",

                        intersect:
                            false,

                        axis:
                            "x"

                    },

                    onHover(
                        evento
                    ) {

                        const elementos =
                            this.getElementsAtEventForMode(
                                evento,
                                "index",
                                {
                                    intersect:
                                        false,

                                    axis:
                                        "x"
                                },
                                true
                            );

                        cambiarCursorGrafico(
                            evento,
                            elementos.length > 0
                        );

                    },

                    onClick(
                        evento
                    ) {

                        const elementos =
                            this.getElementsAtEventForMode(
                                evento,
                                "index",
                                {
                                    intersect:
                                        false,

                                    axis:
                                        "x"
                                },
                                true
                            );

                        if (
                            !elementos.length
                        ) {

                            return;

                        }

                        const fecha =
                            labels[
                                elementos[0].index
                            ];

                        abrirModalDetalle({

                            categoria:
                                "Tendencia diaria",

                            titulo:
                                `Papeletas del ${fecha}`,

                            seleccion:
                                fecha,

                            icono:
                                "fa-solid fa-calendar-day",

                            registros:
                                data.filter(
                                    registro =>
                                        registro.fechaInicio ===
                                        fecha
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
// 👥 PERSONAL
// =====================================================
function renderGraficoPersonas(
    data
) {

    renderGraficoHorizontal({

        chartKey:
            "personas",

        canvasId:
            "personaspapeletas",

        contenidoId:
            "contenidoGraficoPersonas",

        conteo:
            contarPorCampo(
                data,
                "nombre"
            ),

        registros:
            data,

        campo:
            "nombre",

        categoria:
            "Personal con más papeletas",

        icono:
            "fa-solid fa-user",

        label:
            "Cantidad de papeletas"

    });

}


// =====================================================
// 📋 MOTIVOS
// =====================================================
function renderGraficoMotivos(
    data
) {

    renderGraficoHorizontal({

        chartKey:
            "motivos",

        canvasId:
            "motivospapeletas",

        contenidoId:
            "contenidoGraficoMotivos",

        conteo:
            contarPorCampo(
                data,
                "motivo"
            ),

        registros:
            data,

        campo:
            "motivo",

        categoria:
            "Motivos de papeletas",

        icono:
            "fa-solid fa-list-check",

        label:
            "Cantidad de papeletas"

    });

}


// =====================================================
// 📑 TIPOS
// =====================================================
function renderGraficoTipos(
    data
) {

    renderGraficoHorizontal({

        chartKey:
            "tipos",

        canvasId:
            "tipopapeletas",

        contenidoId:
            "contenidoGraficoTipos",

        conteo:
            contarPorCampo(
                data,
                "tipo"
            ),

        registros:
            data,

        campo:
            "tipo",

        categoria:
            "Tipo de papeleta",

        icono:
            "fa-solid fa-file-circle-exclamation",

        label:
            "Cantidad de papeletas"

    });

}


// =====================================================
// 💬 OBSERVACIONES
// =====================================================
function renderGraficoObservaciones(
    data
) {

    renderGraficoHorizontal({

        chartKey:
            "observaciones",

        canvasId:
            "comentarios",

        contenidoId:
            "contenidoGraficoObservaciones",

        conteo:
            contarPorCampo(
                data,
                "observacion"
            ),

        registros:
            data,

        campo:
            "observacion",

        categoria:
            "Observaciones",

        icono:
            "fa-solid fa-comment-dots",

        label:
            "Cantidad de registros"

    });

}


// =====================================================
// 📊 GRÁFICO HORIZONTAL GENÉRICO
// TODA LA FILA ES CLICABLE
// =====================================================
function renderGraficoHorizontal({

    chartKey,

    canvasId,

    contenidoId,

    conteo,

    registros,

    campo,

    categoria,

    icono,

    label

}) {

    destruirChart(
        chartKey
    );

    const canvas =
        document.getElementById(
            canvasId
        );

    const contenido =
        document.getElementById(
            contenidoId
        );

    if (
        !canvas ||
        !contenido
    ) {

        return;

    }

    const entradas =
        Object.entries(
            conteo
        )
            .filter(
                ([valor, cantidad]) =>
                    valor &&
                    cantidad > 0
            )
            .sort(
                (a, b) =>
                    b[1] - a[1] ||
                    ordenarTexto(
                        a[0],
                        b[0]
                    )
            )
            .slice(
                0,
                CONFIG.maximoElementosGrafico
            );

    const labels =
        entradas.map(
            ([valor]) =>
                valor
        );

    const valores =
        entradas.map(
            ([, cantidad]) =>
                cantidad
        );

    ajustarContenidoGrafico(
        contenido,
        labels.length
    );

    estado.charts[
        chartKey
    ] = new Chart(
        canvas,
        {

            type:
                "bar",

            data: {

                labels,

                datasets: [{

                    label,

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

                    mode:
                        "index",

                    intersect:
                        false,

                    axis:
                        "y"

                },

                onHover(
                    evento
                ) {

                    const elementos =
                        this.getElementsAtEventForMode(
                            evento,
                            "index",
                            {
                                intersect:
                                    false,

                                axis:
                                    "y"
                            },
                            true
                        );

                    cambiarCursorGrafico(
                        evento,
                        elementos.length > 0
                    );

                },

                onClick(
                    evento
                ) {

                    const elementos =
                        this.getElementsAtEventForMode(
                            evento,
                            "index",
                            {
                                intersect:
                                    false,

                                axis:
                                    "y"
                            },
                            true
                        );

                    if (
                        !elementos.length
                    ) {

                        return;

                    }

                    const indice =
                        elementos[0].index;

                    const seleccion =
                        labels[indice];

                    const detalle =
                        registros.filter(
                            registro =>
                                registro[campo] ===
                                seleccion
                        );

                    abrirModalDetalle({

                        categoria,

                        titulo:
                            seleccion,

                        seleccion,

                        icono,

                        registros:
                            detalle

                    });

                }

            },

            plugins:
                obtenerPluginsChart()

        }
    );

}


// =====================================================
// 🖱️ CURSOR
// =====================================================
function cambiarCursorGrafico(
    evento,
    activo
) {

    const target =
        evento?.native?.target;

    if (
        target
    ) {

        target.style.cursor =
            activo
                ? "pointer"
                : "default";

    }

}


// =====================================================
// 📏 ALTURA DEL GRÁFICO
// =====================================================
function ajustarContenidoGrafico(
    contenido,
    cantidad
) {

    const altura =
        Math.max(
            CONFIG.alturaMinimaGrafico,
            cantidad *
                CONFIG.alturaPorFilaGrafico
        );

    contenido.style.height =
        `${altura}px`;

    contenido.style.minHeight =
        `${altura}px`;

    contenido.style.position =
        "relative";

    contenido.style.width =
        "100%";

}


// =====================================================
// ⚙️ OPCIONES VERTICALES
// =====================================================
function opcionesVerticales() {

    return {

        responsive:
            true,

        maintainAspectRatio:
            false,

        animation:
            false,

        normalized:
            true,

        layout: {

            padding: {

                top:
                    12,

                right:
                    15,

                left:
                    5

            }

        },

        plugins: {

            legend: {

                display:
                    false

            },

            tooltip: {

                callbacks: {

                    label:
                        contexto =>
                            ` ${formatearNumero(contexto.raw)} papeleta(s)`

                }

            },

            datalabels: {

                color:
                    "#374151",

                anchor:
                    "end",

                align:
                    "top",

                offset:
                    3,

                clamp:
                    true,

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
                        "#475569",

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

                beginAtZero:
                    true,

                ticks: {

                    color:
                        "#475569",

                    precision:
                        0

                },

                grid: {

                    color:
                        "rgba(0,0,0,0.07)"

                }

            }

        }

    };

}


// =====================================================
// ⚙️ OPCIONES HORIZONTALES
// =====================================================
function opcionesHorizontales() {

    return {

        responsive:
            true,

        maintainAspectRatio:
            false,

        animation:
            false,

        normalized:
            true,

        indexAxis:
            "y",

        layout: {

            padding: {

                right:
                    42,

                left:
                    5,

                top:
                    4,

                bottom:
                    4

            }

        },

        plugins: {

            legend: {

                display:
                    false

            },

            tooltip: {

                callbacks: {

                    title:
                        elementos =>
                            elementos[0]?.label || "",

                    label:
                        contexto =>
                            ` ${formatearNumero(contexto.raw)} registro(s)`

                }

            },

            datalabels: {

                color:
                    "#374151",

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
                        valor

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
                        "#475569",

                    precision:
                        0

                },

                grid: {

                    color:
                        "rgba(0,0,0,0.07)"

                }

            },

            y: {

                ticks: {

                    color:
                        "#374151",

                    autoSkip:
                        false,

                    font: {

                        size:
                            11

                    },

                    callback(
                        valor
                    ) {

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


// =====================================================
// 🎨 COLORES
// =====================================================
function generarColoresPorValor(
    valores
) {

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

            if (
                proporcion >= 0.90
            ) {

                return "#E78B8B";

            }

            if (
                proporcion >= 0.80
            ) {

                return "#F0A36B";

            }

            if (
                proporcion >= 0.70
            ) {

                return "#F6BF5B";

            }

            if (
                proporcion >= 0.60
            ) {

                return "#E4D86F";

            }

            if (
                proporcion >= 0.50
            ) {

                return "#C9D97B";

            }

            if (
                proporcion >= 0.40
            ) {

                return "#A9D18E";

            }

            if (
                proporcion >= 0.30
            ) {

                return "#7FC8A9";

            }

            if (
                proporcion >= 0.20
            ) {

                return "#66BFBF";

            }

            return "#8EC5E8";

        }
    );

}


// =====================================================
// 🪟 CONFIGURAR MODAL
// =====================================================
function configurarModal() {

    const modal =
        document.getElementById(
            "modalDetalle"
        );

    const buscador =
        document.getElementById(
            "buscarDetalleModal"
        );

    document
        .getElementById(
            "btnCerrarModalDetalle"
        )
        ?.addEventListener(
            "click",
            cerrarModalDetalle
        );

    document
        .getElementById(
            "btnCerrarModalDetalleInferior"
        )
        ?.addEventListener(
            "click",
            cerrarModalDetalle
        );

    modal?.addEventListener(
        "click",
        evento => {

            if (
                evento.target === modal
            ) {

                cerrarModalDetalle();

            }

        }
    );

    buscador?.addEventListener(
        "input",
        debounce(
            () => {

                filtrarContenidoModal(
                    buscador.value
                );

            },
            140
        )
    );

}


// =====================================================
// 🪟 ABRIR MODAL
// =====================================================
function abrirModalDetalle({

    categoria,

    titulo,

    seleccion,

    icono,

    registros

}) {

    const modal =
        document.getElementById(
            "modalDetalle"
        );

    if (
        !modal
    ) {

        return;

    }

    estado.registrosModal =
        ordenarCronologicamente(
            registros
        );

    setTexto(
        "modalDetalleCategoria",
        categoria
    );

    setTexto(
        "modalDetalleTitulo",
        titulo
    );

    setTexto(
        "modalDetalleSeleccion",
        seleccion
    );

    setTexto(
        "modalDetalleTotal",
        formatearNumero(
            registros.length
        )
    );

    setTexto(
        "modalDetallePersonas",
        formatearNumero(
            contarPersonasUnicas(
                registros
            )
        )
    );

    setTexto(
        "modalDetallePeriodo",
        obtenerPeriodoRegistros(
            registros
        )
    );

    const iconoModal =
        document.getElementById(
            "modalDetalleIcono"
        );

    if (
        iconoModal
    ) {

        iconoModal.className =
            icono;

    }

    const buscador =
        document.getElementById(
            "buscarDetalleModal"
        );

    if (
        buscador
    ) {

        buscador.value =
            "";

    }

    renderizarTablaModal(
        estado.registrosModal
    );

    modal.classList.add(
        "is-open"
    );

    modal.setAttribute(
        "aria-hidden",
        "false"
    );

    document.body.classList.add(
        "modal-open"
    );

    setTimeout(
        () =>
            buscador?.focus(),
        100
    );

}


// =====================================================
// ❌ CERRAR MODAL
// =====================================================
function cerrarModalDetalle() {

    const modal =
        document.getElementById(
            "modalDetalle"
        );

    if (
        !modal ||
        !modal.classList.contains(
            "is-open"
        )
    ) {

        return;

    }

    modal.classList.remove(
        "is-open"
    );

    modal.setAttribute(
        "aria-hidden",
        "true"
    );

    document.body.classList.remove(
        "modal-open"
    );

    estado.registrosModal =
        [];

}


// =====================================================
// 🔍 FILTRAR TABLA DEL MODAL
// =====================================================
function filtrarContenidoModal(
    texto
) {

    const busqueda =
        normalizarBusqueda(
            texto
        );

    if (
        !busqueda
    ) {

        renderizarTablaModal(
            estado.registrosModal
        );

        return;

    }

    const filtrados =
        estado.registrosModal.filter(
            registro => {

                const contenido =
                    normalizarBusqueda(`

                        ${registro.mes}
                        ${registro.nombre}
                        ${registro.dni}
                        ${registro.unidad}
                        ${registro.tipo}
                        ${registro.motivo}
                        ${registro.fechaInicio}
                        ${registro.fechaFin}
                        ${registro.observacion}

                    `);

                return contenido.includes(
                    busqueda
                );

            }
        );

    renderizarTablaModal(
        filtrados
    );

}


// =====================================================
// 📋 TABLA TIPO EXCEL EN EL MODAL
// =====================================================
function renderizarTablaModal(
    registros
) {

    const contenedor =
        document.getElementById(
            "listaDetalleModal"
        );

    const vacio =
        document.getElementById(
            "modalDetalleVacio"
        );

    if (
        !contenedor ||
        !vacio
    ) {

        return;

    }

    contenedor.innerHTML =
        "";

    setTexto(
        "contadorDetalleModal",
        `${formatearNumero(registros.length)} registro${registros.length === 1 ? "" : "s"}`
    );

    if (
        !registros.length
    ) {

        contenedor.hidden =
            true;

        vacio.hidden =
            false;

        return;

    }

    contenedor.hidden =
        false;

    vacio.hidden =
        true;

    const wrapper =
        document.createElement(
            "div"
        );

    wrapper.className =
        "detalle-tabla-responsive";

    const tabla =
        document.createElement(
            "table"
        );

    tabla.className =
        "detalle-tabla-excel";

    tabla.innerHTML = `

        <thead>

            <tr>

                <th class="detalle-columna-numero">
                    N.°
                </th>

                <th>
                    Mes
                </th>

                <th>
                    Apellidos y Nombres
                </th>

                <th>
                    DNI
                </th>

                <th>
                    Unidad Orgánica
                </th>

                <th>
                    Tipo de Papeleta
                </th>

                <th>
                    Motivo de Papeleta
                </th>

                <th>
                    Fecha Inicio
                </th>

                <th>
                    Fecha Fin
                </th>

                <th>
                    Observaciones
                </th>

            </tr>

        </thead>

        <tbody></tbody>

    `;

    const tbody =
        tabla.querySelector(
            "tbody"
        );

    const fragmento =
        document.createDocumentFragment();

    registros.forEach(
        (registro, indice) => {

            const fila =
                document.createElement(
                    "tr"
                );

            fila.innerHTML = `

                <td class="detalle-columna-numero">
                    ${indice + 1}
                </td>

                <td>
                    ${escapeHtml(registro.mes || "-")}
                </td>

                <td class="detalle-celda-nombre">
                    ${escapeHtml(registro.nombre || "Sin nombre registrado")}
                </td>

                <td class="detalle-celda-dni">
                    ${escapeHtml(registro.dni || "-")}
                </td>

                <td>
                    ${escapeHtml(registro.unidad || "-")}
                </td>

                <td>
                    ${escapeHtml(registro.tipo || "-")}
                </td>

                <td>
                    ${escapeHtml(registro.motivo || "-")}
                </td>

                <td class="detalle-celda-fecha">
                    ${escapeHtml(formatearFechaVisible(registro.fechaInicio))}
                </td>

                <td class="detalle-celda-fecha">
                    ${escapeHtml(formatearFechaVisible(registro.fechaFin))}
                </td>

                <td>
                    ${escapeHtml(registro.observacion || "Sin observación")}
                </td>

            `;

            fragmento.appendChild(
                fila
            );

        }
    );

    tbody.appendChild(
        fragmento
    );

    wrapper.appendChild(
        tabla
    );

    contenedor.appendChild(
        wrapper
    );

}


// =====================================================
// 📆 PERIODO DEL MODAL
// =====================================================
function obtenerPeriodoRegistros(
    registros
) {

    const fechas =
        registros
            .map(
                registro =>
                    registro.fechaInicio
            )
            .filter(
                fecha =>
                    fechaATimestamp(fecha) > 0
            )
            .sort(
                (fechaA, fechaB) =>
                    fechaATimestamp(fechaA) -
                    fechaATimestamp(fechaB)
            );

    if (
        !fechas.length
    ) {

        return "Sin fecha registrada";

    }

    const primera =
        formatearFechaVisible(
            fechas[0]
        );

    const ultima =
        formatearFechaVisible(
            fechas[
                fechas.length - 1
            ]
        );

    if (
        primera === ultima
    ) {

        return primera;

    }

    return `${primera} al ${ultima}`;

}


// =====================================================
// 🗂️ ORDEN CRONOLÓGICO
// MÁS RECIENTE PRIMERO
// =====================================================
function ordenarCronologicamente(
    registros
) {

    return [...registros]
        .sort(
            (a, b) => {

                const fechaA =
                    fechaATimestamp(
                        a.fechaInicio
                    );

                const fechaB =
                    fechaATimestamp(
                        b.fechaInicio
                    );

                if (
                    fechaA !== fechaB
                ) {

                    return fechaB -
                        fechaA;

                }

                const comparacionNombre =
                    ordenarTexto(
                        a.nombre,
                        b.nombre
                    );

                if (
                    comparacionNombre !== 0
                ) {

                    return comparacionNombre;

                }

                return ordenarTexto(
                    a.dni,
                    b.dni
                );

            }
        );

}


// =====================================================
// 🔢 PERSONAS ÚNICAS
// =====================================================
function contarPersonasUnicas(
    registros
) {

    const personas =
        new Set();

    registros.forEach(
        registro => {

            const clave =
                registro.dni ||
                normalizarBusqueda(
                    registro.nombre
                );

            if (
                clave
            ) {

                personas.add(
                    clave
                );

            }

        }
    );

    return personas.size;

}


// =====================================================
// 🏷️ RESUMEN DE FILTROS
// =====================================================
function actualizarResumenFiltros(
    data
) {

    setTexto(
        "textoResumenFiltros",
        `Mostrando ${formatearNumero(data.length)} de ${formatearNumero(estado.registrosOriginales.length)} registros`
    );

    const contenedor =
        document.getElementById(
            "chipsFiltros"
        );

    if (
        !contenedor
    ) {

        return;

    }

    contenedor.innerHTML =
        "";

    const fragmento =
        document.createDocumentFragment();

    [

        [
            "mes",
            "Mes"
        ],

        [
            "tipo",
            "Tipo"
        ],

        [
            "motivo",
            "Motivo"
        ],

        [
            "observacion",
            "Observación"
        ]

    ].forEach(
        ([filtro, etiqueta]) => {

            const seleccionados =
                [...estado.filtros[filtro]];

            const total =
                estado.opciones[
                    filtro
                ].length;

            if (
                !seleccionados.length ||
                seleccionados.length === total
            ) {

                return;

            }

            seleccionados.forEach(
                valor => {

                    fragmento.appendChild(
                        crearChip(
                            filtro,
                            etiqueta,
                            valor
                        )
                    );

                }
            );

        }
    );

    if (
        estado.filtros.nombre
    ) {

        const input =
            document.getElementById(
                "filtroNombre"
            );

        fragmento.appendChild(
            crearChip(
                "nombre",
                "Personal",
                input?.value.trim() ||
                estado.filtros.nombre
            )
        );

    }

    contenedor.appendChild(
        fragmento
    );

}


// =====================================================
// 🏷️ CREAR CHIP
// =====================================================
function crearChip(
    filtro,
    etiqueta,
    valor
) {

    const chip =
        document.createElement(
            "button"
        );

    chip.type =
        "button";

    chip.className =
        "filtro-chip";

    chip.innerHTML = `

        <span>

            <strong>
                ${escapeHtml(etiqueta)}:
            </strong>

            ${escapeHtml(valor)}

        </span>

        <i class="fa-solid fa-xmark"></i>

    `;

    chip.addEventListener(
        "click",
        () => {

            if (
                filtro === "nombre"
            ) {

                const input =
                    document.getElementById(
                        "filtroNombre"
                    );

                if (
                    input
                ) {

                    input.value =
                        "";

                }

                estado.filtros.nombre =
                    "";

                estado.personaSeleccionada =
                    null;

                aplicarFiltros();

                return;

            }

            estado.filtros[
                filtro
            ].delete(
                valor
            );

            gestionarCambioFiltro(
                filtro
            );

        }
    );

    return chip;

}


// =====================================================
// 🧹 LIMPIAR FILTROS
// =====================================================
function limpiarTodosLosFiltros() {

    estado.filtros.mes.clear();

    estado.filtros.tipo.clear();

    estado.filtros.motivo.clear();

    estado.filtros.observacion.clear();

    estado.filtros.nombre =
        "";

    estado.personaSeleccionada =
        null;

    const input =
        document.getElementById(
            "filtroNombre"
        );

    if (
        input
    ) {

        input.value =
            "";

    }

    Object.values(
        MULTISELECTS
    ).forEach(
        config => {

            const buscador =
                document.getElementById(
                    config.buscador
                );

            if (
                buscador
            ) {

                buscador.value =
                    "";

            }

        }
    );

    cerrarSugerenciasPersonal();

    recalcularSistema();

}


// =====================================================
// 📚 VALORES ÚNICOS
// =====================================================
function obtenerValoresUnicos(
    registros,
    campo,
    comparador = ordenarTexto
) {

    const valores =
        new Set();

    registros.forEach(
        registro => {

            const valor =
                limpiarTexto(
                    registro[campo]
                );

            if (
                valor
            ) {

                valores.add(
                    valor
                );

            }

        }
    );

    return [...valores]
        .sort(
            comparador
        );

}


// =====================================================
// 🔢 CONTAR POR CAMPO
// =====================================================
function contarPorCampo(
    registros,
    campo
) {

    const conteo = {};

    registros.forEach(
        registro => {

            const valor =
                limpiarTexto(
                    registro[campo]
                );

            if (
                !valor
            ) {

                return;

            }

            conteo[
                valor
            ] = (
                conteo[
                    valor
                ] || 0
            ) + 1;

        }
    );

    return conteo;

}


// =====================================================
// 🥇 OBTENER MAYOR
// =====================================================
function obtenerMayor(
    objeto
) {

    let clave =
        "";

    let valor =
        0;

    Object.entries(
        objeto
    ).forEach(
        ([claveActual, valorActual]) => {

            if (
                valorActual > valor
            ) {

                clave =
                    claveActual;

                valor =
                    valorActual;

            }

        }
    );

    return {

        clave,
        valor

    };

}


// =====================================================
// 📅 NORMALIZAR MES
// =====================================================
function normalizarMes(
    valor
) {

    const texto =
        normalizarClave(
            valor
        );

    if (
        texto === "SEPTIEMBRE"
    ) {

        return "SETIEMBRE";

    }

    if (
        NUMERO_MES[
            texto
        ]
    ) {

        return texto;

    }

    const numero =
        Number(
            texto
        );

    if (
        numero >= 1 &&
        numero <= 12
    ) {

        return MESES[
            numero - 1
        ];

    }

    return texto;

}

function ordenarMeses(
    mesA,
    mesB
) {

    return (
        (
            NUMERO_MES[mesA] || 99
        ) -
        (
            NUMERO_MES[mesB] || 99
        )
    );

}


// =====================================================
// 🗓️ NORMALIZAR FECHAS DD/MM/AAAA
// =====================================================
function normalizarFecha(
    valor
) {

    if (
        valor === undefined ||
        valor === null ||
        valor === ""
    ) {

        return "";

    }


    // =============================================
    // FECHA JAVASCRIPT
    // =============================================
    if (
        valor instanceof Date &&
        !Number.isNaN(
            valor.getTime()
        )
    ) {

        return formatearFecha(
            valor
        );

    }


    // =============================================
    // NÚMERO SERIAL DE EXCEL
    // =============================================
    if (
        typeof valor === "number" &&
        Number.isFinite(valor)
    ) {

        const fechaExcel =
            XLSX.SSF.parse_date_code(
                valor
            );

        if (
            fechaExcel
        ) {

            return construirFechaTexto(

                fechaExcel.d,

                fechaExcel.m,

                fechaExcel.y

            );

        }

    }


    const texto =
        String(
            valor
        )
            .trim();

    if (
        !texto
    ) {

        return "";

    }


    // =============================================
    // QUITAR HORA
    // =============================================
    const textoSinHora =
        texto
            .replace(
                /T.*$/,
                ""
            )
            .split(" ")[0]
            .trim();


    // =============================================
    // DD/MM/AAAA, DD-MM-AAAA O DD.MM.AAAA
    // =============================================
    let coincidencia =
        textoSinHora.match(
            /^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/
        );

    if (
        coincidencia
    ) {

        let dia =
            Number(
                coincidencia[1]
            );

        let mes =
            Number(
                coincidencia[2]
            );

        let anio =
            Number(
                coincidencia[3]
            );

        if (
            anio < 100
        ) {

            anio +=
                anio >= 50
                    ? 1900
                    : 2000;

        }

        if (
            fechaValida(
                dia,
                mes,
                anio
            )
        ) {

            return construirFechaTexto(
                dia,
                mes,
                anio
            );

        }

    }


    // =============================================
    // AAAA/MM/DD, AAAA-MM-DD O AAAA.MM.DD
    // =============================================
    coincidencia =
        textoSinHora.match(
            /^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})$/
        );

    if (
        coincidencia
    ) {

        const anio =
            Number(
                coincidencia[1]
            );

        const mes =
            Number(
                coincidencia[2]
            );

        const dia =
            Number(
                coincidencia[3]
            );

        if (
            fechaValida(
                dia,
                mes,
                anio
            )
        ) {

            return construirFechaTexto(
                dia,
                mes,
                anio
            );

        }

    }


    // =============================================
    // ÚLTIMO INTENTO CONTROLADO
    // =============================================
    const fechaInterpretada =
        new Date(
            texto
        );

    if (
        !Number.isNaN(
            fechaInterpretada.getTime()
        )
    ) {

        return formatearFecha(
            fechaInterpretada
        );

    }

    return "";

}

function construirFechaTexto(
    dia,
    mes,
    anio
) {

    return [

        String(dia)
            .padStart(
                2,
                "0"
            ),

        String(mes)
            .padStart(
                2,
                "0"
            ),

        String(anio)
            .padStart(
                4,
                "0"
            )

    ].join("/");

}

function formatearFecha(
    fecha
) {

    return construirFechaTexto(

        fecha.getDate(),

        fecha.getMonth() + 1,

        fecha.getFullYear()

    );

}

function formatearFechaVisible(
    fecha
) {

    const normalizada =
        normalizarFecha(
            fecha
        );

    return normalizada || "-";

}

function fechaValida(
    dia,
    mes,
    anio
) {

    if (
        !Number.isInteger(dia) ||
        !Number.isInteger(mes) ||
        !Number.isInteger(anio)
    ) {

        return false;

    }

    if (
        anio < 1900 ||
        anio > 2200 ||
        mes < 1 ||
        mes > 12 ||
        dia < 1 ||
        dia > 31
    ) {

        return false;

    }

    const fecha =
        new Date(
            anio,
            mes - 1,
            dia
        );

    return (

        fecha.getFullYear() === anio &&
        fecha.getMonth() === mes - 1 &&
        fecha.getDate() === dia

    );

}

function fechaATimestamp(
    fecha
) {

    const normalizada =
        normalizarFecha(
            fecha
        );

    if (
        !normalizada
    ) {

        return 0;

    }

    const [
        dia,
        mes,
        anio
    ] = normalizada
        .split("/")
        .map(
            Number
        );

    if (
        !fechaValida(
            dia,
            mes,
            anio
        )
    ) {

        return 0;

    }

    return new Date(
        anio,
        mes - 1,
        dia
    ).getTime();

}


// =====================================================
// 🧹 TEXTO Y DNI
// =====================================================
function limpiarTexto(
    valor
) {

    const texto =
        String(
            valor ?? ""
        )
            .replace(
                /\s+/g,
                " "
            )
            .trim();

    const normalizado =
        texto.toUpperCase();

    if (
        !texto ||
        normalizado === "NAN" ||
        normalizado === "NULL" ||
        normalizado === "UNDEFINED"
    ) {

        return "";

    }

    return texto;

}

function limpiarDni(
    valor
) {

    const texto =
        limpiarTexto(
            valor
        );

    if (
        !texto
    ) {

        return "";

    }

    return texto
        .replace(
            /\.0$/,
            ""
        )
        .replace(
            /\s+/g,
            ""
        );

}

function normalizarBusqueda(
    valor
) {

    return String(
        valor ?? ""
    )
        .normalize("NFD")
        .replace(
            /[\u0300-\u036f]/g,
            ""
        )
        .replace(
            /\s+/g,
            " "
        )
        .trim()
        .toLowerCase();

}

function ordenarTexto(
    textoA,
    textoB
) {

    return String(
        textoA || ""
    ).localeCompare(
        String(
            textoB || ""
        ),
        "es",
        {
            sensitivity:
                "base",

            numeric:
                true
        }
    );

}

function truncarTexto(
    texto,
    limite
) {

    const valor =
        String(
            texto ?? ""
        );

    if (
        valor.length <= limite
    ) {

        return valor;

    }

    return `${valor.slice(0, limite - 3)}...`;

}


// =====================================================
// 📉 CHART.JS
// =====================================================
function registrarPluginChart() {

    if (
        typeof Chart === "undefined"
    ) {

        throw new Error(
            "Chart.js no está disponible."
        );

    }

    if (
        window.ChartDataLabels
    ) {

        Chart.register(
            window.ChartDataLabels
        );

    }

}

function obtenerPluginsChart() {

    return window.ChartDataLabels
        ? [
            window.ChartDataLabels
        ]
        : [];

}

function destruirChart(
    clave
) {

    if (
        estado.charts[
            clave
        ]
    ) {

        estado.charts[
            clave
        ].destroy();

        estado.charts[
            clave
        ] = null;

    }

}


// =====================================================
// 🖥️ UI
// =====================================================
function mostrarLoading(
    titulo,
    texto = ""
) {

    setTexto(
        "loadingTitulo",
        titulo
    );

    setTexto(
        "loadingTexto",
        texto
    );

    const overlay =
        document.getElementById(
            "loadingSistema"
        );

    if (
        overlay
    ) {

        overlay.style.display =
            "flex";

    }

}

function ocultarLoading() {

    const overlay =
        document.getElementById(
            "loadingSistema"
        );

    if (
        overlay
    ) {

        overlay.style.display =
            "none";

    }

}

function mostrarError(
    mensaje
) {

    console.error(
        mensaje
    );

    setTexto(
        "kpiTotal",
        "0"
    );

    setTexto(
        "kpiPersonas",
        "0"
    );

    setTexto(
        "kpiUnidad",
        "-"
    );

    setTexto(
        "kpiReincidentes",
        "0"
    );

    setTexto(
        "textoResumenFiltros",
        mensaje
    );

}

function setTexto(
    id,
    valor
) {

    const elemento =
        document.getElementById(
            id
        );

    if (
        elemento
    ) {

        elemento.textContent =
            valor;

    }

}

function asignarTextoSelector(
    contenedor,
    selector,
    texto
) {

    const elemento =
        contenedor.querySelector(
            selector
        );

    if (
        elemento
    ) {

        elemento.textContent =
            texto;

    }

}


// =====================================================
// 🧰 UTILIDADES
// =====================================================
function formatearNumero(
    numero
) {

    return Number(
        numero || 0
    ).toLocaleString(
        "es-PE"
    );

}

function escapeHtml(
    valor
) {

    return String(
        valor ?? ""
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
                () => {

                    funcion(
                        ...argumentos
                    );

                },
                espera
            );

    };

}
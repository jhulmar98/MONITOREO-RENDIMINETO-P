// ============================================================
// REPORTE.JS
// DASHBOARD DIARIO DE DESCANSOS MÉDICOS
//
// FIRESTORE:
// registro-dm/{DD-MM-AAAA}/turnos/{turno-X}/registros/{ID}
// ============================================================


// ============================================================
// IMPORTACIONES
// ============================================================

import * as XLSX from "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/+esm";

import {
    db,
    collection,
    getDocs,
    obtenerArchivoDescansosMedicos,
    descargarExcelStorage
} from "../backend/api.js";
// ============================================================
// CONFIGURACIÓN DE TURNOS
// ============================================================

const TURNOS = {

    "turno-1": {
        nombre: "Turno 1",
        horario: "07:00 a 15:00"
    },

    "turno-2": {
        nombre: "Turno 2",
        horario: "15:00 a 22:00"
    },

    "turno-3": {
        nombre: "Turno 3",
        horario: "22:00 a 07:00"
    }

};
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

const RUTA_ROL_SERVICIO =
    "RENDIMIENTO PERSONAL/ROL DE SERVICIO";

const ARCHIVOS_ROL = {
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

const TURNOS_VALIDOS = new Set([
    "T1",
    "T2",
    "T3"
]);

// ============================================================
// ELEMENTOS HTML
// ============================================================

const filtroFecha =
    document.getElementById("filtroFecha");

const filtroTurno =
    document.getElementById("filtroTurno");

const btnBuscar =
    document.getElementById("btnBuscar");

const btnActualizar =
    document.getElementById("btnActualizar");

const mensajeSistema =
    document.getElementById("mensajeSistema");


// KPI
const kpiTotal =
    document.getElementById("kpiTotal");

const kpiTurno1 =
    document.getElementById("kpiTurno1");

const kpiTurno2 =
    document.getElementById("kpiTurno2");

const kpiTurno3 =
    document.getElementById("kpiTurno3");


// CONTADORES POR TURNO
const totalTarjetaTurno1 =
    document.getElementById("totalTarjetaTurno1");

const totalTarjetaTurno2 =
    document.getElementById("totalTarjetaTurno2");

const totalTarjetaTurno3 =
    document.getElementById("totalTarjetaTurno3");


// LISTAS POR TURNO
const listaResumenTurno1 =
    document.getElementById("listaResumenTurno1");

const listaResumenTurno2 =
    document.getElementById("listaResumenTurno2");

const listaResumenTurno3 =
    document.getElementById("listaResumenTurno3");

const textoFechaResumen =
    document.getElementById("textoFechaResumen");


// PLANTILLA
const templatePersonaResumen =
    document.getElementById("templatePersonaResumen");


// PANTALLA DE CARGA
const pantallaCarga =
    document.getElementById("pantallaCarga");

const textoCarga =
    document.getElementById("textoCarga");


// MODAL DE DETALLE
const modalDetalle =
    document.getElementById("modalDetalle");

const btnCerrarModal =
    document.getElementById("btnCerrarModal");

const detalleNombreRegistrado =
    document.getElementById("detalleNombreRegistrado");

const detalleDniRegistrado =
    document.getElementById("detalleDniRegistrado");

const detalleFecha =
    document.getElementById("detalleFecha");

const detalleHora =
    document.getElementById("detalleHora");

const detalleTurno =
    document.getElementById("detalleTurno");

const detalleReportanteNombre =
    document.getElementById("detalleReportanteNombre");

const detalleReportanteDni =
    document.getElementById("detalleReportanteDni");

const contenedorFotoDetalle =
    document.getElementById("contenedorFotoDetalle");

const detalleFoto =
    document.getElementById("detalleFoto");

const detalleSinFoto =
    document.getElementById("detalleSinFoto");

// MODAL DE ALERTA DM
const modalAlertaDM =
    document.getElementById("modalAlertaDM");

const btnCerrarModalAlerta =
    document.getElementById("btnCerrarModalAlerta");

const tituloModalAlertaDM =
    document.getElementById("tituloModalAlertaDM");

const subtituloModalAlertaDM =
    document.getElementById("subtituloModalAlertaDM");

const alertaNombrePersonal =
    document.getElementById("alertaNombrePersonal");

const alertaDniPersonal =
    document.getElementById("alertaDniPersonal");

const alertaTipoContenedor =
    document.getElementById("alertaTipoContenedor");

const alertaTipoTexto =
    document.getElementById("alertaTipoTexto");

const listaHistorialAlertaDM =
    document.getElementById("listaHistorialAlertaDM");

const alertaTotalHistorial =
    document.getElementById("alertaTotalHistorial");

const alertaDescripcion =
    document.getElementById("alertaDescripcion");
// ============================================================
// ESTADO
// ============================================================

const estado = {

    registros: [],

    registrosVisibles: [],

    descansosMedicos: [],

    rolesPorMes: new Map(),

    alertasPorDNI: new Map(),

    cargandoDatosAlertas: false,

    alertasCargadas: false,

    cargando: false

};


// ============================================================
// INICIO
// ============================================================

document.addEventListener(
    "DOMContentLoaded",
    iniciarSistema
);


async function iniciarSistema() {

    configurarEventos();

    colocarFechaActual();

    await cargarDatosAlertas();

    await consultarRegistros();

}


// ============================================================
// EVENTOS
// ============================================================

function configurarEventos() {

    btnBuscar.addEventListener(
        "click",
        consultarRegistros
    );


    btnActualizar.addEventListener(
        "click",
        async () => {

            if (
                estado.cargando ||
                estado.cargandoDatosAlertas
            ) {
                return;
            }

            estado.alertasCargadas = false;
            estado.descansosMedicos = [];
            estado.rolesPorMes.clear();
            estado.alertasPorDNI.clear();

            await cargarDatosAlertas();
            await consultarRegistros();

        }
    );


    filtroFecha.addEventListener(
        "change",
        consultarRegistros
    );


    filtroTurno.addEventListener(
        "change",
        aplicarFiltroTurno
    );


    btnCerrarModal.addEventListener(
        "click",
        cerrarModalDetalle
    );


    modalDetalle
        .querySelectorAll("[data-cerrar-modal]")
        .forEach(elemento => {

            elemento.addEventListener(
                "click",
                cerrarModalDetalle
            );

        });


    document.addEventListener(
        "keydown",
        evento => {

            if (evento.key !== "Escape") {
                return;
            }

            if (
                modalAlertaDM &&
                !modalAlertaDM.hidden
            ) {

                cerrarModalAlertaDM();
                return;
            }

            if (
                modalDetalle &&
                !modalDetalle.hidden
            ) {

                cerrarModalDetalle();
            }
        }
    );

    if (btnCerrarModalAlerta) {

        btnCerrarModalAlerta.addEventListener(
            "click",
            cerrarModalAlertaDM
        );

    }

    if (modalAlertaDM) {

        modalAlertaDM
            .querySelectorAll("[data-cerrar-alerta]")
            .forEach(elemento => {

                elemento.addEventListener(
                    "click",
                    cerrarModalAlertaDM
                );

            });

    }

}

// ============================================================
// ABRIR MODAL DE ALERTA DM
// ============================================================

// ============================================================
// ABRIR MODAL DE ALERTA DM
// ============================================================

function abrirModalAlertaDM(
    alerta,
    registro
) {

    if (
        !modalAlertaDM ||
        !alerta
    ) {
        return;
    }

    const casos =
        Array.isArray(alerta.casos)
            ? [...alerta.casos]
            : [];

    casos.sort((a, b) => {

        const fechaA =
            a.fechaPosterior instanceof Date
                ? a.fechaPosterior.getTime()
                : 0;

        const fechaB =
            b.fechaPosterior instanceof Date
                ? b.fechaPosterior.getTime()
                : 0;

        return fechaA - fechaB;

    });

    alertaNombrePersonal.textContent =
        registro?.registrado?.nombre ||
        alerta.nombre ||
        "Personal sin nombre";

    alertaDniPersonal.textContent =
        `DNI: ${
            registro?.registrado?.dni ||
            alerta.dni ||
            "-"
        }`;

    alertaTipoContenedor.classList.remove(
        "modal-alerta__tipo--posterior",
        "modal-alerta__tipo--reincidencia"
    );

    if (
        alerta.tipo === "reincidencia"
    ) {

        tituloModalAlertaDM.textContent =
            "Reincidencia posterior al descanso médico";

        subtituloModalAlertaDM.textContent =
            `${casos.length} casos detectados`;

        alertaTipoTexto.textContent =
            "Reincidencia posterior al descanso médico";

        alertaTipoContenedor.classList.add(
            "modal-alerta__tipo--reincidencia"
        );

        

    } else {

        tituloModalAlertaDM.textContent =
            "Posterior al descanso médico";

        subtituloModalAlertaDM.textContent =
            "1 caso detectado";

        alertaTipoTexto.textContent =
            "Posterior al descanso médico";

        alertaTipoContenedor.classList.add(
            "modal-alerta__tipo--posterior"
        );

        alertaDescripcion.textContent =
            "El personal presenta una marca distinta de T1, T2 o T3 al día siguiente de finalizar su descanso médico.";

    }

    renderizarHistorialAlertaDM(
        casos
    );

    modalAlertaDM.hidden =
        false;

    document.body.classList.add(
        "modal-abierto"
    );

}
// ============================================================
// RENDERIZAR TODOS LOS CASOS DE LA ALERTA
// ============================================================

function renderizarHistorialAlertaDM(
    casos
) {

    if (
        !listaHistorialAlertaDM ||
        !alertaTotalHistorial
    ) {
        return;
    }

    const cantidad =
        casos.length;

    alertaTotalHistorial.textContent =
        `${cantidad} ${
            cantidad === 1
                ? "evento"
                : "eventos"
        }`;

    if (!cantidad) {

        listaHistorialAlertaDM.innerHTML = `
            <div class="estado-vacio estado-vacio--pequeno">

                <i class="fa-regular fa-folder-open"></i>

                <span>
                    No existen eventos posteriores registrados.
                </span>

            </div>
        `;

        return;

    }

    listaHistorialAlertaDM.innerHTML =
        casos.map((caso, index) => {

            const motivo =
                caso.motivo ||
                caso.registroDM?.motivo ||
                "NO ESPECIFICADO";

            const marca =
                caso.marcaPosterior ||
                "-";

            return `
                <article class="historial-alerta-item">

                    <div class="historial-alerta-item__numero">
                        ${index + 1}
                    </div>

                    <div class="historial-alerta-item__contenido">

                        <div class="historial-alerta-item__fila">

                            <div>

                                <span>
                                    Inicio del descanso médico
                                </span>

                                <strong>
                                    ${escapeHtmlSeguro(
                                        caso.fechaInicioTexto ||
                                        "-"
                                    )}
                                </strong>

                            </div>

                            <div>

                                <span>
                                    Fin del descanso médico
                                </span>

                                <strong>
                                    ${escapeHtmlSeguro(
                                        caso.fechaFinTexto ||
                                        "-"
                                    )}
                                </strong>

                            </div>

                        </div>

                        <div class="historial-alerta-item__motivo">

                            <span>
                                Motivo o diagnóstico
                            </span>

                            <strong>
                                ${escapeHtmlSeguro(
                                    motivo
                                )}
                            </strong>

                        </div>

                        <div class="historial-alerta-item__fila">

                            <div>

                                <span>
                                    Fecha posterior evaluada
                                </span>

                                <strong>
                                    ${escapeHtmlSeguro(
                                        caso.fechaPosteriorTexto ||
                                        "-"
                                    )}
                                </strong>

                            </div>

                            <div>

                                <span>
                                    Mes
                                </span>

                                <strong>
                                    ${escapeHtmlSeguro(
                                        caso.mesPosterior ||
                                        "-"
                                    )}
                                </strong>

                            </div>

                        </div>

                    </div>

                    <div class="historial-alerta-item__marca">

                        <span>
                            Marca posterior encontrada
                        </span>

                        <strong>
                            ${escapeHtmlSeguro(
                                marca
                            )}
                        </strong>

                        <small>
                            ${escapeHtmlSeguro(
                                obtenerDescripcionMarcaPosterior(
                                    marca
                                )
                            )}
                        </small>

                    </div>

                </article>
            `;

        }).join("");

}

// ============================================================
// CERRAR MODAL DE ALERTA
// ============================================================

function cerrarModalAlertaDM() {

    if (!modalAlertaDM) {
        return;
    }

    modalAlertaDM.hidden = true;

    document.body.classList.remove(
        "modal-abierto"
    );
}


// ============================================================
// DESCRIPCIÓN DE MARCA POSTERIOR
// ============================================================

function obtenerDescripcionMarcaPosterior(
    marca
) {

    const valor =
        normalizarMarcaRol(marca);

    if (!valor) {
        return "-";
    }

    if (
        TURNOS_VALIDOS.has(valor)
    ) {

        return `${valor} — turno normal`;
    }

    return `${valor} — marca distinta de T1, T2 o T3`;
}
// ============================================================
// CONSULTAR FIRESTORE
// ============================================================

async function consultarRegistros() {

    if (estado.cargando) {
        return;
    }


    const fechaISO =
        filtroFecha.value;


    if (!fechaISO) {

        mostrarMensaje(
            "Seleccione una fecha para consultar.",
            "error"
        );

        return;

    }


    estado.cargando = true;

    activarCarga(
        true,
        "Consultando descansos médicos..."
    );

    ocultarMensaje();


    try {

        const fechaDocumento =
            convertirFechaADocumento(fechaISO);


        const consultas =
            Object.keys(TURNOS).map(
                async turno => {

                    const referencia =
                        collection(
                            db,
                            "registro-dm",
                            fechaDocumento,
                            "turnos",
                            turno,
                            "registros"
                        );


                    const resultado =
                        await getDocs(referencia);


                    return resultado.docs.map(
                        documento => {

                            return normalizarRegistro(
                                documento.id,
                                documento.data(),
                                turno,
                                fechaDocumento
                            );

                        }
                    );

                }
            );


        const resultados =
            await Promise.all(consultas);


        estado.registros =
            resultados.flat();


        ordenarRegistros(
            estado.registros
        );


        actualizarTextoFecha(
            fechaDocumento
        );


        aplicarFiltroTurno();


        if (
            estado.registros.length === 0
        ) {

            mostrarMensaje(
                "No existen descansos médicos registrados en esta fecha.",
                "advertencia"
            );

        }

    } catch (error) {

        console.error(
            "Error consultando registros:",
            error
        );


        estado.registros = [];
        estado.registrosVisibles = [];


        actualizarDashboard();


        mostrarMensaje(
            obtenerMensajeError(error),
            "error"
        );

    } finally {

        estado.cargando = false;

        activarCarga(false);

    }

}


// ============================================================
// NORMALIZAR REGISTRO
// ============================================================

function normalizarRegistro(
    id,
    datos,
    turnoRuta,
    fechaRuta
) {

    const registrado =
        datos.registrado ||
        datos.personal_con_descanso_medico ||
        datos.personalDM ||
        {};


    const reportante =
        datos.reportante ||
        datos.personal_que_reporta ||
        {};


    return {

        id,

        fecha: String(
            datos.fecha ||
            fechaRuta ||
            ""
        ).trim(),

        hora: String(
            datos.hora ||
            ""
        ).trim(),

        turno: String(
            datos.turno ||
            turnoRuta ||
            ""
        ).trim(),

        registrado: {

            dni: String(
                registrado.dni ||
                datos.registrado_dni ||
                datos.personal_dni ||
                ""
            ).trim(),

            nombre: String(
                registrado.nombre ||
                registrado.apellidos_nombres ||
                datos.registrado_nombre ||
                datos.personal_nombre ||
                ""
            ).trim()

        },

        reportante: {

            dni: String(
                reportante.dni ||
                datos.reportante_dni ||
                ""
            ).trim(),

            nombre: String(
                reportante.nombre ||
                reportante.apellidos_nombres ||
                datos.reportante_nombre ||
                ""
            ).trim()

        },

        fotoUrl: String(
            datos.foto_url ||
            datos.fotoUrl ||
            datos.imagen_url ||
            ""
        ).trim()

    };

}
// ============================================================
// CARGAR DATOS PARA LAS ALERTAS
// DM.XLSX + ROLES DE SERVICIO
// ============================================================

async function cargarDatosAlertas() {

    if (
        estado.cargandoDatosAlertas ||
        estado.alertasCargadas
    ) {
        return;
    }

    estado.cargandoDatosAlertas = true;

    activarCarga(
        true,
        "Analizando descansos médicos y roles de servicio..."
    );

    try {

        await cargarDescansosMedicos();

        await cargarRolesNecesarios();

        construirAlertasPorDNI();

        estado.alertasCargadas = true;

        console.log(
            "Alertas posteriores al DM cargadas:",
            estado.alertasPorDNI.size
        );

    } catch (error) {

        console.error(
            "Error cargando alertas posteriores al DM:",
            error
        );

        estado.descansosMedicos = [];
        estado.rolesPorMes.clear();
        estado.alertasPorDNI.clear();

        /*
         * El reporte diario continuará funcionando aunque
         * falle la descarga de DM.xlsx o de los roles.
         */
        mostrarMensaje(
            "Los registros diarios se cargarán, pero no fue posible analizar las alertas posteriores al descanso médico.",
            "advertencia"
        );

    } finally {

        estado.cargandoDatosAlertas = false;

        activarCarga(false);

    }

}


// ============================================================
// CARGAR DM.XLSX
// ============================================================

async function cargarDescansosMedicos() {

    const buffer =
        await obtenerArchivoDescansosMedicos();

    if (
        !(buffer instanceof ArrayBuffer) ||
        buffer.byteLength === 0
    ) {

        throw new Error(
            "El archivo DM.xlsx está vacío o no es válido."
        );

    }

    const workbook = XLSX.read(
        buffer,
        {
            type: "array",
            cellDates: true
        }
    );

    if (!workbook.SheetNames.length) {

        throw new Error(
            "DM.xlsx no contiene hojas."
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

    const filas =
        XLSX.utils.sheet_to_json(
            hoja,
            {
                defval: "",
                raw: true
            }
        );

    const registros =
        filas
            .map((fila, indice) =>
                normalizarDescansoMedico(
                    fila,
                    indice + 2
                )
            )
            .filter(registro =>
                registro.dni &&
                registro.fechaInicio instanceof Date &&
                registro.fechaFin instanceof Date
            );

    estado.descansosMedicos =
        eliminarDescansosDuplicados(
            registros
        );

    console.log(
        "DM.xlsx procesado:",
        {
            filas: filas.length,
            registrosValidos: registros.length,
            registrosSinDuplicados:
                estado.descansosMedicos.length
        }
    );

}
// ============================================================
// NORMALIZAR DESCANSO MÉDICO
// ============================================================

function normalizarDescansoMedico(
    fila,
    filaExcel
) {

    const dni =
        limpiarDocumento(
            obtenerCampoFlexible(
                fila,
                [
                    "numdocumento",
                    "NUMDOCUMENTO",
                    "NUM_DOCUMENTO",
                    "DNI",
                    "dni",
                    "DOCUMENTO",
                    "Documento"
                ]
            )
        );

    const nombre =
        normalizarTexto(
            obtenerCampoFlexible(
                fila,
                [
                    "des_nomapell",
                    "DES_NOMAPELL",
                    "Des_NomApell",
                    "APELLIDOS Y NOMBRES",
                    "NOMBRES",
                    "PERSONAL"
                ]
            )
        );

    const fechaInicio =
        convertirFechaExcel(
            obtenerCampoFlexible(
                fila,
                [
                    "fec_inicio",
                    "FEC_INICIO",
                    "fec inicio",
                    "FECHA DE INICIO",
                    "FECHA INICIO"
                ]
            )
        );

    const fechaFin =
        convertirFechaExcel(
            obtenerCampoFlexible(
                fila,
                [
                    "fec_termino",
                    "FEC_TERMINO",
                    "fec termino",
                    "FECHA DE FIN",
                    "FECHA FIN",
                    "FECHA TÉRMINO",
                    "FECHA TERMINO"
                ]
            )
        );

    const numeroDias =
        Number(
            obtenerCampoFlexible(
                fila,
                [
                    "cant. Días",
                    "cant. días",
                    "CANT. DÍAS",
                    "CANT. DIAS",
                    "cant_dias",
                    "CANT_DIAS",
                    "DIAS",
                    "DÍAS"
                ]
            )
        ) || 0;

    const motivo =
        normalizarTexto(
            obtenerCampoFlexible(
                fila,
                [
                    "txtotrosdiagnosticos",
                    "TXTOTROSDIAGNOSTICOS",
                    "txt otros diagnosticos",
                    "OTROS DIAGNOSTICOS",
                    "OTROS DIAGNÓSTICOS",
                    "DIAGNOSTICO",
                    "DIAGNÓSTICO",
                    "MOTIVO",
                    "MOTIVO DEL DESCANSO",
                    "MOTIVO DESCANSO"
                ]
            )
        );

    return {

        filaExcel,

        dni,

        dniClave:
            normalizarDNIComparacion(
                dni
            ),

        nombre,

        fechaInicio,

        fechaFin,

        fechaInicioTexto:
            formatearFechaCorta(
                fechaInicio
            ),

        fechaFinTexto:
            formatearFechaCorta(
                fechaFin
            ),

        numeroDias,

        motivo:
            motivo ||
            "NO ESPECIFICADO"

    };

}

// ============================================================
// ELIMINAR DESCANSOS DUPLICADOS
// ============================================================

function eliminarDescansosDuplicados(
    registros
) {

    const mapa = new Map();

    registros.forEach(registro => {

        const clave = [
            registro.dniClave,
            registro.fechaInicioTexto,
            registro.fechaFinTexto,
            registro.numeroDias,
            registro.motivo
        ].join("|");

        if (!mapa.has(clave)) {

            mapa.set(
                clave,
                registro
            );

        }

    });

    return [...mapa.values()];

}


// ============================================================
// CARGAR ROLES NECESARIOS
// ============================================================

async function cargarRolesNecesarios() {

    estado.rolesPorMes.clear();

    const mesesNecesarios =
        obtenerMesesNecesariosParaAlertas();

    if (!mesesNecesarios.length) {
        return;
    }

    const tareas =
        mesesNecesarios.map(
            mesNumero =>
                cargarRolMensual(
                    mesNumero
                )
        );

    await Promise.all(tareas);

}


// ============================================================
// OBTENER MESES NECESARIOS
// ============================================================

function obtenerMesesNecesariosParaAlertas() {

    const meses = new Set();

    estado.descansosMedicos.forEach(
        descanso => {

            if (
                !(descanso.fechaFin instanceof Date) ||
                Number.isNaN(
                    descanso.fechaFin.getTime()
                )
            ) {
                return;
            }

            const fechaPosterior =
                sumarDiasFecha(
                    descanso.fechaFin,
                    1
                );

            meses.add(
                fechaPosterior.getMonth() + 1
            );

        }
    );

    return [...meses]
        .sort((a, b) => a - b);

}


// ============================================================
// CARGAR ROL DE UN MES
// ============================================================

async function cargarRolMensual(
    mesNumero
) {

    if (
        estado.rolesPorMes.has(
            mesNumero
        )
    ) {

        return estado.rolesPorMes.get(
            mesNumero
        );

    }

    const nombresArchivo =
        ARCHIVOS_ROL[mesNumero] || [];

    for (const nombreArchivo of nombresArchivo) {

        const ruta =
            `${RUTA_ROL_SERVICIO}/${nombreArchivo}`;

        try {

            const buffer =
                await descargarExcelStorage(
                    ruta
                );

            if (
                !(buffer instanceof ArrayBuffer) ||
                buffer.byteLength === 0
            ) {

                continue;

            }

            const personas =
                procesarArchivoRol(
                    buffer,
                    mesNumero,
                    ruta
                );

            estado.rolesPorMes.set(
                mesNumero,
                personas
            );

            console.log(
                `Rol cargado: ${ruta}`,
                `Personas: ${personas.size}`
            );

            return personas;

        } catch (error) {

            console.warn(
                `No se pudo cargar ${ruta}:`,
                error
            );

        }

    }

    const mapaVacio =
        new Map();

    estado.rolesPorMes.set(
        mesNumero,
        mapaVacio
    );

    console.warn(
        `No se encontró el rol de ${MESES_ORDEN[mesNumero - 1]}.`
    );

    return mapaVacio;

}


// ============================================================
// PROCESAR ARCHIVO DEL ROL
// ============================================================

function procesarArchivoRol(
    buffer,
    mesNumero,
    rutaArchivo
) {

    const workbook =
        XLSX.read(
            buffer,
            {
                type: "array",
                cellDates: false
            }
        );

    const mapaPersonas =
        new Map();

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
                        header: 1,
                        defval: "",
                        raw: false
                    }
                );

            if (!filas.length) {
                return;
            }

            const indiceCabecera =
                localizarCabeceraRol(
                    filas
                );

            if (indiceCabecera < 0) {
                return;
            }

            const cabeceraOriginal =
                filas[indiceCabecera] || [];

            const cabeceraNormalizada =
                cabeceraOriginal.map(
                    normalizarCabeceraRol
                );

            const indiceDNI =
                buscarIndiceCabeceraRol(
                    cabeceraNormalizada,
                    [
                        "DNI",
                        "NUMDOCUMENTO",
                        "NUM DOCUMENTO",
                        "DOCUMENTO"
                    ]
                );

            const indiceNombre =
                buscarIndiceCabeceraRol(
                    cabeceraNormalizada,
                    [
                        "NOMBRE",
                        "NOMBRES",
                        "APELLIDOS Y NOMBRES",
                        "PERSONAL",
                        "DES NOMAPELL"
                    ]
                );

            if (indiceDNI < 0) {
                return;
            }

            const columnasDias =
                obtenerColumnasDiasRol(
                    cabeceraOriginal
                );

            for (
                let indiceFila =
                    indiceCabecera + 1;

                indiceFila <
                    filas.length;

                indiceFila++
            ) {

                const fila =
                    filas[indiceFila] || [];

                const dni =
                    limpiarDocumento(
                        fila[indiceDNI]
                    );

                const dniClave =
                    normalizarDNIComparacion(
                        dni
                    );

                if (!dniClave) {
                    continue;
                }

                const nombre =
                    indiceNombre >= 0
                        ? normalizarTexto(
                            fila[indiceNombre]
                        )
                        : "";

                const dias = {};

                columnasDias.forEach(
                    ({ indice, dia }) => {

                        dias[dia] =
                            normalizarMarcaRol(
                                fila[indice]
                            );

                    }
                );

                if (
                    !mapaPersonas.has(
                        dniClave
                    )
                ) {

                    mapaPersonas.set(
                        dniClave,
                        {
                            dni,
                            nombre,
                            mesNumero,
                            dias,
                            hoja: nombreHoja,
                            filaExcel:
                                indiceFila + 1,
                            rutaArchivo
                        }
                    );

                    continue;

                }

                const existente =
                    mapaPersonas.get(
                        dniClave
                    );

                Object.entries(dias)
                    .forEach(
                        ([dia, marca]) => {

                            if (
                                marca &&
                                !existente.dias[dia]
                            ) {

                                existente.dias[dia] =
                                    marca;

                            }

                        }
                    );

                if (
                    !existente.nombre &&
                    nombre
                ) {

                    existente.nombre =
                        nombre;

                }

            }

        }
    );

    return mapaPersonas;

}


// ============================================================
// LOCALIZAR CABECERA DEL ROL
// ============================================================

function localizarCabeceraRol(
    filas
) {

    const limite =
        Math.min(
            filas.length,
            30
        );

    for (
        let indice = 0;
        indice < limite;
        indice++
    ) {

        const fila =
            filas[indice] || [];

        const normalizada =
            fila.map(
                normalizarCabeceraRol
            );

        const tieneDNI =
            normalizada.some(
                valor =>
                    [
                        "DNI",
                        "NUMDOCUMENTO",
                        "NUM DOCUMENTO",
                        "DOCUMENTO"
                    ].includes(valor)
            );

        const cantidadDias =
            fila.filter(valor => {

                const dia =
                    obtenerNumeroDiaRol(
                        valor
                    );

                return (
                    dia >= 1 &&
                    dia <= 31
                );

            }).length;

        if (
            tieneDNI &&
            cantidadDias >= 5
        ) {

            return indice;

        }

    }

    return -1;

}


// ============================================================
// OBTENER COLUMNAS DE DÍAS DEL ROL
// ============================================================

function obtenerColumnasDiasRol(
    cabecera
) {

    const columnas = [];

    cabecera.forEach(
        (valor, indice) => {

            const dia =
                obtenerNumeroDiaRol(
                    valor
                );

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


// ============================================================
// OBTENER NÚMERO DEL DÍA
// ============================================================

function obtenerNumeroDiaRol(
    valor
) {

    const texto =
        String(valor ?? "")
            .trim()
            .replace(/\.0$/, "");

    if (
        !/^\d{1,2}$/.test(
            texto
        )
    ) {

        return 0;

    }

    const numero =
        Number(texto);

    return (
        numero >= 1 &&
        numero <= 31
    )
        ? numero
        : 0;

}


// ============================================================
// BUSCAR COLUMNA DE CABECERA
// ============================================================

function buscarIndiceCabeceraRol(
    cabecera,
    alternativas
) {

    for (
        const alternativa
        of alternativas
    ) {

        const buscado =
            normalizarCabeceraRol(
                alternativa
            );

        const indice =
            cabecera.indexOf(
                buscado
            );

        if (indice >= 0) {
            return indice;
        }

    }

    return -1;

}


// ============================================================
// NORMALIZAR CABECERA DEL ROL
// ============================================================

function normalizarCabeceraRol(
    valor
) {

    return String(valor ?? "")
        .trim()
        .replace(/[_-]+/g, " ")
        .replace(/\s+/g, " ")
        .normalize("NFD")
        .replace(
            /[\u0300-\u036f]/g,
            ""
        )
        .toUpperCase();

}


// ============================================================
// NORMALIZAR MARCA DEL ROL
// ============================================================

function normalizarMarcaRol(
    valor
) {

    return String(valor ?? "")
        .trim()
        .replace(/\s+/g, "")
        .toUpperCase();

}


// ============================================================
// CONSTRUIR ALERTAS POR DNI
// ============================================================

function construirAlertasPorDNI() {

    estado.alertasPorDNI.clear();

    const casos =
        detectarCasosPosterioresDM();

    const agrupados =
        new Map();

    casos.forEach(caso => {

        const dniClave =
            normalizarDNIComparacion(
                caso.dni
            );

        if (!dniClave) {
            return;
        }

        if (
            !agrupados.has(
                dniClave
            )
        ) {

            agrupados.set(
                dniClave,
                {
                    dni: caso.dni,
                    dniClave,
                    nombre:
                        caso.nombre ||
                        "NO ESPECIFICADO",
                    cantidad: 0,
                    casos: []
                }
            );

        }

        const persona =
            agrupados.get(
                dniClave
            );

        persona.cantidad += 1;

        persona.casos.push(
            caso
        );

    });

    agrupados.forEach(
        persona => {

            persona.casos.sort(
                (a, b) =>
                    a.fechaPosterior.getTime() -
                    b.fechaPosterior.getTime()
            );

            const ultimoCaso =
                persona.casos[
                    persona.casos.length - 1
                ];

            estado.alertasPorDNI.set(
                persona.dniClave,
                {
                    ...persona,

                    tipo:
                        persona.cantidad >= 1
                            ? "reincidencia"
                            : "posterior",

                    ultimoCaso
                }
            );

        }
    );

}


// ============================================================
// DETECTAR CASOS POSTERIORES AL DM
// ============================================================

function detectarCasosPosterioresDM() {

    const resultados = [];

    const clavesAgregadas =
        new Set();

    estado.descansosMedicos.forEach(
        descanso => {

            if (
                !descanso.dniClave ||
                !(descanso.fechaFin instanceof Date) ||
                Number.isNaN(
                    descanso.fechaFin.getTime()
                )
            ) {

                return;

            }

            const fechaPosterior =
                sumarDiasFecha(
                    descanso.fechaFin,
                    1
                );

            const mesPosterior =
                fechaPosterior.getMonth() + 1;

            const diaPosterior =
                fechaPosterior.getDate();

            const rolMes =
                estado.rolesPorMes.get(
                    mesPosterior
                );

            if (
                !(rolMes instanceof Map) ||
                !rolMes.size
            ) {

                return;

            }

            const personaRol =
                rolMes.get(
                    descanso.dniClave
                );

            if (!personaRol) {
                return;
            }

            const marcaPosterior =
                normalizarMarcaRol(
                    personaRol.dias[
                        diaPosterior
                    ]
                );

            /*
             * Misma lógica que el dashboard principal:
             *
             * - Celda vacía: no genera alerta.
             * - T1, T2 o T3: no genera alerta.
             * - Cualquier otra marca: genera alerta.
             */
            if (!marcaPosterior) {
                return;
            }

            if (
                TURNOS_VALIDOS.has(
                    marcaPosterior
                )
            ) {

                return;

            }

            const fechaPosteriorTexto =
                formatearFechaCorta(
                    fechaPosterior
                );

            const clave = [
                descanso.dniClave,
                descanso.fechaInicioTexto,
                descanso.fechaFinTexto,
                fechaPosteriorTexto,
                marcaPosterior
            ].join("|");

            if (
                clavesAgregadas.has(
                    clave
                )
            ) {

                return;

            }

            clavesAgregadas.add(
                clave
            );

            resultados.push({

                dni:
                    descanso.dni,

                dniClave:
                    descanso.dniClave,


                nombre:
                    descanso.nombre ||
                    personaRol.nombre ||
                    "NO ESPECIFICADO",
                 motivo:

                    descanso.motivo ||

                    "NO ESPECIFICADO",

                fechaInicio:
                    descanso.fechaInicio,

                fechaFin:
                    descanso.fechaFin,

                fechaInicioTexto:
                    descanso.fechaInicioTexto,

                fechaFinTexto:
                    descanso.fechaFinTexto,

                fechaPosterior,

                fechaPosteriorTexto,

                mesPosteriorNumero:
                    mesPosterior,

                mesPosterior:
                    MESES_ORDEN[
                        mesPosterior - 1
                    ],

                diaPosterior,

                marcaPosterior,

                registroDM:
                    descanso

            });

        }
    );

    return resultados.sort(
        (a, b) => {

            const diferenciaFecha =
                a.fechaPosterior.getTime() -
                b.fechaPosterior.getTime();

            if (
                diferenciaFecha !== 0
            ) {

                return diferenciaFecha;

            }

            return a.nombre.localeCompare(
                b.nombre,
                "es"
            );

        }
    );

}


// ============================================================
// OBTENER ALERTA DE UN DNI
// ============================================================

function obtenerAlertaPorDNI(
    dni
) {

    const dniClave =
        normalizarDNIComparacion(
            dni
        );

    if (!dniClave) {
        return null;
    }

    return (
        estado.alertasPorDNI.get(
            dniClave
        ) || null
    );

}


// ============================================================
// OBTENER CAMPO FLEXIBLE
// ============================================================

function obtenerCampoFlexible(
    objeto,
    alternativas
) {

    for (
        const alternativa
        of alternativas
    ) {

        if (
            objeto[alternativa] !==
                undefined &&
            objeto[alternativa] !==
                null &&
            String(
                objeto[alternativa]
            ).trim() !== ""
        ) {

            return objeto[
                alternativa
            ];

        }

    }

    /*
     * Segundo intento:
     * compara cabeceras sin tildes,
     * guiones, mayúsculas o espacios.
     */
    const mapaCampos =
        new Map();

    Object.keys(objeto)
        .forEach(campo => {

            mapaCampos.set(
                normalizarCabeceraRol(
                    campo
                ),
                campo
            );

        });

    for (
        const alternativa
        of alternativas
    ) {

        const campoReal =
            mapaCampos.get(
                normalizarCabeceraRol(
                    alternativa
                )
            );

        if (
            campoReal &&
            objeto[campoReal] !==
                undefined &&
            objeto[campoReal] !==
                null &&
            String(
                objeto[campoReal]
            ).trim() !== ""
        ) {

            return objeto[
                campoReal
            ];

        }

    }

    return "";

}


// ============================================================
// NORMALIZAR TEXTO
// ============================================================

function normalizarTexto(
    valor
) {

    return String(valor || "")
        .trim()
        .replace(/\s+/g, " ")
        .toUpperCase();

}


// ============================================================
// LIMPIAR DOCUMENTO
// ============================================================

function limpiarDocumento(
    valor
) {

    return String(valor || "")
        .trim()
        .replace(/\D/g, "");

}


// ============================================================
// NORMALIZAR DNI PARA COMPARACIÓN
// ============================================================

function normalizarDNIComparacion(
    valor
) {

    const documento =
        limpiarDocumento(
            valor
        );

    if (!documento) {
        return "";
    }

    return (
        documento.replace(
            /^0+/,
            ""
        ) || "0"
    );

}


// ============================================================
// CONVERTIR FECHA DE EXCEL
// ============================================================

function convertirFechaExcel(
    valor
) {

    if (!valor) {
        return null;
    }

    if (
        valor instanceof Date &&
        !Number.isNaN(
            valor.getTime()
        )
    ) {

        return normalizarFechaDate(
            valor
        );

    }

    if (
        typeof valor === "number"
    ) {

        const fecha =
            new Date(
                Math.round(
                    (
                        valor -
                        25569
                    ) *
                    86400 *
                    1000
                )
            );

        return normalizarFechaDate(
            fecha
        );

    }

    const texto =
        String(valor)
            .trim()
            .split(" ")[0]
            .replaceAll("-", "/");

    const partes =
        texto.split("/");

    if (
        partes.length !== 3
    ) {

        return null;

    }

    let dia;
    let mes;
    let anio;

    /*
     * Formato YYYY/MM/DD
     */
    if (
        partes[0].length === 4
    ) {

        anio = Number(partes[0]);
        mes = Number(partes[1]);
        dia = Number(partes[2]);

    } else {

        /*
         * Formato DD/MM/YYYY
         */
        dia = Number(partes[0]);
        mes = Number(partes[1]);
        anio = Number(partes[2]);

    }

    if (
        anio < 100
    ) {

        anio += 2000;

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
        Number.isNaN(
            fecha.getTime()
        )
    ) {

        return null;

    }

    return normalizarFechaDate(
        fecha
    );

}


// ============================================================
// NORMALIZAR DATE
// ============================================================

function normalizarFechaDate(
    fecha
) {

    return new Date(
        fecha.getFullYear(),
        fecha.getMonth(),
        fecha.getDate()
    );

}


// ============================================================
// FORMATEAR FECHA CORTA
// ============================================================

function formatearFechaCorta(
    fecha
) {

    if (
        !(fecha instanceof Date) ||
        Number.isNaN(
            fecha.getTime()
        )
    ) {

        return "";

    }

    return new Intl.DateTimeFormat(
        "es-PE",
        {
            day: "2-digit",
            month: "2-digit",
            year: "numeric"
        }
    ).format(fecha);

}


// ============================================================
// SUMAR DÍAS A UNA FECHA
// ============================================================

function sumarDiasFecha(
    fecha,
    cantidad
) {

    const nuevaFecha =
        new Date(
            fecha.getFullYear(),
            fecha.getMonth(),
            fecha.getDate()
        );

    nuevaFecha.setDate(
        nuevaFecha.getDate() +
        cantidad
    );

    return nuevaFecha;

}

// ============================================================
// FILTRAR POR TURNO
// ============================================================

function aplicarFiltroTurno() {

    const turnoSeleccionado =
        filtroTurno.value;


    if (
        turnoSeleccionado === "todos"
    ) {

        estado.registrosVisibles =
            [...estado.registros];

    } else {

        estado.registrosVisibles =
            estado.registros.filter(
                registro =>
                    registro.turno ===
                    turnoSeleccionado
            );

    }


    actualizarDashboard();

}


// ============================================================
// ACTUALIZAR DASHBOARD
// ============================================================

function actualizarDashboard() {

    actualizarIndicadores();

    actualizarListasTurnos();

}


// ============================================================
// INDICADORES
// ============================================================

function actualizarIndicadores() {

    const totalTurno1 =
        contarPorTurno(
            estado.registros,
            "turno-1"
        );

    const totalTurno2 =
        contarPorTurno(
            estado.registros,
            "turno-2"
        );

    const totalTurno3 =
        contarPorTurno(
            estado.registros,
            "turno-3"
        );


    kpiTotal.textContent =
        estado.registros.length;

    kpiTurno1.textContent =
        totalTurno1;

    kpiTurno2.textContent =
        totalTurno2;

    kpiTurno3.textContent =
        totalTurno3;


    totalTarjetaTurno1.textContent =
        totalTurno1;

    totalTarjetaTurno2.textContent =
        totalTurno2;

    totalTarjetaTurno3.textContent =
        totalTurno3;

}


// ============================================================
// ACTUALIZAR LISTAS
// ============================================================

function actualizarListasTurnos() {

    const turnoSeleccionado =
        filtroTurno.value;


    const mostrarTurno1 =
        turnoSeleccionado === "todos" ||
        turnoSeleccionado === "turno-1";

    const mostrarTurno2 =
        turnoSeleccionado === "todos" ||
        turnoSeleccionado === "turno-2";

    const mostrarTurno3 =
        turnoSeleccionado === "todos" ||
        turnoSeleccionado === "turno-3";


    const tarjetaTurno1 =
        listaResumenTurno1.closest(
            ".tarjeta-turno"
        );

    const tarjetaTurno2 =
        listaResumenTurno2.closest(
            ".tarjeta-turno"
        );

    const tarjetaTurno3 =
        listaResumenTurno3.closest(
            ".tarjeta-turno"
        );


    tarjetaTurno1.hidden =
        !mostrarTurno1;

    tarjetaTurno2.hidden =
        !mostrarTurno2;

    tarjetaTurno3.hidden =
        !mostrarTurno3;


    renderizarListaTurno(
        listaResumenTurno1,
        obtenerRegistrosTurno("turno-1")
    );

    renderizarListaTurno(
        listaResumenTurno2,
        obtenerRegistrosTurno("turno-2")
    );

    renderizarListaTurno(
        listaResumenTurno3,
        obtenerRegistrosTurno("turno-3")
    );

}


// ============================================================
// OBTENER REGISTROS DE UN TURNO
// ============================================================

function obtenerRegistrosTurno(turno) {

    return estado.registros.filter(
        registro =>
            registro.turno === turno
    );

}


// ============================================================
// RENDERIZAR PERSONAS
// ============================================================
function renderizarListaTurno(
    contenedor,
    registros
) {

    contenedor.innerHTML = "";

    if (registros.length === 0) {

        contenedor.innerHTML = `
            <div class="estado-vacio estado-vacio--pequeno">
                <i class="fa-regular fa-folder-open"></i>
                <span>Sin registros</span>
            </div>
        `;

        return;
    }

    const fragmento =
        document.createDocumentFragment();

    registros.forEach(registro => {

        const clon =
            templatePersonaResumen.content
                .cloneNode(true);

        const boton =
            clon.querySelector(".resumen-persona");

        const nombre =
            clon.querySelector(".resumen-persona__nombre");

        const dni =
            clon.querySelector(".resumen-persona__dni");

        const hora =
            clon.querySelector(".resumen-persona__hora");

        const alertaElemento =
            clon.querySelector(".alerta-dm");

        const contadorAlerta =
            clon.querySelector(".alerta-dm__contador");

        nombre.textContent =
            registro.registrado.nombre ||
            "Personal sin nombre";

        dni.textContent =
            `DNI: ${registro.registrado.dni || "-"}`;

        hora.textContent =
            formatearHora(registro.hora);

        const alertaPersonal =
            obtenerAlertaPorDNI(
                registro.registrado.dni
            );

        if (
            alertaElemento &&
            alertaPersonal
        ) {

            alertaElemento.hidden = false;

            alertaElemento.classList.remove(
                "alerta-dm--posterior",
                "alerta-dm--reincidencia"
            );

            if (
                alertaPersonal.tipo ===
                "reincidencia"
            ) {

                alertaElemento.classList.add(
                    "alerta-dm--reincidencia"
                );

                alertaElemento.title =
                    `${alertaPersonal.cantidad} casos posteriores al descanso médico`;

                alertaElemento.setAttribute(
                    "aria-label",
                    `Reincidencia posterior al descanso médico. ${alertaPersonal.cantidad} casos detectados.`
                );

                if (contadorAlerta) {

                    contadorAlerta.hidden = false;

                    contadorAlerta.textContent =
                        alertaPersonal.cantidad;
                }

            } else {

                alertaElemento.classList.add(
                    "alerta-dm--posterior"
                );

                alertaElemento.title =
                    "Posterior al descanso médico";

                alertaElemento.setAttribute(
                    "aria-label",
                    "Alerta posterior al descanso médico"
                );

                if (contadorAlerta) {

                    contadorAlerta.hidden = true;

                    contadorAlerta.textContent = "1";
                }
            }

            alertaElemento.addEventListener(
                "click",
                evento => {

                    evento.preventDefault();
                    evento.stopPropagation();

                    abrirModalAlertaDM(
                        alertaPersonal,
                        registro
                    );
                }
            );

            alertaElemento.addEventListener(
                "keydown",
                evento => {

                    if (
                        evento.key !== "Enter" &&
                        evento.key !== " "
                    ) {
                        return;
                    }

                    evento.preventDefault();
                    evento.stopPropagation();

                    abrirModalAlertaDM(
                        alertaPersonal,
                        registro
                    );
                }
            );

        } else if (alertaElemento) {

            alertaElemento.hidden = true;
        }

        boton.addEventListener(
            "click",
            () => abrirModalDetalle(registro)
        );

        fragmento.appendChild(clon);
    });

    contenedor.appendChild(fragmento);
}


// ============================================================
// ABRIR MODAL
// ============================================================

function abrirModalDetalle(registro) {

    detalleNombreRegistrado.textContent =
        registro.registrado.nombre ||
        "Personal sin nombre";


    detalleDniRegistrado.textContent =
        `DNI: ${
            registro.registrado.dni ||
            "-"
        }`;


    detalleFecha.textContent =
        formatearFechaDocumento(
            registro.fecha
        );


    detalleHora.textContent =
        formatearHora(
            registro.hora
        );


    detalleTurno.textContent =
        `${obtenerNombreTurno(
            registro.turno
        )} — ${obtenerHorarioTurno(
            registro.turno
        )}`;


    detalleReportanteNombre.textContent =
        registro.reportante.nombre ||
        "Sin información";


    detalleReportanteDni.textContent =
        `DNI: ${
            registro.reportante.dni ||
            "-"
        }`;


    if (registro.fotoUrl) {

        detalleFoto.src =
            registro.fotoUrl;

        detalleFoto.alt =
            `Fotografía del descanso médico de ${
                registro.registrado.nombre ||
                "personal registrado"
            }`;


        contenedorFotoDetalle.hidden =
            false;

        detalleSinFoto.hidden =
            true;

    } else {

        detalleFoto.src = "";

        contenedorFotoDetalle.hidden =
            true;

        detalleSinFoto.hidden =
            false;

    }


    modalDetalle.hidden =
        false;


    document.body.classList.add(
        "modal-abierto"
    );

}


// ============================================================
// CERRAR MODAL
// ============================================================

function cerrarModalDetalle() {

    modalDetalle.hidden =
        true;


    detalleFoto.src =
        "";


    document.body.classList.remove(
        "modal-abierto"
    );

}


// ============================================================
// FECHAS
// ============================================================

function colocarFechaActual() {

    filtroFecha.value =
        obtenerFechaPeru();

}


function obtenerFechaPeru() {

    const partes =
        new Intl.DateTimeFormat(
            "en-CA",
            {
                timeZone:
                    "America/Lima",

                year:
                    "numeric",

                month:
                    "2-digit",

                day:
                    "2-digit"
            }
        ).formatToParts(new Date());


    const valores = {};


    partes.forEach(parte => {

        if (
            parte.type !== "literal"
        ) {

            valores[parte.type] =
                parte.value;

        }

    });


    return (
        `${valores.year}-` +
        `${valores.month}-` +
        `${valores.day}`
    );

}


function convertirFechaADocumento(
    fechaISO
) {

    const [
        anio,
        mes,
        dia
    ] = fechaISO.split("-");


    return `${dia}-${mes}-${anio}`;

}


function formatearFechaDocumento(
    fechaDocumento
) {

    const partes =
        String(fechaDocumento || "")
            .split("-");


    if (
        partes.length !== 3
    ) {

        return fechaDocumento || "-";

    }


    const [
        dia,
        mes,
        anio
    ] = partes;


    const fecha =
        new Date(
            Number(anio),
            Number(mes) - 1,
            Number(dia)
        );


    if (
        Number.isNaN(
            fecha.getTime()
        )
    ) {

        return fechaDocumento;

    }


    return new Intl.DateTimeFormat(
        "es-PE",
        {
            day:
                "2-digit",

            month:
                "long",

            year:
                "numeric"
        }
    ).format(fecha);

}


function actualizarTextoFecha(
    fechaDocumento
) {

    textoFechaResumen.textContent =
        `Registros correspondientes al ${
            formatearFechaDocumento(
                fechaDocumento
            )
        }. Pulse sobre una persona para ver el detalle.`;

}


// ============================================================
// HORA
// ============================================================

function formatearHora(hora) {

    const valor =
        String(hora || "").trim();


    if (!valor) {
        return "-";
    }


    const partes =
        valor.split(":");


    if (
        partes.length < 2
    ) {

        return valor;

    }


    const horas =
        Number(partes[0]);

    const minutos =
        Number(partes[1]);


    if (
        Number.isNaN(horas) ||
        Number.isNaN(minutos)
    ) {

        return valor;

    }


    const fecha =
        new Date();


    fecha.setHours(
        horas,
        minutos,
        0,
        0
    );


    return new Intl.DateTimeFormat(
        "es-PE",
        {
            hour:
                "2-digit",

            minute:
                "2-digit",

            hour12:
                true
        }
    ).format(fecha);

}


// ============================================================
// TURNOS
// ============================================================

function obtenerNombreTurno(turno) {

    return (
        TURNOS[turno]?.nombre ||
        turno ||
        "Sin turno"
    );

}


function obtenerHorarioTurno(turno) {

    return (
        TURNOS[turno]?.horario ||
        "-"
    );

}


function contarPorTurno(
    registros,
    turno
) {

    return registros.filter(
        registro =>
            registro.turno === turno
    ).length;

}


// ============================================================
// ORDENAR REGISTROS
// ============================================================

function ordenarRegistros(registros) {

    const ordenTurnos = {

        "turno-1": 1,
        "turno-2": 2,
        "turno-3": 3

    };


    registros.sort((a, b) => {

        const diferenciaTurno =
            (
                ordenTurnos[a.turno] ||
                99
            ) -
            (
                ordenTurnos[b.turno] ||
                99
            );


        if (
            diferenciaTurno !== 0
        ) {

            return diferenciaTurno;

        }


        return (
            convertirHoraSegundos(
                a.hora
            ) -
            convertirHoraSegundos(
                b.hora
            )
        );

    });

}


function convertirHoraSegundos(hora) {

    const partes =
        String(hora || "")
            .split(":")
            .map(Number);


    if (
        partes.length < 2 ||
        partes.some(Number.isNaN)
    ) {

        return 0;

    }


    const [
        horas,
        minutos,
        segundos = 0
    ] = partes;


    return (
        horas * 3600 +
        minutos * 60 +
        segundos
    );

}

// ============================================================
// ESCAPAR CONTENIDO PARA HTML
// ============================================================

function escapeHtmlSeguro(
    valor
) {

    return String(valor ?? "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");

}
// ============================================================
// MENSAJES
// ============================================================

function mostrarMensaje(
    texto,
    tipo = "informacion"
) {

    mensajeSistema.textContent =
        texto;


    mensajeSistema.className =
        `mensaje-sistema mensaje-sistema--${tipo}`;


    mensajeSistema.hidden =
        false;

}


function ocultarMensaje() {

    mensajeSistema.hidden =
        true;


    mensajeSistema.textContent =
        "";

}


// ============================================================
// CARGA
// ============================================================

function activarCarga(
    mostrar,
    texto = "Cargando..."
) {

    pantallaCarga.hidden =
        !mostrar;


    textoCarga.textContent =
        texto;


    btnBuscar.disabled =
        mostrar;


    btnActualizar.disabled =
        mostrar;


    filtroFecha.disabled =
        mostrar;


    filtroTurno.disabled =
        mostrar;

}


// ============================================================
// ERRORES
// ============================================================

function obtenerMensajeError(error) {

    const codigo =
        error?.code || "";


    switch (codigo) {

        case "permission-denied":
        case "firestore/permission-denied":

            return (
                "No tiene permisos para consultar los registros. " +
                "Revise las reglas de Firestore."
            );


        case "unavailable":
        case "firestore/unavailable":

            return (
                "Firebase no está disponible. " +
                "Revise su conexión a internet."
            );


        case "failed-precondition":
        case "firestore/failed-precondition":

            return (
                "La consulta requiere una configuración adicional en Firestore."
            );


        default:

            return (
                error?.message ||
                "No se pudieron cargar los descansos médicos."
            );

    }

}
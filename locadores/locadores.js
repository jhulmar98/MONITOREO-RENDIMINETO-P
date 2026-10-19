
// =====================================================
// FALTAS LOCADORES
// FIREBASE STORAGE + EXCEL + CHART.JS
// =====================================================

import * as XLSX from
    "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/+esm";

import {
    descargarExcelStorage,
    RUTAS_STORAGE
} from "../backend/api.js";


// =====================================================
// CONFIGURACIÓN
// =====================================================

const CONFIG = {
    ANIO: 2026,
    HOJA: "Gral",

    RUTA:
        RUTAS_STORAGE.faltasLocadores ||
        "RENDIMIENTO PERSONAL/locadores/locadores faltas.xlsx",

    ALTO_RANKING: 490,
    ALTO_FILA: 43
};

const COLOR = {
    oliva: "#556B2F",
    olivaOscuro: "#3D4D1F",
    rojo: "#DC2626",
    azul: "#2563EB",
    texto: "#1E293B",
    gris: "#64748B",
    blanco: "#FFFFFF"
};

const MESES = [
    "ENERO", "FEBRERO", "MARZO", "ABRIL",
    "MAYO", "JUNIO", "JULIO", "AGOSTO",
    "SETIEMBRE", "OCTUBRE", "NOVIEMBRE",
    "DICIEMBRE"
];

const MESES_CORTOS = [
    "ENE", "FEB", "MAR", "ABR", "MAY", "JUN",
    "JUL", "AGO", "SET", "OCT", "NOV", "DIC"
];

const DIAS = [
    "DOMINGO", "LUNES", "MARTES",
    "MIÉRCOLES", "JUEVES", "VIERNES",
    "SÁBADO"
];


// =====================================================
// ESTADO GLOBAL
// =====================================================

const estado = {
    personas: [],
    filtradas: [],
    fechas: [],
    meses: [],
    sectores: [],
    ranking: [],
    desglose: [],

    filtros: {
        nombre: "",
        meses: [],
        sectores: []
    },

    charts: {},
    filtroAbierto: null,
    indiceRankingTeclado: 0
};


// =====================================================
// INICIO
// =====================================================

document.addEventListener("DOMContentLoaded", async () => {
    configurarEventos();
    configurarFiltrosFlotantes();
    await iniciar();
});

async function iniciar() {
    mostrarCarga(true);
    ocultarError();

    try {
        if (!window.Chart) {
            throw new Error("No se pudo cargar Chart.js.");
        }

        await cargarExcel();
        inicializarFiltros();
        aplicarFiltros();

        mostrarEstado("Información cargada correctamente.");
    } catch (error) {
        console.error("Error Locadores:", error);
        mostrarError(error.message);
    } finally {
        mostrarCarga(false);
    }
}


// =====================================================
// UTILIDADES
// =====================================================

function normalizar(valor) {
    return String(valor ?? "")
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .toUpperCase()
        .replace(/[^A-Z0-9]+/g, " ")
        .trim();
}

function escapar(valor) {
    return String(valor ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
}

function texto(id, valor) {
    const el = document.getElementById(id);
    if (el) el.textContent = valor;
}

function obtener(fila, indice) {
    if (indice < 0 || indice === undefined) return "";
    return String(fila[indice] ?? "").trim();
}

function nombreSector(valor) {
    const v = String(valor || "").trim();

    if (!v) return "SIN SECTOR";

    return normalizar(v).startsWith("SECTOR ")
        ? v
        : "SECTOR " + v;
}


// =====================================================
// RECONOCER FALTAS Y DESCANSOS
//
// F/f = FALTA
// D/d = DESCANSO MÉDICO
// DM/dm = DESCANSO MÉDICO
// =====================================================

function identificarMarca(valor) {
    const marca = normalizar(valor);

    if ([
        "F", "FA", "FALTA", "FALTAS",
        "FALTO", "INASISTENCIA", "AUSENTE"
    ].includes(marca)) {
        return {
            codigo: "F",
            tipo: "FALTA"
        };
    }

    if ([
        "D", "DM", "D M",
        "DESCANSO MEDICO",
        "DESCANSOS MEDICOS",
        "DESC MEDICO",
        "DESC MED",
        "DES MED"
    ].includes(marca)) {
        return {
            codigo: "D",
            tipo: "DESCANSO MÉDICO"
        };
    }

    return null;
}


// =====================================================
// FECHAS
// =====================================================

function crearFecha(anio, mes, dia) {
    const f = new Date(Date.UTC(anio, mes, dia));

    if (
        f.getUTCFullYear() !== anio ||
        f.getUTCMonth() !== mes ||
        f.getUTCDate() !== dia
    ) return null;

    return f;
}

function convertirFecha(valor) {
    if (valor === null || valor === undefined || valor === "") {
        return null;
    }

    if (valor instanceof Date) {
        if (isNaN(valor.getTime())) return null;

        return crearFecha(
            valor.getUTCFullYear(),
            valor.getUTCMonth(),
            valor.getUTCDate()
        );
    }

    if (
        typeof valor === "number" ||
        /^\d{5}$/.test(String(valor).trim())
    ) {
        const d = XLSX.SSF.parse_date_code(Number(valor));

        return d ? crearFecha(d.y, d.m - 1, d.d) : null;
    }

    const cadena = String(valor).trim();

    const ddmmyyyy = cadena.match(
        /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/
    );

    if (ddmmyyyy) {
        let anio = Number(ddmmyyyy[3]);
        if (anio < 100) anio += 2000;

        return crearFecha(
            anio,
            Number(ddmmyyyy[2]) - 1,
            Number(ddmmyyyy[1])
        );
    }

    const iso = cadena.match(/^(\d{4})-(\d{2})-(\d{2})/);

    if (iso) {
        return crearFecha(
            Number(iso[1]),
            Number(iso[2]) - 1,
            Number(iso[3])
        );
    }

    const escrita = normalizar(cadena).match(
        /(\d{1,2}) DE ([A-Z]+) DE (\d{4})/
    );

    if (escrita) {
        const nombreMes = escrita[2] === "SEPTIEMBRE"
            ? "SETIEMBRE"
            : escrita[2];

        const mes = MESES.indexOf(nombreMes);

        if (mes >= 0) {
            return crearFecha(
                Number(escrita[3]),
                mes,
                Number(escrita[1])
            );
        }
    }

    return null;
}

function fechaISO(fecha) {
    return fecha.toISOString().slice(0, 10);
}

function formatoFecha(fecha) {
    return new Intl.DateTimeFormat("es-PE", {
        timeZone: "UTC",
        day: "2-digit",
        month: "2-digit",
        year: "numeric"
    }).format(fecha);
}

function diaSemana(fecha) {
    return DIAS[fecha.getUTCDay()];
}

function etiquetaDia(fecha, variosMeses) {
    return variosMeses
        ? [String(fecha.dia), MESES_CORTOS[fecha.mes]]
        : String(fecha.dia);
}


// =====================================================
// DESCARGAR EXCEL DESDE FIREBASE
// =====================================================

async function cargarExcel() {
    const buffer = await descargarExcelStorage(CONFIG.RUTA);

    if (!buffer || !buffer.byteLength) {
        throw new Error("El archivo Excel está vacío.");
    }

    const libro = XLSX.read(buffer, {
        type: "array",
        cellDates: true
    });

    const nombreHoja = libro.SheetNames.find(
        n => normalizar(n) === normalizar(CONFIG.HOJA)
    ) || libro.SheetNames[0];

    if (!nombreHoja) {
        throw new Error("No se encontró ninguna hoja.");
    }

    const matriz = XLSX.utils.sheet_to_json(
        libro.Sheets[nombreHoja],
        {
            header: 1,
            defval: "",
            raw: true
        }
    );

    procesarExcel(matriz);

    texto(
        "fechaActualizacion",
        "Consultado: " + new Date().toLocaleString("es-PE")
    );
}


// =====================================================
// PROCESAR EXCEL
// =====================================================

function procesarExcel(matriz) {
    const filaCabecera = matriz.findIndex(fila => {
        const c = fila.map(normalizar);

        return c.includes("APELLIDOS Y NOMBRES") &&
               c.includes("SECTOR") &&
               c.includes("CARGO");
    });

    if (filaCabecera < 0) {
        throw new Error("No se encontró la cabecera del Excel.");
    }

    const encabezados = matriz[filaCabecera];

    function columna(...nombres) {
        const claves = nombres.map(normalizar);

        return encabezados.findIndex(
            v => claves.includes(normalizar(v))
        );
    }

    const c = {
        nombre: columna("APELLIDOS Y NOMBRES"),
        sector: columna("SECTOR"),
        cargo: columna("CARGO"),
        documento: columna("NUM. DOC.", "NUM DOC"),
        funcion: columna("FUNCION")
    };

    // DETECTAR FECHAS REALES DEL AÑO 2026
    const mapaFechas = new Map();

    encabezados.forEach((valor, indice) => {
        if (indice <= c.cargo) return;

        const fecha = convertirFecha(valor);

        if (!fecha || fecha.getUTCFullYear() !== CONFIG.ANIO) {
            return;
        }

        const iso = fechaISO(fecha);

        if (!mapaFechas.has(iso)) {
            mapaFechas.set(iso, {
                indice,
                fecha,
                iso,
                dia: fecha.getUTCDate(),
                mes: fecha.getUTCMonth(),
                diaSemana: diaSemana(fecha)
            });
        }
    });

    estado.fechas = [...mapaFechas.values()]
        .sort((a, b) => a.fecha - b.fecha);

    if (!estado.fechas.length) {
        throw new Error("No se encontraron fechas de 2026.");
    }

    const mapaPersonas = new Map();

    for (let i = filaCabecera + 1; i < matriz.length; i++) {
        const fila = matriz[i];

        const nombre = obtener(fila, c.nombre);
        const sector = obtener(fila, c.sector);
        const documento = obtener(fila, c.documento);

        if (!nombre) continue;

        if (["TOTAL", "SUBTOTAL", "RESUMEN", "TOTAL GENERAL"]
            .includes(normalizar(nombre))) {
            continue;
        }

        const claveSector = normalizar(sector);

        const clave = documento
            ? "DOC:" + normalizar(documento)
            : "NOMBRE:" + normalizar(nombre) + "|" + claveSector;

        if (!mapaPersonas.has(clave)) {
            mapaPersonas.set(clave, {
                id: mapaPersonas.size + 1,
                nombre,
                sector,
                sectorClave: claveSector,
                funcion: obtener(fila, c.funcion),
                incidencias: []
            });
        }

        const persona = mapaPersonas.get(clave);

        estado.fechas.forEach(col => {
            const marca = identificarMarca(fila[col.indice]);

            if (!marca) return;

            const repetida = persona.incidencias.some(
                x => x.iso === col.iso &&
                     x.codigo === marca.codigo
            );

            if (repetida) return;

            persona.incidencias.push({
                ...marca,
                fecha: col.fecha,
                iso: col.iso,
                dia: col.dia,
                mes: col.mes,
                diaSemana: col.diaSemana
            });
        });
    }

    estado.personas = [...mapaPersonas.values()];

    estado.meses = [...new Set(
        estado.fechas.map(x => x.mes)
    )].sort((a, b) => a - b);

    const sectores = new Map();

    estado.personas.forEach(p => {
        if (p.sectorClave && !sectores.has(p.sectorClave)) {
            sectores.set(p.sectorClave, p.sector);
        }
    });

    estado.sectores = [...sectores].map(([clave, nombre]) => ({
        clave,
        nombre
    })).sort((a, b) =>
        a.nombre.localeCompare(b.nombre, "es", { numeric: true })
    );

    console.log("Locadores cargados:", estado.personas.length);
}


// =====================================================
// FILTROS
// =====================================================

function inicializarFiltros() {
    const datalist = document.getElementById("listaNombresPersonal");

    if (datalist) {
        datalist.replaceChildren();

        estado.personas.forEach(p => {
            const option = document.createElement("option");
            option.value = p.nombre;
            datalist.appendChild(option);
        });
    }

    const ultimoMes = estado.meses.at(-1);

    estado.filtros = {
        nombre: "",
        meses: ultimoMes === undefined ? [] : [ultimoMes],
        sectores: []
    };

    actualizarTextosFiltros();
}

function actualizarTextosFiltros() {
    const m = estado.filtros.meses;
    const s = estado.filtros.sectores;

    texto(
        "textoFiltroMes",
        m.length === 1
            ? MESES[m[0]]
            : m.length > 1
                ? `${m.length} MESES`
                : "MESES DISPONIBLES"
    );

    let nombre = "SECTORES";

    if (s.length === 1) {
        const sector = estado.sectores.find(
            x => x.clave === s[0]
        );

        nombre = nombreSector(sector?.nombre || s[0]);
    } else if (s.length > 1) {
        nombre = `${s.length} SECTORES`;
    }

    texto("textoFiltroSector", nombre);
}

function configurarFiltrosFlotantes() {
    [
        ["boxFiltroMes", "mes"],
        ["boxFiltroSector", "sector"]
    ].forEach(([id, tipo]) => {
        const summary = document.querySelector(`#${id} summary`);

        summary?.addEventListener("click", e => {
            e.preventDefault();

            if (estado.filtroAbierto === tipo) {
                cerrarFiltro();
            } else {
                abrirFiltro(tipo, summary);
            }
        });
    });

    document.addEventListener("pointerdown", e => {
        if (!estado.filtroAbierto) return;

        const panel = document.getElementById("panelFiltroFlotante");
        const id = estado.filtroAbierto === "mes"
            ? "boxFiltroMes"
            : "boxFiltroSector";

        if (
            panel?.contains(e.target) ||
            document.getElementById(id)?.contains(e.target)
        ) return;

        cerrarFiltro();
    });

    window.addEventListener("resize", cerrarFiltro);
}

function abrirFiltro(tipo, referencia) {
    cerrarFiltro();

    const panel = document.getElementById("panelFiltroFlotante");
    if (!panel) return;

    estado.filtroAbierto = tipo;
    panel.replaceChildren();

    const opciones = tipo === "mes"
        ? estado.meses.map(m => ({
            valor: m,
            texto: MESES[m]
        }))
        : estado.sectores.map(s => ({
            valor: s.clave,
            texto: nombreSector(s.nombre)
        }));

    const rect = referencia.getBoundingClientRect();
    const ancho = 260;
    const alto = Math.min(340, opciones.length * 43 + 25);

    const izquierda = Math.max(
        8,
        Math.min(rect.left, window.innerWidth - ancho - 8)
    );

    const abajo = window.innerHeight - rect.bottom;

    Object.assign(panel.style, {
        position: "fixed",
        zIndex: "20000",
        display: "block",
        width: `${ancho}px`,
        maxHeight: "340px",
        overflowY: "auto",
        left: `${izquierda}px`,
        top: abajo >= alto
            ? `${rect.bottom + 7}px`
            : `${Math.max(8, rect.top - alto - 7)}px`
    });

    panel.hidden = false;

    opciones.forEach(opcion => {
        const label = document.createElement("label");
        label.className = "check-option";

        const input = document.createElement("input");
        input.type = "checkbox";

        const valores = tipo === "mes"
            ? estado.filtros.meses
            : estado.filtros.sectores;

        input.checked = valores.includes(opcion.valor);

        const span = document.createElement("span");
        span.textContent = opcion.texto;

        label.append(input, span);
        panel.appendChild(label);

        input.addEventListener("change", () => {
            const arr = tipo === "mes"
                ? estado.filtros.meses
                : estado.filtros.sectores;

            const posicion = arr.indexOf(opcion.valor);

            if (input.checked && posicion < 0) {
                arr.push(opcion.valor);
            } else if (!input.checked && posicion >= 0) {
                arr.splice(posicion, 1);
            }

            actualizarTextosFiltros();
            aplicarFiltros();
        });
    });
}

function cerrarFiltro() {
    const panel = document.getElementById("panelFiltroFlotante");

    if (panel) {
        panel.hidden = true;
        panel.style.display = "none";
    }

    estado.filtroAbierto = null;
}

function limpiarFiltros() {
    estado.filtros = {
        nombre: "",
        meses: [],
        sectores: []
    };

    const buscador = document.getElementById("filtroNombreGeneral");
    if (buscador) buscador.value = "";

    cerrarFiltro();
    actualizarTextosFiltros();
    aplicarFiltros();
}


// =====================================================
// INFORMACIÓN FILTRADA
// =====================================================

function fechasFiltradas() {
    const meses = estado.filtros.meses;

    return estado.fechas.filter(
        f => !meses.length || meses.includes(f.mes)
    );
}

function aplicarFiltros() {
    const nombre = normalizar(estado.filtros.nombre);
    const meses = estado.filtros.meses;
    const sectores = estado.filtros.sectores;

    estado.filtradas = estado.personas
        .filter(p => {
            if (nombre && !normalizar(p.nombre).includes(nombre)) {
                return false;
            }

            if (
                sectores.length &&
                !sectores.includes(p.sectorClave)
            ) {
                return false;
            }

            return true;
        })
        .map(p => {
            const incidencias = p.incidencias.filter(
                x => !meses.length || meses.includes(x.mes)
            );

            const faltas = incidencias.filter(
                x => x.codigo === "F"
            ).length;

            const descansos = incidencias.filter(
                x => x.codigo === "D"
            ).length;

            return {
                ...p,
                incidenciasFiltradas: incidencias,
                faltas,
                descansos,
                total: faltas + descansos
            };
        });

    renderDashboard();
}

function registrosFiltrados(condicion = () => true) {
    const registros = [];

    estado.filtradas.forEach(persona => {
        persona.incidenciasFiltradas.forEach(incidencia => {
            if (condicion(persona, incidencia)) {
                registros.push({ persona, incidencia });
            }
        });
    });

    return registros.sort(
        (a, b) =>
            a.incidencia.fecha - b.incidencia.fecha ||
            a.persona.nombre.localeCompare(b.persona.nombre, "es")
    );
}


// =====================================================
// KPI
// =====================================================

function actualizarKPI() {
    const personas = estado.filtradas;

    const total = personas.length;

    const conFaltas = personas.filter(
        p => p.faltas > 0
    ).length;

    const conDescanso = personas.filter(
        p => p.descansos > 0
    ).length;

    texto("kpiTotalLocadores", total);
    texto("kpiLocadoresConFaltas", conFaltas);
    texto("kpiDescansosMedicos", conDescanso);

    texto(
        "kpiPorcentajeFaltas",
        total ? (conFaltas / total * 100).toFixed(1) + "%" : "0%"
    );

    texto(
        "kpiPorcentajeDescansos",
        total ? (conDescanso / total * 100).toFixed(1) + "%" : "0%"
    );
}


// =====================================================
// COLORES DE GRÁFICOS
// MISMA ESCALA DE RESTRICCIONES MÉDICAS
// =====================================================

function generarColores(valores) {
    const max = Math.max(0, ...valores);

    return valores.map(v => {
        const r = max ? v / max : 0;

        if (r >= .90) return "#E78B8B";
        if (r >= .80) return "#F0A36B";
        if (r >= .70) return "#F6BF5B";
        if (r >= .60) return "#E4D86F";
        if (r >= .50) return "#C9D97B";
        if (r >= .40) return "#A9D18E";
        if (r >= .30) return "#7FC8A9";
        if (r >= .20) return "#66BFBF";

        return "#8EC5E8";
    });
}


// =====================================================
// ETIQUETAS SIEMPRE VISIBLES
// =====================================================

const etiquetasFallback = {
    id: "etiquetasFallbackLocadores",

    afterDatasetsDraw(chart) {
        if (window.ChartDataLabels) return;

        const ctx = chart.ctx;
        ctx.save();

        ctx.font = "bold 12px Arial";
        ctx.fillStyle = COLOR.texto;

        chart.data.datasets.forEach((dataset, datasetIndex) => {
            const meta = chart.getDatasetMeta(datasetIndex);

            meta.data.forEach((el, i) => {
                const v = Number(dataset.data[i]) || 0;
                if (v <= 0) return;

                if (chart.config.type === "doughnut") {
                    const angulo = (el.startAngle + el.endAngle) / 2;
                    const radio = (el.innerRadius + el.outerRadius) / 2;

                    ctx.textAlign = "center";
                    ctx.fillStyle = "#FFFFFF";

                    ctx.fillText(
                        String(v),
                        el.x + Math.cos(angulo) * radio,
                        el.y + Math.sin(angulo) * radio
                    );
                } else if (chart.options.indexAxis === "y") {
                    ctx.textAlign = "left";
                    ctx.fillText(String(v), el.x + 8, el.y + 4);
                } else {
                    ctx.textAlign = "center";
                    ctx.fillText(String(v), el.x, el.y - 8);
                }
            });
        });

        ctx.restore();
    }
};

function destruirChart(id) {
    if (estado.charts[id]) {
        estado.charts[id].destroy();
        delete estado.charts[id];
    }
}

function crearChart(id, config) {
    destruirChart(id);

    const canvas = document.getElementById(id);
    if (!canvas) return;

    estado.charts[id] = new Chart(canvas, {
        ...config,
        plugins: [
            window.ChartDataLabels || etiquetasFallback,
            ...(config.plugins || [])
        ]
    });
}

function cursorGrafico(evento, elementos, chart) {
    if (chart?.canvas) {
        chart.canvas.style.cursor =
            elementos.length ? "pointer" : "default";
    }
}

function opcionesBarras(horizontal = false) {
    return {
        responsive: true,
        maintainAspectRatio: false,
        indexAxis: horizontal ? "y" : "x",

        layout: {
            padding: {
                top: 28,
                right: 43,
                bottom: 8,
                left: 5
            }
        },

        plugins: {
            legend: { display: false },

            datalabels: {
                display: c =>
                    Number(c.dataset.data[c.dataIndex]) > 0,

                color: COLOR.texto,
                anchor: "end",
                align: "end",
                offset: 5,
                clip: false,
                clamp: true,

                font: {
                    weight: "bold",
                    size: 12
                },

                formatter: v => String(v)
            },

            tooltip: {
                backgroundColor: "rgba(15,23,42,.96)",
                padding: 13,
                displayColors: false
            }
        },

        scales: {
            x: {
                beginAtZero: true,
                grid: { color: "#EDF1E7" },
                ticks: {
                    color: COLOR.gris,
                    font: { weight: "bold", size: 11 }
                }
            },

            y: {
                beginAtZero: true,
                grid: { color: "#EDF1E7" },
                ticks: {
                    color: COLOR.gris,
                    precision: 0,
                    font: { weight: "bold", size: 11 }
                }
            }
        }
    };
}


// =====================================================
// RENDER GENERAL
// =====================================================

function renderDashboard() {
    actualizarKPI();
    renderFaltasPorDia();
    renderFaltasPorSector();
    renderDistribucion();
    renderRanking();
}


// =====================================================
// GRÁFICO PRINCIPAL: FALTAS POR DÍA
//
// CLIC -> FALTAS POR SECTOR DE ESE DÍA
// =====================================================

function renderFaltasPorDia() {
    const fechas = fechasFiltradas();

    const registros = registrosFiltrados(
        (_, inc) => inc.codigo === "F"
    );

    const mapa = new Map(
        fechas.map(f => [f.iso, 0])
    );

    registros.forEach(({ incidencia }) => {
        mapa.set(
            incidencia.iso,
            (mapa.get(incidencia.iso) || 0) + 1
        );
    });

    const varios = new Set(fechas.map(f => f.mes)).size > 1;

    const valores = fechas.map(
        f => mapa.get(f.iso) || 0
    );

    const labels = fechas.map(
        f => etiquetaDia(f, varios)
    );

    const meses = [...new Set(fechas.map(f => f.mes))];

    texto(
        "textoPeriodoGrafico",
        meses.length === 1
            ? MESES[meses[0]]
            : `${meses.length} MESES`
    );

    crearChart("graficoFaltasPorDia", {
        type: "bar",

        data: {
            labels,
            datasets: [{
                label: "Faltas",
                data: valores,
                backgroundColor: generarColores(valores),
                borderRadius: 9,
                borderSkipped: false,
                maxBarThickness: 36
            }]
        },

        options: {
            ...opcionesBarras(),

            onClick: (_, elementos) => {
                if (!elementos.length) return;

                const fecha = fechas[elementos[0].index];
                if (fecha) abrirDesgloseDia(fecha);
            },

            onHover: cursorGrafico,

            scales: {
                x: {
                    grid: { display: false },
                    ticks: {
                        color: COLOR.gris,
                        autoSkip: false,
                        maxRotation: 0,
                        minRotation: 0,
                        font: { size: 10, weight: "bold" }
                    }
                },

                y: {
                    beginAtZero: true,
                    grace: "20%",
                    ticks: { precision: 0 },
                    grid: { color: "#EDF1E7" }
                }
            }
        }
    });
}


// =====================================================
// GRÁFICO PRINCIPAL: FALTAS POR SECTOR
//
// CLIC -> FALTAS POR DÍA DE ESE SECTOR
// =====================================================

function renderFaltasPorSector() {
    const mapa = new Map();

    estado.filtradas.forEach(p => {
        const clave = p.sectorClave || "SIN SECTOR";

        if (!mapa.has(clave)) {
            mapa.set(clave, {
                clave,
                nombre: nombreSector(p.sector),
                cantidad: 0
            });
        }

        mapa.get(clave).cantidad += p.faltas;
    });

    const grupos = [...mapa.values()]
        .sort((a, b) =>
            a.nombre.localeCompare(
                b.nombre, "es", { numeric: true }
            )
        );

    const valores = grupos.map(g => g.cantidad);

    crearChart("graficoFaltasPorSector", {
        type: "bar",

        data: {
            labels: grupos.map(g => g.nombre),
            datasets: [{
                label: "Faltas",
                data: valores,
                backgroundColor: generarColores(valores),
                borderRadius: 10,
                borderSkipped: false,
                maxBarThickness: 65
            }]
        },

        options: {
            ...opcionesBarras(),

            onClick: (_, elementos) => {
                if (!elementos.length) return;

                const grupo = grupos[elementos[0].index];

                if (grupo) {
                    abrirDesgloseSector(grupo);
                }
            },

            onHover: cursorGrafico,

            scales: {
                x: {
                    grid: { display: false },
                    ticks: {
                        color: COLOR.gris,
                        font: { weight: "bold", size: 11 }
                    }
                },

                y: {
                    beginAtZero: true,
                    grace: "20%",
                    ticks: { precision: 0 },
                    grid: { color: "#EDF1E7" }
                }
            }
        }
    });
}


// =====================================================
// GRÁFICO CIRCULAR
//
// SE CONSERVA EL REPORTE DE INCIDENCIAS
// =====================================================

function renderDistribucion() {
    const registros = registrosFiltrados();

    const faltas = registros.filter(
        r => r.incidencia.codigo === "F"
    ).length;

    const descansos = registros.filter(
        r => r.incidencia.codigo === "D"
    ).length;

    const total = faltas + descansos;

    crearChart("graficoDistribucionIncidencias", {
        type: "doughnut",

        data: {
            labels: ["FALTAS", "DESCANSOS MÉDICOS"],

            datasets: [{
                data: [faltas, descansos],
                backgroundColor: [COLOR.rojo, COLOR.azul],
                borderColor: COLOR.blanco,
                borderWidth: 4,
                hoverOffset: 10
            }]
        },

        options: {
            responsive: true,
            maintainAspectRatio: false,
            cutout: "60%",

            layout: {
                padding: 25
            },

            onClick: (_, elementos) => {
                if (!elementos.length) return;

                const codigo = elementos[0].index === 0
                    ? "F" : "D";

                const detalle = registrosFiltrados(
                    (_, inc) => inc.codigo === codigo
                );

                abrirReporteGrafico({
                    titulo: codigo === "F"
                        ? "REPORTE DE FALTAS"
                        : "REPORTE DE DESCANSOS MÉDICOS",

                    subtitulo: "Detalle de incidencias registradas",
                    registros: detalle
                });
            },

            onHover: cursorGrafico,

            plugins: {
                legend: { display: false },

                datalabels: {
                    display: c =>
                        Number(c.dataset.data[c.dataIndex]) > 0,

                    color: "#FFFFFF",

                    font: {
                        weight: "bold",
                        size: 14
                    },

                    textAlign: "center",

                    formatter: valor => [
                        String(valor),
                        (total ? valor / total * 100 : 0)
                            .toFixed(1) + "%"
                    ]
                },

                tooltip: {
                    backgroundColor: "rgba(15,23,42,.96)",
                    padding: 13
                }
            }
        }
    });
}


// =====================================================
// RANKING DE FALTAS
//
// 10 FILAS VISIBLES + SCROLL INTERNO
// CLIC -> REPORTE INDIVIDUAL
// =====================================================

function renderRanking() {
    const ranking = [...estado.filtradas]
        .filter(p => p.faltas > 0)
        .sort((a, b) =>
            b.faltas - a.faltas ||
            a.nombre.localeCompare(b.nombre, "es")
        );

    estado.ranking = ranking;
    estado.indiceRankingTeclado = 0;

    texto("totalRankingLocadores", ranking.length);

    const contenedor =
        document.getElementById("contenedorRanking");

    const wrap =
        document.getElementById("rankingCanvasWrap");

    if (contenedor) {
        contenedor.style.setProperty(
            "height",
            CONFIG.ALTO_RANKING + "px",
            "important"
        );

        contenedor.style.setProperty(
            "max-height",
            CONFIG.ALTO_RANKING + "px",
            "important"
        );

        contenedor.style.overflowY = "auto";
        contenedor.scrollTop = 0;
    }

    if (wrap) {
        const altura = Math.max(
            430,
            ranking.length * CONFIG.ALTO_FILA + 65
        );

        wrap.style.height = altura + "px";
        wrap.style.minHeight = altura + "px";
    }

    const valores = ranking.map(p => p.faltas);

    crearChart("graficoRankingLocadores", {
        type: "bar",

        data: {
            labels: ranking.map(p =>
                p.nombre.length > 29
                    ? p.nombre.slice(0, 27) + "..."
                    : p.nombre
            ),

            datasets: [{
                label: "Faltas",
                data: valores,
                backgroundColor: generarColores(valores),
                borderRadius: 8,
                borderSkipped: false,
                maxBarThickness: 28
            }]
        },

        options: {
            ...opcionesBarras(true),

            onClick: (_, elementos) => {
                if (!elementos.length) return;

                const persona = ranking[elementos[0].index];

                if (persona) {
                    abrirReporteIndividual(persona.id);
                }
            },

            onHover: cursorGrafico,

            plugins: {
                legend: { display: false },

                datalabels: {
                    display: true,
                    color: COLOR.texto,
                    anchor: "end",
                    align: "end",
                    offset: 5,

                    font: {
                        weight: "bold",
                        size: 12
                    }
                },

                tooltip: {
                    backgroundColor: "rgba(15,23,42,.96)",
                    padding: 13,

                    callbacks: {
                        title: ctx =>
                            ranking[ctx[0]?.dataIndex]?.nombre || "",

                        label: ctx =>
                            "Faltas: " + ctx.raw
                    }
                }
            },

            scales: {
                x: {
                    beginAtZero: true,
                    grace: "20%",
                    ticks: { precision: 0 },
                    grid: { color: "#EDF1E7" }
                },

                y: {
                    grid: { display: false },
                    ticks: {
                        color: COLOR.texto,
                        font: { size: 11, weight: "bold" }
                    }
                }
            }
        }
    });
}


// =====================================================
// DESGLOSE 1
//
// CLIC EN FALTAS POR DÍA
// -> GRÁFICO DE FALTAS POR SECTOR DE ESE DÍA
// -> VER PERSONAS DEL DÍA Y SECTOR
// =====================================================

function abrirDesgloseDia(fecha) {
    const registros = registrosFiltrados(
        (_, inc) =>
            inc.codigo === "F" &&
            inc.iso === fecha.iso
    );

    const mapa = new Map();

    registros.forEach(registro => {
        const p = registro.persona;
        const clave = p.sectorClave || "SIN SECTOR";

        if (!mapa.has(clave)) {
            mapa.set(clave, {
                etiqueta: nombreSector(p.sector),
                registros: []
            });
        }

        mapa.get(clave).registros.push(registro);
    });

    const grupos = [...mapa.values()]
        .sort((a, b) =>
            a.etiqueta.localeCompare(
                b.etiqueta, "es", { numeric: true }
            )
        );

    abrirDesglose({
        titulo: "FALTAS POR SECTOR",
        subtitulo:
            fecha.diaSemana + " - " + formatoFecha(fecha.fecha),

        tituloGrafico: "FALTAS POR SECTOR DEL DÍA",
        tituloTabla: "RESUMEN POR SECTOR",
        encabezado: "SECTOR",
        grupos,
        horizontal: false
    });
}


// =====================================================
// DESGLOSE 2
//
// CLIC EN FALTAS POR SECTOR
// -> GRÁFICO DE FALTAS POR DÍA DE ESE SECTOR
// -> VER PERSONAS DEL DÍA Y SECTOR
// =====================================================

function abrirDesgloseSector(sector) {
    const fechas = fechasFiltradas();

    const registros = registrosFiltrados(
        (p, inc) =>
            p.sectorClave === sector.clave &&
            inc.codigo === "F"
    );

    const mapa = new Map();

    fechas.forEach(fecha => {
        mapa.set(fecha.iso, {
            etiqueta: String(fecha.dia),
            fecha,
            registros: []
        });
    });

    registros.forEach(registro => {
        const grupo = mapa.get(registro.incidencia.iso);

        if (grupo) {
            grupo.registros.push(registro);
        }
    });

    const grupos = [...mapa.values()];

    const variosMeses =
        new Set(fechas.map(f => f.mes)).size > 1;

    grupos.forEach(grupo => {
        grupo.etiqueta = variosMeses
            ? grupo.fecha.dia + " " +
              MESES_CORTOS[grupo.fecha.mes]
            : String(grupo.fecha.dia);
    });

    abrirDesglose({
        titulo: "FALTAS POR DÍA",
        subtitulo: sector.nombre,
        tituloGrafico: "FALTAS DIARIAS - " + sector.nombre,
        tituloTabla: "RESUMEN POR FECHA",
        encabezado: "FECHA",
        grupos,
        horizontal: false
    });
}


// =====================================================
// MODAL DE DESGLOSE
// UTILIZADO POR AMBOS GRÁFICOS
// =====================================================

function abrirDesglose({
    titulo,
    subtitulo,
    tituloGrafico,
    tituloTabla,
    encabezado,
    grupos,
    horizontal
}) {
    estado.desglose = grupos;

    texto("tituloModalDesglose", titulo);
    texto("subtituloModalDesglose", subtitulo);
    texto("tituloGraficoDesglose", tituloGrafico);
    texto("tituloTablaDesglose", tituloTabla);

    const panel = document.getElementById(
        "panelPersonasDesglose"
    );

    if (panel) panel.hidden = true;

    const encabezados = document.getElementById(
        "encabezadosTablaDesglose"
    );

    if (encabezados) {
        encabezados.innerHTML = `
            <th>${escapar(encabezado)}</th>
            <th>FALTAS</th>
            <th>REPORTE</th>
        `;
    }

    const tbody = document.getElementById(
        "tablaResumenDesglose"
    );

    if (tbody) {
        const positivos = grupos
            .map((g, indice) => ({ ...g, indice }))
            .filter(g => g.registros.length > 0);

        tbody.innerHTML = positivos.length
            ? positivos.map(g => `
                <tr>
                    <td><strong>${escapar(
                        g.fecha
                            ? formatoFecha(g.fecha.fecha)
                            : g.etiqueta
                    )}</strong></td>

                    <td class="texto-rojo">
                        <strong>${g.registros.length}</strong>
                    </td>

                    <td>
                        <button
                            type="button"
                            class="btn-reporte-individual"
                            data-grupo="${g.indice}"
                        >
                            <i class="fa-solid fa-eye"></i>
                            VER
                        </button>
                    </td>
                </tr>
            `).join("")
            : `
                <tr>
                    <td colspan="3" class="sin-datos">
                        No existen faltas registradas.
                    </td>
                </tr>
            `;
    }

    // ABRIR MODAL ANTES DE CREAR EL GRÁFICO
    abrirModal("modalDesglose");

    const valores = grupos.map(
        g => g.registros.length
    );

    crearChart("graficoDesglose", {
        type: "bar",

        data: {
            labels: grupos.map(g => g.etiqueta),

            datasets: [{
                label: "Faltas",
                data: valores,
                backgroundColor: generarColores(valores),
                borderRadius: 9,
                borderSkipped: false,
                maxBarThickness: 55
            }]
        },

        options: {
            ...opcionesBarras(horizontal),

            onClick: (_, elementos) => {
                if (!elementos.length) return;

                const grupo = grupos[elementos[0].index];

                if (grupo && grupo.registros.length) {
                    mostrarPersonasDesglose(grupo);
                }
            },

            onHover: cursorGrafico,

            scales: {
                x: {
                    grid: { display: false },
                    ticks: {
                        color: COLOR.gris,
                        autoSkip: !horizontal,
                        maxRotation: 0,
                        font: { size: 10, weight: "bold" }
                    }
                },

                y: {
                    beginAtZero: true,
                    grace: "20%",
                    ticks: {
                        precision: 0,
                        color: COLOR.gris
                    },
                    grid: { color: "#EDF1E7" }
                }
            }
        }
    });
}


// =====================================================
// BOTÓN VER DEL DESGLOSE
//
// MUESTRA QUIÉNES FALTARON EN ESE DÍA Y SECTOR
// =====================================================

function mostrarPersonasDesglose(grupo) {
    const panel = document.getElementById(
        "panelPersonasDesglose"
    );

    const tbody = document.getElementById(
        "tablaPersonasDesglose"
    );

    if (!panel || !tbody) return;

    texto(
        "tituloPersonasDesglose",
        "PERSONAL CON FALTAS - " +
        (grupo.fecha
            ? formatoFecha(grupo.fecha.fecha)
            : grupo.etiqueta)
    );

    tbody.innerHTML = grupo.registros.map(
        (r, indice) => {
            const p = r.persona;
            const inc = r.incidencia;

            return `
                <tr>
                    <td>${indice + 1}</td>

                    <td>
                        <strong>${escapar(p.nombre)}</strong>
                    </td>

                    <td>${escapar(nombreSector(p.sector))}</td>

                    <td>${formatoFecha(inc.fecha)}</td>

                    <td>${escapar(inc.diaSemana)}</td>

                    <td class="texto-rojo">
                        <strong>F</strong>
                    </td>

                    <td>
                        <button
                            type="button"
                            class="btn-reporte-individual"
                            data-persona-id="${p.id}"
                        >
                            <i class="fa-solid fa-eye"></i>
                            VER
                        </button>
                    </td>
                </tr>
            `;
        }
    ).join("");

    panel.hidden = false;

    panel.scrollIntoView({
        behavior: "smooth",
        block: "nearest"
    });
}


// =====================================================
// REPORTE ORIGINAL DEL GRÁFICO CIRCULAR
// =====================================================

function abrirReporteGrafico({
    titulo,
    subtitulo,
    registros
}) {
    texto("tituloModalGrafico", titulo);
    texto("subtituloModalGrafico", subtitulo);
    texto("totalModalGrafico", registros.length);

    const tbody = document.getElementById(
        "tablaDetalleGrafico"
    );

    if (!tbody) return;

    tbody.innerHTML = registros.length
        ? registros.map((r, indice) => {
            const p = r.persona;
            const inc = r.incidencia;

            const clase = inc.codigo === "F"
                ? "texto-rojo"
                : "texto-azul";

            return `
                <tr>
                    <td>${indice + 1}</td>

                    <td>
                        <strong>${escapar(p.nombre)}</strong>
                    </td>

                    <td>${escapar(nombreSector(p.sector))}</td>

                    <td>${formatoFecha(inc.fecha)}</td>

                    <td>${escapar(inc.diaSemana)}</td>

                    <td class="${clase}">
                        <strong>${inc.codigo}</strong>
                    </td>

                    <td class="${clase}">
                        <strong>${escapar(inc.tipo)}</strong>
                    </td>

                    <td>
                        <button
                            type="button"
                            class="btn-reporte-individual"
                            data-persona-id="${p.id}"
                        >
                            <i class="fa-solid fa-eye"></i>
                            VER
                        </button>
                    </td>
                </tr>
            `;
        }).join("")
        : `
            <tr>
                <td colspan="8" class="sin-datos">
                    No existen incidencias registradas.
                </td>
            </tr>
        `;

    abrirModal("modalDetalleGrafico");
}


// =====================================================
// REPORTE INDIVIDUAL
//
// SOLO:
// - NOMBRE
// - SECTOR
// - FUNCIÓN
//
// SE CONSERVA HISTORIAL CON DÍA DE LA SEMANA
// =====================================================

function abrirReporteIndividual(id) {
    const p = estado.filtradas.find(
        persona => persona.id === Number(id)
    );

    if (!p) return;

    texto("tituloModalLocador", "REPORTE INDIVIDUAL");
    texto("subtituloModalLocador", p.nombre);

    texto("detalleNombre", p.nombre);
    texto("detalleSector", nombreSector(p.sector));
    texto("detalleFuncion", p.funcion || "-");

    texto("detalleTotalFaltas", p.faltas);
    texto("detalleTotalDescansos", p.descansos);
    texto("detalleTotalIncidencias", p.total);

    const historial = [...p.incidenciasFiltradas]
        .sort((a, b) => a.fecha - b.fecha);

    const tbody = document.getElementById(
        "tablaHistorialLocador"
    );

    if (!tbody) return;

    tbody.innerHTML = historial.length
        ? historial.map(inc => {
            const clase = inc.codigo === "F"
                ? "texto-rojo"
                : "texto-azul";

            return `
                <tr>
                    <td>${formatoFecha(inc.fecha)}</td>

                    <td>
                        <strong>${escapar(inc.diaSemana)}</strong>
                    </td>

                    <td class="${clase}">
                        <strong>${inc.codigo}</strong>
                    </td>

                    <td class="${clase}">
                        <strong>${escapar(inc.tipo)}</strong>
                    </td>
                </tr>
            `;
        }).join("")
        : `
            <tr>
                <td colspan="4" class="sin-datos">
                    Sin incidencias registradas.
                </td>
            </tr>
        `;

    abrirModal("modalDetalleLocador");
}


// =====================================================
// ADMINISTRACIÓN DE MODALES
// =====================================================

function abrirModal(id) {
    const modal = document.getElementById(id);
    if (!modal) return;

    modal.hidden = false;
    modal.style.display = "flex";

    modal.style.zIndex = id === "modalDetalleLocador"
        ? "100003"
        : "100000";

    modal.classList.add("modal-detalle--show");
    document.body.style.overflow = "hidden";
}

function cerrarModal(id) {
    const modal = document.getElementById(id);
    if (!modal) return;

    modal.classList.remove("modal-detalle--show");
    modal.hidden = true;
    modal.style.display = "none";

    if (id === "modalDesglose") {
        destruirChart("graficoDesglose");
    }

    const hayOtro = [
        "modalDesglose",
        "modalDetalleGrafico",
        "modalDetalleLocador"
    ].some(mid => {
        const el = document.getElementById(mid);
        return el && !el.hidden;
    });

    if (!hayOtro) {
        document.body.style.overflow = "";
    }
}


// =====================================================
// EVENTOS
// =====================================================

function configurarEventos() {

    // BUSCAR SOLO POR NOMBRE
    document.getElementById(
        "filtroNombreGeneral"
    )?.addEventListener("input", e => {
        estado.filtros.nombre = e.target.value;
        aplicarFiltros();
    });

    // LIMPIAR
    document.getElementById(
        "btnLimpiarTodo"
    )?.addEventListener("click", limpiarFiltros);

    // CERRAR MODALES
    [
        ["btnCerrarModalDesglose", "modalDesglose"],
        ["btnCerrarModalGrafico", "modalDetalleGrafico"],
        ["btnCerrarModalLocador", "modalDetalleLocador"]
    ].forEach(([boton, modal]) => {
        document.getElementById(boton)
            ?.addEventListener("click", () => cerrarModal(modal));
    });

    // CERRAR DESDE FONDO
    [
        "modalDesglose",
        "modalDetalleGrafico",
        "modalDetalleLocador"
    ].forEach(id => {
        document.querySelector(
            `#${id} .modal-detalle__overlay`
        )?.addEventListener(
            "click",
            () => cerrarModal(id)
        );
    });

    // BOTÓN VER: SECTOR O FECHA DEL DESGLOSE
    document.getElementById(
        "tablaResumenDesglose"
    )?.addEventListener("click", e => {
        const boton = e.target.closest("[data-grupo]");
        if (!boton) return;

        const grupo = estado.desglose[
            Number(boton.dataset.grupo)
        ];

        if (grupo) mostrarPersonasDesglose(grupo);
    });

    // VER INDIVIDUAL DESDE LISTADO DE PERSONAS
    document.getElementById(
        "tablaPersonasDesglose"
    )?.addEventListener("click", e => {
        const boton = e.target.closest("[data-persona-id]");
        if (!boton) return;

        abrirReporteIndividual(boton.dataset.personaId);
    });

    // VER INDIVIDUAL DESDE GRÁFICO CIRCULAR
    document.getElementById(
        "tablaDetalleGrafico"
    )?.addEventListener("click", e => {
        const boton = e.target.closest("[data-persona-id]");
        if (!boton) return;

        abrirReporteIndividual(boton.dataset.personaId);
    });

    // CERRAR CON ESC
    document.addEventListener("keydown", e => {
        if (e.key !== "Escape") return;

        const orden = [
            "modalDetalleLocador",
            "modalDesglose",
            "modalDetalleGrafico"
        ];

        const abierto = orden.find(id => {
            const el = document.getElementById(id);
            return el && !el.hidden;
        });

        if (abierto) {
            cerrarModal(abierto);
        } else {
            cerrarFiltro();
        }
    });

    // RANKING CON TECLADO
    document.getElementById(
        "graficoRankingLocadores"
    )?.addEventListener("keydown", e => {
        if (!estado.ranking.length) return;

        if (e.key === "ArrowDown") {
            e.preventDefault();

            estado.indiceRankingTeclado = Math.min(
                estado.indiceRankingTeclado + 1,
                estado.ranking.length - 1
            );
        } else if (e.key === "ArrowUp") {
            e.preventDefault();

            estado.indiceRankingTeclado = Math.max(
                estado.indiceRankingTeclado - 1,
                0
            );
        } else if (e.key === "Enter" || e.key === " ") {
            e.preventDefault();

            const p = estado.ranking[
                estado.indiceRankingTeclado
            ];

            if (p) abrirReporteIndividual(p.id);
        }
    });

    // REINTENTAR
    document.getElementById(
        "btnReintentarCarga"
    )?.addEventListener("click", iniciar);
}


// =====================================================
// ESTADOS DE CARGA Y ERROR
// =====================================================

function mostrarCarga(activo) {
    const el = document.getElementById("loadingOverlay");
    if (!el) return;

    el.hidden = !activo;
    el.style.display = activo ? "flex" : "none";
}

function ocultarError() {
    const el = document.getElementById("mensajeError");
    if (!el) return;

    el.hidden = true;
    el.style.display = "none";
}

function mostrarError(mensaje) {
    texto("textoMensajeError", mensaje);

    const el = document.getElementById("mensajeError");
    if (!el) return;

    el.hidden = false;
    el.style.display = "block";
}

let avisoTimer;

function mostrarEstado(mensaje) {
    const el = document.getElementById("estadoCarga");
    if (!el) return;

    clearTimeout(avisoTimer);

    el.textContent = mensaje;
    el.hidden = false;
    el.style.display = "flex";

    avisoTimer = setTimeout(() => {
        el.hidden = true;
        el.style.display = "none";
    }, 2500);
}


// =====================================================
// FIN DEL ARCHIVO LOCADORES.JS
// =====================================================

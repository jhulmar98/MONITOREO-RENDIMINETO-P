// =====================================================
// FIREBASE API GENERAL
// AUTHENTICATION + FIRESTORE + STORAGE
// =====================================================


// =====================================================
// FIREBASE APP
// =====================================================
import {
    initializeApp,
    getApps,
    getApp
} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js";


// =====================================================
// AUTHENTICATION
// =====================================================
import {
    getAuth
} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js";


// =====================================================
// FIRESTORE
// =====================================================
import {
    getFirestore,
    collection,
    getDocs,
    getDoc,
    doc,
    addDoc,
    setDoc,
    updateDoc,
    deleteDoc,
    query,
    where,
    orderBy,
    limit
} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js";


// =====================================================
// FIREBASE STORAGE
// =====================================================
import {
    getStorage,
    ref,
    getBytes,
    getDownloadURL,
    listAll
} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-storage.js";


// =====================================================
// CONFIGURACIÓN FIREBASE
// =====================================================
const firebaseConfig = {
    apiKey: "AIzaSyD_hz2Y0qajgRMPeb0L3Ky75jQJIebywJM",
    authDomain: "red-de-patas.firebaseapp.com",
    projectId: "red-de-patas",
    storageBucket: "red-de-patas.firebasestorage.app",
    messagingSenderId: "812893065625",
    appId: "1:812893065625:web:08b1c067911872edd14308",
    measurementId: "G-M1NH2DJF29"
};


// =====================================================
// INICIALIZACIÓN ÚNICA
// =====================================================
const app = getApps().length
    ? getApp()
    : initializeApp(firebaseConfig);

const auth = getAuth(app);
const db = getFirestore(app);
const storage = getStorage(app);


// =====================================================
// EXPORTACIONES GENERALES
// =====================================================
export {
    app,
    auth,
    db,
    storage,

    // Firestore
    collection,
    getDocs,
    getDoc,
    doc,
    addDoc,
    setDoc,
    updateDoc,
    deleteDoc,
    query,
    where,
    orderBy,
    limit,

    // Storage
    ref,
    getBytes,
    getDownloadURL,
    listAll
};


// =====================================================
// RUTAS PRINCIPALES DE STORAGE
// Puedes agregar más rutas cuando crees nuevos módulos.
// =====================================================
export const RUTAS_STORAGE = {

    descansoMedico:
        "RENDIMIENTO PERSONAL/DESCANSO MEDICOS/DM.xlsx",

    carpetaDescansosMedicos:
        "RENDIMIENTO PERSONAL/DESCANSO MEDICOS",

    papeletas:
        "RENDIMIENTO PERSONAL/PAPELETAS",

    rolServicio:
        "RENDIMIENTO PERSONAL/ROL DE SERVICIO",

    restriccionesMedicas:

        "RESTRICCIONES MEDICAS/restricciones.xlsx",
};


// =====================================================
// OBTENER REFERENCIA DE CUALQUIER RUTA
// =====================================================
export function obtenerReferenciaStorage(rutaArchivo) {

    const ruta = limpiarRutaStorage(rutaArchivo);

    return ref(storage, ruta);
}


// =====================================================
// DESCARGAR CUALQUIER ARCHIVO COMO ARRAYBUFFER
// Sirve para Excel, PDF, imágenes y otros binarios.
// =====================================================
export async function descargarArchivoStorage(
    rutaArchivo,
    maximoBytes = 100 * 1024 * 1024
) {

    const ruta = limpiarRutaStorage(rutaArchivo);

    try {

        const archivoRef =
            obtenerReferenciaStorage(ruta);

        /*
         * getBytes() devuelve un ArrayBuffer.
         */
        const buffer = await getBytes(
            archivoRef,
            maximoBytes
        );

        if (
            !(buffer instanceof ArrayBuffer) ||
            buffer.byteLength === 0
        ) {

            throw new Error(
                `El archivo "${ruta}" está vacío o no es válido.`
            );
        }

        console.log(
            `Archivo descargado: ${ruta}`,
            `${buffer.byteLength} bytes`
        );

        return buffer;

    } catch (error) {

        console.error(
            `Error descargando "${ruta}":`,
            error
        );

        throw new Error(
            obtenerMensajeErrorStorage(
                error,
                ruta
            )
        );
    }
}


// =====================================================
// DESCARGAR CUALQUIER EXCEL
// Compatible con XLSX.read(buffer, { type: "array" })
// =====================================================
export async function descargarExcelStorage(rutaExcel) {

    const ruta = limpiarRutaStorage(rutaExcel);

    const rutaMinuscula =
        ruta.toLowerCase();

    const extensionValida =
        rutaMinuscula.endsWith(".xlsx") ||
        rutaMinuscula.endsWith(".xls");

    if (!extensionValida) {

        throw new Error(
            `El archivo no tiene una extensión Excel válida: "${ruta}".`
        );
    }

    return descargarArchivoStorage(ruta);
}


// =====================================================
// DESCARGAR DM.XLSX
// =====================================================
export async function obtenerArchivoDescansosMedicos() {

    return descargarExcelStorage(
        RUTAS_STORAGE.descansoMedico
    );
}


// =====================================================
// DESCARGAR ARCHIVO INDICANDO CARPETA Y NOMBRE
// Ejemplo:
// descargarArchivoPorCarpeta(
//     "RENDIMIENTO PERSONAL/PAPELETAS",
//     "papeletas.xlsx"
// );
// =====================================================
export async function descargarArchivoPorCarpeta(
    rutaCarpeta,
    nombreArchivo
) {

    const carpeta =
        limpiarRutaStorage(rutaCarpeta);

    const nombre =
        limpiarNombreArchivo(nombreArchivo);

    return descargarArchivoStorage(
        `${carpeta}/${nombre}`
    );
}


// =====================================================
// DESCARGAR EXCEL INDICANDO CARPETA Y NOMBRE
// =====================================================
export async function descargarExcelPorCarpeta(
    rutaCarpeta,
    nombreArchivo
) {

    const carpeta =
        limpiarRutaStorage(rutaCarpeta);

    const nombre =
        limpiarNombreArchivo(nombreArchivo);

    return descargarExcelStorage(
        `${carpeta}/${nombre}`
    );
}


// =====================================================
// OBTENER URL DE CUALQUIER ARCHIVO
// Útil para imágenes, PDF o enlaces de descarga.
// =====================================================
export async function obtenerUrlArchivoStorage(
    rutaArchivo
) {

    const ruta = limpiarRutaStorage(rutaArchivo);

    try {

        const archivoRef =
            obtenerReferenciaStorage(ruta);

        return await getDownloadURL(
            archivoRef
        );

    } catch (error) {

        console.error(
            `Error obteniendo URL de "${ruta}":`,
            error
        );

        throw new Error(
            obtenerMensajeErrorStorage(
                error,
                ruta
            )
        );
    }
}


// =====================================================
// LISTAR ARCHIVOS Y SUBCARPETAS
// =====================================================
export async function listarArchivosStorage(
    rutaCarpeta
) {

    const ruta = limpiarRutaStorage(rutaCarpeta);

    try {

        const carpetaRef =
            obtenerReferenciaStorage(ruta);

        const resultado =
            await listAll(carpetaRef);

        return {
            archivos: resultado.items.map(item => ({
                nombre: item.name,
                ruta: item.fullPath
            })),

            carpetas: resultado.prefixes.map(item => ({
                nombre: item.name,
                ruta: item.fullPath
            }))
        };

    } catch (error) {

        console.error(
            `Error listando "${ruta}":`,
            error
        );

        throw new Error(
            obtenerMensajeErrorStorage(
                error,
                ruta
            )
        );
    }
}


// =====================================================
// BUSCAR ARCHIVO POR NOMBRE DENTRO DE UNA CARPETA
// =====================================================
export async function buscarArchivoEnCarpeta(
    rutaCarpeta,
    nombreArchivo
) {

    const nombreBuscado =
        limpiarNombreArchivo(nombreArchivo)
            .toLowerCase();

    const contenido =
        await listarArchivosStorage(rutaCarpeta);

    return contenido.archivos.find(
        archivo =>
            archivo.nombre.toLowerCase() ===
            nombreBuscado
    ) || null;
}


// =====================================================
// LIMPIAR RUTA DE STORAGE
// =====================================================
function limpiarRutaStorage(rutaArchivo) {

    const ruta = String(rutaArchivo || "")
        .trim()
        .replace(/^\/+/, "")
        .replace(/\/+$/, "")
        .replace(/\/{2,}/g, "/");

    if (!ruta) {

        throw new Error(
            "La ruta de Firebase Storage está vacía."
        );
    }

    return ruta;
}


// =====================================================
// LIMPIAR NOMBRE DE ARCHIVO
// =====================================================
function limpiarNombreArchivo(nombreArchivo) {

    const nombre = String(nombreArchivo || "")
        .trim()
        .replace(/^\/+/, "")
        .replace(/\/+$/, "");

    if (!nombre) {

        throw new Error(
            "El nombre del archivo está vacío."
        );
    }

    return nombre;
}


// =====================================================
// MENSAJES DE ERROR DE STORAGE
// =====================================================
function obtenerMensajeErrorStorage(
    error,
    rutaArchivo
) {

    const codigo =
        error?.code || "";

    const mensaje =
        error?.message || "";

    switch (codigo) {

        case "storage/object-not-found":
            return (
                `No se encontró el archivo: "${rutaArchivo}".`
            );

        case "storage/unauthorized":
            return (
                `No tienes permiso para acceder a "${rutaArchivo}". ` +
                "Revisa las reglas de Firebase Storage."
            );

        case "storage/unauthenticated":
            return (
                `Debes iniciar sesión para acceder a "${rutaArchivo}".`
            );

        case "storage/canceled":
            return (
                `La descarga de "${rutaArchivo}" fue cancelada.`
            );

        case "storage/retry-limit-exceeded":
            return (
                `Se agotó el tiempo de descarga de "${rutaArchivo}".`
            );

        case "storage/download-size-exceeded":
            return (
                `El archivo "${rutaArchivo}" supera el tamaño permitido.`
            );

        case "storage/bucket-not-found":
            return (
                "No se encontró el bucket de Firebase Storage."
            );

        case "storage/project-not-found":
            return (
                "No se encontró el proyecto de Firebase."
            );

        case "storage/quota-exceeded":
            return (
                "Se superó la cuota de Firebase Storage."
            );

        default:

            if (
                mensaje.includes(
                    "Access-Control-Allow-Origin"
                ) ||
                mensaje.includes(
                    "Load failed"
                ) ||
                mensaje.includes(
                    "Failed to fetch"
                ) ||
                mensaje.includes(
                    "XMLHttpRequest"
                )
            ) {

                return (
                    `La descarga de "${rutaArchivo}" fue bloqueada por CORS.`
                );
            }

            return (
                mensaje ||
                `No se pudo acceder a "${rutaArchivo}".`
            );
    }
}
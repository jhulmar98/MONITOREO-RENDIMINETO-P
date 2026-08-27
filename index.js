// =====================================================
// 🔐 INDEX.JS
// SISTEMA DE MONITOREO INTEGRAL
// LISTO PARA PRODUCCIÓN
//
// FUNCIONES:
// - Verificación de autenticación
// - Protección de la página principal
// - Cierre seguro de sesión
// - Prevención de acceso mediante caché
// - Navegación mediante enlaces HTML
// =====================================================

import { auth } from "./backend/api.js";

import {
    onAuthStateChanged,
    signOut
} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js";


// =====================================================
// ⚙️ CONFIGURACIÓN
// =====================================================
const CONFIG = {

    rutaLogin:
        "loging/loging.html",

    idBotonLogout:
        "btnLogout"

};


// =====================================================
// 🌍 ESTADO
// =====================================================
const estado = {

    autenticado:
        false,

    cerrandoSesion:
        false,

    verificacionInicialCompleta:
        false

};


// =====================================================
// 🚀 INICIO
// =====================================================
document.addEventListener(
    "DOMContentLoaded",
    iniciarSistema
);

function iniciarSistema() {

    prepararInterfaz();

    configurarBotonLogout();

    protegerSesion();

    configurarProteccionCache();

}


// =====================================================
// 🖥️ PREPARAR INTERFAZ
// =====================================================
function prepararInterfaz() {

    document.body.style.visibility =
        "hidden";

    document.body.setAttribute(
        "aria-busy",
        "true"
    );

}


// =====================================================
// 🔐 PROTEGER SESIÓN
// =====================================================
function protegerSesion() {

    onAuthStateChanged(
        auth,
        usuario => {

            estado.verificacionInicialCompleta =
                true;

            if (
                !usuario
            ) {

                estado.autenticado =
                    false;

                redirigirAlLogin();

                return;

            }

            estado.autenticado =
                true;

            mostrarSistema();

            console.info(
                "✅ Sesión autenticada:",
                usuario.email || "Usuario sin correo"
            );

        },
        error => {

            console.error(
                "❌ Error verificando autenticación:",
                error
            );

            estado.autenticado =
                false;

            redirigirAlLogin();

        }
    );

}


// =====================================================
// 👁️ MOSTRAR SISTEMA
// =====================================================
function mostrarSistema() {

    document.body.style.visibility =
        "visible";

    document.body.removeAttribute(
        "aria-busy"
    );

}


// =====================================================
// 🔓 CONFIGURAR CIERRE DE SESIÓN
// =====================================================
function configurarBotonLogout() {

    const botonLogout =
        document.getElementById(
            CONFIG.idBotonLogout
        );

    if (
        !botonLogout
    ) {

        console.warn(
            "⚠️ No se encontró el botón de cierre de sesión."
        );

        return;

    }

    botonLogout.setAttribute(
        "role",
        "button"
    );

    botonLogout.setAttribute(
        "tabindex",
        "0"
    );

    botonLogout.addEventListener(
        "click",
        cerrarSesion
    );

    botonLogout.addEventListener(
        "keydown",
        evento => {

            if (
                evento.key === "Enter" ||
                evento.key === " "
            ) {

                evento.preventDefault();

                cerrarSesion();

            }

        }
    );

}


// =====================================================
// 🚪 CERRAR SESIÓN
// =====================================================
async function cerrarSesion() {

    if (
        estado.cerrandoSesion
    ) {

        return;

    }

    estado.cerrandoSesion =
        true;

    const botonLogout =
        document.getElementById(
            CONFIG.idBotonLogout
        );

    const contenidoOriginal =
        botonLogout?.innerHTML || "";

    try {

        bloquearBotonLogout(
            botonLogout
        );

        await signOut(
            auth
        );

        estado.autenticado =
            false;

        limpiarDatosLocales();

        redirigirAlLogin();

    } catch (error) {

        console.error(
            "❌ Error al cerrar sesión:",
            error
        );

        restaurarBotonLogout(
            botonLogout,
            contenidoOriginal
        );

        estado.cerrandoSesion =
            false;

    }

}


// =====================================================
// 🔒 BLOQUEAR BOTÓN DE LOGOUT
// =====================================================
function bloquearBotonLogout(
    boton
) {

    if (
        !boton
    ) {

        return;

    }

    boton.setAttribute(
        "aria-disabled",
        "true"
    );

    boton.style.pointerEvents =
        "none";

    boton.innerHTML = `

        <i class="fa-solid fa-spinner fa-spin"></i>

        Cerrando sesión...

    `;

}


// =====================================================
// 🔓 RESTAURAR BOTÓN
// =====================================================
function restaurarBotonLogout(
    boton,
    contenidoOriginal
) {

    if (
        !boton
    ) {

        return;

    }

    boton.removeAttribute(
        "aria-disabled"
    );

    boton.style.pointerEvents =
        "";

    boton.innerHTML =
        contenidoOriginal;

}


// =====================================================
// 🧹 LIMPIAR DATOS LOCALES
// =====================================================
function limpiarDatosLocales() {

    try {

        sessionStorage.clear();

        /*
         * No se ejecuta localStorage.clear() porque podría
         * eliminar preferencias o información necesaria de
         * otros módulos. Elimina únicamente las claves propias
         * del sistema cuando corresponda.
         */

        localStorage.removeItem(
            "usuario"
        );

        localStorage.removeItem(
            "sesion"
        );

        localStorage.removeItem(
            "token"
        );

    } catch (error) {

        console.warn(
            "⚠️ No se pudieron limpiar todos los datos locales:",
            error
        );

    }

}


// =====================================================
// 🔀 REDIRECCIÓN SEGURA AL LOGIN
// =====================================================
function redirigirAlLogin() {

    const rutaActual =
        window.location.pathname;

    if (
        rutaActual.endsWith(
            "/loging/loging.html"
        )
    ) {

        return;

    }

    window.location.replace(
        CONFIG.rutaLogin
    );

}


// =====================================================
// 🔒 PROTECCIÓN CONTRA CACHÉ DEL NAVEGADOR
// =====================================================
function configurarProteccionCache() {

    window.addEventListener(
        "pageshow",
        evento => {

            /*
             * Si la página fue recuperada desde el Back-Forward
             * Cache, se vuelve a comprobar la autenticación.
             */
            if (
                evento.persisted
            ) {

                verificarSesionActual();

            }

        }
    );

    document.addEventListener(
        "visibilitychange",
        () => {

            /*
             * Cuando el usuario vuelve a la pestaña, se valida
             * nuevamente la sesión si la página ya fue iniciada.
             */
            if (
                document.visibilityState === "visible" &&
                estado.verificacionInicialCompleta
            ) {

                verificarSesionActual();

            }

        }
    );

}


// =====================================================
// 🔍 VERIFICAR SESIÓN ACTUAL
// =====================================================
function verificarSesionActual() {

    const usuario =
        auth.currentUser;

    if (
        !usuario
    ) {

        estado.autenticado =
            false;

        redirigirAlLogin();

        return;

    }

    estado.autenticado =
        true;

    mostrarSistema();

}


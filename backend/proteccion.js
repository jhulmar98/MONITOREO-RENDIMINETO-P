// =====================================================
// 🔐 PROTECCION.JS
// PROTEGE LOS MÓDULOS PRIVADOS
// NO CIERRA SESIÓN
// =====================================================

import {
    auth
} from "./api.js";

import {
    onAuthStateChanged
} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js";


// =====================================================
// 🔒 OCULTAR LA PÁGINA MIENTRAS FIREBASE VALIDA
// =====================================================
document.documentElement.style.visibility =
    "hidden";


// =====================================================
// 🔍 VALIDAR SESIÓN
// =====================================================
onAuthStateChanged(

    auth,

    usuario => {

        if (!usuario) {

            console.warn(
                "Acceso bloqueado: no existe una sesión activa."
            );

            window.location.replace(
                "../loging/loging.html"
            );

            return;
        }

        console.info(
            "Acceso autorizado:",
            usuario.email || "Usuario autenticado"
        );

        document.documentElement.style.visibility =
            "visible";

    },

    error => {

        console.error(
            "Error verificando la sesión:",
            error
        );

        window.location.replace(
            "../loging/loging.html"
        );

    }

);


// =====================================================
// 🔒 REVISAR SESIÓN AL VOLVER DESDE EL HISTORIAL
// =====================================================
window.addEventListener(
    "pageshow",
    evento => {

        if (
            evento.persisted &&
            !auth.currentUser
        ) {

            window.location.replace(
                "../loging/loging.html"
            );

        }

    }
);
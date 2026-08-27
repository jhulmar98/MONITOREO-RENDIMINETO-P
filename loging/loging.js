// =====================================================
// 🔐 LOGIN FIREBASE AUTH - PRODUCCIÓN SEGURA
// =====================================================

import { auth } from "../backend/api.js";

import {
    signInWithEmailAndPassword,
    onAuthStateChanged,
    setPersistence,
    browserLocalPersistence
} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js";


// =====================================================
// 📌 REFERENCIAS DOM
// =====================================================
const form = document.getElementById("formLogin");

const errorDiv = document.getElementById("error");

const btn = document.getElementById("btnLogin");


// =====================================================
// 🔐 PERSISTENCIA
// Mantiene sesión aunque cierre navegador
// =====================================================
await setPersistence(
    auth,
    browserLocalPersistence
);


// =====================================================
// 🔒 EVITAR RETROCESO
// =====================================================
window.history.pushState(
    null,
    "",
    window.location.href
);

window.addEventListener("popstate", () => {

    window.history.pushState(
        null,
        "",
        window.location.href
    );

});


// =====================================================
// 🔍 VALIDAR SI YA EXISTE SESIÓN
// =====================================================
onAuthStateChanged(auth, (user) => {

    // ✅ YA LOGEADO
    if (user) {

        console.log(
            "🔓 Sesión activa:",
            user.email
        );

        window.location.replace(
            "../index.html"
        );

        return;
    }

    // ✅ MOSTRAR LOGIN
    document.body.style.visibility = "visible";

});


// =====================================================
// 🚀 LOGIN
// =====================================================
form.addEventListener("submit", async (e) => {

    e.preventDefault();

    const email = document
        .getElementById("email")
        .value
        .trim();

    const password = document
        .getElementById("password")
        .value
        .trim();

    try {

        activarCarga(true);

        limpiarError();

        // 🔐 LOGIN FIREBASE
        await signInWithEmailAndPassword(
            auth,
            email,
            password
        );

        // ✅ REDIRECCIÓN SEGURA
        window.location.replace(
            "../index.html"
        );

    } catch (error) {

        console.error(
            "❌ Error login:",
            error
        );

        mostrarError(
            traducirError(error.code)
        );

        activarCarga(false);
    }

});


// =====================================================
// 🎯 ESTADO BOTÓN
// =====================================================
function activarCarga(estado) {

    if (estado) {

        btn.innerHTML = `
            <i class="fa fa-spinner fa-spin"></i>
            Ingresando...
        `;

        btn.disabled = true;

        return;
    }

    btn.innerHTML = `
        <i class="fa fa-right-to-bracket"></i>
        Iniciar Sesión
    `;

    btn.disabled = false;
}


// =====================================================
// 🚨 MOSTRAR ERROR
// =====================================================
function mostrarError(mensaje) {

    errorDiv.textContent = mensaje;

    errorDiv.style.opacity = "1";

}


// =====================================================
// 🧹 LIMPIAR ERROR
// =====================================================
function limpiarError() {

    errorDiv.textContent = "";

    errorDiv.style.opacity = "0";

}


// =====================================================
// 🧠 TRADUCTOR ERRORES FIREBASE
// =====================================================
function traducirError(code) {

    switch (code) {

        case "auth/user-not-found":
            return "Usuario no encontrado";

        case "auth/wrong-password":
            return "Contraseña incorrecta";

        case "auth/invalid-email":
            return "Correo inválido";

        case "auth/invalid-credential":
            return "Credenciales inválidas";

        case "auth/too-many-requests":
            return "Demasiados intentos. Intenta más tarde";

        case "auth/network-request-failed":
            return "Sin conexión a internet";

        default:
            return "Error al iniciar sesión";
    }

}
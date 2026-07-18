const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbygJCQfpZO2JsOKdd3PPlnBCK9wsPjwv4zXv4oytS32LKqXCzmvYopAK5eteaRYEM-7yA/exec";

document.addEventListener("DOMContentLoaded", () => {
    // 1. Inicializar componentes
    initNavigation();
    initScrollReveal();
    cargarNovedades3D(); // Corregido el nombre para que coincida con la función

    // 2. Configurar Formulario de Turnos
    const form = document.getElementById('formTurno');
    if (form) form.addEventListener('submit', enviarFormulario);
});

function initScrollReveal() {
    const observerOptions = {
        threshold: 0.1,
        rootMargin: "0px 0px -50px 0px"
    };

    const revealObserver = new IntersectionObserver((entries) => {
        entries.forEach(entry => {
            if (entry.isIntersecting) {
                entry.target.classList.add('revealed');
                revealObserver.unobserve(entry.target); // Solo anima una vez
            }
        });
    }, observerOptions);

    // Aplicar a secciones clave
    document.querySelectorAll('.service-card-modern, .form-container-modern, .cinzel-title').forEach(el => {
        el.classList.add('reveal-hidden');
        revealObserver.observe(el);
    });
}

/* Manejo de la navegación y efectos del Header */
function initNavigation() {
    const navbar = document.querySelector('.custom-navbar');
    const menuElement = document.getElementById('menuLateral');
    
    window.addEventListener('scroll', () => {
        if (window.scrollY > 50) {
            navbar.classList.add('navbar-scrolled');
        } else {
            navbar.classList.remove('navbar-scrolled');
        }
    });

    document.querySelectorAll('.nav-item-modern').forEach(link => {
        link.addEventListener('click', function(e) {
            const href = this.getAttribute('href');

            const bsOffcanvas = bootstrap.Offcanvas.getOrCreateInstance(menuElement);
            if (bsOffcanvas) {
                bsOffcanvas.hide();
            }
            
            if (href.includes('.html') || href === '/') {
                e.preventDefault();
                setTimeout(() => {
                    window.location.href = href; 
                }, 300);
                return;
            }

            if (href.startsWith('#')) {
                e.preventDefault();
                const targetId = href.substring(1);
                const targetElement = document.getElementById(targetId);
                
                if (targetElement) {
                    setTimeout(() => {
                        targetElement.scrollIntoView({ behavior: 'smooth' });
                    }, 300);
                }
            }
        });
    });
}

/* Carga dinámica de Novedades (Corregido y Cerrado) */
async function cargarNovedades3D() {
    const contenedor = document.getElementById('contenedorCards3D');
    if (!contenedor) return;

    try {
        const urlFinal = `${WEB_APP_URL}?action=getNovedades&t=${Date.now()}`;
        const response = await fetch(urlFinal, { method: "GET", redirect: "follow" });
        const novedades = await response.json();

        if (!novedades || novedades.length === 0) {
            contenedor.innerHTML = `<div class="cinzel-title text-white">No hay novedades disponibles.</div>`;
            return;
        }

        contenedor.innerHTML = ""; // Limpiamos el loader

        const totalTarjetas = novedades.length * 2;
        const tiempoCiclo = 29; 
        const intervaloDelay = tiempoCiclo / totalTarjetas;

        let delayAcumulado = 0;

        novedades.forEach((nov, index) => {
            // 1. CREAR TARJETA DE IMAGEN
            const cardImagen = document.createElement('div');
            cardImagen.className = 'tarjeta img-dinamica';
            cardImagen.style.animationDelay = `-${delayAcumulado}s`;
            
            let imgUrl = nov.imagen || "https://images.unsplash.com/photo-1589829545856-d10d557cf95f?w=800";
            cardImagen.style.backgroundImage = `url('${imgUrl}')`;
            cardImagen.innerHTML = `<h3 style="display:none">${nov.titulo}</h3>`;
            
            contenedor.appendChild(cardImagen);
            delayAcumulado += intervaloDelay;

            // 2. CREAR TARJETA DE TEXTO
            const cardTexto = document.createElement('div');
            cardTexto.className = 'tarjeta txt-dinamico';
            cardTexto.style.animationDelay = `-${delayAcumulado}s`;
            
            cardTexto.innerHTML = `
                <span class="badge-bronce" style="font-size: 10px;">${nov.etiqueta || 'Novedad'}</span>
                <h2 class="cinzel-title" style="color: gold;">${nov.titulo}</h2>
                <div class="title-underline-premium" style="width: 50%; height: 2px; background: gold; margin: 10px 0;"></div>
                <p class="lead-text" style="font-size:0.85rem; color:white;">${nov.descripcion.substring(0, 120)}...</p>
                <a href="page-2.html" class="btn btn-sm btn-outline-light mt-2" style="font-size: 10px;">Ver más</a>
            `;
            
            contenedor.appendChild(cardTexto);
            delayAcumulado += intervaloDelay;
        });

    } catch (err) {
        console.error("Error al cargar novedades:", err);
        contenedor.innerHTML = `<div class="cinzel-title text-white">Error al sincronizar con el gestor de contenidos.</div>`;
    }
}

/* Procesamiento del Formulario de Turnos */
async function enviarFormulario(e) {
    e.preventDefault();
    
    // Animación de carga con SweetAlert2
    Swal.fire({
        title: 'Procesando Solicitud',
        text: 'Por favor aguarde un instante...',
        allowOutsideClick: false,
        didOpen: () => { Swal.showLoading(); }
    });

    const datos = {
        action: "addTurno",
        nombre: document.getElementById('nombre').value,
        telefono: document.getElementById('telefono').value,
        motivo: document.getElementById('motivo').value
    };

    try {
        const response = await fetch(WEB_APP_URL, {
            method: "POST",
            mode: "no-cors", // Requerido para llamadas directas a Apps Script
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(datos)
        });

        Swal.fire({
            icon: 'success',
            title: '¡Solicitud Recibida!',
            text: 'Nos comunicaremos a la brevedad vía WhatsApp para confirmar el horario.',
            customClass: { popup: 'swal-custom-soft' }
        });
        
        document.getElementById('formTurno').reset();

    } catch (error) {
        console.error("Error al enviar formulario:", error);
        Swal.fire({
            icon: 'error',
            title: 'Ups...',
            text: 'No pudimos registrar tu turno. Por favor, intentá nuevamente o contactanos por WhatsApp.'
        });
    }
}
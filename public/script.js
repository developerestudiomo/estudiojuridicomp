const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbygJCQfpZO2JsOKdd3PPlnBCK9wsPjwv4zXv4oytS32LKqXCzmvYopAK5eteaRYEM-7yA/exec";

document.addEventListener("DOMContentLoaded", () => {
    // 1. Inicializar componentes
    initNavigation();
    initScrollReveal();
    cargarNovedades();

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

/*Manejo de la navegación y efectos del Header */
function initNavigation() {
    const navbar = document.querySelector('.custom-navbar');
    const menuElement = document.getElementById('menuLateral');
    
    // 1. Efecto de scroll en el header
    window.addEventListener('scroll', () => {
        if (window.scrollY > 50) {
            navbar.classList.add('navbar-scrolled');
        } else {
            navbar.classList.remove('navbar-scrolled');
        }
    });

    // 2. Control total de clics en el Menú
    document.querySelectorAll('.nav-item-modern').forEach(link => {
        link.addEventListener('click', function(e) {
            const href = this.getAttribute('href');

            // --- PASO A: CERRAR EL MENÚ LATERAL ---
            // Usamos la API de Bootstrap para asegurar que se cierre y quite el fondo oscuro
            const bsOffcanvas = bootstrap.Offcanvas.getOrCreateInstance(menuElement);
            if (bsOffcanvas) {
                bsOffcanvas.hide();
            }

            // --- PASO B: MANEJO DE LA REDIRECCIÓN ---
            
            // Si el enlace es a una página (contiene .html o es la raíz /)
            if (href.includes('.html') || href === '/') {
                e.preventDefault(); // Detenemos cualquier script "fantasma"
                
                // Pequeño delay de 300ms para dejar que el menú se guarde antes de saltar
                setTimeout(() => {
                    window.location.href = href; 
                }, 300);
                return;
            }

            // Si es un ancla interna (#servicios, #contacto)
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

/*Carga dinámica de Novedades (Optimizado)*/
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

        // Configuramos la distribución de las tarjetas
        // Si tienes 4 novedades, generaremos 8 tarjetas (Imagen + Texto por cada una)
        const totalTarjetas = novedades.length * 2;
        const tiempoCiclo = 29; // Segundos de tu animación CSS
        const intervaloDelay = tiempoCiclo / totalTarjetas;

        let delayAcumulado = 0;

        novedades.forEach((nov, index) => {
            // 1. CREAR TARJETA DE IMAGEN
            const cardImagen = document.createElement('div');
            cardImagen.className = 'tarjeta img-dinamica';
            cardImagen.style.animationDelay = `-${delayAcumulado}s`;
            
            let imgUrl = nov.imagen || "https://images.unsplash.com/photo-1589829545856-d10d557cf95f?w=800";
            cardImagen.style.backgroundImage = `url('${imgUrl}')`;
            cardImagen.innerHTML = `<h3 style="display:none">${nov.titulo}</h3>`; // SEO/Accesibilidad
            
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
                <p class="lead-text">${nov.descripcion.substring(0, 120)}...</p>
                <a href="page-2.html" class="btn btn-sm btn-outline-light mt-2" style="font-size: 10px;">Ver más</a>
            `;
            
            contenedor.appendChild(cardTexto);
            delayAcumulado += intervaloDelay;
        });

    } catch (err) {
        console.error("Error en el fetch 3D:", err);
        contenedor.innerHTML = `<div class="text-danger">Error de sincronización 3D.</div>`;
    }
}

// Ejecutamos la carga
cargarNovedades3D();


/*Envío de Formulario con Feedback de alta calidad */
async function enviarFormulario(e) {
    e.preventDefault();
    const btn = e.target.querySelector('button');
    const originalText = btn.innerHTML;

    btn.disabled = true;
    btn.innerHTML = `<span class="spinner-border spinner-border-sm"></span> PROCESANDO...`;

    const payload = {
        nombre: document.getElementById('nombre').value,
        telefono: document.getElementById('telefono').value,
        motivo: document.getElementById('motivo').value,
        origen: "Web Premium MP"
    };

    try {
        await fetch(WEB_APP_URL, {
            method: "POST",
            mode: "no-cors",
            headers: { "Content-Type": "text/plain" },
            body: JSON.stringify(payload)
        });

        Swal.fire({
            title: '¡Solicitud Enviada!',
            text: 'La Dra. Peralta se comunicará con usted a la brevedad.',
            icon: 'success',
            confirmButtonColor: '#4A0E0E',
            customClass: { popup: 'swal-custom-soft' }
        });
        e.target.reset();
    } catch (err) {
        Swal.fire('Error', 'No pudimos procesar la solicitud. Intente nuevamente.', 'error');
    } finally {
        btn.disabled = false;
        btn.innerHTML = originalText;
    }
}

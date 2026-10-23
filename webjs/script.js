// HQL — interacciones del sitio

document.addEventListener('DOMContentLoaded', function () {
  // Menú móvil
  var navToggle = document.getElementById('navToggle');
  var mobileNav = document.getElementById('mobileNav');

  if (navToggle && mobileNav) {
    navToggle.addEventListener('click', function () {
      var isOpen = document.body.classList.toggle('nav-open');
      navToggle.setAttribute('aria-expanded', isOpen ? 'true' : 'false');
    });

    mobileNav.querySelectorAll('a').forEach(function (link) {
      link.addEventListener('click', function () {
        document.body.classList.remove('nav-open');
        navToggle.setAttribute('aria-expanded', 'false');
      });
    });
  }

  // Año en el footer
  var yearEl = document.getElementById('year');
  if (yearEl) {
    yearEl.textContent = new Date().getFullYear();
  }

  // Formulario de contacto (sin backend: confirma en pantalla)
  var form = document.getElementById('contactForm');
  var formNote = document.getElementById('formNote');

  if (form && formNote) {
    form.addEventListener('submit', function (e) {
      e.preventDefault();

      var nombre = form.nombre.value.trim();
      var email = form.email.value.trim();
      var empresa = form.empresa.value.trim();

      if (!nombre || !email || !empresa) {
        formNote.textContent = 'Completá nombre, email y empresa para continuar.';
        return;
      }

      formNote.textContent = 'Gracias, ' + nombre.split(' ')[0] + '. Te vamos a contactar a la brevedad.';
      form.reset();
    });
  }
});

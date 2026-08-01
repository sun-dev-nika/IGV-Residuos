(() => {
  "use strict";

  /* ============================================================
     CONFIGURACIÓN — pegar acá las URLs CSV publicadas de Google Sheets.
     Para obtener: Sheet → Archivo → Compartir → Publicar en la web →
     Hoja específica → CSV → Publicar → Copiar link.
     ============================================================ */
  const CONFIG = {
    PUNTOS_CSV_URL: "", // ← pegar URL CSV de la pestaña 'puntos_limpios'
    ALIADOS_CSV_URL: "", // ← pegar URL CSV de la pestaña 'aliados' (opcional)
    FALLBACK_URL: "data/fallback.json",
  };

  /* Fallback inline — se usa cuando el browser bloquea fetch() de archivos
     locales (al abrir con file://). En hosting real se intenta primero
     fallback.json y solo se usa esto si todo lo demás falla. */
  const FALLBACK_INLINE = {
    puntos_limpios: [
      {
        nombre: "Punto Limpio Arauco Express Las Brujas",
        comuna: "La Reina",
        direccion: "Carlos Silva Vildósola 9073, Arauco Express Las Brujas",
        estado: "Inaugurado",
        horarios: "Lunes a viernes: 9:30 a 13:00 · 14:00 a 19:30 hrs\nMartes: cerrado\nSábado a domingo: 9:00 a 13:00 · 14:00 a 17:00 hrs",
        descripcion: "Reciclaje cercano, ordenado y trazable para la comunidad. Recepción de residuos reciclables con señalética clara, acompañamiento a la comunidad y trazabilidad de la gestión.",
        residuos_aceptados: "PET 1, PEAD 2, vidrio, carton, papel, tetra pak, latas, aluminio",
        foto_url: "assets/imagenes/punto-las-brujas.jpg",
        activo: "si",
      },
      {
        nombre: "Punto Limpio Jorge Alessandri",
        comuna: "La Reina",
        direccion: "Jorge Alessandri 680, La Reina",
        estado: "Inaugurado",
        horarios: "Lunes: cerrado\nMartes a viernes: 8:30 a 19:30 hrs\nSábado y domingo: 9:00 a 17:00 hrs",
        descripcion: "Centro de reciclaje, tratamiento de residuos y educación ambiental en La Reina. Recepción, correcta separación y procesamiento de materiales, acercando la educación ambiental a la comunidad.",
        residuos_aceptados: "PET 1, PEAD 2, PEBD 4, PP 5, vidrio, carton, papel, tetra pak, latas, aluminio, libros, pilas",
        foto_url: "assets/imagenes/punto-jorge-alessandri.jpg",
        activo: "si",
      },
      {
        nombre: "Reciclómetro",
        comuna: "Plataforma",
        direccion: "",
        estado: "Plataforma",
        fecha_inauguracion: "",
        descripcion: "Reciclaje medible, participativo y conectado con beneficios. Herramienta de participación ciudadana que conecta reciclaje, trazabilidad y beneficios, permitiendo visibilizar el esfuerzo de las personas que reciclan.",
        residuos_aceptados: "",
        foto_url: "assets/imagenes/reciclometro.jpg",
        activo: "si",
      },
    ],
    aliados: [
      { orden: 1,  nombre: "Municipalidad de La Reina",      logo_url: "assets/imagenes/aliados/1.jpg"  },
      { orden: 2,  nombre: "CCU",                             logo_url: "assets/imagenes/aliados/2.jpg"  },
      { orden: 3,  nombre: "Parque Padre Hurtado",            logo_url: "assets/imagenes/aliados/3.jpg"  },
      { orden: 4,  nombre: "Arauco Express",                  logo_url: "assets/imagenes/aliados/4.jpg"  },
      { orden: 5,  nombre: "Municipalidad de Huechuraba",     logo_url: "assets/imagenes/aliados/5.jpg"  },
      { orden: 6,  nombre: "Municipalidad de Cabo de Hornos", logo_url: "assets/imagenes/aliados/6.jpg"  },
      { orden: 7,  nombre: "Municipalidad de Las Condes",     logo_url: "assets/imagenes/aliados/7.jpg"  },
      { orden: 8,  nombre: "Municipalidad de Temuco",         logo_url: "assets/imagenes/aliados/8.jpg"  },
      { orden: 9,  nombre: "Quiñenco",                        logo_url: "assets/imagenes/aliados/9.jpg"  },
      { orden: 10, nombre: "Shell",                           logo_url: "assets/imagenes/aliados/10.jpg" },
    ],
  };

  /* ============================================================
     UTILIDADES
     ============================================================ */

  const $ = (sel, root = document) => root.querySelector(sel);
  const $$ = (sel, root = document) => root.querySelectorAll(sel);

  const slug = (s) =>
    String(s || "")
      .toLowerCase()
      .normalize("NFD")
      .replace(/[̀-ͯ]/g, "")
      .replace(/[^a-z0-9]+/g, "-")
      .replace(/^-+|-+$/g, "");

  /**
   * Convierte un link de Google Drive a una URL de imagen directa visualizable.
   * Soporta los formatos:
   *   - https://drive.google.com/file/d/{ID}/view?usp=sharing
   *   - https://drive.google.com/open?id={ID}
   *   - https://drive.google.com/uc?id={ID}
   * Devuelve la URL original si no es de Drive (permite usar rutas locales).
   */
  function normalizarFotoUrl(url) {
    if (!url) return "";
    const u = String(url).trim();
    if (!u) return "";

    const matchFile = u.match(/drive\.google\.com\/file\/d\/([a-zA-Z0-9_-]+)/);
    if (matchFile) {
      return `https://drive.google.com/uc?export=view&id=${matchFile[1]}`;
    }
    const matchOpen = u.match(/drive\.google\.com\/(?:open|uc)\?(?:[^&]*&)*id=([a-zA-Z0-9_-]+)/);
    if (matchOpen) {
      return `https://drive.google.com/uc?export=view&id=${matchOpen[1]}`;
    }
    return u;
  }

  /** Convierte string "si"/"sí"/"true"/"1" a boolean. */
  function esActivo(valor) {
    if (valor == null) return true;
    const v = String(valor).trim().toLowerCase();
    if (!v) return true;
    return ["si", "sí", "true", "1", "yes", "x"].includes(v);
  }

  /** Slug del estado para usar como clase CSS del badge. */
  function badgeClass(estado) {
    const s = slug(estado);
    if (s.includes("inaugur")) return "badge--inaugurado";
    if (s.includes("proxim") || s.includes("pronto")) return "badge--proximo";
    if (s.includes("operac") || s.includes("activo")) return "badge--operacion";
    if (s.includes("servicio") || s.includes("plataform") || s.includes("iniciativ")) return "badge--servicio";
    return "";
  }

  /**
   * Formatea una fecha string que puede venir como "2026-03-15", "2026-07",
   * "marzo 2026" o cualquier formato libre. Si parsea como fecha válida,
   * devuelve "Inaugurado en marzo 2026". Si no, devuelve el string tal cual.
   */
  function formatearFecha(fecha, estado) {
    if (!fecha) return "";
    const raw = String(fecha).trim();
    if (!raw) return "";

    const meses = [
      "enero", "febrero", "marzo", "abril", "mayo", "junio",
      "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre",
    ];

    const m = raw.match(/^(\d{4})-(\d{1,2})(?:-(\d{1,2}))?$/);
    let textoFecha = raw;
    if (m) {
      const ano = m[1];
      const mes = meses[parseInt(m[2], 10) - 1] || "";
      textoFecha = mes ? `${mes} ${ano}` : ano;
    }

    const e = String(estado || "").toLowerCase();
    if (e.includes("proxim") || e.includes("pronto")) {
      return `Próximamente · ${textoFecha}`;
    }
    return `Inaugurado en ${textoFecha}`;
  }

  /* ============================================================
     CARGA DE DATOS
     ============================================================ */

  async function cargarPuntosLimpios() {
    if (CONFIG.PUNTOS_CSV_URL) {
      try {
        const data = await fetchCSV(CONFIG.PUNTOS_CSV_URL);
        if (data && data.length) return { puntos: data, fuente: "sheets" };
      } catch (err) {
        console.warn("Falló fetch de Google Sheets, usando fallback:", err);
      }
    }
    try {
      const fallback = await fetchJSON(CONFIG.FALLBACK_URL);
      return { puntos: fallback.puntos_limpios || [], fuente: "fallback" };
    } catch (err) {
      console.warn("Falló fetch de fallback.json (file:// CORS?), usando inline:", err);
      return { puntos: FALLBACK_INLINE.puntos_limpios, fuente: "inline" };
    }
  }

  async function cargarAliados() {
    if (CONFIG.ALIADOS_CSV_URL) {
      try {
        const data = await fetchCSV(CONFIG.ALIADOS_CSV_URL);
        if (data && data.length) return data;
      } catch (err) {
        console.warn("Falló fetch de aliados, usando fallback:", err);
      }
    }
    try {
      const fallback = await fetchJSON(CONFIG.FALLBACK_URL);
      return fallback.aliados || FALLBACK_INLINE.aliados;
    } catch (err) {
      console.warn("Falló fetch de fallback.json (file:// CORS?), usando inline:", err);
      return FALLBACK_INLINE.aliados;
    }
  }

  function fetchCSV(url) {
    return new Promise((resolve, reject) => {
      if (typeof Papa === "undefined") {
        reject(new Error("PapaParse no está disponible"));
        return;
      }
      const cacheBuster = `${url}${url.includes("?") ? "&" : "?"}t=${Date.now()}`;
      Papa.parse(cacheBuster, {
        download: true,
        header: true,
        skipEmptyLines: "greedy",
        complete: (results) => {
          if (results.errors && results.errors.length) {
            console.warn("Errores parseando CSV:", results.errors);
          }
          resolve(results.data || []);
        },
        error: (err) => reject(err),
      });
    });
  }

  async function fetchJSON(url) {
    const res = await fetch(url, { cache: "no-cache" });
    if (!res.ok) throw new Error(`HTTP ${res.status} al leer ${url}`);
    return res.json();
  }

  /* ============================================================
     RENDER: PUNTOS LIMPIOS
     ============================================================ */

  function renderPuntosLimpios(puntos, fuente) {
    const grid = $("#puntosGrid");
    const feedback = $("#puntosFeedback");
    const tpl = $("#puntoCardTpl");
    if (!grid || !tpl) return;

    const visibles = puntos.filter((p) => esActivo(p.activo)).sort((a, b) => {
      const fa = String(a.fecha_inauguracion || "");
      const fb = String(b.fecha_inauguracion || "");
      return fb.localeCompare(fa);
    });

    grid.innerHTML = "";
    grid.setAttribute("aria-busy", "false");

    if (!visibles.length) {
      grid.innerHTML = "";
      if (feedback) {
        feedback.hidden = false;
        feedback.textContent = "Aún no hay puntos limpios publicados.";
      }
      return;
    }

    const fragment = document.createDocumentFragment();
    visibles.forEach((punto, idx) => {
      const node = tpl.content.cloneNode(true);
      const img = node.querySelector("img");
      const badge = node.querySelector(".punto-card__badge");
      const codeEl = node.querySelector(".punto-card__code");
      const comunaEl = node.querySelector(".punto-card__comuna");
      const nombreEl = node.querySelector(".punto-card__nombre");
      const fechaEl = node.querySelector(".punto-card__fecha");
      const descEl = node.querySelector(".punto-card__desc");
      const residuosEl = node.querySelector(".punto-card__residuos");
      const ctaEl = node.querySelector(".punto-card__cta");
      const metaCells = node.querySelectorAll(".punto-card__meta-cell");

      const fotoUrl = normalizarFotoUrl(punto.foto_url);
      if (fotoUrl) {
        img.src = fotoUrl;
        img.alt = `${punto.nombre} — Punto limpio en ${punto.comuna}`;
        img.addEventListener("error", () => {
          img.removeAttribute("src");
          img.style.display = "none";
        });
      } else {
        img.style.display = "none";
      }

      const cls = badgeClass(punto.estado);
      badge.textContent = punto.estado || "Activo";
      if (cls) badge.classList.add(cls);

      const codigo = `PL-${String(idx + 1).padStart(3, "0")}`;
      if (codeEl) codeEl.textContent = codigo;

      comunaEl.textContent = `[ ${String(punto.comuna || "").toUpperCase()} ]`;
      nombreEl.textContent = punto.nombre || "";

      const direccionEl = node.querySelector(".punto-card__direccion");
      if (direccionEl) {
        if (punto.direccion && String(punto.direccion).trim()) {
          direccionEl.textContent = punto.direccion;
        } else {
          direccionEl.remove();
        }
      }

      const horariosRaw = String(punto.horarios || "").trim();
      if (horariosRaw) {
        fechaEl.classList.add("punto-card__horarios");
        fechaEl.textContent = "";
        horariosRaw
          .split(/\r?\n/)
          .map((l) => l.trim())
          .filter(Boolean)
          .forEach((linea) => {
            const span = document.createElement("span");
            span.textContent = linea;
            fechaEl.appendChild(span);
          });
      } else {
        const fechaText = formatearFecha(punto.fecha_inauguracion, punto.estado);
        if (fechaText) {
          fechaEl.textContent = fechaText;
        } else {
          fechaEl.remove();
        }
      }

      descEl.textContent = punto.descripcion || "";

      const residuos = String(punto.residuos_aceptados || "")
        .split(/[,;]/)
        .map((r) => r.trim())
        .filter(Boolean);
      residuos.forEach((r) => {
        const li = document.createElement("li");
        li.textContent = r;
        residuosEl.appendChild(li);
      });
      if (!residuos.length) residuosEl.remove();

      if (ctaEl) {
        if (String(punto.estado || "").trim() === "Plataforma") {
          ctaEl.textContent = "Conocer más →";
        } else {
          ctaEl.remove();
        }
      }

      if (metaCells.length >= 3) {
        metaCells[0].textContent = codigo;
        metaCells[1].textContent = `EST: ${(punto.estado || "OK").substring(0, 8).toUpperCase()}`;
        metaCells[2].textContent = String(punto.fecha_inauguracion || "—").replace(/-/g, ".");
      }

      fragment.appendChild(node);
    });

    grid.appendChild(fragment);

    if (feedback) {
      if (fuente === "fallback" && CONFIG.PUNTOS_CSV_URL) {
        feedback.hidden = false;
        feedback.textContent = "Mostrando datos de respaldo (no se pudo conectar con la fuente en vivo).";
      } else {
        feedback.hidden = true;
      }
    }
  }

  /* ============================================================
     RENDER: ALIADOS
     ============================================================ */

  function renderAliados(aliados) {
    const grid = $("#aliadosGrid");
    const tpl = $("#aliadoTpl");
    if (!grid || !tpl) return;

    grid.innerHTML = "";
    const visibles = aliados.filter((a) => a && a.nombre);
    if (!visibles.length) return;

    const ordenados = [...visibles].sort((a, b) => {
      const oa = Number(a.orden) || 999;
      const ob = Number(b.orden) || 999;
      return oa - ob;
    });

    const fragment = document.createDocumentFragment();

    ordenados.forEach((aliado, idx) => {
      const node = tpl.content.cloneNode(true);
      const card = node.querySelector(".aliado-card");
      const img = node.querySelector("img");
      const numEl = node.querySelector(".aliado-card__num");
      const nameEl = node.querySelector(".aliado-card__name");

      const numero = String(Number(aliado.orden) || idx + 1).padStart(2, "0");
      if (numEl) numEl.textContent = `[ ${numero} ]`;
      if (nameEl) nameEl.textContent = aliado.nombre;

      const logoUrl = normalizarFotoUrl(aliado.logo_url);
      if (logoUrl) {
        img.src = logoUrl;
        img.alt = aliado.nombre;
        img.addEventListener("error", () => {
          card.classList.add("no-image");
        });
      } else {
        card.classList.add("no-image");
      }

      fragment.appendChild(node);
    });

    grid.appendChild(fragment);
  }

  /* ============================================================
     UI: NAV, REVEAL, STICKY HEADER
     ============================================================ */

  function initUI() {
    const header = $("#siteHeader");
    const navToggle = $("#navToggle");
    const mobileNav = $("#mobileNav");
    const yearEl = $("#year");

    if (yearEl) yearEl.textContent = new Date().getFullYear();

    const onScroll = () => {
      if (!header) return;
      header.classList.toggle("scrolled", window.scrollY > 8);
    };
    onScroll();
    window.addEventListener("scroll", onScroll, { passive: true });

    if (navToggle && mobileNav) {
      const closeMenu = () => {
        navToggle.setAttribute("aria-expanded", "false");
        navToggle.setAttribute("aria-label", "Abrir menú");
        mobileNav.classList.remove("open");
      };

      navToggle.addEventListener("click", () => {
        const isOpen = mobileNav.classList.toggle("open");
        navToggle.setAttribute("aria-expanded", String(isOpen));
        navToggle.setAttribute("aria-label", isOpen ? "Cerrar menú" : "Abrir menú");
      });

      mobileNav.querySelectorAll("a").forEach((link) => {
        link.addEventListener("click", closeMenu);
      });

      document.addEventListener("keydown", (e) => {
        if (e.key === "Escape") closeMenu();
      });
    }

    const revealEls = $$(".reveal");
    if ("IntersectionObserver" in window && revealEls.length) {
      const io = new IntersectionObserver(
        (entries) => {
          entries.forEach((entry) => {
            if (entry.isIntersecting) {
              entry.target.classList.add("is-visible");
              io.unobserve(entry.target);
            }
          });
        },
        { threshold: 0.12, rootMargin: "0px 0px -40px 0px" }
      );
      revealEls.forEach((el) => io.observe(el));
    } else {
      revealEls.forEach((el) => el.classList.add("is-visible"));
    }
  }

  /* ============================================================
     ARRANQUE
     ============================================================ */

  async function start() {
    initUI();

    try {
      const { puntos, fuente } = await cargarPuntosLimpios();
      renderPuntosLimpios(puntos, fuente);
    } catch (err) {
      console.error("Error cargando puntos limpios:", err);
      const feedback = $("#puntosFeedback");
      if (feedback) {
        feedback.hidden = false;
        feedback.textContent = "No se pudieron cargar los puntos limpios en este momento.";
      }
      $("#puntosGrid")?.setAttribute("aria-busy", "false");
    }

    try {
      const aliados = await cargarAliados();
      renderAliados(aliados);
    } catch (err) {
      console.error("Error cargando aliados:", err);
    }
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", start);
  } else {
    start();
  }
})();

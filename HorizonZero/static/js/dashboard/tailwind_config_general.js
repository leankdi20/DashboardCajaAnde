tailwind.config = {
  darkMode: "class",
  theme: {
    extend: {
      colors: {
        // ── Colores institucionales principales ──
        "primary":              "#003fb7",   // Azul institucional
        "primary-dark":         "#002d85",   // Azul oscuro (hover)
        "primary-light":        "#e8eeff",   // Azul muy claro (fondos)
        "primary-fixed":        "#c8d6ff",
        "primary-fixed-dim":    "#99b0ff",
        "on-primary":           "#ffffff",
        "on-primary-fixed":     "#001257",
        "primary-container":    "#0050e6",   // Azul medio (botones secundarios)
        "on-primary-container": "#ffffff",

        // ── Amarillo institucional (acento) ──
        "secondary":              "#ffc900",  // Amarillo institucional
        "secondary-dark":         "#e6b400",  // Amarillo oscuro (hover)
        "secondary-light":        "#fff8e0",  // Amarillo muy claro (fondos)
        "secondary-container":    "#ffd740",
        "on-secondary":           "#1a1000",  // Texto sobre amarillo (negro)
        "on-secondary-container": "#1a1000",

        // ── Superficies y fondos ──
        "background":                   "#f4f6fb",
        "background-dark":              "#0f1521",
        "surface":                      "#ffffff",
        "surface-dim":                  "#dde3f0",
        "surface-bright":               "#f9faff",
        "surface-container-lowest":     "#ffffff",
        "surface-container-low":        "#f0f3fb",
        "surface-container":            "#e8ecf7",
        "surface-container-high":       "#dde3f0",
        "surface-container-highest":    "#d2d8ec",
        "on-surface":                   "#191c24",
        "on-surface-variant":           "#44495a",
        "on-background":                "#191c24",
        "inverse-surface":              "#2e3140",
        "inverse-on-surface":           "#f0f3fb",

        // ── Bordes y outlines ──
        "outline":          "#74788a",
        "outline-variant":  "#c4c8dc",

        // ── Terciario (azul claro para acentos secundarios) ──
        "tertiary":           "#0065a3",
        "tertiary-container": "#cce5ff",
        "on-tertiary":        "#ffffff",

        // ── Estados ──
        "error":              "#ba1a1a",
        "error-container":    "#ffdad6",
        "on-error":           "#ffffff",
        "on-error-container": "#410002",

        // ── Atajos semánticos (para usar en clases como bg-accent) ──
        "accent":    "#ffc900",   // = secondary, para conveniencia
        "muted":     "#74788a",
        "subtle":    "#f0f3fb",
      },

      fontFamily: {
        "headline": ["Public Sans", "sans-serif"],
        "body":     ["Public Sans", "sans-serif"],
        "label":    ["Public Sans", "sans-serif"],
      },

      borderRadius: {
        "DEFAULT": "0.375rem",
        "lg":      "0.625rem",
        "xl":      "0.875rem",
        "2xl":     "1.25rem",
        "full":    "9999px",
      },
    },
  },
}
 tailwind.config = {
            darkMode: "class",
            theme: {
                extend: {
                    colors: {
                        // primary: "#003fb7",
                        // accent: "#ffc900",
                        "primary": "#003fb7",
                        "background-light": "#f8f6f6",
                        "background-dark": "#ffc900",
                    },
                    fontFamily: {
                        "display": ["Public Sans"]
                    },
                    borderRadius: {
                        "DEFAULT": "0.25rem",
                        "lg": "0.5rem",
                        "xl": "0.75rem",
                        "full": "9999px"
                    },
                },
            },
        }

function toggleMenu(menuId, iconId) {
    const menu = document.getElementById(menuId);
    const icon = document.getElementById(iconId);

    menu.classList.toggle("hidden");
    icon.classList.toggle("rotate-80");
    }
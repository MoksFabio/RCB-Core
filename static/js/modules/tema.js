 /**
 * Módulo: Tema e Acessibilidade
 * Resumo: Controla o tema visual (claro/escuro/alto contraste), cores personalizadas e tamanho da fonte.
 */
window.App = window.App || {};

 window.App.theme = {
      elements: {
        temaSelect: document.getElementById("tema"),
        customColorPickerContainer: document.getElementById("customColorPickerContainer"),
        customColorPicker: document.getElementById("customColorPicker"),
        salvarTemaBtn: document.getElementById("salvarTema"),
        altoContrasteCheckbox: document.getElementById("altoContraste"),
        fontSizeSlider: document.getElementById("fontSizeSlider"),
        fontSizeValue: document.getElementById("fontSizeValue"),
        salvarAcessibilidadeBtn: document.getElementById("salvarAcessibilidade"),
        acessibilidadeSaveStatus: document.getElementById("acessibilidadeSaveStatus"),
      },
      init() {
        // Re-query elements strictly on init to avoid stale references
        this.elements.temaSelect = document.getElementById("tema");
        this.elements.customColorPickerContainer = document.getElementById("customColorPickerContainer");
        this.elements.customColorPicker = document.getElementById("customColorPicker");
        this.elements.salvarTemaBtn = document.getElementById("salvarTema");
        this.elements.altoContrasteCheckbox = document.getElementById("altoContraste");
        this.elements.fontSizeSlider = document.getElementById("fontSizeSlider");
        this.elements.fontSizeValue = document.getElementById("fontSizeValue");
        this.elements.salvarAcessibilidadeBtn = document.getElementById("salvarAcessibilidade");
        this.elements.acessibilidadeSaveStatus = document.getElementById("acessibilidadeSaveStatus");

        this.loadSettings();
        this.addListeners();
      },
      addListeners() {
        const { temaSelect, salvarTemaBtn, salvarAcessibilidadeBtn, fontSizeSlider, altoContrasteCheckbox } = this.elements;

        if (temaSelect) temaSelect.addEventListener("change", () => this.toggleColorPickerVisibility());
        if (salvarTemaBtn) salvarTemaBtn.addEventListener("click", () => this.saveTheme());
        if (salvarAcessibilidadeBtn) salvarAcessibilidadeBtn.addEventListener("click", () => this.saveAccessibility());
        if (fontSizeSlider) fontSizeSlider.addEventListener("input", () => this.applyFontSize(fontSizeSlider.value));
        if (altoContrasteCheckbox) altoContrasteCheckbox.addEventListener("change", () => {
          this.applyAccessibilitySettings(altoContrasteCheckbox.checked, fontSizeSlider?.value || 100);
        });
      },
      darkenColor(hex, percent) {
        hex = hex.replace(/^#/, "");
        let r = parseInt(hex.substring(0, 2), 16);
        let g = parseInt(hex.substring(2, 4), 16);
        let b = parseInt(hex.substring(4, 6), 16);
        r = Math.max(0, Math.floor(r * (1 - percent / 100)));
        g = Math.max(0, Math.floor(g * (1 - percent / 100)));
        b = Math.max(0, Math.floor(b * (1 - percent / 100)));
        return `#${r.toString(16).padStart(2, "0")}${g.toString(16).padStart(2, "0")}${b.toString(16).padStart(2, "0")}`;
      },
      toggleColorPickerVisibility() {
        const { temaSelect, customColorPickerContainer } = this.elements;
        if (!temaSelect || !customColorPickerContainer) return;
        customColorPickerContainer.style.display = temaSelect.value === "personalizado" ? "block" : "none";
      },
      applyTheme(theme) {
        const body = document.body;
        const { temaSelect, customColorPicker } = this.elements;
        const rootStyle = document.documentElement.style;

        if (body.classList.contains("alto-contraste") && theme !== "alto-contraste") {
          body.classList.remove("alto-contraste");
        }
        body.classList.remove("dark-theme");
        rootStyle.removeProperty("--cor-principal");
        rootStyle.removeProperty("--cor-principal-hover");
        rootStyle.removeProperty("--cor-principal-gradient");

        if (theme === "escuro") {
          body.classList.add("dark-theme");
        } else if (theme === "personalizado") {
          const customColor = customColorPicker ? customColorPicker.value : "#f97316";
          const customColorHover = this.darkenColor(customColor, 10);
          rootStyle.setProperty("--cor-principal", customColor);
          rootStyle.setProperty("--cor-principal-hover", customColorHover);
          rootStyle.setProperty("--cor-principal-gradient", customColor);
        }

        if (temaSelect) {
          temaSelect.value = (theme === "alto-contraste") ? "claro" : theme;
        }
        this.toggleColorPickerVisibility();
      },
      applyFontSize(size) {
        const { fontSizeValue, fontSizeSlider } = this.elements;
        const sizePercentage = parseInt(size, 10);
        document.documentElement.style.setProperty("--tamanho-texto-base", `${sizePercentage / 100}rem`);
        if (fontSizeValue) fontSizeValue.textContent = `${sizePercentage}`;
        if (fontSizeSlider) fontSizeSlider.value = sizePercentage;
      },
      applyAccessibilitySettings(altoContraste, fontSize) {
        const body = document.body;
        const { temaSelect, customColorPickerContainer, altoContrasteCheckbox } = this.elements;

        if (altoContraste) {
          body.classList.add("alto-contraste");
          body.classList.remove("dark-theme");
          document.documentElement.style.removeProperty("--cor-principal");
          document.documentElement.style.removeProperty("--cor-principal-hover");
          document.documentElement.style.removeProperty("--cor-principal-gradient");
          if (temaSelect) temaSelect.value = "claro";
          if (customColorPickerContainer) customColorPickerContainer.style.display = "none";
        } else if (body.classList.contains("alto-contraste")) {
          body.classList.remove("alto-contraste");
          const savedTheme = localStorage.getItem("theme") || "claro";
          this.applyTheme(savedTheme);
        }

        this.applyFontSize(fontSize);
        if (altoContrasteCheckbox) altoContrasteCheckbox.checked = altoContraste;
      },
      loadSettings() {
        const savedTheme = localStorage.getItem("theme") || "claro";
        const savedAltoContraste = localStorage.getItem("altoContraste") === "true";
        const savedFontSize = localStorage.getItem("fontSize") || "80";
        const savedCustomColor = localStorage.getItem("customThemeColor") || "#f97316";

        if (this.elements.customColorPicker) this.elements.customColorPicker.value = savedCustomColor;

        this.applyAccessibilitySettings(savedAltoContraste, savedFontSize);

        if (!savedAltoContraste) {
          this.applyTheme(savedTheme);
        } else {
          if (this.elements.temaSelect) this.elements.temaSelect.value = "claro";
        }

        this.toggleColorPickerVisibility();
      },
      saveTheme() {
        const { temaSelect, customColorPicker, altoContrasteCheckbox, fontSizeSlider } = this.elements;
        if (!temaSelect) return;
        const theme = temaSelect.value;

        if (document.body.classList.contains("alto-contraste")) {
          if (altoContrasteCheckbox) altoContrasteCheckbox.checked = false;
          this.applyAccessibilitySettings(false, fontSizeSlider?.value || 100);
        }

        this.applyTheme(theme);
        localStorage.setItem("theme", theme);
        if (theme === "personalizado" && customColorPicker) {
          localStorage.setItem("customThemeColor", customColorPicker.value);
        }
      },
      saveAccessibility() {
        const { altoContrasteCheckbox, fontSizeSlider, acessibilidadeSaveStatus } = this.elements;
        const altoContraste = altoContrasteCheckbox ? altoContrasteCheckbox.checked : false;
        const newFontSize = fontSizeSlider ? fontSizeSlider.value : "100";

        localStorage.setItem("altoContraste", String(altoContraste));
        localStorage.setItem("fontSize", newFontSize);
        this.applyAccessibilitySettings(altoContraste, newFontSize);

        if (acessibilidadeSaveStatus) {
          acessibilidadeSaveStatus.textContent = "Configurações salvas!";
          acessibilidadeSaveStatus.classList.remove("opacity-0");
          setTimeout(() => {
            acessibilidadeSaveStatus.classList.add("opacity-0");
          }, 2000);
        }
      },
    };

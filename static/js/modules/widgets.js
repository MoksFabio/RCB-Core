 /**
 * Módulo: Widgets Diversos
 * Resumo: Componentes menores como o Relógio em tempo real e o bloco de Notas Rápidas.
 */
window.App = window.App || {};

 window.App.clock = {
      init() {
        this.update();
        setInterval(() => this.update(), 1000);
      },
      update() {
        const el = document.getElementById("currentTime");
        if (!el) return;
        const now = new Date();
        try {
          el.textContent = now.toLocaleTimeString("pt-BR", {
            timeZone: "America/Recife", hour: "2-digit", minute: "2-digit", second: "2-digit", hour12: false
          });
        } catch (e) {
          el.textContent = now.toLocaleTimeString("pt-BR");
        }
      },
    };

 window.App.quickNotes = {
      init() {
        const notesTextarea = document.getElementById("notesTextarea");
        if (!notesTextarea) return; // SEGURANÇA

        this.load();
        notesTextarea.addEventListener("input", () => {
          clearTimeout(App.state.notesSaveTimeout);
          const notesSaveStatus = document.getElementById("notesSaveStatus");
          if (notesSaveStatus) {
            notesSaveStatus.textContent = "Salvando...";
            notesSaveStatus.classList.remove("opacity-0", "text-green-600", "text-red-600");
            notesSaveStatus.classList.add("text-gray-500");
            notesSaveStatus.classList.remove("opacity-0");
          }
          App.state.notesSaveTimeout = setTimeout(() => {
            this.save(notesTextarea.value);
          }, 1000);
        });
      },
      load() {
        const notesTextarea = document.getElementById("notesTextarea");
        const savedNotes = localStorage.getItem("quickNotes");
        if (savedNotes && notesTextarea) {
          notesTextarea.value = savedNotes;
        }
      },
      save(notes) {
        const notesSaveStatus = document.getElementById("notesSaveStatus");
        localStorage.setItem("quickNotes", notes);
        if (notesSaveStatus) {
          notesSaveStatus.textContent = "Salvo!";
          notesSaveStatus.classList.remove("text-gray-500");
          notesSaveStatus.classList.add("text-green-600");
          setTimeout(() => {
            notesSaveStatus.classList.add("opacity-0");
          }, 2000);
        }
      },
    };

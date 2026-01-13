 /**
 * Módulo: Painel (Dashboard)
 * Resumo: Gerencia widgets específicos da dashboard, como a agenda de compromissos e resumo de eventos.
 */
window.App = window.App || {};

 window.App.dashboardAgenda = {
      elements: {
        // on demand
      },

      init() {
        // SEGURANÇA: Só roda se a lista de eventos existir na página
        const lista = document.getElementById("dashboard-eventos-lista");
        if (!lista) return;

        this.addListeners();
        this.loadEvents();

        App.modals.addModalListener("dashboard-novo-evento-btn", "dashboardEventoModal", () => this.openNewEventModal());
        App.modals.addModalListener("closeDashboardEventoModal", "dashboardEventoModal");
      },

      addListeners() {
        const form = document.getElementById("dashboard-evento-form");
        if (form) form.addEventListener("submit", (e) => {
          e.preventDefault();
          this.saveEvent();
        });

        const lista = document.getElementById("dashboard-eventos-lista");
        if (lista) lista.addEventListener("click", (e) => {
          const target = e.target;
          const item = target.closest(".dashboard-event-item");
          if (!item) return;

          const eventId = item.dataset.id;

          if (target.closest("[data-action='toggle-status']")) {
            const currentStatus = item.dataset.status;
            const newStatus = (currentStatus === "Agendado") ? "Concluído" : "Agendado";
            this.updateEventStatus(eventId, newStatus);
          }

          if (target.closest("[data-action='cancel-event']")) {
            App.modals.confirm(
                "Cancelar Evento",
                "Tem certeza que deseja cancelar este evento?",
                () => this.updateEventStatus(eventId, "Cancelado")
            );
          }
        });
      },

      async loadEvents() {
        try {
          const response = await fetch("/api/compromissos");
          if (!response.ok) throw new Error("Falha ao carregar agenda.");
          const data = await response.json();
          this.render(data.compromissos || []);
        } catch (error) {
          console.error("Erro ao carregar eventos da agenda:", error);
          const lista = document.getElementById("dashboard-eventos-lista");
          if (lista) {
            lista.innerHTML = `<p class="text-sm text-red-500">Erro ao carregar eventos.</p>`;
          }
        }
      },

      render(events) {
        const lista = document.getElementById("dashboard-eventos-lista");
        const empty = document.getElementById("dashboard-eventos-empty");
        if (!lista || !empty) return;

        lista.innerHTML = "";

        const hoje = new Date().setHours(0, 0, 0, 0);

        const eventosFiltrados = events
          .filter(e => e.status !== "Cancelado" && new Date(e.date + "T00:00:00") >= hoje)
          .sort((a, b) => new Date(a.date + "T" + a.start_time) - new Date(b.date + "T" + b.start_time))
          .slice(0, 5);

        if (eventosFiltrados.length === 0) {
          empty.classList.remove("hidden");
          return;
        }
        empty.classList.add("hidden");

        eventosFiltrados.forEach(evento => {
          const li = document.createElement("li");
          li.className = "dashboard-event-item";
          li.dataset.id = evento.id;
          li.dataset.status = evento.status;

          const isConcluido = evento.status === "Concluído";

          li.innerHTML = `
            <span class="dashboard-event-item__time">${evento.start_time}</span>
            <div class="dashboard-event-item__details">
              <p title="${evento.title}">${evento.title}</p>
              <p class="text-sm" title="${evento.description || ''}">${evento.description || 'Sem descrição'}</p>
            </div>
            <div class="dashboard-event-item__controls">
              <button class="btn btn-icon btn-ghost btn-sm" data-action="toggle-status" title="${isConcluido ? 'Marcar como Agendado' : 'Marcar como Concluído'}">
                <span class="material-icons-outlined">${isConcluido ? 'check_box' : 'check_box_outline_blank'}</span>
              </button>
              <button class="btn btn-icon btn-ghost btn-sm text-red-500" data-action="cancel-event" title="Cancelar Evento">
                <span class="material-icons-outlined">event_busy</span>
              </button>
            </div>
          `;
          lista.appendChild(li);
        });
      },

      openNewEventModal() {
        const form = document.getElementById("dashboard-evento-form");
        const idInput = document.getElementById("dashboard-evento-id");
        const dateInput = document.getElementById("dashboard-evento-date");

        if (form) form.reset();
        if (idInput) idInput.value = "";
        if (dateInput) dateInput.value = new Date().toISOString().split('T')[0];

        App.modals.open("dashboardEventoModal");
      },

      async saveEvent() {
        const id = document.getElementById("dashboard-evento-id")?.value;
        const title = document.getElementById("dashboard-evento-title")?.value;
        const date = document.getElementById("dashboard-evento-date")?.value;
        const time = document.getElementById("dashboard-evento-time")?.value;
        const desc = document.getElementById("dashboard-evento-desc")?.value;

        const data = {
          title: title,
          date: date,
          start_time: time,
          description: desc,
          status: "Agendado"
        };

        const url = id ? `/api/compromissos/${id}/` : "/api/compromissos/";
        const method = id ? "PUT" : "POST";

        try {
          const response = await fetch(url, {
            method: method,
            headers: { 
                "Content-Type": "application/json",
                "X-CSRFToken": RCBUtils.getCookie('csrftoken')
            },
            body: JSON.stringify(data),
          });
          if (!response.ok) {
              const errorData = await response.json();
              throw new Error(errorData.message || "Erro ao salvar o evento.");
          }

          App.modals.close("dashboardEventoModal");
          this.loadEvents();
        } catch (error) {
          console.error("Erro ao salvar evento:", error);
          alert("Não foi possível salvar o evento.");
        }
      },

      async updateEventStatus(eventId, newStatus) {
        try {
          const responseGet = await fetch(`/api/compromissos/${eventId}/`);
          if (!responseGet.ok) throw new Error("Erro ao buscar dados do evento.");
          const eventoOriginal = await responseGet.json();

          eventoOriginal.status = newStatus;

          const response = await fetch(`/api/compromissos/${eventId}/`, {
            method: "PUT",
            headers: { 
                "Content-Type": "application/json",
                "X-CSRFToken": RCBUtils.getCookie('csrftoken')
            },
            body: JSON.stringify(eventoOriginal),
          });
          if (!response.ok) throw new Error("Erro ao atualizar status.");

          this.loadEvents();
        } catch (error) {
          console.error("Erro ao atualizar status:", error);
          alert("Não foi possível atualizar o status do evento.");
        }
      }
    };

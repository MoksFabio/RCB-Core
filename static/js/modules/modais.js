 /**
 * Módulo: Gerenciador de Modais
 * Resumo: Sistema centralizado para abrir, fechar e gerenciar janelas modais, incluindo foco e acessibilidade.
 */
window.App = window.App || {};

 window.App.modals = {
      init() {
        this.safeAddListener("settings-link", "settingsModal", () => App.theme.loadSettings());
        this.safeAddListener("closeSettingsModal", "settingsModal", null, true);

        this.safeAddListener("notifications-button", "notificationsModal", () => App.notifications.onOpen());
        this.safeAddListener("closeNotificationsModal", "notificationsModal", null, true);

        this.safeAddListener("profile-link", "profileModal", () => App.profile.load());
        this.safeAddListener("closeProfileModal", "profileModal", () => App.profile.cancelEdit(), true);

        this.safeAddListener("command-palette-toggle-btn", "command-palette-modal", () => App.commandPalette.open());
        this.safeAddListener("close-command-palette-modal", "command-palette-modal", null, true);

        // Kanban Modals
        if (document.getElementById("kanbanNewItemModal")) {
          this.safeAddListener("add-item-btn", "kanbanNewItemModal", () => document.getElementById("kanban-item-name")?.focus());
          this.safeAddListener("closeKanbanNewItemModal", "kanbanNewItemModal", null, true);
        }

        if (document.getElementById("kanbanNewColumnModal")) {
          this.safeAddListener("add-column-btn", "kanbanNewColumnModal", () => document.getElementById("kanban-column-title")?.focus());
          this.safeAddListener("closeKanbanNewColumnModal", "kanbanNewColumnModal", null, true);
        }

        // Dashboard Modals
        if (document.getElementById("dashboardEventoModal")) {
          this.safeAddListener("dashboard-novo-evento-btn", "dashboardEventoModal", () => App.dashboardAgenda.openNewEventModal());
          this.safeAddListener("closeDashboardEventoModal", "dashboardEventoModal", null, true);
        }

        window.addEventListener("click", (event) => {
          if (event.target.classList.contains("modal")) {
            this.close(event.target.id);
            if (event.target.id === "profileModal") App.profile.cancelEdit();
          }
        });

        window.addEventListener("keydown", (event) => {
          if (event.key === "Escape") {
            document.querySelectorAll(".modal").forEach(modal => {
              if (modal.style.display === "block") {
                this.close(modal.id);
                if (modal.id === "profileModal") App.profile.cancelEdit();
              }
            });
            const chatWidget = document.getElementById("hannah-widget");
            if (chatWidget && !chatWidget.classList.contains("hidden")) {
              chatWidget.classList.add("hidden");
            }
          }
        });
      },
      
      addModalListener(triggerId, modalId, callback) {
          if (callback) {
              this.safeAddListener(triggerId, modalId, callback, false);
          } else {
              this.safeAddListener(triggerId, modalId, null, true);
          }
      },

      safeAddListener(triggerId, modalId, openCallback, isCloseBtn = false) {
        const trigger = document.getElementById(triggerId);
        if (trigger) {
          console.log(`[Modals] Adding listener to ${triggerId} for ${modalId}`);
          trigger.addEventListener("click", (e) => {
            console.log(`[Modals] Clicked ${triggerId}`);
            e.preventDefault();
            if (isCloseBtn) this.close(modalId);
            else this.open(modalId, openCallback);
          });
        } else {
             console.warn(`[Modals] Trigger element not found: ${triggerId}`);
        }
      },

      open(modalId, onOpen = () => { }) {
        const modal = document.getElementById(modalId);
        if (!modal) {
            console.error(`[Modals] Modal element not found: ${modalId}`);
            return;
        }

        if (onOpen) {
          try {
            console.log(`[Modals] Executing callback for ${modalId}`);
            onOpen();
          } catch (error) {
            console.error(`[Modals] Error in open callback for ${modalId}:`, error);
          }
        }
        
        // Remove 'hidden' class explicitly as it might contain !important or conflict with inline style
        modal.classList.remove("hidden");
        modal.style.display = "block";
        console.log(`[Modals] Opened ${modalId} (display: block, hidden class removed)`);

        const focusableElements = modal.querySelectorAll(
          'button, [href], input, select, textarea, [tabindex]:not([tabindex="-1"])'
        );
        if (focusableElements.length > 0) {
          const firstFocusable = focusableElements[0];
          const lastFocusable = focusableElements[focusableElements.length - 1];

          modal.focusTrapHandler = (e) => this.handleFocusTrap(e, firstFocusable, lastFocusable);
          modal.addEventListener('keydown', modal.focusTrapHandler);

          setTimeout(() => firstFocusable.focus(), 50);
        } else {
          const closeButton = modal.querySelector(".modal-close-btn");
          if (closeButton) setTimeout(() => closeButton.focus(), 50);
        }
      },
      close(modalId) {
        const modal = document.getElementById(modalId);
        if (modal) {
          modal.style.display = "none";
          modal.classList.add("hidden"); // Restore hidden class
          if (modal.focusTrapHandler) {
            modal.removeEventListener('keydown', modal.focusTrapHandler);
          }
          console.log(`[Modals] Closed ${modalId}`);
        }
      },
      handleFocusTrap(e, firstFocusable, lastFocusable) {
        if (e.key !== 'Tab') return;

        if (e.shiftKey) {
          if (document.activeElement === firstFocusable) {
            e.preventDefault();
            lastFocusable.focus();
          }
        } else {
          if (document.activeElement === lastFocusable) {
            e.preventDefault();
            firstFocusable.focus();
          }
        }
      },

      confirm(title, message, onConfirm) {
          const modal = document.getElementById("confirmation-modal");
          const titleEl = document.getElementById("confirmation-modal-title");
          const msgEl = document.getElementById("confirmation-modal-message");
          const confirmBtn = document.getElementById("confirmation-modal-confirm");
          const cancelBtn = document.getElementById("confirmation-modal-cancel");

          if (!modal || !confirmBtn || !cancelBtn) {
              // Fallback if modal elements missing
              if (window.confirm(message)) onConfirm();
              return;
          }

          if (titleEl) titleEl.textContent = title;
          if (msgEl) msgEl.textContent = message;

          // Clone buttons to remove old listeners
          const newConfirmBtn = confirmBtn.cloneNode(true);
          const newCancelBtn = cancelBtn.cloneNode(true);
          confirmBtn.parentNode.replaceChild(newConfirmBtn, confirmBtn);
          cancelBtn.parentNode.replaceChild(newCancelBtn, cancelBtn);

          newConfirmBtn.addEventListener("click", () => {
              onConfirm();
              this.close("confirmation-modal");
          });

          newCancelBtn.addEventListener("click", () => {
              this.close("confirmation-modal");
          });

          this.open("confirmation-modal");
      }
    };

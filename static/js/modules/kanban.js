 /**
 * Módulo: Kanban (Status do Sistema)
 * Resumo: Gerencia o quadro Kanban, permitindo criar, mover e excluir colunas e cartões de status.
 */
window.App = window.App || {};

 window.App.kanban = {
      elements: {
        container: () => document.getElementById("status-container"),
        loadingSkeleton: () => document.getElementById("kanban-loading-skeleton"),
        saveStatus: () => document.getElementById("kanbanSaveStatus"),
        newItemForm: () => document.getElementById("kanban-new-item-form"),
        itemNameInput: () => document.getElementById("kanban-item-name"),
        newColumnForm: () => document.getElementById("kanban-new-column-form"),
        columnTitleInput: () => document.getElementById("kanban-column-title"),
      },
      init() {
        // Segurança: Só executa se o container existir
        if (!this.elements.container()) return;

        this.addListeners();
        this.load();
      },
      addListeners() {
        const container = this.elements.container();

        const itemForm = this.elements.newItemForm();
        if (itemForm) itemForm.addEventListener('submit', (e) => { e.preventDefault(); this.addItem(); });

        const colForm = this.elements.newColumnForm();
        if (colForm) colForm.addEventListener('submit', (e) => { e.preventDefault(); this.addColumn(); });

        container?.addEventListener("click", (e) => {
          if (e.target.closest(".delete-column-btn")) {
            this.deleteColumn(e.target.closest(".delete-column-btn").dataset.statusId);
          }
          if (e.target.closest(".delete-item-btn")) {
            this.deleteItem(e.target.closest(".delete-item-btn").dataset.funcId);
          }
        });

        container?.addEventListener("dragstart", (e) => this.handleDragStart(e));
        container?.addEventListener("dragend", (e) => this.handleDragEnd(e));
        container?.addEventListener("dragover", (e) => this.handleDragOver(e));
        container?.addEventListener("dragleave", (e) => this.handleDragLeave(e));
        container?.addEventListener("drop", (e) => this.handleDrop(e));
      },
      async saveData() {
        const saveStatus = this.elements.saveStatus();
        if (saveStatus) {
          clearTimeout(App.state.kanbanSaveTimeout);
          saveStatus.textContent = "Salvando...";
          saveStatus.className = "text-sm text-gray-500 dark:text-gray-400 transition-opacity duration-300";
          saveStatus.classList.remove("opacity-0");
        }

        try {
          await fetch(App.config.statusApiUrl, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              statuses: App.state.kanbanStatuses,
              functionalities: App.state.kanbanFunctionalities,
            }),
          });

          if (saveStatus) {
            saveStatus.textContent = "Salvo!";
            saveStatus.classList.add("text-green-600");
            App.state.kanbanSaveTimeout = setTimeout(() => saveStatus.classList.add("opacity-0"), 2000);
          }
        } catch (error) {
          console.error("Falha ao salvar os dados:", error);
          if (saveStatus) {
            saveStatus.textContent = "Erro ao salvar!";
            saveStatus.classList.remove("text-gray-500", "text-green-600");
            saveStatus.classList.add("text-red-600");
            App.state.kanbanSaveTimeout = setTimeout(() => saveStatus.classList.add("opacity-0"), 3000);
          }
        }
      },
      render() {
        const container = this.elements.container();
        if (!container) return;

        container.innerHTML = "";
        App.state.kanbanStatuses.forEach((status) => {
          const column = document.createElement("div");
          column.className = `status-column ${status.colorClass || ""}`;
          column.dataset.statusId = status.id;
          column.innerHTML = `
            <div class="status-column__header">
              <h3 class="status-column__title">
                <span class="material-icons-outlined">${status.icon}</span>
                <span>${status.title}</span>
              </h3>
              <div class="status-column__actions">
                <button class="delete-column-btn" data-status-id="${status.id}" title="Excluir esta coluna">
                  <span class="material-icons-outlined">delete_outline</span>
                </button>
              </div>
            </div>
            <ul class="status-column__list" data-status-id="${status.id}"></ul>
          `;
          container.appendChild(column);
        });

        App.state.kanbanFunctionalities.forEach((func) => {
          const list = container.querySelector(
            `.status-column__list[data-status-id="${func.statusId}"]`
          );
          if (list) {
            const listItem = document.createElement("li");
            listItem.className = "status-column__list-item glassmorphism-card";
            listItem.dataset.id = func.id;
            listItem.draggable = true;
            listItem.innerHTML = `
              <span class="flex-grow-1 pr-2">${func.name}</span>
              <button class="delete-item-btn" data-func-id="${func.id}" title="Excluir item">
                <span class="material-icons-outlined">close</span>
              </button>
            `;
            list.appendChild(listItem);
          }
        });
      },
      handleDragStart(e) {
        if (e.target.classList.contains("status-column__list-item")) {
          App.state.kanbanDraggedItem = e.target;
          setTimeout(() => e.target.classList.add("dragging"), 0);
        }
      },
      handleDragEnd(e) {
        if (App.state.kanbanDraggedItem) {
          e.target.classList.remove("dragging");
          App.state.kanbanDraggedItem = null;
        }
      },
      handleDragOver(e) {
        e.preventDefault();
        const zone = e.target.closest(".status-column__list");
        if (zone) {
          zone.classList.add("drag-over");
        }
      },
      handleDragLeave(e) {
        const zone = e.target.closest(".status-column__list");
        if (zone) {
          zone.classList.remove("drag-over");
        }
      },
      async handleDrop(e) {
        e.preventDefault();
        const zone = e.target.closest(".status-column__list");
        if (zone) {
          zone.classList.remove("drag-over");
          const { kanbanDraggedItem } = App.state;
          if (kanbanDraggedItem) {
            const newStatusId = zone.dataset.statusId;
            const itemId = kanbanDraggedItem.dataset.id;
            zone.appendChild(kanbanDraggedItem);

            const functionality = App.state.kanbanFunctionalities.find((f) => f.id === itemId);
            if (functionality) {
              functionality.statusId = newStatusId;
              await this.saveData();
            }
          }
        }
      },
      async deleteColumn(statusIdToDelete) {
        App.modals.confirm(
            "Excluir Coluna",
            "Tem certeza que deseja excluir esta coluna? Todos os itens nela serão movidos para a primeira coluna.",
            async () => {
                if (App.state.kanbanStatuses.length <= 1) {
                  alert("Não é possível excluir a última coluna.");
                  return;
                }
                const firstStatusId = App.state.kanbanStatuses[0].id;
                App.state.kanbanFunctionalities.forEach((func) => {
                  if (func.statusId === statusIdToDelete) {
                    func.statusId = firstStatusId;
                  }
                });
                App.state.kanbanStatuses = App.state.kanbanStatuses.filter(
                  (status) => status.id !== statusIdToDelete
                );
                await this.saveData();
                this.render();
            }
        );
      },
      async deleteItem(funcIdToDelete) {
          App.modals.confirm(
            "Excluir Item",
            "Tem certeza que deseja excluir este item?",
            async () => {
              App.state.kanbanFunctionalities = App.state.kanbanFunctionalities.filter(
                (func) => func.id !== funcIdToDelete
              );
              await this.saveData();
              this.render();
            }
          );
      },
      async addColumn() {
        const input = this.elements.columnTitleInput();
        const title = input.value.trim();
        if (title) {
          const newStatus = {
            id: `status-${Date.now()}`,
            title: title,
            icon: "label",
            colorClass: "",
          };
          App.state.kanbanStatuses.push(newStatus);
          await this.saveData();
          this.render();
          App.modals.close("kanbanNewColumnModal");
          this.elements.newColumnForm()?.reset();
        }
      },
      async addItem() {
        const input = this.elements.itemNameInput();
        const name = input.value.trim();
        if (name && App.state.kanbanStatuses.length > 0) {
          const newItem = {
            id: `func-${Date.now()}`,
            name: name,
            statusId: App.state.kanbanStatuses[0].id,
          };
          App.state.kanbanFunctionalities.push(newItem);
          await this.saveData();
          this.render();
          App.modals.close("kanbanNewItemModal");
          this.elements.newItemForm()?.reset();
        } else if (App.state.kanbanStatuses.length === 0) {
          alert("Crie pelo menos uma coluna de status antes de adicionar um item.");
        }
      },
      async load() {
        const skeleton = this.elements.loadingSkeleton();
        const container = this.elements.container();
        if (skeleton) skeleton.classList.remove("hidden");
        if (container) container.classList.add("hidden");

        try {
          const response = await fetch(App.config.statusApiUrl);
          if (!response.ok) throw new Error("Falha ao carregar os dados do painel.");
          const data = await response.json();
          App.state.kanbanStatuses = data.statuses || [];
          App.state.kanbanFunctionalities = data.functionalities || [];
          this.render();
        } catch (error) {
          console.error(error);
          if (container) {
            container.innerHTML = `<p class="text-red-500">Não foi possível carregar o painel de status.</p>`;
          }
        } finally {
          if (skeleton) skeleton.classList.add("hidden");
          if (container) container.classList.remove("hidden");
        }
      },
    };

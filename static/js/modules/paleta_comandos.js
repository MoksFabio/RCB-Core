 /**
 * Módulo: Paleta de Comandos (Ctrl+K)
 * Resumo: Implementa a busca rápida global e navegação por atalhos de teclado para todas as funcionalidades.
 */
window.App = window.App || {};

 window.App.commandPalette = {
      elements: {
        // On-demand
      },
      init() {
        this.gatherLinks();

        const input = document.getElementById("command-palette-input");
        const list = document.getElementById("command-palette-results");

        if (input) {
          input.addEventListener("input", () => this.filter());
          input.addEventListener("keydown", (e) => this.navigate(e));
        }
        if (list) {
          list.addEventListener("click", (e) => {
            const item = e.target.closest(".command-palette-item");
            if (item && item.href) {
              window.location.href = item.href;
              App.modals.close("command-palette-modal");
            }
          });
        }

        document.addEventListener("keydown", (e) => {
          if (e.ctrlKey && e.key === "k") {
            e.preventDefault();
            this.open();
          }
        });
      },
      gatherLinks() {
        console.log("PortalJS v3.1 Loaded - CommandPalette"); 
        App.state.commandPaletteLinks = [];
        
        // Manual entries to ensure EVERYTHING from sidebar is here and correct
        const items = [
            // Dashboard
            { name: "Dashboard", href: "/portal_de_conexoes/", icon: "dashboard", category: "Navegação" },

            // Funções
            { name: "Congestionamento", action: "modal-congestionamento", icon: "traffic", category: "Funções" },
            { name: "Cota de Óleo Diesel", href: "/cota_de_oleo_diesel/", icon: "local_gas_station", category: "Funções" },
            { name: "Ouvidorias (SAC)", action: "modal-ouvidorias", icon: "support_agent", category: "Funções" },
            { name: "Passageiro Integrado", href: "/passageiro_integrado/", icon: "group", category: "Funções" },

            // Remuneração
            { name: "Remuneração: Parâmetros", action: "modal-parametros-remuneracao", icon: "settings", category: "Remuneração" },
            { name: "Remuneração: Dados de Frota", action: "modal-dados-frota-remuneracao", icon: "directions_bus", category: "Remuneração" },
            { name: "Remuneração: Bilhetagem | Efetivo", href: "#", icon: "receipt", category: "Remuneração" },
            { name: "Remuneração: Programação", href: "#", icon: "schedule", category: "Remuneração" },
            { name: "Remuneração: Cálculo da Remuneração", href: "#", icon: "calculate", category: "Remuneração" },
            { name: "Remuneração: Relatórios", href: "#", icon: "description", category: "Remuneração" },
            { name: "Remuneração: Relatórios SINFORME", href: "#", icon: "description", category: "Remuneração" },

            // CCT
            { name: "CCT: Parâmetros", href: "#", icon: "work", category: "CCT" },
            { name: "CCT: Dados de Frota", action: "modal-dados-frota-remuneracao", icon: "directions_bus", category: "CCT" },
            { name: "CCT: Bilhetagem | Efetivo", href: "#", icon: "receipt", category: "CCT" },
            { name: "CCT: Programação", href: "#", icon: "schedule", category: "CCT" },
            { name: "CCT: Cálculo da CCT", href: "#", icon: "calculate", category: "CCT" },
            { name: "CCT: Consultas", href: "#", icon: "search", category: "CCT" },
            { name: "CCT: Relatórios SINFORME", href: "#", icon: "description", category: "CCT" },

            // Fator de Utilização
            { name: "Fator de Utilização", href: "#", icon: "pie_chart", category: "Fator de Util." },
            { name: "Encargos Sociais", href: "#", icon: "people", category: "Fator de Util." },
            { name: "Relatórios (Fator/Encargos)", href: "#", icon: "description", category: "Fator de Util." },

            // Tabelas
            { name: "Tabelas do Sistema", href: "#", icon: "table_chart", category: "Dados" },

            // Ferramentas - Importação
            { name: "Importação: Remuneração", href: "#", icon: "upload", category: "Importação" },
            { name: "Importação: CCT", href: "#", icon: "upload", category: "Importação" },
            { name: "Importação: Custos", href: "#", icon: "upload", category: "Importação" },
            { name: "Importação: Avaliação", href: "#", icon: "upload", category: "Importação" },
            { name: "Importação: Fator Util. e Encargos", href: "#", icon: "upload", category: "Importação" },
            { name: "Importação: Indicadores", href: "#", icon: "upload", category: "Importação" },
            { name: "Importação: Indicadores QualiÔnibus", href: "#", icon: "upload", category: "Importação" },

             // Ferramentas - Exportação
            { name: "Exportação: Remuneração", href: "#", icon: "download", category: "Exportação" },
            { name: "Exportação: CCT", href: "#", icon: "download", category: "Exportação" },
            { name: "Exportação: Custos", href: "#", icon: "download", category: "Exportação" },
            { name: "Exportação: Avaliação", href: "#", icon: "download", category: "Exportação" },
            { name: "Exportação: Fator Util. e Encargos", href: "#", icon: "download", category: "Exportação" },

            // Sistema
            { name: "Aprovações", action: "modal-aprovar_registros", icon: "fact_check", category: "Sistema" },
            { name: "Meu Perfil", action: "profile", icon: "person", category: "Sistema" },
            { name: "Configurações", action: "settings", icon: "settings_suggest", category: "Sistema" },
            { name: "Assistente Hannah", action: "hannah", icon: "psychology", category: "Ajuda" },
            { name: "Sair (Logout)", action: "logout", icon: "logout", category: "Sistema" }
        ];

        App.state.commandPaletteLinks = items;
      },
      open() {
        App.modals.open("command-palette-modal", () => {
          const input = document.getElementById("command-palette-input");
          if (input) {
            input.value = "";
            this.filter();
            input.focus();
          }
        });
      },
      filter() {
        const input = document.getElementById("command-palette-input");
        const list = document.getElementById("command-palette-results");
        if (!input || !list) return;

        const searchTerm = input.value.toLowerCase();
        list.innerHTML = "";

        const filtered = App.state.commandPaletteLinks.filter(link =>
          link.name.toLowerCase().includes(searchTerm)
        );

        filtered.forEach(link => {
          const li = document.createElement("li");
          // Ensure href is present even for actions
          const href = link.href || '#';
          
          li.innerHTML = `
                <a href="${href}" target="${link.target || '_self'}" class="command-palette-item" data-action="${link.action ? 'true' : 'false'}">
                    <div class="flex items-center">
                        <span class="material-icons-outlined">${link.icon}</span>
                        <span>${link.name}</span>
                    </div>
                    ${link.target === '_blank' ? '<span class="material-icons-outlined text-sm">open_in_new</span>' : ''}
                </a>
            `;
          
          if (link.action) {
            li.querySelector('a').addEventListener('click', (e) => {
              e.preventDefault();
              if (typeof link.action === 'function') {
                  link.action();
              } else if (typeof link.action === 'string') {
                  // Handle string actions (modal IDs or special keywords)
                  if (link.action === 'profile') {
                      App.modals.open("profileModal", () => App.profile.load());
                  } else if (link.action === 'settings') {
                      App.modals.open("settingsModal", () => App.theme.loadSettings());
                  } else if (link.action === 'hannah') {
                      App.chat.open();
                  } else if (link.action === 'logout') {
                      document.getElementById('logout-form').submit();
                  } else {
                       // Assume it's a feature modal ID
                      openFeatureModal(link.action);
                  }
              }
              App.modals.close("command-palette-modal");
            });
          }
          list.appendChild(li);
        });

        App.state.commandPaletteIndex = -1;
        if (list.children.length > 0) {
          App.state.commandPaletteIndex = 0;
          list.children[0].querySelector('a').classList.add("active");
        }
      },
      navigate(e) {
        const list = document.getElementById("command-palette-results");
        if (!list) return;
        const items = list.querySelectorAll("a");
        if (items.length === 0) return;

        let { commandPaletteIndex } = App.state;

        if (e.key === "ArrowDown") {
          e.preventDefault();
          items[commandPaletteIndex]?.classList.remove("active");
          commandPaletteIndex = (commandPaletteIndex + 1) % items.length;
          items[commandPaletteIndex]?.classList.add("active");
        } else if (e.key === "ArrowUp") {
          e.preventDefault();
          items[commandPaletteIndex]?.classList.remove("active");
          commandPaletteIndex = (commandPaletteIndex - 1 + items.length) % items.length;
          items[commandPaletteIndex]?.classList.add("active");
        } else if (e.key === "Enter") {
          e.preventDefault();
          items[commandPaletteIndex]?.click();
        }
        App.state.commandPaletteIndex = commandPaletteIndex;
      }
    };

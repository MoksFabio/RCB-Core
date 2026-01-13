 /**
 * Módulo: Menu Lateral (Sidebar)
 * Resumo: Controla a expansão/recolhimento da barra lateral e o comportamento responsivo dos menus.
 */
window.App = window.App || {};

 window.App.sidebar = {
      init() {
        const sidebar = document.getElementById("sidebar");
        const sidebarToggleBtn = document.getElementById("sidebar-toggle-btn");
        const mainContentWrapper = document.getElementById("main-content-wrapper");
        const sidebarLogoLink = document.getElementById("sidebar-logo-link");

        if (!sidebar || !mainContentWrapper) {
          console.warn("Elementos críticos da sidebar (sidebar ou wrapper) não encontrados.");
          return;
        }

        // Recupera estado salvo (Desktop)
        let isCollapsed = localStorage.getItem("sidebarCollapsed") === "true";

        // Aplica estado inicial (apenas se for Desktop)
        if (window.innerWidth >= 1024) {
            if (isCollapsed) {
              sidebar.classList.add("sidebar-collapsed");
              mainContentWrapper.classList.add("sidebar-collapsed");
            } else {
              sidebar.classList.remove("sidebar-collapsed");
              mainContentWrapper.classList.remove("sidebar-collapsed");
            }
        }
        
        this.updateIcon(isCollapsed);

        // Desabilita dropdowns se estiver colapsado (Desktop)
        if (sidebarToggleBtn) { // Only if button exists
             sidebar.querySelectorAll("[data-collapse-toggle]").forEach(button => {
               button.disabled = isCollapsed && window.innerWidth >= 1024;
               
               // Listener para rotação da seta (Visual only)
               // Removemos listeners antigos para evitar duplicidade se init for chamado múltiplas vezes
               button.removeEventListener('click', button._rotateHandler);
               button._rotateHandler = () => {
                   if (sidebar.classList.contains("sidebar-collapsed") && window.innerWidth >= 1024) return;
                   const arrow = button.querySelector(".material-icons-outlined:last-child");
                   if (arrow && arrow.textContent.trim() === 'expand_more') {
                       arrow.classList.toggle("rotate-180");
                   }
               };
               button.addEventListener('click', button._rotateHandler);
             });

            // Event Listener do Botão Toggle Desktop
            sidebarToggleBtn.onclick = (e) => {
              e.preventDefault();
              e.stopPropagation();
              this.toggle();
            };
        }

        // Listener do Logo
        if (sidebarLogoLink) {
          sidebarLogoLink.addEventListener("click", (e) => {
            if (window.innerWidth < 1024) {
               // No mobile, logo pode fechar se já aberto, ou ir para home
               // Se estiver aberto (mobile-open), fecha.
               if (sidebar.classList.contains("mobile-open")) {
                   e.preventDefault();
                   this.closeMobile();
               }
            } else {
                // Desktop Logic
                const isDashboardActive = sidebar.querySelector("a[data-page='dashboard']")?.classList.contains("active");
                if (sidebar.classList.contains("sidebar-collapsed")) {
                    e.preventDefault();
                    this.toggle();
                } else if (isDashboardActive) {
                    e.preventDefault();
                }
            }
          });
        }

        // Listener para links de navegação (Mobile: fecha ao clicar / Desktop: previne reload se ativo)
        const nav = sidebar.querySelector("nav");
        if (nav) {
          nav.addEventListener("click", (e) => {
            const link = e.target.closest("a.menu-item");
            if (link) {
                if (window.innerWidth < 1024) {
                    // Mobile: fecha sidebar ao navegar (UX padrão)
                    // Aguarda um pouco para animação de clique se quiser, mas instantâneo é melhor
                    this.closeMobile();
                }
                
                if (link.classList.contains("active")) {
                  e.preventDefault();
                }
            }
          });
        }
        
        // Listener de Resize para limpar estados conflitantes
        window.addEventListener('resize', () => {
            if (window.innerWidth >= 1024) {
                this.closeMobile(); 
                // Restaura estado desktop visualmente
                // (O estado JS já está correto em localStorage)
            }
        });
      },

      toggle() {
        if (window.innerWidth < 1024) {
            this.toggleMobile();
        } else {
            this.toggleDesktop();
        }
      },

      // Logic removed as Sidebar is no longer used on mobile.
      toggleMobile() {
         console.warn("Sidebar Disabled on Mobile.");
      },

      closeMobile() {
         // No-op
      },

      toggleDesktop() {
        const sidebar = document.getElementById("sidebar");
        const mainContentWrapper = document.getElementById("main-content-wrapper");

        if (!sidebar || !mainContentWrapper) return;

        const isCollapsed = sidebar.classList.toggle("sidebar-collapsed");
        mainContentWrapper.classList.toggle("sidebar-collapsed", isCollapsed);

        localStorage.setItem("sidebarCollapsed", isCollapsed);

        this.updateIcon(isCollapsed);

        sidebar.querySelectorAll("[data-collapse-toggle]").forEach(button => {
          button.disabled = isCollapsed;
          if (isCollapsed) {
            const dropdownId = button.getAttribute("data-collapse-toggle");
            const dropdown = document.getElementById(dropdownId);
            if (dropdown && !dropdown.classList.contains("hidden")) {
              dropdown.classList.add("hidden");
            }
          }
        });
      },

      toggleMobileMenu() {
        const sheet = document.getElementById('mobile-menu-sheet');
        if (sheet) sheet.classList.toggle('hidden');
      },

      toggleAccordion(button) {
          const content = button.nextElementSibling;
          if (content) {
              const isOpen = content.classList.contains('open');
              
              if (isOpen) {
                  content.classList.remove('open');
                  button.setAttribute('aria-expanded', 'false');
              } else {
                  content.classList.add('open');
                  button.setAttribute('aria-expanded', 'true');
              }
          }
      },

      updateIcon(isCollapsed) {
        const btn = document.getElementById("sidebar-toggle-btn");
        if (!btn) return;

        const icon = btn.querySelector(".material-icons-outlined");
        if (icon) {
          // No mobile, sempre mostra 'menu' se fechado, 'menu_open' (ou close) se aberto
          // Mas como o botão está DENTRO da sidebar no design atual...
          // Precisamos ver se o design tem botão fora.
          // O base.html mostra o botão DENTRO do sidebar footer.
          // ISSO É UM PROBLEMA PARA MOBILE SE A SIDEBAR ESTIVER ESCONDIDA.
          // Precisamos mover/duplicar o botão para o Header no Mobile.
          
          icon.textContent = isCollapsed ? "menu" : "menu_open";
        }
      }
    };

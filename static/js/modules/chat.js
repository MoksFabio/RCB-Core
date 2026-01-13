 /**
 * Módulo: Chat (Hannah)
 * Resumo: Gerencia o widget de chat da assistente virtual, incluindo histórico, envio de mensagens e interface arrastável.
 */
window.App = window.App || {};

 window.App.chat = {
      elements: {
        // On-demand
      },
      init() {
        console.log("Hannah: Inicializando...");
        const btn = document.getElementById("hannah-assistant-button");
        const close = document.getElementById("close-hannah-widget-btn");
        const send = document.getElementById("hannah-send-button");
        const input = document.getElementById("hannah-chat-input");
        const header = document.getElementById("hannah-widget-header");
        
        // NOVO: Referência ao botão de limpar
        const clearBtn = document.getElementById("clear-hannah-history-btn");

        // 1. Carrega histórico de mensagens
        this.loadHistory();

        // 2. Verifica se ela deveria estar aberta
        this.restoreState();

        if (btn) btn.addEventListener("click", (e) => { e.preventDefault(); this.open(); });
        if (close) close.addEventListener("click", () => this.close());
        if (send) send.addEventListener("click", () => this.sendMessage());
        if (input) input.addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); this.sendMessage(); } });

        // NOVO: Evento de clique para limpar histórico
        if (clearBtn) clearBtn.addEventListener("click", (e) => { 
            e.preventDefault(); 
            this.clearHistory(); 
        });

        if (header) {
          header.addEventListener("mousedown", (e) => this.dragStart(e));
          document.addEventListener("mousemove", (e) => this.dragMove(e));
          document.addEventListener("mouseup", () => this.dragEnd());
        }
      },

      loadHistory() {
        const chatBox = document.getElementById("hannah-chat-box");
        const widget = document.getElementById("hannah-widget");
        const headerUsername = document.getElementById("header-username");
        
        if (!chatBox) return;

        // Determine current user
        let currentUser = "Usuário";
        if (headerUsername && headerUsername.textContent.trim() !== 'Usuário') {
            currentUser = headerUsername.textContent.trim();
        } else if (widget && widget.dataset.username) {
            currentUser = widget.dataset.username;
        }

        // Check stored user
        const storedUser = localStorage.getItem("hannah_chat_user");

        if (storedUser !== currentUser) {
            // User changed, clear history
            console.log(`Hannah: User changed from ${storedUser} to ${currentUser}. Clearing history.`);
            localStorage.removeItem("hannah_chat_history");
            localStorage.setItem("hannah_chat_user", currentUser);
        }

        const rawHistory = localStorage.getItem("hannah_chat_history");
        const history = JSON.parse(rawHistory || "[]");

        if (history.length > 0) {
          chatBox.innerHTML = "";
          history.forEach(msg => {
            this.addMessageToChat(msg.message, msg.sender, false);
          });
          this.scrollToBottom();
        } else {
            // If no history (new session), show welcome message
             // But open() handles the welcome message if empty. 
             // We can just leave it empty here.
        }
      },

      // NOVA FUNÇÃO: Restaura o estado aberto/fechado
      restoreState() {
        const wasOpen = localStorage.getItem("hannah_widget_is_open") === "true";
        const widget = document.getElementById("hannah-widget");

        if (wasOpen && widget) {
          console.log("Hannah: Restaurando estado ABERTO.");
          widget.classList.remove("hidden");
          this.scrollToBottom();
          // Opcional: Dar foco no input automaticamente ao trocar de página
          // setTimeout(() => document.getElementById("hannah-chat-input")?.focus(), 100);
        }
      },

      clearHistory() {
          App.modals.confirm(
              "Limpar Histórico",
              "Tem certeza que deseja apagar todo o histórico da conversa com a Hannah?",
              () => {
                // 1. Limpa o LocalStorage
                localStorage.removeItem("hannah_chat_history");

                // 2. Limpa o visual
                const chatBox = document.getElementById("hannah-chat-box");
                if (chatBox) chatBox.innerHTML = "";

                // 3. Adiciona uma mensagem de reinício
                const widget = document.getElementById("hannah-widget");
                const headerUsername = document.getElementById("header-username");
                
                let nomeUsuario = "Usuário";
                if (headerUsername && headerUsername.textContent && headerUsername.textContent !== 'Usuário') {
                    nomeUsuario = headerUsername.textContent;
                } else if (widget && widget.dataset.username) {
                    nomeUsuario = widget.dataset.username;
                }

                this.addMessageToChat(`Olá, ${nomeUsuario}. Como posso ajudar?`, "assistant", true);
              }
          );
      },

      open() {
        const widget = document.getElementById("hannah-widget");
        const chatBox = document.getElementById("hannah-chat-box");
        const headerUsername = document.getElementById("header-username");

        if (!widget) return;

        widget.classList.remove("hidden");

        // SALVA O ESTADO COMO ABERTO
        localStorage.setItem("hannah_widget_is_open", "true");

        if (chatBox && chatBox.children.length === 0) {
          let nomeUsuario = "Usuário";
          if (headerUsername && headerUsername.textContent && headerUsername.textContent !== 'Usuário') {
            nomeUsuario = headerUsername.textContent;
          } else if (widget.dataset.username && widget.dataset.username !== 'Usuário') {
            nomeUsuario = widget.dataset.username;
          }
          const msg = `Olá, ${nomeUsuario}! Meu nome é Hannah e estou aqui para te ajudar.`;
          this.addMessageToChat(msg, "assistant", true);
        }

        setTimeout(() => document.getElementById("hannah-chat-input")?.focus(), 50);
        this.scrollToBottom();
      },

      close() {
        // SALVA O ESTADO COMO FECHADO
        localStorage.setItem("hannah_widget_is_open", "false");
        document.getElementById("hannah-widget")?.classList.add("hidden");
      },

      scrollToBottom() {
        const chatBox = document.getElementById("hannah-chat-box");
        if (chatBox) chatBox.scrollTop = chatBox.scrollHeight;
      },

      handleAction(action) {
          console.log("Hannah Action:", action);
          // Small delay to let the user read the message first
          setTimeout(() => {
              // --- 1. MODAL OPENING ACTIONS ---
              if (action.startsWith('OPEN_MODAL_')) {
                  const modalType = action.replace('OPEN_MODAL_', '');
                  let modalId = '';
                  
                  switch(modalType) {
                      case 'CONGESTIONAMENTO': modalId = 'modal-congestionamento'; break;
                      case 'OUVIDORIAS': modalId = 'modal-ouvidorias'; break;
                      case 'COTA': modalId = 'modal-cota_oleo_diesel'; break;
                      case 'APROVAR': modalId = 'modal-aprovar_registros'; break;
                      case 'PARAMETROS': modalId = 'modal-parametros-remuneracao'; break;
                      case 'FROTA': modalId = 'dados_frota_remuneracao_modal'; break;
                      default: console.warn("Tipo de modal desconhecido:", modalType); return;
                  }
                  
                  if (typeof window.openFeatureModal === 'function') {
                      window.openFeatureModal(modalId);
                  } else {
                      console.error("Função openFeatureModal não disponível globalmente.");
                  }

              // --- 2. SCROLL ACTIONS ---
              } else if (action.startsWith('SCROLL_TO_')) {
                  const scrollType = action.replace('SCROLL_TO_', '');
                  let targetId = '';

                  switch(scrollType) {
                      case 'AGENDA': targetId = 'proximos-eventos'; break;
                      case 'NOTAS': targetId = 'anotacoes-rapidas'; break;
                      default: return;
                  }

                  const element = document.getElementById(targetId);
                  if (element) {
                      element.scrollIntoView({ behavior: 'smooth', block: 'center' });
                      element.classList.add('animate__animated', 'animate__pulse');
                      setTimeout(() => element.classList.remove('animate__animated', 'animate__pulse'), 1000);
                      
                      // Auto-open widget content if collapsed
                      if (typeof window.toggleWidget === 'function') {
                          const content = element.querySelector('.widget-content');
                          if (content && content.classList.contains('hidden')) {
                              // Simulate click on header to open
                              const header = element.querySelector('.card-title')?.parentElement;
                              if (header) header.click();
                          }
                      }
                  }

              // --- 3. URL NAVIGATION (Fallback/Legacy) ---
              } else {
                  switch(action) {
                      case 'NAVIGATE_PASSAGEIRO':
                          window.location.href = "/passageiro_integrado/";
                          break;
                      case 'NAVIGATE_SABE':
                          alert("Funcionalidade SABE: URL não configurada.");
                          break;
                      case 'NAVIGATE_PROFILE':
                          if(window.App && window.App.profile && window.App.profile.openProfileModal) {
                               window.App.profile.openProfileModal();
                          }
                          break;
                      default:
                          console.warn("Ação desconhecida ou não mapeada:", action);
                  }
              }
          }, 1500);
      },

      addMessageToChat(message, sender, save = true) {
        const chatBox = document.getElementById("hannah-chat-box");
        if (!chatBox) return;

        const messageElement = document.createElement("div");
        messageElement.classList.add("chat-message", `${sender}-message`);

        if (sender === "assistant" && message === "typing") {
          messageElement.innerHTML = `<div class="typing-indicator"><span></span><span></span><span></span></div>`;
          messageElement.id = "typing-indicator";
          save = false;
        } else {
          messageElement.textContent = message;
        }

        chatBox.appendChild(messageElement);
        this.scrollToBottom();

        if (save) {
          try {
            const history = JSON.parse(localStorage.getItem("hannah_chat_history") || "[]");
            history.push({ message, sender, timestamp: new Date().toISOString() });
            if (history.length > 50) history.shift();
            localStorage.setItem("hannah_chat_history", JSON.stringify(history));
          } catch (e) { console.error(e); }
        }

        return messageElement;
      },

      async sendMessage() {
        const input = document.getElementById("hannah-chat-input");
        const sendBtn = document.getElementById("hannah-send-button");
        if (!input || !input.value.trim()) return;

        const message = input.value.trim();
        this.addMessageToChat(message, "user", true);

        input.value = "";
        input.disabled = true;
        if (sendBtn) sendBtn.disabled = true;

        const typingIndicator = this.addMessageToChat("typing", "assistant");

        try {
          const response = await fetch(App.config.hannahChatApiUrl, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ message: message }),
          });

          if (typingIndicator) typingIndicator.remove();

          if (!response.ok) {
            const errorData = await response.json();
            throw new Error(errorData.reply || "Erro na comunicação com o servidor.");
          }

          const data = await response.json();
          
          this.addMessageToChat(data.reply, "assistant", true);

          // HANDLE ACTIONS IF ANY
          if (data.action) {
              this.handleAction(data.action);
          }
        } catch (error) {
          if (typingIndicator) typingIndicator.remove();
          this.addMessageToChat(`Erro: ${error.message}`, "assistant", true);
        } finally {
          input.disabled = false;
          if (sendBtn) sendBtn.disabled = false;
          input.focus();
        }
      },

      dragStart(e) {
        const widget = document.getElementById("hannah-widget");
        const header = document.getElementById("hannah-widget-header");

        if (e.target === header || e.target.closest(".hannah-widget__title")) {
          App.state.isHannahDragging = true;
          const rect = widget.getBoundingClientRect();
          App.state.hannahDragOffset = {
            x: e.clientX - rect.left,
            y: e.clientY - rect.top,
          };
          widget.style.cursor = "grabbing";
          App.elements.body.style.userSelect = "none";
        }
      },
      dragMove(e) {
        if (!App.state.isHannahDragging) return;
        const widget = document.getElementById("hannah-widget");
        e.preventDefault();

        let newX = e.clientX - App.state.hannahDragOffset.x;
        let newY = e.clientY - App.state.hannahDragOffset.y;

        const maxX = window.innerWidth - widget.offsetWidth;
        const maxY = window.innerHeight - widget.offsetHeight;

        newX = Math.max(0, Math.min(newX, maxX));
        newY = Math.max(0, Math.min(newY, maxY));

        widget.style.transform = `translate3d(${newX}px, ${newY}px, 0)`;
        widget.style.bottom = "auto";
        widget.style.right = "auto";
        widget.style.top = "0px";
        widget.style.left = "0px";
      },
      dragEnd() {
        if (App.state.isHannahDragging) {
          App.state.isHannahDragging = false;
          const widget = document.getElementById("hannah-widget");
          if (widget) widget.style.cursor = "default";
          App.elements.body.style.userSelect = "auto";
        }
      },
    };

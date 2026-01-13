 /**
 * Módulo: Notificações
 * Resumo: Monitora novas notificações em tempo real (polling), atualiza o contador e destaca itens não lidos.
 */
window.App = window.App || {};

 window.App.notifications = {
      pollingInterval: null,

      init() {
        const clearBtn = document.getElementById("clear-notifications-button");
        if (clearBtn) clearBtn.addEventListener("click", () => this.clearAll());
        
        // Initial Check for Alerts
        this.checkAlert();
        
        // Polling for real-time alerts
        this.startPolling();

        this.load();
        
        this.initManagement();
      },
      onOpen() {
        console.log("[Notifications] onOpen called");
        try {
            this.markAsRead();
            this.load();
        } catch (e) {
            console.error("[Notifications] Error in onOpen:", e);
        }
      },
      load() {
        this.render();
      },
      startPolling() {
        if (this.pollingInterval) clearInterval(this.pollingInterval);
        
        // CHECK IMMEDIATELY on load/start
        this.checkAlert();
        
        this.pollingInterval = setInterval(() => {
            this.checkAlert();
        }, 300000); // Check every 5 minutes
      },
      async checkAlert() {
          const btn = document.getElementById("notifications-button");
          // Fetch latest ID from API
          try {
            const response = await fetch('/api/latest-notification/');
            const data = await response.json();
            const latestId = data.latest_id;
            
            // Allow override from template if API fails or for initial render
            const templateIdInput = document.getElementById("latest-notification-id");
            const templateId = templateIdInput ? parseInt(templateIdInput.value) : 0;
            
            const effectiveLatestId = Math.max(latestId, templateId);
            const userId = window.currentUserId || '';
            const storageKey = 'lastSeenNotificationId_' + userId;
            const lastSeenId = parseInt(localStorage.getItem(storageKey) || "0");
            
            if (btn) {
                const icon = btn.querySelector('.material-icons-outlined');
                // FORCE RED COLOR direct on element style with !important priority
                if (effectiveLatestId > lastSeenId) {
                    btn.classList.add("notification-alert");
                    if(icon) {
                        icon.style.setProperty('color', '#ef4444', 'important');
                        icon.parentElement.style.setProperty('color', '#ef4444', 'important');
                    }
                } else {
                    btn.classList.remove("notification-alert");
                    if(icon) {
                        icon.style.removeProperty('color');
                        icon.parentElement.style.removeProperty('color');
                    }
                }
            }
            
            // Only update highlights if we are NOT currently looking at the list (to avoid it jumping/disappearing while reading)
            // Or act smart: Highlight anything > lastSeenId
            this.highlightUnreadItems(lastSeenId);

          } catch (error) {
            console.error("Error fetching latest notification:", error);
          }
      },
      highlightUnreadItems(lastSeenId) {
          document.querySelectorAll('.notification-item').forEach(item => {
              const id = parseInt(item.dataset.id || "0");
              const badge = item.querySelector('.new-badge');
              if (id > lastSeenId) {
                  // It's new/unread
                  if(badge) badge.classList.remove('hidden');
                  item.classList.add('bg-orange-50/50', 'dark:bg-orange-900/10', 'border-orange-200', 'dark:border-orange-800');
                  item.classList.remove('border-transparent');
              } else {
                  // It's read
                  if(badge) badge.classList.add('hidden');
                  item.classList.remove('bg-orange-50/50', 'dark:bg-orange-900/10', 'border-orange-200', 'dark:border-orange-800');
                  item.classList.add('border-transparent');
              }
          });
      },
      markAsRead() {
          // Update the localStorage but DO NOT wipe the visuals immediately.
          // This allows the "New" badges to stay visible while the user is reading.
          // They will disappear on the NEXT polling cycle or next page load, which is acceptable.
          
          const latestIdInput = document.getElementById("latest-notification-id");
          let effectiveLatestId = latestIdInput ? parseInt(latestIdInput.value) : 0;
          
          fetch('/api/latest-notification/').then(r => r.json()).then(d => {
              const id = d.latest_id;
              if (id > 0) {
                  // Store it
                  const userId = window.currentUserId || '';
                  localStorage.setItem("lastSeenNotificationId_" + userId, id);
                  
                  // We do NOT call highlightUnreadItems(id) here. 
                  // Let the badges stay until next refresh so user sees what is new.
              }
          });
          
          // Clear the global bell alert immediately though, as the modal is open
          const btn = document.getElementById("notifications-button");
          if (btn) {
              btn.classList.remove("notification-alert");
              const icon = btn.querySelector('.material-icons-outlined');
              if(icon) {
                  icon.style.removeProperty('color');
                  icon.parentElement.style.removeProperty('color');
              }
          }
      },
      updateBadge() {
        const badge = document.getElementById("notification-count-badge");
        if (!badge) return;
        badge.style.display = "none";
        badge.textContent = "0";
      },
      render() {
        const list = document.getElementById("notification-list");
        const noNotif = document.getElementById("no-notifications");
        if (!list || !noNotif) return;
        // In a real app we would fetch the list of notifications here
        // For now leveraging server-side rendered list
        this.updateBadge();
      },

      
      // Management Functions (Moved from base.html)
      initManagement() {
           this.quill = null;
           
           // Initialize Quill if Editor exists
           if (document.getElementById('notif-editor')) {
                // Ensure Quill is loaded first (check window.Quill)
                if (typeof Quill !== 'undefined') {
                    this.quill = new Quill('#notif-editor', {
                        theme: 'snow',
                        placeholder: 'Escreva sua mensagem...',
                        modules: {
                            toolbar: [
                                ['bold', 'underline'],
                                [{ 'color': [] }, { 'background': [] }],
                                ['clean']
                            ]
                        }
                    });
                } else {
                    console.warn("Quill library not loaded.");
                }
           }
           
           // Check for reopen flag
            if (sessionStorage.getItem('reopen_notifications_modal') === 'true') {
                if (window.App.modals && window.App.modals.open) {
                     // Assuming we have a generic open method or use the direct element if simple
                     const modal = document.getElementById('notificationsModal');
                     if(modal) modal.style.display = 'block';
                }
                sessionStorage.removeItem('reopen_notifications_modal');
            }
      },

      toggleForm() {
        const form = document.getElementById('new-notification-form');
        if (!form) return;
        
        form.classList.toggle('hidden');
        if (!form.classList.contains('hidden')) {
            // Reset form for new entry if opening
            if (document.getElementById('notif-action').value !== 'edit') {
                this.resetForm();
            }
        } else {
            this.resetForm();
        }
      },

      resetForm() {
        if(!document.getElementById('notif-id')) return;
        document.getElementById('notif-id').value = '';
        document.getElementById('notif-action').value = 'add';
        document.getElementById('notif-title').value = '';
        if (this.quill) this.quill.setText('');
        const titleEl = document.getElementById('notif-form-title');
        if (titleEl) titleEl.innerText = 'Nova Notificação';
      },
      
      startEdit(id, title, message) {
        this.toggleForm(); // Open it
        // Ensure it's visible (toggle might have closed it if it was open but hidden logic handles it)
        document.getElementById('new-notification-form').classList.remove('hidden');
        
        document.getElementById('notif-id').value = id;
        document.getElementById('notif-action').value = 'edit';
        document.getElementById('notif-title').value = title;
        if (this.quill) this.quill.clipboard.dangerouslyPasteHTML(0, message);
        const titleEl = document.getElementById('notif-form-title');
        if (titleEl) titleEl.innerText = 'Editar Notificação';
      },
      
      async send() {
        const id = document.getElementById('notif-id').value;
        const action = document.getElementById('notif-action').value;
        const title = document.getElementById('notif-title').value;
        const message = this.quill ? this.quill.root.innerHTML : '';
        const textContent = this.quill ? this.quill.getText().trim() : '';

        if (!title || (!message || textContent.length === 0)) {
            alert('Por favor, preencha título e mensagem.');
            return;
        }
        
        // CSRF Token Helper
        const getCookie = (name) => {
            let cookieValue = null;
            if (document.cookie && document.cookie !== '') {
                const cookies = document.cookie.split(';');
                for (let i = 0; i < cookies.length; i++) {
                    const cookie = cookies[i].trim();
                    if (cookie.substring(0, name.length + 1) === (name + '=')) {
                        cookieValue = decodeURIComponent(cookie.substring(name.length + 1));
                        break;
                    }
                }
            }
            return cookieValue;
        };
        const csrftoken = getCookie('csrftoken');

        try {
            const response = await fetch('/manage_notification/', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'X-CSRFToken': csrftoken
                },
                body: JSON.stringify({ action, id, title, message })
            });

            const data = await response.json();
            
            if (data.status === 'success') {
                this.toggleForm();
                
                // Real-time DOM Update instead of reload
                const list = document.getElementById('notification-list');
                const noNotif = document.getElementById('no-notifications');
                if (noNotif) noNotif.remove();

                if (action === 'add') {
                    // Create new item HTML
                    const today = new Date();
                    const dateStr = today.toLocaleDateString('pt-BR') + ' às ' + today.toLocaleTimeString('pt-BR', {hour: '2-digit', minute:'2-digit'});
                    const newId = data.new_id || Date.now(); 
                    
                    // Note: Escaping quotes for onclick arguments is crucial. 
                    const safeTitle = title.replace(/'/g, "\\'");
                    const safeMessage = message.replace(/'/g, "\\'").replace(/\n/g, "");
                    
                    const newHtml = `
                    <div class="group p-6 hover:bg-white/80 dark:hover:bg-gray-800/80 transition-all duration-200 relative border-l-4 border-transparent hover:border-orange-500 notification-item" data-id="${newId}">
                      <div class="flex justify-between items-start mb-2">
                          <div class="flex items-center gap-3">
                              <div class="notification-icon w-10 h-10 rounded-full bg-orange-100 dark:bg-orange-900/30 flex items-center justify-center text-orange-600 dark:text-orange-400">
                                  <span class="material-icons-outlined">campaign</span>
                              </div>
                              <div>
                                  <h4 class="font-bold text-base text-gray-800 dark:text-gray-100 leading-tight flex items-center gap-2 notification-title">
                                    ${title}
                                    <span class="new-badge hidden bg-red-500 text-white text-[10px] px-2 py-0.5 rounded-full font-bold uppercase tracking-wide animate-pulse">Novo</span>
                                  </h4>
                                  <span class="text-xs text-gray-400 flex items-center gap-1 mt-0.5">
                                      <span class="material-icons-outlined text-[10px]">schedule</span>
                                      ${dateStr}
                                  </span>
                              </div>
                          </div>
                          
                          <div class="flex flex-col items-end gap-1">
                              <div class="opacity-0 group-hover:opacity-100 transition-opacity flex gap-1">
                                  <button onclick="App.notifications.startEdit('${newId}', '${safeTitle}', '${safeMessage}')" class="p-1 text-blue-500 hover:bg-blue-50 rounded" title="Editar">
                                      <span class="material-icons-outlined text-lg">edit</span>
                                  </button>
                                  <button onclick="App.notifications.delete('${newId}')" class="p-1 text-red-500 hover:bg-red-50 rounded" title="Excluir">
                                      <span class="material-icons-outlined text-lg">delete</span>
                                  </button>
                              </div>
                          </div>
                      </div>
                      <div class="text-sm text-gray-600 dark:text-gray-300 prose prose-sm dark:prose-invert max-w-none pl-[3.25rem] notification-message">
                          ${message}
                      </div>
                    </div>`;
                    
                    list.insertAdjacentHTML('afterbegin', newHtml);
                    
                    // Preventing Admin Self-Notification
                    if (newId) {
                        const userId = window.currentUserId || '';
                        localStorage.setItem('lastSeenNotificationId_' + userId, newId);
                        this.checkAlert();
                    }
                    
                } else if (action === 'edit') {
                    // Update existing item
                    const item = document.querySelector(`.notification-item[data-id="${id}"]`);
                    if (item) {
                        item.querySelector('.notification-title').childNodes[0].textContent = title + ' ';
                        item.querySelector('.notification-message').innerHTML = message;
                        
                        // Update edit button parameters
                        const editBtn = item.querySelector('button[title="Editar"]');
                        if (editBtn) {
                             const safeTitle = title.replace(/'/g, "\\'");
                             const safeMessage = message.replace(/'/g, "\\'").replace(/\n/g, "");
                             editBtn.setAttribute('onclick', `App.notifications.startEdit('${id}', '${safeTitle}', '${safeMessage}')`);
                        }
                    }
                }

            } else {
                alert('Erro: ' + data.message);
            }
        } catch (error) {
            console.error('Error:', error);
            alert('Erro ao enviar notificação');
        }
      },
      
      delete(id) {
          if (window.App.modals && window.App.modals.confirm) {
            window.App.modals.confirm(
                "Excluir Notificação",
                "Tem certeza que deseja excluir esta notificação?",
                () => this.performDelete(id)
            );
        } else {
             if (confirm('Tem certeza que deseja excluir esta notificação?')) {
                 this.performDelete(id);
             }
        }
      },
      
      async performDelete(id) {
          // CSRF Token Helper (duplicated here ideally moved to utils)
          const getCookie = (name) => {
            let cookieValue = null;
            if (document.cookie && document.cookie !== '') {
                const cookies = document.cookie.split(';');
                for (let i = 0; i < cookies.length; i++) {
                    const cookie = cookies[i].trim();
                     if (cookie.substring(0, name.length + 1) === (name + '=')) {
                        cookieValue = decodeURIComponent(cookie.substring(name.length + 1));
                        break;
                    }
                }
            }
            return cookieValue;
        };
        const csrftoken = getCookie('csrftoken');

        try {
            const response = await fetch('/manage_notification/', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'X-CSRFToken': csrftoken
                },
                body: JSON.stringify({ action: 'delete', id })
            });

            const data = await response.json();
            
            if (data.status === 'success') {
                // Real-time DOM removal
                const item = document.querySelector(`.notification-item[data-id="${id}"]`);
                if (item) {
                    item.remove();
                    
                    // Check if list empty
                    const list = document.getElementById('notification-list');
                    if (list.children.length === 0) {
                        list.innerHTML = `
                        <div id="no-notifications" class="flex flex-col items-center justify-center py-20 text-center">
                           <div class="w-24 h-24 bg-gray-100 dark:bg-gray-800 rounded-full flex items-center justify-center mb-4">
                               <span class="material-icons-outlined text-5xl text-gray-300 dark:text-gray-600">notifications_off</span>
                           </div>
                           <h3 class="text-lg font-medium text-gray-900 dark:text-gray-100">Tudo limpo por aqui!</h3>
                           <p class="text-gray-500 dark:text-gray-400 max-w-xs mt-2">Nenhuma notificação global foi encontrada no momento.</p>
                        </div>`;
                    }
                }
            } else {
                alert('Erro: ' + data.message);
            }
        } catch (error) {
            console.error('Error:', error);
            alert('Erro ao excluir notificação');
        }
      },
      
      clearAll() {
           if (window.App.modals && window.App.modals.confirm) {
                window.App.modals.confirm(
                    "Limpar tudo",
                    "Tem certeza que deseja apagar todas as notificações?",
                    () => this.performClearAll()
                );
            } else {
                 if(confirm('Tem certeza que deseja apagar todas as notificações?')) {
                     this.performClearAll();
                 }
            }
      },
      
      async performClearAll() {
           // CSRF Token Helper
            const getCookie = (name) => {
            let cookieValue = null;
            if (document.cookie && document.cookie !== '') {
                const cookies = document.cookie.split(';');
                for (let i = 0; i < cookies.length; i++) {
                    const cookie = cookies[i].trim();
                     if (cookie.substring(0, name.length + 1) === (name + '=')) {
                        cookieValue = decodeURIComponent(cookie.substring(name.length + 1));
                        break;
                    }
                }
            }
            return cookieValue;
        };
        const csrftoken = getCookie('csrftoken');
          
        try {
            const response = await fetch('/manage_notification/', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'X-CSRFToken': csrftoken
                },
                body: JSON.stringify({ action: 'clear_all' })
            });
            
            const data = await response.json();
            if (data.status === 'success') {
                // Clear list
                const list = document.getElementById('notification-list');
                list.innerHTML = `
                <div id="no-notifications" class="flex flex-col items-center justify-center py-16 text-center px-6">
                   <div class="w-16 h-16 bg-gray-50 dark:bg-gray-800 rounded-full flex items-center justify-center mb-3">
                       <span class="material-icons-outlined text-3xl text-gray-300 dark:text-gray-600">notifications_none</span>
                   </div>
                   <h3 class="text-sm font-medium text-gray-900 dark:text-gray-100">Nada por aqui</h3>
                   <p class="text-xs text-gray-400 mt-1">Você não tem novas notificações.</p>
                </div>`;
                
                // Remove bell badge
                const badge = document.getElementById('notification-badge');
                if (badge) badge.remove();
                
            } else {
                alert('Erro ao limpar notificações: ' + data.message);
            }
        } catch(e) {
            console.error(e);
            alert('Erro de conexão.');
        }
      },

      stopPolling() {
          if (this.pollingInterval) clearInterval(this.pollingInterval);
      }
    };


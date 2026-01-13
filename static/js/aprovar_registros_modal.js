// Aprovar Registros Module Logic with AJAX and Confirmation

document.addEventListener("DOMContentLoaded", () => {
    
    // Check if AprovarRegistros is already defined to avoid redeclaration issues in SPA
    if (window.AprovarRegistrosInitialized) return;
    window.AprovarRegistrosInitialized = true;

    const AprovarRegistros = {
        init() {
            this.addListeners();
            this.handleExistingFlashMessages();
        },

        handleExistingFlashMessages() {
            const messages = document.querySelectorAll('#flash-messages-container-aprov .flash-message');
            messages.forEach(msg => {
                setTimeout(() => {
                    msg.classList.remove('animate__fadeIn');
                    msg.classList.add('animate__fadeOut');
                    setTimeout(() => msg.remove(), 500);
                }, 5000);
            });
        },

        addListeners() {
            // Edit User Button Click (Delegation)
            document.body.addEventListener('click', (e) => {
                const editBtn = e.target.closest('.edit-user-btn-aprov');
                if (editBtn) {
                    const userId = editBtn.dataset.userId;
                    const username = editBtn.dataset.userUsername;
                    this.openEditUserModal(userId, username);
                }
            });

            // Form Submissions (Delegation)
            document.body.addEventListener('submit', async (e) => {
                const form = e.target;
                
                // Check if the form belongs to Aprovar Registros modal actions
                if (form.closest('#modal-aprovar_registros') || form.id === 'editUserForm-aprov') {
                    // Determine action type
                    const isReject = form.querySelector('button[title="Rejeitar"]');
                    const isRemove = form.querySelector('button[title="Remover Usuário"]');
                    const isApprove = form.querySelector('button[title="Aprovar"]');

                    let confirmTitle = '';
                    let confirmMessage = '';

                    if (isReject) {
                        confirmTitle = 'Rejeitar Registro';
                        confirmMessage = 'Tem certeza que deseja rejeitar este registro? Esta operação não pode ser desfeita.';
                    } else if (isRemove) {
                        confirmTitle = 'Remover Usuário';
                        confirmMessage = 'Tem certeza que deseja remover este usuário? Esta operação não pode ser desfeita.';
                    } else if (isApprove) {
                        confirmTitle = 'Aprovar Usuário';
                        confirmMessage = 'Tem certeza que deseja aprovar este usuário?';
                    }

                    if (confirmTitle) {
                        e.preventDefault();
                        if (window.App && window.App.modals && window.App.modals.confirm) {
                            window.App.modals.confirm(confirmTitle, confirmMessage, () => {
                                this.handleFormSubmit(form);
                            });
                        } else if (confirm(confirmMessage)) {
                            // Fallback if App.modals not available
                             await this.handleFormSubmit(form);
                        }
                        return;
                    }

                    e.preventDefault();
                    await this.handleFormSubmit(form);
                }
            });
        },

        openEditUserModal(userId, username) {
            const modal = document.getElementById('editUserModal-aprov');
            const form = document.getElementById('editUserForm-aprov');
            const usernameInput = document.getElementById('newUsername-aprov');

            if (!modal || !form || !usernameInput) return;

            form.action = `/editar_usuario/${userId}/`;
            usernameInput.value = username;

            if (typeof openFeatureModal === 'function') {
                modal.style.display = 'flex';
                setTimeout(() => {
                    modal.classList.add('active');
                    modal.setAttribute('aria-hidden', 'false');
                }, 10);
            } else {
                modal.style.display = 'block';
            }
        },

        async handleFormSubmit(form) {
            const submitBtn = form.querySelector('button[type="submit"]');
            const originalBtnContent = submitBtn ? submitBtn.innerHTML : '';
            
            if (submitBtn) {
                submitBtn.disabled = true;
                submitBtn.innerHTML = '<span class="material-icons-outlined animate-spin text-sm">refresh</span> Processando...';
            }

            try {
                const formData = new FormData(form);
                const response = await fetch(form.action, {
                    method: 'POST',
                    body: formData,
                    headers: {
                        'X-Requested-With': 'XMLHttpRequest'
                    }
                });

                const data = await response.json();

                if (response.ok && data.status === 'success') {
                    this.showNotification(data.message || 'Operação realizada com sucesso!', 'success');
                    
                    if (form.id === 'editUserForm-aprov') {
                        closeModal('editUserModal-aprov');
                        // Ideally update the specific row data here if possible
                         if (data.username && data.id) {
                            const nameCell = document.querySelector(`button[data-user-id="${data.id}"]`)?.closest('tr')?.querySelector('td:nth-child(2)');
                            if (nameCell) nameCell.textContent = data.username;
                         }
                    } else {
                        // Handle Row Actions (Approve, Reject, Remove)
                        const row = form.closest('tr');
                        const actionUrl = form.action;
                        
                        if (row) {
                            // Check if it's an Approval action
                            if (actionUrl.includes('aprovar_registro')) {
                                // Move to Approved Table
                                this.moveRowToApproved(data, row);
                            } else {
                                // Reject or Remove - just delete
                                this.removeRow(row);
                            }

                            // Update Dashboard KPI
                            this.updateDashboardKPI();
                        }
                    }
                    
                } else {
                   this.showNotification(data.message || 'Ocorreu um erro.', 'error');
                }

            } catch (error) {
                console.error('Erro:', error);
                this.showNotification('Erro de conexão ou servidor.', 'error');
            } finally {
                 if (submitBtn) {
                    submitBtn.disabled = false;
                    submitBtn.innerHTML = originalBtnContent;
                }
            }
        },

        moveRowToApproved(data, oldRow) {
            // 1. Remove from Pending
            this.removeRow(oldRow);

            // 2. Add to Approved Table
            // We need to find the approved table body. 
            // Since there are multiple tables, we look for the one in the "Usuários Aprovados" card.
            // A robust way is to select by the card title or structure, but simpler is assuming order or ID.
            // Let's look for the table that contains "Administrador" or "Editar" buttons, or simply the second table.
            
            const approvedTableBody = document.querySelectorAll('.data-table tbody')[1];
            if (approvedTableBody) {
                // Remove "Empty" message if present
                const emptyRow = approvedTableBody.querySelector('tr td[colspan="3"]')?.parentElement;
                if (emptyRow) emptyRow.remove();

                const newRow = document.createElement('tr');
                newRow.className = "border-b dark:border-gray-700 hover:bg-gray-50 dark:hover:bg-gray-600 animate__animated animate__fadeIn";
                newRow.innerHTML = `
                    <td class="px-6 py-4">${data.id}</td>
                    <td class="px-6 py-4 font-medium text-gray-900 dark:text-white">${data.username}</td>
                    <td class="px-6 py-4 text-center space-x-2">
                        <button type="button" class="btn btn-secondary btn-sm edit-user-btn-aprov inline-flex items-center gap-1" title="Editar Usuário"
                                data-user-id="${data.id}" data-user-username="${data.username}">
                            <span class="material-icons-outlined text-sm">edit</span> Editar
                        </button>
                        <form action="/remover_usuario/${data.id}/" method="POST" class="inline-block">
                            <input type="hidden" name="csrfmiddlewaretoken" value="${document.querySelector('[name=csrfmiddlewaretoken]').value}">
                            <button type="submit" class="btn btn-danger btn-sm inline-flex items-center gap-1" title="Remover Usuário">
                                <span class="material-icons-outlined text-sm">delete</span> Remover
                            </button>
                        </form>
                    </td>
                `;
                approvedTableBody.appendChild(newRow);
                
                // Reindex both tables
                this.reindexTable(approvedTableBody);
                
                // Update Dashboard KPI (Increment Active Users)
                this.incrementActiveUsersKPI();
            }
        },

        removeRow(row) {
            const tbody = row.parentElement;
            const isPendingTable = tbody.closest('.card').querySelector('h2').textContent.includes('Pendentes');
            
            row.remove();
            
            // Check if backend table is empty
            if (tbody.children.length === 0) {
                const colSpan = isPendingTable ? 4 : 3;
                const msg = isPendingTable ? 'Nenhum registro pendente.' : 'Nenhum usuário aprovado.';
                
                const emptyRow = document.createElement('tr');
                emptyRow.innerHTML = `<td colspan="${colSpan}" class="text-center py-6 text-gray-500">${msg}</td>`;
                tbody.appendChild(emptyRow);
            } else {
                this.reindexTable(tbody);
            }

            // Update KPIs based on which table was modified
            if (!isPendingTable) {
                console.log('Updating Active Users KPI...');
                this.updateActiveUsersKPI();
            }
        },
        updateDashboardKPI() {
            const kpiCount = document.getElementById('kpi-pending-count');
            const kpiBadge = document.getElementById('kpi-pending-badge');
            
            if (kpiCount) {
                let currentCount = parseInt(kpiCount.textContent) || 0;
                let newCount = Math.max(0, currentCount - 1);
                kpiCount.textContent = newCount;

                if (newCount === 0 && kpiBadge) {
                    kpiBadge.style.display = 'none';
                }
            }
        },

        updateActiveUsersKPI() {
            const kpiCount = document.getElementById('kpi-users-active-count');
            
            if (kpiCount) {
                let currentCount = parseInt(kpiCount.textContent) || 0;
                let newCount = Math.max(0, currentCount - 1);
                kpiCount.textContent = newCount;
            }
        },

        incrementActiveUsersKPI() {
            const kpiCount = document.getElementById('kpi-users-active-count');
            
            if (kpiCount) {
                let currentCount = parseInt(kpiCount.textContent) || 0;
                let newCount = currentCount + 1;
                kpiCount.textContent = newCount;
            }
        },

        reindexTable(tbody) {
            if (!tbody) return;
            const rows = tbody.querySelectorAll('tr');
            let counter = 1;
            rows.forEach(row => {
                // Skip empty message row
                if (row.querySelector('td[colspan]')) return;
                
                const firstCell = row.querySelector('td:first-child');
                if (firstCell) {
                    firstCell.textContent = counter++;
                }
            });
        },

        showNotification(message, type = 'info') {
            // Check if there is a global notification function
            // If not, use simple alert or inject into existing container
            // The template has #flash-messages-container-aprov
            
            const container = document.getElementById('flash-messages-container-aprov');
            if (container) {
                const msgDiv = document.createElement('div');
                msgDiv.className = `flash-message ${type} animate__animated animate__fadeIn`;
                msgDiv.innerHTML = `
                    <span class="material-icons-outlined mr-3">
                        ${type === 'success' ? 'check_circle' : type === 'error' ? 'warning' : 'info'}
                    </span>
                    <span>${message}</span>
                `;
                container.prepend(msgDiv);
                
                // Auto remove after 5s
                setTimeout(() => {
                    msgDiv.classList.remove('animate__fadeIn');
                    msgDiv.classList.add('animate__fadeOut');
                    setTimeout(() => msgDiv.remove(), 500);
                }, 5000);
            } else {
                alert(message);
            }
        }
    };

    AprovarRegistros.init();
});

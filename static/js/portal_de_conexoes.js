
function openServiceModal(action, id = '', name = '', status = 'operando') {
    document.getElementById('service-modal').style.display = 'block';
    document.getElementById('service-action').value = action;
    document.getElementById('service-id').value = id;
    document.getElementById('service-name').value = name;
    document.getElementById('service-status').value = status;
    document.getElementById('service-modal-title').innerText = action === 'add' ? 'Novo Serviço' : 'Editar Serviço';
}

function closeServiceModal() {
    document.getElementById('service-modal').style.display = 'none';
}

async function handleServiceSubmit(event) {
    event.preventDefault();
    const action = document.getElementById('service-action').value;
    const id = document.getElementById('service-id').value;
    const name = document.getElementById('service-name').value;
    const status = document.getElementById('service-status').value;
    
    // Get CSRF token from the cookie or meta tag since we can't use template tags in external JS
    // Assuming standard Django CSRF cookie name
    function getCookie(name) {
        let cookieValue = null;
        if (document.cookie && document.cookie !== '') {
            const cookies = document.cookie.split(';');
            for (let i = 0; i < cookies.length; i++) {
                const cookie = cookies[i].trim();
                // Does this cookie string begin with the name we want?
                if (cookie.substring(0, name.length + 1) === (name + '=')) {
                    cookieValue = decodeURIComponent(cookie.substring(name.length + 1));
                    break;
                }
            }
        }
        return cookieValue;
    }
    const csrftoken = getCookie('csrftoken');

    try {
        const response = await fetch('/manage_service/', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'X-CSRFToken': csrftoken
            },
            body: JSON.stringify({ action, id, name, status })
        });
        
        const data = await response.json();
        if (data.status === 'success') {
            closeServiceModal();
            const list = document.getElementById('services-list');
            
            // Define colors
            let statusColor = 'bg-red-500';
            let statusTextColor = 'text-red-500 dark:text-red-400';
            let statusDisplay = 'Offline';
            let statusPulse = '';
            
            if (status === 'operando') {
                statusColor = 'bg-green-500';
                statusTextColor = 'text-green-600 dark:text-green-400';
                statusDisplay = 'Operando';
            } else if (status === 'instavel') {
                statusColor = 'bg-yellow-500';
                statusTextColor = 'text-yellow-600 dark:text-yellow-400';
                statusDisplay = 'Instável';
                statusPulse = 'animate-pulse';
            }
            
            if (action === 'add') {
                const newItem = document.createElement('div');
                newItem.className = 'flex items-center justify-between text-sm group/item';
                newItem.dataset.id = data.id; // Use ID from backend in future
                
                // Construct HTML manually for the new item
                newItem.innerHTML = `
                    <span class="text-gray-600 dark:text-gray-300 flex items-center gap-2">
                         <span class="w-2 h-2 rounded-full ${statusColor} ${statusPulse}"></span> 
                         ${name}
                    </span>
                    
                    <div class="flex items-center gap-2">
                        <span class="font-medium text-xs ${statusTextColor}">
                            ${statusDisplay}
                        </span>
                        
                        <div class="flex items-center transition-opacity">
                            <button onclick="openServiceModal('edit', '${data.id}', '${name}', '${status}')" class="p-1 hover:bg-gray-200 dark:hover:bg-gray-700 rounded text-blue-500">
                                <span class="material-icons-outlined text-sm">edit</span>
                            </button>
                            <button onclick="deleteService('${data.id}')" class="p-1 hover:bg-gray-200 dark:hover:bg-gray-700 rounded text-red-500">
                                <span class="material-icons-outlined text-sm">delete</span>
                            </button>
                        </div>
                    </div>
                `;
                
                // Clear "empty" message if it exists
                const emptyMsg = list.querySelector('p.text-center');
                if (emptyMsg) emptyMsg.remove();
                
                list.appendChild(newItem);
                
            } else if (action === 'edit') {
                const item = list.querySelector(`[data-id="${id}"]`);
                if (item) {
                    // Update Status Dot and Name
                    const nameSpan = item.querySelector('span.text-gray-600');
                    if(nameSpan) {
                         nameSpan.innerHTML = `
                             <span class="w-2 h-2 rounded-full ${statusColor} ${statusPulse}"></span> 
                             ${name}
                         `;
                    }
                    
                    // Update Status Text
                    const statusTextSpan = item.querySelector('.font-medium');
                    if(statusTextSpan) {
                        statusTextSpan.className = `font-medium text-xs ${statusTextColor}`;
                        statusTextSpan.innerText = statusDisplay;
                    }
                    
                    // Update Edit Button onclick to reflect new values
                    const editBtn = item.querySelector('button[onclick^="openServiceModal"]');
                    if(editBtn) {
                        editBtn.setAttribute('onclick', `openServiceModal('edit', '${id}', '${name}', '${status}')`);
                    }
                }
            }
        } else {
            alert('Erro: ' + data.message);
        }
    } catch (error) {
        console.error('Error:', error);
        alert('Erro ao processar requisição');
    }
}

async function deleteService(id) {
    if (typeof window.App !== 'undefined' && window.App.modals && window.App.modals.confirm) {
        window.App.modals.confirm(
            "Remover Serviço",
            "Tem certeza que deseja remover este serviço?",
            () => performDeleteService(id)
        );
    } else {
        if (confirm('Tem certeza que deseja remover este serviço?')) {
            performDeleteService(id);
        }
    }
}

async function performDeleteService(id) {
    // Helper to get cookie (duplicated here for now, better to move to a utils file later)
    function getCookie(name) {
        let cookieValue = null;
        if (document.cookie && document.cookie !== '') {
            const cookies = document.cookie.split(';');
            for (let i = 0; i < cookies.length; i++) {
                const cookie = cookies[i].trim();
                // Does this cookie string begin with the name we want?
                if (cookie.substring(0, name.length + 1) === (name + '=')) {
                    cookieValue = decodeURIComponent(cookie.substring(name.length + 1));
                    break;
                }
            }
        }
        return cookieValue;
    }
    const csrftoken = getCookie('csrftoken');

    try {
        const response = await fetch('/manage_service/', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'X-CSRFToken': csrftoken
            },
            body: JSON.stringify({ action: 'delete', id })
        });
        
        const data = await response.json();
        if (data.status === 'success') {
            const item = document.querySelector(`[data-id="${id}"]`);
            if (item) item.remove();
            
            const list = document.getElementById('services-list');
            if (list && list.children.length === 0) {
                 list.innerHTML = '<p class="text-xs text-center text-gray-500 py-2">Nenhum serviço monitorado.</p>';
            }
        } else {
            alert('Erro: ' + data.message);
        }
    } catch (error) {
        console.error('Error:', error);
        alert('Erro ao processar requisição');
    }
}

// Global Modal Management
function openFeatureModal(modalId) {
    // Close all feature modals first
    const modals = document.querySelectorAll('.feature-modal');
    modals.forEach(modal => {
        modal.classList.remove('active');
        modal.setAttribute('aria-hidden', 'true');
        modal.style.display = 'none'; // Ensure it's hidden
    });

    // Open the target modal
    const targetModal = document.getElementById(modalId);
    if (targetModal) {
        targetModal.classList.remove('hidden'); // Ensure hidden class is removed (added by modais.js)
        targetModal.style.display = 'flex'; // Use flex to center if using flex layout
        // Small timeout to allow display change before adding active class for animation
        setTimeout(() => {
            targetModal.classList.add('active');
            targetModal.setAttribute('aria-hidden', 'false');
        }, 10);
    }
}

function closeFeatureModal(modalId) {
    const modal = document.getElementById(modalId);
    if (modal) {
        modal.classList.remove('active');
        modal.setAttribute('aria-hidden', 'true');
        setTimeout(() => {
            modal.style.display = 'none';
        }, 300); // Wait for animation
    }
}

// Close specific modals
function closeModal(modalId) {
    const modal = document.getElementById(modalId);
    if (modal) {
        modal.style.display = 'none';
        modal.setAttribute('aria-hidden', 'true'); 
    }
}

// Event Listeners for closing on click outside (optional, carefully applied)
window.onclick = function(event) {
    if (event.target.classList.contains('feature-modal')) {
        closeFeatureModal(event.target.id);
    }
}

// Widget Toggle Universal
function toggleWidget(contentId, headerElement) {
    const content = document.getElementById(contentId);
    // Derive arrow ID from content ID (e.g. 'anotacoes-content' -> 'anotacoes-arrow')
    const arrowId = contentId.replace('-content', '-arrow');
    const arrow = document.getElementById(arrowId);

    if (content.classList.contains('hidden')) {
        // Open
        content.classList.remove('hidden');
        if (arrow) {
            arrow.style.transform = 'rotate(180deg)';
        }
    } else {
        // Close
        content.classList.add('hidden');
        if (arrow) {
            arrow.style.transform = 'rotate(0deg)';
        }
    }
}

// Initialize Widgets based on screen size
document.addEventListener('DOMContentLoaded', function() {
    const isMobile = window.innerWidth < 1024;
    const widgetContents = document.querySelectorAll('.widget-content');
    
    widgetContents.forEach(content => {
        // Derive arrow ID from content ID
        const arrowId = content.id.replace('-content', '-arrow');
        const arrow = document.getElementById(arrowId);

        if (isMobile) {
            content.classList.add('hidden');
            if (arrow) arrow.style.transform = 'rotate(0deg)';
        } else {
            content.classList.remove('hidden');
            // Desktop starts OPEN, so rotate arrows to 180 (point UP)
            if (arrow) arrow.style.transform = 'rotate(180deg)';
        }
    });
});

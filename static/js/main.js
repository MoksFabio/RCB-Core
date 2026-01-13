/* Main Entry Point */

// Ensure App Object exists and has State/Config
window.App = window.App || {};

// Merge default state/config if not already present (handled by modules mostly, but good to have base)
window.App.elements = {
    body: document.body,
    html: document.documentElement,
    sidebar: document.getElementById("sidebar"),
    mainContent: document.getElementById("main-content-wrapper"),
};

window.App.state = {
    sidebarCollapsed: false,
    currentProfileData: null,
    isEditingProfile: false,
    selectedProfileImageFile: null,
    kanbanStatuses: [],
    kanbanFunctionalities: [],
    kanbanDraggedItem: null,
    kanbanSaveTimeout: null,
    notesSaveTimeout: null,
    commandPaletteLinks: [],
    commandPaletteIndex: -1,
    isHannahDragging: false,
    hannahDragOffset: { x: 0, y: 0 },
};

window.App.config = {
    updateProfileUrl: window.AppDjangoConfig?.urls?.updateProfile || "/api/update_user_profile/",
    getProfileUrl: window.AppDjangoConfig?.urls?.getProfile || "/api/get_user_profile/",
    hannahChatApiUrl: window.AppDjangoConfig?.urls?.hannahChat || "/api/hannah/chat/",
    statusApiUrl: window.AppDjangoConfig?.urls?.statusApi || "/api/service-status/",
    profileImagesUrl: window.AppDjangoConfig?.profileImagesUrl || "/media/profile_images/",
    latestNotificationUrl: window.AppDjangoConfig?.urls?.latestNotification || "/api/latest-notification/",
    manageNotificationUrl: window.AppDjangoConfig?.urls?.manageNotification || "/manage_notification/",
};


window.App.init = function() {
    console.log("PortalJS v3.0 Modules Initializing...");
    
    // Initialize all modules that exist
    if (this.theme) this.theme.init();
    if (this.sidebar) this.sidebar.init();
    if (this.clock) this.clock.init();
    if (this.notifications) this.notifications.init();
    if (this.profile) this.profile.init();
    if (this.modals) this.modals.init();
    if (this.kanban) this.kanban.init();
    if (this.chat) this.chat.init();
    if (this.commandPalette) this.commandPalette.init();
    if (this.dashboardAgenda) this.dashboardAgenda.init();
    if (this.quickNotes) this.quickNotes.init();
    
    // Generic Global Event Listeners
    this.initGlobalEvents();

    console.log("PortalJS v3.0 Modules Initialized.");
};

window.App.initGlobalEvents = function() {
    // Force reload if page is loaded from back/forward cache
    window.addEventListener( "pageshow", function ( event ) {
        var historyTraversal = event.persisted || 
                                ( typeof window.performance != "undefined" && 
                                    window.performance.navigation.type === 2 );
        if ( historyTraversal ) {
            // Handle page restore.
            window.location.reload();
        }
    });

    // Nuclear Auth Check: Call API to verify session on every load (even cache)
    // Only if not on login/public page (naive check, better to rely on backend redirects but this is a double check requested)
    if (!window.location.pathname.includes('/login')) {
        fetch('/api/check_session/')
        .then(response => response.json())
        .then(data => {
            if (!data.is_authenticated) {
                window.location.href = '/login/';
            }
        })
        .catch(err => {
            console.error("Auth check failed", err);
        });
    }
};


// Auto-start on DOMContentLoaded
document.addEventListener("DOMContentLoaded", () => {
    window.App.init();
});

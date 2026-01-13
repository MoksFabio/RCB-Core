document.addEventListener("DOMContentLoaded", function () {
  console.log("JS: Documento carregado. Iniciando script Login.js.");

  function getCookie(name) {
    let cookieValue = null;
    if (document.cookie && document.cookie !== "") {
      const cookies = document.cookie.split(";");
      for (let i = 0; i < cookies.length; i++) {
        const cookie = cookies[i].trim();
        if (cookie.substring(0, name.length + 1) === name + "=") {
          cookieValue = decodeURIComponent(cookie.substring(name.length + 1));
          break;
        }
      }
    }
    return cookieValue;
  }
  const csrftoken = getCookie("csrftoken");
  console.log("JS: Token CSRF encontrado:", csrftoken ? "Sim" : "Não");

  const savedTheme = localStorage.getItem("theme") || "claro";
  const savedAltoContraste =
    localStorage.getItem("altoContraste") === "true";

  if (savedAltoContraste) {
    document.body.classList.add("alto-contraste");
  } else if (savedTheme === "escuro") {
    document.body.classList.add("dark-theme");
  }

  function setupPasswordToggle(inputId, toggleId) {
    const passwordInput = document.getElementById(inputId);
    const togglePassword = document.getElementById(toggleId);

    if (passwordInput && togglePassword) {
      togglePassword.addEventListener("click", function () {
        const type =
          passwordInput.getAttribute("type") === "password"
            ? "text"
            : "password";
        passwordInput.setAttribute("type", type);
        this.textContent =
          type === "password" ? "visibility_off" : "visibility";
      });
    }
  }

  setupPasswordToggle("login-password", "toggle-login-password");
  setupPasswordToggle("register-password", "toggle-register-password");
  setupPasswordToggle(
    "register-confirm-password",
    "toggle-register-confirm-password"
  );
  console.log("JS: Procurando elementos do flipper...");
  const flipper = document.getElementById("validator-flipper");
  console.log("JS: 'validator-flipper' encontrado:", flipper);

  const toggleToRegister = document.getElementById("toggle-to-register");
  console.log("JS: 'toggle-to-register' encontrado:", toggleToRegister);

  const toggleToLogin = document.getElementById("toggle-to-login");
  console.log("JS: 'toggle-to-login' encontrado:", toggleToLogin);

  const frontFace = document.querySelector(".validator-front");
  console.log("JS: '.validator-front' encontrado:", frontFace);

  const backFace = document.querySelector(".validator-back");
  console.log("JS: '.validator-back' encontrado:", backFace);

  if (flipper && toggleToRegister && toggleToLogin && frontFace && backFace) {
    console.log("JS: SUCESSO! Elementos do flipper encontrados. Adicionando 'escutas' de clique.");

    function adjustFlipperHeight() {
      if (flipper.classList.contains("is-flipped")) {
        console.log("JS: Ajustando altura para FACE TRASEIRA. Altura medida:", backFace.offsetHeight);
        flipper.style.height = backFace.offsetHeight + "px";
      } else {
        console.log("JS: Ajustando altura para FACE FRONTAL. Altura medida:", frontFace.offsetHeight);
        flipper.style.height = frontFace.offsetHeight + "px";
      }
    }

    toggleToRegister.addEventListener("click", function (e) {
      e.preventDefault();
      console.log("JS: CLIQUE! Girando para Registro.");
      flipper.classList.add("is-flipped");
      adjustFlipperHeight();
    });

    toggleToLogin.addEventListener("click", function (e) {
      e.preventDefault();
      console.log("JS: CLIQUE! Girando para Login.");
      flipper.classList.remove("is-flipped");
      adjustFlipperHeight();
    });

    window.addEventListener("resize", adjustFlipperHeight);
    
    adjustFlipperHeight(); 
    setTimeout(() => {
        if (flipper) { 
            flipper.classList.add("height-transition-enabled");
            console.log("JS: Transição de altura ativada (para cliques futuros).");
        }
    }, 50);

  } else {
    console.error("JS: ERRO CRÍTICO! Um ou mais elementos do flipper não foram encontrados. A animação de giro não vai funcionar.");
    console.error("JS: Verifique se os IDs/Classes no Login.html estão corretos.");
  }

  const loginForm = document.getElementById("login-form");
  const cardButton = document.getElementById("submit-card");
  const cardWrapper = document.getElementById("card-wrapper");
  const nfcButton = document.getElementById("nfc-submit-button");
  const loginUsernameInput = document.getElementById("login-username");
  const loginPasswordInput = document.getElementById("login-password");
  const lightRed = document.getElementById("light-red");
  const lightGreen = document.getElementById("light-green");

  function triggerSuccess(redirectUrl) {
    if (lightGreen) {
      lightGreen.classList.add("active-green");
    }
    if (cardButton) {
      cardButton.disabled = true;
    }

    setTimeout(function () {
      window.location.href = redirectUrl;
    }, 1000);
  }

  function triggerError(message, containerId, buttonsToEnable = []) {
    const errorContainer = document.getElementById(containerId);

    if (errorContainer) {
      errorContainer.innerHTML = "";
    }

    const errorDiv = document.createElement("div");
    errorDiv.className = "flash-error";
    errorDiv.setAttribute("role", "alert");

    const iconSpan = document.createElement("span");
    iconSpan.className = "material-icons-outlined";
    iconSpan.textContent = "error_outline";

    errorDiv.appendChild(iconSpan);
    errorDiv.appendChild(document.createTextNode(" " + message));

    if (errorContainer) {
      errorContainer.appendChild(errorDiv);
      errorContainer.style.display = "block";
    }

    if (lightRed) {
      lightRed.classList.add("active-red");
    }

    setTimeout(function () {
      if (lightRed) {
        lightRed.classList.remove("active-red");
      }
      if (errorContainer) {
        errorContainer.style.display = "none";
        errorContainer.innerHTML = "";
      }

      buttonsToEnable.forEach(button => {
        if (button) {
          button.disabled = false;
        }
      });
      
      if (cardWrapper && cardButton) {
        cardWrapper.classList.remove("is-visible");
        cardButton.classList.remove("tapping");
      }
    }, 3000);
  }

  async function submitLogin() {
    const username = loginUsernameInput.value;
    const password = loginPasswordInput.value;
    const loginUrl = loginForm.dataset.loginUrl;

    if (!csrftoken) {
      triggerError("Erro de segurança. Recarregue a página.", "login-flash-container", [cardButton, nfcButton]);
      return;
    }
    
    if (!loginUrl) {
      triggerError("Erro de configuração. URL de login não encontrada.", "login-flash-container", [cardButton, nfcButton]);
      return;
    }

    try {
      const response = await fetch(loginUrl, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "Accept": "application/json",
          "X-CSRFToken": csrftoken,
        },
        body: JSON.stringify({ username: username, password: password }),
      });

      const data = await response.json();

      if (response.ok && data.status === "success") {
        triggerSuccess(data.redirect_url);
      } else {
        triggerError(data.message || "Erro desconhecido", "login-flash-container", [cardButton, nfcButton]);
      }
    } catch (error) {
      console.error("Erro ao tentar fazer login:", error);
      triggerError("Erro de conexão. Tente novamente.", "login-flash-container", [cardButton, nfcButton]);
    }
  }

  if (
    loginForm &&
    cardButton &&
    cardWrapper &&
    nfcButton &&
    loginPasswordInput &&
    loginUsernameInput
  ) {
    function runLoginAnimation() {
      if (cardButton.disabled) return;

      cardButton.disabled = true;
      nfcButton.disabled = true;
      cardWrapper.classList.add("is-visible");

      setTimeout(function () {
        cardButton.classList.add("tapping");

        setTimeout(function () {
          submitLogin();
        }, 700);
      }, 600);
    }

    nfcButton.addEventListener("click", function (event) {
      event.preventDefault();
      runLoginAnimation();
    });

    loginUsernameInput.addEventListener("keydown", function (event) {
      if (event.key === "Enter") {
        event.preventDefault();
        runLoginAnimation();
      }
    });

    loginPasswordInput.addEventListener("keydown", function (event) {
      if (event.key === "Enter") {
        event.preventDefault();
        runLoginAnimation();
      }
    });
  }
  
  const registerForm = document.getElementById("register-form");
  const registerSubmitButton = document.getElementById("register-submit-button");
  const registerUsernameInput = document.getElementById("register-username");
  const registerPasswordInput = document.getElementById("register-password");
  const registerConfirmPasswordInput = document.getElementById("register-confirm-password");

  async function submitRegister() {
    const username = registerUsernameInput.value;
    const password = registerPasswordInput.value;
    const confirm_password = registerConfirmPasswordInput.value;
    const registerUrl = registerForm.dataset.registerUrl;

    if (!csrftoken) {
      triggerError("Erro de segurança. Recarregue a página.", "register-flash-container", [registerSubmitButton]);
      return;
    }

    if (!registerUrl) {
      triggerError("Erro de configuração. URL de registo não encontrada.", "register-flash-container", [registerSubmitButton]);
      return;
    }
    
    registerSubmitButton.disabled = true;

    try {
      const response = await fetch(registerUrl, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "Accept": "application/json",
          "X-CSRFToken": csrftoken,
        },
        body: JSON.stringify({ 
          username: username, 
          password: password, 
          confirm_password: confirm_password 
        }),
      });

      const data = await response.json();

      if (response.ok && data.status === "success") {
        triggerError(data.message, "register-flash-container", [registerSubmitButton]);
        registerForm.reset();
      } else {
        triggerError(data.message || "Erro desconhecido", "register-flash-container", [registerSubmitButton]);
      }
    } catch (error) {
      console.error("Erro ao tentar registrar:", error);
      triggerError("Erro de conexão. Tente novamente.", "register-flash-container", [registerSubmitButton]);
    }
  }

  if (registerForm && registerSubmitButton) {
    registerSubmitButton.addEventListener("click", function (event) {
      event.preventDefault();
      submitRegister();
    });
  }
  
  // --- Início do Código do Modal da Trajetória ---
  console.log("JS: Procurando elementos do Modal da Trajetória...");
  const openModalButton = document.getElementById("open-history-modal");
  const closeModalButton = document.getElementById("close-history-modal");
  const modal = document.getElementById("history-modal");
  const modalOverlay = document.getElementById("modal-overlay");
  
  const pagesWrapper = document.getElementById("modal-pages-wrapper");
  const btnPrev = document.getElementById("modal-btn-prev");
  const btnNext = document.getElementById("modal-btn-next");
  const pageIndicator = document.getElementById("modal-page-indicator");
  const modalPages = document.querySelectorAll(".modal-page");
  
  let currentPageIndex = 0;
  const totalPages = modalPages.length;

  function updateModalPages() {
    if (!pagesWrapper) return;
    
    pagesWrapper.style.transform = `translateX(-${currentPageIndex * 100}%)`;
    
    if (pageIndicator) {
      pageIndicator.textContent = `Página ${currentPageIndex + 1} de ${totalPages}`;
    }
    
    if (btnPrev) {
      btnPrev.style.visibility = currentPageIndex === 0 ? "hidden" : "visible";
    }
    
    if (btnNext) {
      if (currentPageIndex === totalPages - 1) {
        btnNext.textContent = "Fechar";
        btnNext.removeEventListener("click", goToNextPage);
        btnNext.addEventListener("click", closeModal);
      } else {
        btnNext.innerHTML = 'Próximo <span class="material-icons-outlined">chevron_right</span>';
        btnNext.removeEventListener("click", closeModal);
        btnNext.addEventListener("click", goToNextPage);
      }
    }
  }
  
  function goToNextPage() {
    if (currentPageIndex < totalPages - 1) {
      currentPageIndex++;
      updateModalPages();
    }
  }
  
  function goToPrevPage() {
    if (currentPageIndex > 0) {
      currentPageIndex--;
      updateModalPages();
    }
  }

  function openModal() {
    if (modal) {
      console.log("JS: Abrindo modal.");
      modal.classList.remove("hidden");
      currentPageIndex = 0;
      updateModalPages();
    }
  }

  function closeModal() {
    if (modal) {
      console.log("JS: Fechando modal.");
      modal.classList.add("hidden");
    }
  }

  if (openModalButton && closeModalButton && modal && modalOverlay && btnPrev && btnNext && pagesWrapper) {
    console.log("JS: SUCESSO! Elementos do modal carrossel encontrados. Adicionando 'escutas' de clique.");
    openModalButton.addEventListener("click", openModal);
    closeModalButton.addEventListener("click", closeModal);
    modalOverlay.addEventListener("click", closeModal);
    
    btnPrev.addEventListener("click", goToPrevPage);
    
    updateModalPages();
  } else {
      console.warn("JS: AVISO! Elementos do modal carrossel 'Saiba mais' não foram encontrados. Verifique os IDs no HTML.");
  }
  // --- Fim do Código do Modal da Trajetória ---
  
});
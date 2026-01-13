 /**
 * Módulo: Perfil do Usuário
 * Resumo: Gerencia a visualização e edição dos dados do usuário, incluindo upload de foto e salvamento de informações.
 */
window.App = window.App || {};

 window.App.profile = {
      elements: {
        // Mapeamento sob demanda
      },
      init() {
        this.addListeners();
      },
      addListeners() {
        const editBtn = document.getElementById("editProfileButton");
        const cancelBtn = document.getElementById("cancelEditProfileButton");
        const saveBtn = document.getElementById("saveProfileButton");
        const imgInput = document.getElementById("profileImageInput");

        if (editBtn) editBtn.addEventListener("click", () => this.toggleEditMode(true));
        if (cancelBtn) cancelBtn.addEventListener("click", () => this.toggleEditMode(false));
        if (saveBtn) saveBtn.addEventListener("click", () => this.save());
        if (imgInput) imgInput.addEventListener("change", (e) => this.previewImage(e));
      },
      toggleEditMode(enableEdit) {
        App.state.isEditingProfile = enableEdit;

        const displayIds = ["editProfileButton"]; // Don't hide display text on card
        const editIds = ["profileEditSidePanel", "profileEditButtons"]; // Show side panel and buttons

        const setDisplay = (ids, action) => {
          ids.forEach(id => {
            const el = document.getElementById(id);
            if (el) action === 'hide' ? el.classList.add("hidden") : el.classList.remove("hidden");
          });
        };

        if (enableEdit) {
          setDisplay(displayIds, 'hide');
          setDisplay(editIds, 'show');

          const data = App.state.currentProfileData;
          if (data) {
            const setVal = (id, val) => { const el = document.getElementById(id); if (el) el.value = val || ""; };
            setVal("employeeNameInput", data.name);
            setVal("employeeRoleInput", data.role);
            setVal("employeeEmailInput", data.email);
            setVal("employeePhoneInput", data.phone);
            setVal("employeeBioEdit", data.bio);

            const hireEdit = document.getElementById("hireDateValueEdit");
            if (hireEdit && data.hireDate) {
              try {
                const date = new Date(data.hireDate + "T00:00:00").toLocaleDateString("pt-BR");
                hireEdit.innerHTML = `${date} <span class="text-xs italic">(Não editável)</span>`;
              } catch (e) { }
            }
          }
          document.getElementById("employeeNameInput")?.focus();
        } else {
          setDisplay(displayIds, 'show');
          setDisplay(editIds, 'hide');
          const saveStatus = document.getElementById("profileSaveStatus");
          if (saveStatus) saveStatus.classList.add("hidden");

          const imgInput = document.getElementById("profileImageInput");
          if (imgInput) imgInput.value = null;
          App.state.selectedProfileImageFile = null;

          if (App.state.currentProfileData) {
            this.setImage(App.state.currentProfileData.imageUrl, App.state.currentProfileData.name);
          }
        }
      },
      setImage(imageUrl, name = "Usuário") {
        const image = document.getElementById("profileImage");
        if (!image) return;

        let src = "";

        if (imageUrl && typeof imageUrl === "string" && imageUrl.trim() !== "") {
          if (!imageUrl.startsWith("http") && !imageUrl.startsWith("/static/")) {
            src = App.config.profileImagesUrl + imageUrl.split("/").pop();
          } else {
            src = imageUrl;
          }
        } else {
          const initials = (name || "U").split(" ").map((n) => n[0]).slice(0, 2).join("").toUpperCase();
          const themeColor = getComputedStyle(document.documentElement).getPropertyValue("--cor-principal").trim().substring(1) || "f97316";
          src = `https://via.placeholder.com/150/${themeColor}/ffffff?text=${initials || "?"}`;
        }

        image.src = src;
      },
      load() {
        const saveStatus = document.getElementById("profileSaveStatus");
        if (saveStatus) saveStatus.classList.add("hidden");

        const loading = document.getElementById("profileLoadingView");
        const infoLoading = document.getElementById("profileInfoLoadingView");
        const imgCont = document.getElementById("profileImageContainer");
        const display = document.getElementById("profileDisplayView");
        const infoCont = document.getElementById("profileInfoContainer");
        const editBtn = document.getElementById("editProfileButton");

        if (loading) loading.classList.remove("hidden");
        if (infoLoading) infoLoading.classList.remove("hidden");
        if (imgCont) imgCont.classList.add("hidden");
        if (display) display.classList.add("hidden");
        if (infoCont) infoCont.classList.add("hidden");
        if (editBtn) editBtn.classList.add("hidden");

        fetch(App.config.getProfileUrl)
          .then(response => {
            if (!response.ok) return response.json().then(err => { throw new Error(err.message || `Erro ${response.status}`); }).catch(() => { throw new Error(`Erro ${response.status}`); });
            return response.json();
          })
          .then(profileData => {
            App.state.currentProfileData = profileData;

            const setText = (id, txt) => { const el = document.getElementById(id); if (el) { el.textContent = txt; el.title = txt; } };

            setText("employeeNameDisplay", profileData.name || "Nome não definido");
            setText("employeeRoleDisplay", profileData.role || "Cargo não definido");
            setText("employeeRoleView", profileData.role || "--");
            setText("employeeEmailView", profileData.email || "--");
            setText("employeePhoneView", profileData.phone || "--");
            setText("employeeBioView", profileData.bio || "Nenhuma biografia adicionada.");

            if (profileData.hireDate) {
              const hireView = document.getElementById("hireDateValueView");
              const hireEdit = document.getElementById("hireDateValueEdit");
              try {
                const date = new Date(profileData.hireDate + "T00:00:00").toLocaleDateString("pt-BR");
                if (hireView) hireView.textContent = date;
                if (hireEdit) hireEdit.innerHTML = `${date} <span class="text-xs italic">(Não editável)</span>`;
              } catch (e) { console.error("Error parsing hire date:", e); }
            }

            this.setImage(profileData.imageUrl, profileData.name);
            this.toggleEditMode(false);
          })
          .catch(error => {
            console.error("Erro ao carregar perfil:", error);
            const nameDisplay = document.getElementById("employeeNameDisplay");
            if (nameDisplay) nameDisplay.textContent = "Erro";
          })
          .finally(() => {
            if (loading) loading.classList.add("hidden");
            if (infoLoading) infoLoading.classList.add("hidden");
            if (imgCont) imgCont.classList.remove("hidden");
            if (display) display.classList.remove("hidden");
            if (infoCont) infoCont.classList.remove("hidden");
            if (editBtn) editBtn.classList.remove("hidden");
          });
      },
      save() {
        const saveStatus = document.getElementById("profileSaveStatus");
        const saveBtn = document.getElementById("saveProfileButton");
        const cancelBtn = document.getElementById("cancelEditProfileButton");

        if (saveStatus) {
          saveStatus.textContent = "Salvando...";
          saveStatus.className = "text-sm mt-4 font-medium text-gray-500 dark:text-gray-400 h-5 text-left";
          saveStatus.classList.remove("hidden");
        }
        if (saveBtn) saveBtn.disabled = true;
        if (cancelBtn) cancelBtn.disabled = true;

        let bodyContent;
        let headers = {};

        const getVal = (id) => document.getElementById(id)?.value.trim() || "";

        if (App.state.selectedProfileImageFile) {
          bodyContent = new FormData();
          bodyContent.append("profileImage", App.state.selectedProfileImageFile);
          bodyContent.append("name", getVal("employeeNameInput"));
          bodyContent.append("role", getVal("employeeRoleInput"));
          bodyContent.append("email", getVal("employeeEmailInput"));
          bodyContent.append("phone", getVal("employeePhoneInput"));
          bodyContent.append("bio", getVal("employeeBioEdit"));
        } else {
          bodyContent = JSON.stringify({
            name: getVal("employeeNameInput"),
            role: getVal("employeeRoleInput"),
            email: getVal("employeeEmailInput"),
            phone: getVal("employeePhoneInput"),
            bio: getVal("employeeBioEdit"),
          });
          headers["Content-Type"] = "application/json";
        }

        fetch(App.config.updateProfileUrl, { method: "PUT", headers: headers, body: bodyContent })
          .then(response => response.json().then(data => ({ status: response.status, ok: response.ok, body: data })))
          .then(result => {
            if (!result.ok) throw new Error(result.body.message || `Erro ${result.status}`);

            const savedProfile = result.body;
            App.state.currentProfileData = savedProfile;
            App.state.selectedProfileImageFile = null;

            // Update UI
            const setText = (id, txt) => { const el = document.getElementById(id); if (el) { el.textContent = txt; } };
            setText("employeeNameDisplay", savedProfile.name || "N/D");
            setText("employeeRoleDisplay", savedProfile.role || "N/D");
            setText("employeeRoleView", savedProfile.role || "--");

            this.setImage(savedProfile.imageUrl || App.state.currentProfileData.imageUrl, savedProfile.name);

            if (saveStatus) {
              saveStatus.textContent = "Perfil salvo com sucesso!";
              saveStatus.className = "text-sm mt-4 font-medium text-green-600 h-5 text-left";
              saveStatus.classList.remove("hidden");
            }
            setTimeout(() => this.toggleEditMode(false), 1500);
          })
          .catch(error => {
            console.error("Erro ao salvar perfil:", error);
            if (saveStatus) {
              saveStatus.textContent = `Erro ao salvar: ${error.message}`;
              saveStatus.className = "text-sm mt-4 font-medium text-red-600 h-5 text-left";
              saveStatus.classList.remove("hidden");
            }
          })
          .finally(() => {
            if (saveBtn) saveBtn.disabled = false;
            if (cancelBtn) cancelBtn.disabled = false;
          });
      },
      cancelEdit() {
        if (App.state.isEditingProfile) {
          this.toggleEditMode(false);
        }
      },
      previewImage(event) {
        const file = event.target.files[0];
        const image = document.getElementById("profileImage");
        if (file && image) {
          App.state.selectedProfileImageFile = file;
          const reader = new FileReader();
          reader.onload = (e) => {
            image.src = e.target.result;
          };
          reader.readAsDataURL(file);
        }
      },
    };

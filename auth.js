(function () {
  "use strict";

  const U = window.CTUtils;

  async function initLoginPage() {
    const form = U.byId("loginForm");
    if (!form || !window.CTSupabase) return;

    const session = await window.CTSupabase.getSession().catch(function () {
      return null;
    });

    if (session) {
      const profile = await window.CTSupabase.getProfile().catch(function () {
        return null;
      });
      window.location.href = profile && profile.role === "admin" ? "admin.html" : "index.html";
      return;
    }

    form.addEventListener("submit", async function (event) {
      event.preventDefault();
      U.clearMessage("loginMessage");

      const email = U.byId("email").value.trim();
      const password = U.byId("password").value;

      if (!email || !password) {
        U.showMessage("loginMessage", "Informe e-mail e senha para continuar.", "warning");
        return;
      }

      try {
        const button = U.byId("loginBtn");
        if (button) button.disabled = true;

        await window.CTSupabase.signIn(email, password);
        const profile = await window.CTSupabase.getProfile();

        if (!profile || !profile.role) {
          U.showMessage("loginMessage", "Perfil inválido. Verifique o cadastro do usuário.", "error");
          await window.CTSupabase.signOut();
          return;
        }

        window.location.href = profile.role === "admin" ? "admin.html" : "index.html";
      } catch (error) {
        U.showMessage("loginMessage", error.message || "Não foi possível entrar no sistema.", "error");
      } finally {
        const button = U.byId("loginBtn");
        if (button) button.disabled = false;
      }
    });
  }

  window.CTAuth = { initLoginPage };
})();

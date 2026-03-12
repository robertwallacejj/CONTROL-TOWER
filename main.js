(function () {
  "use strict";

  document.addEventListener("DOMContentLoaded", function () {
    const path = (window.location.pathname.split("/").pop() || "").toLowerCase();

    if (path === "login.html" && window.CTAuth) {
      window.CTAuth.initLoginPage().catch(function (error) {
        console.error("Falha ao iniciar login.", error);
      });
      return;
    }

    if (path === "admin.html" && window.CTAdmin) {
      window.CTAdmin.initAdminPage().catch(function (error) {
        console.error("Falha ao iniciar área admin.", error);
      });
      return;
    }

    if ((path === "index.html" || path === "") && window.CTDashboard) {
      window.CTDashboard.initIndexPage().catch(function (error) {
        console.error("Falha ao iniciar painel.", error);
      });
    }
  });
})();

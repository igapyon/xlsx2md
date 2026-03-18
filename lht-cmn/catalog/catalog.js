const demoTemplates = {
  "hero-basic": {
    afterRender(preview) {
      const menu = preview.querySelector("lht-page-menu");
      if (menu) {
        const menuPanel = menu.querySelector(".md-menu-panel");
        if (menuPanel) menuPanel.classList.add("md-hidden");
      }
    }
  },
  "switch-basic": {
    afterRender() {
      window.catalogHandleSwitchChange = () => {
        const input = document.getElementById("catalogDemoSwitch");
        const status = document.getElementById("switch-status");
        if (!input || !status) return;
        status.textContent = `Switch status: ${input.checked ? "on" : "off"}`;
      };
      window.catalogHandleSwitchChange();
    }
  },
  "command-basic": {
    afterRender(preview) {
      const command = preview.querySelector("#catalogCommand");
      if (command) {
        command.textContent = "git diff --stat origin/main...feature/catalog-page";
      }
    }
  },
  "file-select-basic": {
    afterRender(preview) {
      const element = preview.querySelector("lht-file-select");
      const log = document.getElementById("file-select-log");
      if (!element || !log) return;

      element.addEventListener("lht-file-select:before-open", (event) => {
        log.textContent = `before-open: autoOpen=${String(event.detail.autoOpen)}`;
      });
      element.addEventListener("lht-file-select:change", (event) => {
        const names = Array.isArray(event.detail.names) ? event.detail.names.join(", ") : "";
        log.textContent = names ? `change: ${names}` : "change: no file names";
      });
    }
  },
  "preview-basic": {
    afterRender(preview) {
      const output = preview.querySelector("#catalogPreviewOutput");
      if (!output) return;

      preview.querySelector('[data-action="preview-fill"]')?.addEventListener("click", () => {
        output.setText("<lht-preview-output> can update text via setText().");
      });
      preview.querySelector('[data-action="preview-clear"]')?.addEventListener("click", () => {
        output.clear();
      });
    }
  },
  "input-mode-basic": {
    afterRender() {
      window.catalogHandleModeChange = (mode) => {
        const status = document.getElementById("input-mode-status");
        if (!status) return;
        status.textContent = `Current mode: ${mode}`;
      };
      window.catalogHandleModeChange("file");
    }
  },
  "loading-basic": {
    afterRender(preview) {
      const overlay = preview.querySelector("#catalogLoadingOverlay");
      const trigger = preview.querySelector('[data-action="loading-run"]');
      if (!overlay || !trigger) return;

      trigger.addEventListener("click", async () => {
        overlay.setActive(true);
        await overlay.waitForNextPaint();
        await new Promise((resolve) => window.setTimeout(resolve, 900));
        overlay.setActive(false);
      });
    }
  },
  "toast-basic": {
    afterRender(preview) {
      const toast = preview.querySelector("#catalogToast");
      const trigger = preview.querySelector('[data-action="toast-show"]');
      if (!toast || !trigger) return;
      trigger.addEventListener("click", () => {
        toast.show("Catalog toast fired.", 1600);
      });
    }
  },
  "alert-basic": {
    afterRender(preview) {
      const alert = preview.querySelector("#catalogAlert");
      if (!alert) return;

      preview.querySelector('[data-action="alert-error"]')?.addEventListener("click", () => {
        alert.setAttribute("variant", "error");
        alert.show("Error: invalid input.");
      });
      preview.querySelector('[data-action="alert-warning"]')?.addEventListener("click", () => {
        alert.setAttribute("variant", "warning");
        alert.show("Warning: review your options.");
      });
      preview.querySelector('[data-action="alert-info"]')?.addEventListener("click", () => {
        alert.setAttribute("variant", "info");
        alert.show("Info: generation completed.");
      });
      preview.querySelector('[data-action="alert-clear"]')?.addEventListener("click", () => {
        alert.clear();
      });
    }
  }
};

function normalizeIndent(text) {
  const lines = text.replace(/^\n+|\n+$/g, "").split("\n");
  const indents = lines
    .filter((line) => line.trim().length > 0)
    .map((line) => line.match(/^\s*/)?.[0].length ?? 0);
  const minIndent = indents.length ? Math.min(...indents) : 0;
  return lines.map((line) => line.slice(minIndent)).join("\n");
}

function renderCatalogCard(card) {
  const key = card.dataset.demo;
  const template = document.querySelector(`[data-demo-template="${key}"]`);
  const preview = card.querySelector(".catalog-preview");
  const code = card.querySelector(".catalog-code code");
  if (!template || !preview || !code) return;

  const fragment = template.content.cloneNode(true);
  const source = normalizeIndent(template.innerHTML);

  preview.appendChild(fragment);
  code.textContent = source;

  const controller = demoTemplates[key];
  if (controller?.afterRender) {
    controller.afterRender(preview);
  }
}

function bootstrapCatalog() {
  document.querySelectorAll(".catalog-card").forEach((card) => {
    renderCatalogCard(card);
  });
}

window.addEventListener("DOMContentLoaded", bootstrapCatalog);

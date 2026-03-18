/*
 * lht-cmn components.js
 * Version: v20260308
 * Copyright 2026 Toshiki Iga
 * Licensed under the Apache License, Version 2.0
 */

class LhtHelpTooltip extends HTMLElement {
  connectedCallback() {
    if (this.dataset.initialized === "true") return;
    this.dataset.initialized = "true";

    const label = this.getAttribute("label") || "説明";
    const isWide = this.hasAttribute("wide");
    const helpContentHtml = this.innerHTML.trim();
    const placement = this._normalizePlacement(this.getAttribute("placement"));

    this.textContent = "";

    const group = document.createElement("span");
    group.className = "md-tooltip-group";

    const hasMdIconButton = !!(window.customElements && window.customElements.get("md-icon-button"));
    const button = document.createElement(hasMdIconButton ? "md-icon-button" : "button");
    button.className = "md-help-icon-button";
    if (!hasMdIconButton) {
      button.type = "button";
      button.classList.add("md-help-icon-button--fallback");
    }
    button.setAttribute("aria-label", label);
    button.innerHTML = '<svg aria-hidden="true" viewBox="0 0 24 24" class="md-info-icon" fill="none"><circle cx="12" cy="12" r="9" fill="#cbbcf0"/><rect x="11" y="10" width="2" height="7" rx="1" fill="#ffffff"/><circle cx="12" cy="7.5" r="1" fill="#ffffff"/></svg>';

    const tooltip = document.createElement("span");
    tooltip.className = `md-tooltip-content md-tooltip md-tooltip--rich${isWide ? " md-tooltip--wide" : ""}`;
    tooltip.innerHTML = helpContentHtml;
    tooltip.dataset.placement = placement;

    group.appendChild(button);
    group.appendChild(tooltip);
    this.appendChild(group);

    this._group = group;
    this._tooltip = tooltip;
    this._activeTooltip = false;
    this._handleTooltipEnter = () => {
      this._activeTooltip = true;
      group.removeAttribute("data-force-hidden");
      this._applyTooltipPlacement();
    };
    this._handleTooltipLeave = () => {
      this._activeTooltip = false;
      group.removeAttribute("data-force-hidden");
      this._resetTooltipPlacement();
    };
    this._handleTooltipKeydown = (event) => {
      if (event.key !== "Escape") return;
      event.preventDefault();
      event.stopPropagation();
      this._activeTooltip = false;
      group.setAttribute("data-force-hidden", "true");
      this._resetTooltipPlacement();
      if (document.activeElement && group.contains(document.activeElement)) {
        document.activeElement.blur();
      }
    };
    this._handleTooltipResize = () => {
      if (!this._activeTooltip) return;
      this._applyTooltipPlacement();
    };

    group.addEventListener("mouseenter", this._handleTooltipEnter);
    group.addEventListener("focusin", this._handleTooltipEnter);
    group.addEventListener("mouseleave", this._handleTooltipLeave);
    group.addEventListener("keydown", this._handleTooltipKeydown);
    group.addEventListener("focusout", () => {
      requestAnimationFrame(() => {
        if (!group.matches(":focus-within")) {
          this._handleTooltipLeave();
        }
      });
    });
    window.addEventListener("resize", this._handleTooltipResize);
  }

  disconnectedCallback() {
    if (this._group && this._handleTooltipEnter) {
      this._group.removeEventListener("mouseenter", this._handleTooltipEnter);
      this._group.removeEventListener("focusin", this._handleTooltipEnter);
      this._group.removeEventListener("mouseleave", this._handleTooltipLeave);
      this._group.removeEventListener("keydown", this._handleTooltipKeydown);
    }
    if (this._handleTooltipResize) {
      window.removeEventListener("resize", this._handleTooltipResize);
    }
  }

  _normalizePlacement(rawPlacement) {
    const normalized = (rawPlacement || "auto").trim().toLowerCase();
    return ["auto", "left", "right", "top", "bottom"].includes(normalized) ? normalized : "auto";
  }

  _applyTooltipPlacement() {
    const tooltip = this._tooltip;
    const group = this._group;
    if (!tooltip || !group) return;

    const viewportWidth = window.innerWidth || document.documentElement.clientWidth || 0;
    const viewportHeight = window.innerHeight || document.documentElement.clientHeight || 0;
    const safeWidth = Math.max(120, viewportWidth - 32);
    tooltip.style.maxWidth = `${safeWidth}px`;
    tooltip.style.visibility = "hidden";
    tooltip.style.display = "block";

    const anchorRect = group.getBoundingClientRect();
    const tooltipRect = tooltip.getBoundingClientRect();
    const placement = this._normalizePlacement(this.getAttribute("placement"));
    const appliedPlacement = placement === "auto"
      ? this._pickAutoPlacement(anchorRect, tooltipRect, viewportWidth, viewportHeight)
      : placement;
    const position = this._computeTooltipPosition(appliedPlacement, anchorRect, tooltipRect, viewportWidth, viewportHeight);
    const relativeLeft = position.left - anchorRect.left;
    const relativeTop = position.top - anchorRect.top;

    tooltip.dataset.placement = appliedPlacement;
    tooltip.style.left = `${relativeLeft}px`;
    tooltip.style.top = `${relativeTop}px`;
    tooltip.style.right = "auto";
    tooltip.style.bottom = "auto";
    tooltip.style.transform = "none";
    tooltip.style.marginTop = "0";
    tooltip.style.visibility = "";
  }

  _resetTooltipPlacement() {
    const tooltip = this._tooltip;
    if (!tooltip) return;
    tooltip.dataset.placement = this._normalizePlacement(this.getAttribute("placement"));
    tooltip.style.left = "";
    tooltip.style.top = "";
    tooltip.style.right = "";
    tooltip.style.bottom = "";
    tooltip.style.transform = "";
    tooltip.style.marginTop = "";
    tooltip.style.maxWidth = "";
    tooltip.style.visibility = "";
    tooltip.style.display = "";
  }

  _pickAutoPlacement(anchorRect, tooltipRect, viewportWidth, viewportHeight) {
    const candidates = ["right", "left", "bottom", "top"];
    let bestPlacement = "right";
    let bestScore = Number.POSITIVE_INFINITY;

    for (const candidate of candidates) {
      const position = this._computeTooltipPosition(candidate, anchorRect, tooltipRect, viewportWidth, viewportHeight);
      const score = this._computeOverflowScore(position, tooltipRect, viewportWidth, viewportHeight);
      if (score < bestScore) {
        bestScore = score;
        bestPlacement = candidate;
      }
    }

    return bestPlacement;
  }

  _computeTooltipPosition(placement, anchorRect, tooltipRect, viewportWidth, viewportHeight) {
    const gap = 8;
    const minInset = 16;
    const maxLeft = Math.max(minInset, viewportWidth - tooltipRect.width - minInset);
    const maxTop = Math.max(minInset, viewportHeight - tooltipRect.height - minInset);

    if (placement === "left") {
      return {
        left: Math.max(minInset, anchorRect.left - tooltipRect.width - gap),
        top: this._clamp(anchorRect.top + (anchorRect.height - tooltipRect.height) / 2, minInset, maxTop)
      };
    }
    if (placement === "right") {
      return {
        left: Math.min(maxLeft, anchorRect.right + gap),
        top: this._clamp(anchorRect.top + (anchorRect.height - tooltipRect.height) / 2, minInset, maxTop)
      };
    }
    if (placement === "top") {
      return {
        left: this._clamp(anchorRect.left + (anchorRect.width - tooltipRect.width) / 2, minInset, maxLeft),
        top: Math.max(minInset, anchorRect.top - tooltipRect.height - gap)
      };
    }
    return {
      left: this._clamp(anchorRect.left + (anchorRect.width - tooltipRect.width) / 2, minInset, maxLeft),
      top: Math.min(maxTop, anchorRect.bottom + gap)
    };
  }

  _computeOverflowScore(position, tooltipRect, viewportWidth, viewportHeight) {
    const overflowLeft = Math.max(0, 16 - position.left);
    const overflowRight = Math.max(0, position.left + tooltipRect.width + 16 - viewportWidth);
    const overflowTop = Math.max(0, 16 - position.top);
    const overflowBottom = Math.max(0, position.top + tooltipRect.height + 16 - viewportHeight);
    return overflowLeft + overflowRight + overflowTop + overflowBottom;
  }

  _clamp(value, min, max) {
    return Math.min(max, Math.max(min, value));
  }
}

class LhtTextFieldHelp extends HTMLElement {
  connectedCallback() {
    if (this.dataset.initialized === "true") return;
    this.dataset.initialized = "true";

    const fieldId = (this.getAttribute("field-id") || "").trim();
    if (!fieldId) return;

    const hasMdOutlinedTextField = !!(window.customElements && window.customElements.get("md-outlined-text-field"));
    const isTextarea = (this.getAttribute("type") || "").trim().toLowerCase() === "textarea" || this.hasAttribute("rows");
    const isClearable = this.hasAttribute("clearable") && !isTextarea;
    const field = hasMdOutlinedTextField
      ? document.createElement("md-outlined-text-field")
      : document.createElement(isTextarea ? "textarea" : "input");
    field.id = fieldId;
    this._isFallbackTextField = !hasMdOutlinedTextField;
    let fallbackWrapper = null;
    let fallbackSupportingText = null;
    let controlWrap = null;
    let clearButton = null;

    const label = (this.getAttribute("label") || "").trim();
    if (label) {
      if (this._isFallbackTextField) {
        field.setAttribute("aria-label", label);
      } else {
        field.setAttribute("label", label);
      }
    }

    const placeholder = this.getAttribute("placeholder");
    if (placeholder != null) field.setAttribute("placeholder", placeholder);

    const autocomplete = this.getAttribute("autocomplete");
    if (autocomplete != null) field.setAttribute("autocomplete", autocomplete);

    const type = this.getAttribute("type");
    if (isTextarea && !this._isFallbackTextField) {
      field.setAttribute("type", "textarea");
    } else if (type != null && !isTextarea) {
      field.setAttribute("type", type);
    }

    const min = this.getAttribute("min");
    if (min != null) field.setAttribute("min", min);

    const max = this.getAttribute("max");
    if (max != null) field.setAttribute("max", max);

    const step = this.getAttribute("step");
    if (step != null) field.setAttribute("step", step);

    const rows = this.getAttribute("rows");
    if (rows != null) field.setAttribute("rows", rows);

    const value = this.getAttribute("value");
    if (value != null) {
      if (this._isFallbackTextField) {
        field.value = value;
      } else {
        field.setAttribute("value", value);
      }
    }

    const fieldClass = (this.getAttribute("field-class") || "").trim();
    if (fieldClass) {
      fieldClass.split(/\s+/).filter(Boolean).forEach((name) => field.classList.add(name));
    }
    field.classList.add(this._isFallbackTextField ? "lht-text-field-help__fallback" : "md-outlined-field");
    if (isClearable) {
      field.classList.add("lht-text-field-help__field--clearable");
    }

    if (this.hasAttribute("required")) {
      field.required = true;
      field.setAttribute("required", "");
    }
    if (this.hasAttribute("readonly")) {
      if (this._isFallbackTextField) {
        field.readOnly = true;
      } else {
        field.setAttribute("readonly", "");
      }
    }
    if (this.hasAttribute("disabled")) field.disabled = true;

    if (isClearable) {
      controlWrap = document.createElement("div");
      controlWrap.className = "lht-text-field-help__control-wrap";

      clearButton = document.createElement("button");
      clearButton.type = "button";
      clearButton.className = "lht-text-field-help__clear-button";
      clearButton.hidden = true;
      clearButton.setAttribute("aria-label", `${label || fieldId}をクリア`);

      const syncClearButtonVisibility = () => {
        const currentValue = String(field.value || "");
        clearButton.hidden = currentValue.length === 0 || !!field.disabled;
      };

      clearButton.addEventListener("click", (event) => {
        event.preventDefault();
        event.stopPropagation();
        field.value = "";
        if (!this._isFallbackTextField) {
          field.setAttribute("value", "");
        }
        syncClearButtonVisibility();
        field.focus();
        field.dispatchEvent(new Event("input", { bubbles: true, composed: true }));
        field.dispatchEvent(new Event("change", { bubbles: true, composed: true }));
      });

      field.addEventListener("input", syncClearButtonVisibility);
      field.addEventListener("change", syncClearButtonVisibility);
      queueMicrotask(syncClearButtonVisibility);
    }

    const helpText = (this.getAttribute("help-text") || "").trim();
    const hideDelayMsAttr = this.getAttribute("hide-delay-ms");
    const hideDelayMsRaw = hideDelayMsAttr == null ? Number.NaN : Number(hideDelayMsAttr);
    const hideDelayMs = Number.isFinite(hideDelayMsRaw) && hideDelayMsRaw >= 0 ? hideDelayMsRaw : 160;
    if (helpText) {
      if (this._isFallbackTextField) {
        field.title = helpText;
        fallbackWrapper = document.createElement("div");
        fallbackWrapper.className = "lht-text-field-help__fallback-wrap";
        fallbackSupportingText = document.createElement("div");
        fallbackSupportingText.className = "lht-text-field-help__supporting-text";
        fallbackSupportingText.textContent = helpText;
        fallbackSupportingText.hidden = true;
        fallbackSupportingText.setAttribute("aria-hidden", "true");
        fallbackSupportingText.setAttribute("aria-live", "polite");
      } else {
      }
      let blurHideTimer = null;
      field.addEventListener("focus", () => {
        if (blurHideTimer) {
          clearTimeout(blurHideTimer);
          blurHideTimer = null;
        }
        if (this._isFallbackTextField) {
          fallbackSupportingText.hidden = false;
          fallbackSupportingText.setAttribute("aria-hidden", "false");
        } else {
          field.supportingText = helpText;
        }
      });
      field.addEventListener("blur", () => {
        if (blurHideTimer) {
          clearTimeout(blurHideTimer);
        }
        blurHideTimer = setTimeout(() => {
          if (this._isFallbackTextField) {
            fallbackSupportingText.hidden = true;
            fallbackSupportingText.setAttribute("aria-hidden", "true");
          } else {
            field.supportingText = "";
          }
          blurHideTimer = null;
        }, hideDelayMs);
      });
    }

    this.textContent = "";
    if (fallbackWrapper) {
      if (controlWrap) {
        controlWrap.appendChild(field);
        controlWrap.appendChild(clearButton);
        fallbackWrapper.appendChild(controlWrap);
      } else {
        fallbackWrapper.appendChild(field);
      }
      fallbackWrapper.appendChild(fallbackSupportingText);
      this.appendChild(fallbackWrapper);
      return;
    }
    if (controlWrap) {
      controlWrap.appendChild(field);
      controlWrap.appendChild(clearButton);
      this.appendChild(controlWrap);
      return;
    }
    this.appendChild(field);
  }
}

class LhtSelectHelp extends HTMLElement {
  connectedCallback() {
    if (this.dataset.initialized === "true") return;
    this.dataset.initialized = "true";

    const fieldId = (this.getAttribute("field-id") || "").trim();
    if (!fieldId) return;
    const hasDeclarativeOptions = this._hasDeclarativeOptions();

    const hasMdOutlinedSelect = !!(window.customElements && window.customElements.get("md-outlined-select"));
    const field = document.createElement(hasMdOutlinedSelect ? "md-outlined-select" : "select");
    field.id = fieldId;
    this._lhtField = field;
    this._isFallbackSelect = !hasMdOutlinedSelect;
    let fallbackWrapper = null;
    let fallbackSupportingText = null;

    const label = (this.getAttribute("label") || "").trim();
    if (label) {
      if (this._isFallbackSelect) {
        field.setAttribute("aria-label", label);
      } else {
        field.setAttribute("label", label);
      }
    }

    const value = this.getAttribute("value");
    if (value != null) field.value = value;

    const fieldClass = (this.getAttribute("field-class") || "").trim();
    if (fieldClass) {
      fieldClass.split(/\s+/).filter(Boolean).forEach((name) => field.classList.add(name));
    }
    if (this._isFallbackSelect) {
      field.classList.add("lht-select-help__fallback");
    } else {
      field.classList.add("md-outlined-field");
    }

    if (this.hasAttribute("required")) {
      field.required = true;
      field.setAttribute("required", "");
    }
    if (this.hasAttribute("disabled")) field.disabled = true;

    const helpText = (this.getAttribute("help-text") || "").trim();
    const hideDelayMsAttr = this.getAttribute("hide-delay-ms");
    const hideDelayMsRaw = hideDelayMsAttr == null ? Number.NaN : Number(hideDelayMsAttr);
    const hideDelayMs = Number.isFinite(hideDelayMsRaw) && hideDelayMsRaw >= 0 ? hideDelayMsRaw : 160;
    if (helpText) {
      if (this._isFallbackSelect) {
        field.title = helpText;
        fallbackWrapper = document.createElement("div");
        fallbackWrapper.className = "lht-select-help__fallback-wrap";
        fallbackSupportingText = document.createElement("div");
        fallbackSupportingText.className = "lht-select-help__supporting-text";
        fallbackSupportingText.textContent = helpText;
        fallbackSupportingText.hidden = true;
        fallbackSupportingText.setAttribute("aria-hidden", "true");
        fallbackSupportingText.setAttribute("aria-live", "polite");
      } else {
      }
      let blurHideTimer = null;
      field.addEventListener("focus", () => {
        if (blurHideTimer) {
          clearTimeout(blurHideTimer);
          blurHideTimer = null;
        }
        if (this._isFallbackSelect) {
          fallbackSupportingText.hidden = false;
          fallbackSupportingText.setAttribute("aria-hidden", "false");
        } else {
          field.supportingText = helpText;
        }
      });
      field.addEventListener("blur", () => {
        if (blurHideTimer) {
          clearTimeout(blurHideTimer);
        }
        blurHideTimer = setTimeout(() => {
          if (this._isFallbackSelect) {
            fallbackSupportingText.hidden = true;
            fallbackSupportingText.setAttribute("aria-hidden", "true");
          } else {
            field.supportingText = "";
          }
          blurHideTimer = null;
        }, hideDelayMs);
      });
    }

    if (fallbackWrapper) {
      fallbackWrapper.appendChild(field);
      fallbackWrapper.appendChild(fallbackSupportingText);
      this.appendChild(fallbackWrapper);
    } else {
      this.appendChild(field);
    }
    this.hydrateOptions();

    if (!hasDeclarativeOptions) {
      this._optionsObserver = new MutationObserver(() => {
        this.hydrateOptions();
      });
      this._optionsObserver.observe(this, { childList: true, subtree: true });
      requestAnimationFrame(() => this.hydrateOptions());
    }
  }

  disconnectedCallback() {
    if (this._optionsObserver) {
      this._optionsObserver.disconnect();
      this._optionsObserver = null;
    }
  }

  _hasDeclarativeOptions() {
    return this.hasAttribute("options") || !!this.querySelector("script[type='application/json'][slot='options']");
  }

  _normalizeOptions(rawOptions) {
    return rawOptions
      .map((entry) => {
        const value = String(entry?.value ?? entry?.label ?? "");
        const label = String(entry?.label ?? entry?.text ?? entry?.value ?? "");
        return {
          value,
          label,
          selected: !!entry?.selected,
          disabled: !!entry?.disabled
        };
      })
      .filter((entry) => entry.value || entry.label);
  }

  _readDeclarativeOptions() {
    const optionsJson = (this.getAttribute("options") || "").trim();
    if (optionsJson) {
      try {
        const parsed = JSON.parse(optionsJson);
        if (Array.isArray(parsed)) return this._normalizeOptions(parsed);
      } catch (_) {
        // JSON 不正時は次の入力ソースへフォールバック
      }
    }

    const script = this.querySelector("script[type='application/json'][slot='options']");
    if (script) {
      try {
        const parsed = JSON.parse(script.textContent || "[]");
        if (Array.isArray(parsed)) return this._normalizeOptions(parsed);
      } catch (_) {
        // JSON 不正時は空扱い
      }
      return [];
    }

    return null;
  }

  _readChildOptionElements() {
    const sourceOptions = Array.from(this.querySelectorAll("option"));
    if (sourceOptions.length === 0) return [];
    return sourceOptions.map((sourceOption) => ({
      value: sourceOption.getAttribute("value") ?? sourceOption.textContent ?? "",
      label: sourceOption.textContent ?? "",
      selected: sourceOption.hasAttribute("selected"),
      disabled: sourceOption.hasAttribute("disabled")
    }));
  }

  _setFieldOptions(options) {
    const field = this._lhtField;
    if (!field) return;
    const previousValue = field.value;
    field.innerHTML = "";

    for (const entry of options) {
      if (this._isFallbackSelect) {
        const option = document.createElement("option");
        option.value = entry.value;
        option.textContent = entry.label;
        if (entry.disabled) option.disabled = true;
        if (entry.selected) {
          option.selected = true;
          field.value = entry.value;
        }
        field.appendChild(option);
      } else {
        const option = document.createElement("md-select-option");
        option.value = entry.value;
        if (entry.disabled) option.disabled = true;
        if (entry.selected) {
          option.selected = true;
          option.setAttribute("selected", "");
          field.value = entry.value;
        }
        const headline = document.createElement("div");
        headline.slot = "headline";
        headline.textContent = entry.label;
        option.appendChild(headline);
        field.appendChild(option);
      }
    }

    if (!field.value && previousValue) {
      field.value = previousValue;
    }
  }

  hydrateOptions() {
    const declarativeOptions = this._readDeclarativeOptions();
    if (Array.isArray(declarativeOptions)) {
      this._setFieldOptions(declarativeOptions);
      const jsonScript = this.querySelector("script[type='application/json'][slot='options']");
      if (jsonScript) jsonScript.remove();
      if (this._optionsObserver) {
        this._optionsObserver.disconnect();
        this._optionsObserver = null;
      }
      return;
    }

    const optionsFromChildren = this._readChildOptionElements();
    if (optionsFromChildren.length === 0) return;
    this._setFieldOptions(optionsFromChildren);
    this.querySelectorAll("option").forEach((option) => option.remove());
    if (this._optionsObserver) {
      this._optionsObserver.disconnect();
      this._optionsObserver = null;
    }
  }

  setOptions(rawOptions, config = {}) {
    const options = this._normalizeOptions(Array.isArray(rawOptions) ? rawOptions : []);
    const preserveValue = config?.preserveValue !== false;
    const field = this._lhtField;
    const previousValue = preserveValue ? (field?.value || "") : "";

    if (field) {
      field.innerHTML = "";
    }
    this.querySelectorAll("option, script[type='application/json'][slot='options']").forEach((node) => node.remove());
    if (this._optionsObserver) {
      this._optionsObserver.disconnect();
      this._optionsObserver = null;
    }

    const nextOptions = preserveValue && previousValue
      ? options.map((entry) => ({
          ...entry,
          selected: entry.value === previousValue || entry.selected
        }))
      : options;

    this._setFieldOptions(nextOptions);
    if (field && previousValue && !nextOptions.some((entry) => entry.value === previousValue)) {
      field.value = "";
    }
  }

  getValue() {
    return this._lhtField?.value ?? "";
  }

  setValue(value) {
    if (!this._lhtField) return;
    this._lhtField.value = value == null ? "" : String(value);
  }
}

class LhtLoadingOverlay extends HTMLElement {
  static get observedAttributes() {
    return ["active", "text"];
  }

  connectedCallback() {
    if (this.dataset.initialized === "true") return;
    this.dataset.initialized = "true";

    this.setAttribute("role", "status");
    this.setAttribute("aria-live", "polite");

    const text = (this.getAttribute("text") || "Loading...").trim();

    this.textContent = "";

    const dialog = document.createElement("div");
    dialog.className = "lht-loading-overlay__dialog";

    const spinner = document.createElement("div");
    spinner.className = "lht-loading-overlay__spinner";
    spinner.setAttribute("aria-hidden", "true");

    const message = document.createElement("p");
    message.className = "lht-loading-overlay__text";
    message.textContent = text;

    dialog.appendChild(spinner);
    dialog.appendChild(message);
    this.appendChild(dialog);

    this._messageNode = message;
    this.setActive(this.hasAttribute("active"));
  }

  attributeChangedCallback(name, _oldValue, newValue) {
    if (name === "text" && this._messageNode) {
      const text = (newValue || "Loading...").trim();
      this._messageNode.textContent = text || "Loading...";
      return;
    }
    if (name === "active") {
      this.setActive(newValue !== null);
    }
  }

  isActive() {
    return this.hasAttribute("active");
  }

  setActive(inProgress) {
    const next = !!inProgress;
    this.toggleAttribute("active", next);
    this.setAttribute("aria-hidden", next ? "false" : "true");

    const busyTargetId = (this.getAttribute("busy-target-id") || "").trim();
    if (busyTargetId) {
      const target = document.getElementById(busyTargetId);
      if (target) target.setAttribute("aria-busy", next ? "true" : "false");
    }

    const disableTargetIds = (this.getAttribute("disable-target-ids") || "")
      .split(",")
      .map((id) => id.trim())
      .filter(Boolean);
    for (const id of disableTargetIds) {
      const element = document.getElementById(id);
      if (!element || !("disabled" in element)) continue;
      element.disabled = next;
    }
  }

  waitForNextPaint() {
    return new Promise((resolve) => {
      requestAnimationFrame(() => resolve());
    });
  }
}

class LhtToast extends HTMLElement {
  static get observedAttributes() {
    return ["text", "active"];
  }

  connectedCallback() {
    if (this.dataset.initialized === "true") return;
    this.dataset.initialized = "true";

    this.setAttribute("role", "status");
    this.setAttribute("aria-live", "polite");
    this.setAttribute("aria-atomic", "true");

    const initialText = (this.getAttribute("text") || this.textContent || "完了").trim();
    const text = initialText || "完了";

    this.textContent = "";

    const body = document.createElement("div");
    body.className = "lht-toast__body";
    body.textContent = text;
    this.appendChild(body);
    this._body = body;

    this.setActive(this.hasAttribute("active"));

    if (typeof window.showToast !== "function") {
      window.showToast = (message, durationMs) => {
        this.show(message, durationMs);
      };
    }
  }

  disconnectedCallback() {
    if (this._hideTimer) {
      clearTimeout(this._hideTimer);
      this._hideTimer = null;
    }
  }

  attributeChangedCallback(name, _oldValue, newValue) {
    if (name === "text") {
      if (!this._body) return;
      const text = (newValue || "").trim();
      if (text) this._body.textContent = text;
      return;
    }
    if (name === "active") {
      this.setActive(newValue !== null);
    }
  }

  show(message, durationMs) {
    if (this._hideTimer) {
      clearTimeout(this._hideTimer);
      this._hideTimer = null;
    }

    const defaultDurationMs = Number(this.getAttribute("duration-ms"));
    const fallbackDuration = Number.isFinite(defaultDurationMs) && defaultDurationMs > 0 ? defaultDurationMs : 1600;
    const nextDuration = Number(durationMs);
    const hideAfterMs = Number.isFinite(nextDuration) && nextDuration > 0 ? nextDuration : fallbackDuration;

    const text = (message || this.getAttribute("text") || this._body?.textContent || "完了").trim();
    if (this._body) this._body.textContent = text || "完了";

    this.setActive(true);

    this._hideTimer = setTimeout(() => {
      this.hide();
    }, hideAfterMs);
  }

  hide() {
    if (this._hideTimer) {
      clearTimeout(this._hideTimer);
      this._hideTimer = null;
    }
    this.setActive(false);
  }

  setActive(active) {
    const next = !!active;
    this.toggleAttribute("active", next);
    this.setAttribute("data-visible", next ? "true" : "false");
    this.setAttribute("aria-hidden", next ? "false" : "true");
  }
}

class LhtErrorAlert extends HTMLElement {
  static get observedAttributes() {
    return ["text", "active", "variant"];
  }

  connectedCallback() {
    if (this.dataset.initialized === "true") return;
    this.dataset.initialized = "true";

    this.setAttribute("aria-atomic", "true");

    const initialText = (this.getAttribute("text") || this.textContent || "").trim();

    this.textContent = "";

    const body = document.createElement("p");
    body.className = "lht-error-alert__body";
    body.textContent = initialText;
    this.appendChild(body);
    this._body = body;

    this._syncVariant();
    this.setActive(this.hasAttribute("active"));
  }

  attributeChangedCallback(name, _oldValue, newValue) {
    if (name === "text") {
      const text = (newValue || "").trim();
      if (this._body) this._body.textContent = text;
      return;
    }
    if (name === "variant") {
      this._syncVariant();
      return;
    }
    if (name === "active") {
      this.setActive(newValue !== null);
    }
  }

  isVisible() {
    return this.getAttribute("data-visible") === "true";
  }

  show(message) {
    const text = (message || this.getAttribute("text") || "").trim();
    if (this._body) this._body.textContent = text;
    this.setActive(text.length > 0);
  }

  clear() {
    if (this._body) this._body.textContent = "";
    this.hide();
  }

  hide() {
    this.setActive(false);
  }

  setActive(active) {
    const next = !!active;
    this.toggleAttribute("active", next);
    this.setAttribute("data-visible", next ? "true" : "false");
    this.setAttribute("aria-hidden", next ? "false" : "true");
  }

  _normalizeVariant(value) {
    const normalized = (value || "error").trim().toLowerCase();
    return ["error", "warning", "info"].includes(normalized) ? normalized : "error";
  }

  _syncVariant() {
    const variant = this._normalizeVariant(this.getAttribute("variant"));
    if (this.getAttribute("variant") !== variant) {
      this.setAttribute("variant", variant);
      return;
    }
    this.setAttribute("data-variant", variant);

    if (variant === "error") {
      this.setAttribute("role", "alert");
      this.setAttribute("aria-live", "assertive");
      return;
    }

    this.setAttribute("role", "status");
    this.setAttribute("aria-live", "polite");
  }
}

class LhtInputModeToggle extends HTMLElement {
  connectedCallback() {
    if (this.dataset.initialized === "true") return;
    this.dataset.initialized = "true";

    const groupLabel = (this.getAttribute("group-label") || "入力方式").trim();
    const groupName = (this.getAttribute("name") || "inputMode").trim();
    const fileId = (this.getAttribute("file-id") || "inputModeFile").trim();
    const sourceId = (this.getAttribute("source-id") || "inputModeSource").trim();
    const fileLabel = (this.getAttribute("file-label") || "ファイル読込").trim();
    const sourceLabel = (this.getAttribute("source-label") || "ソースコード入力").trim();
    const defaultMode = (this.getAttribute("default-mode") || "file").trim().toLowerCase();
    const disabled = this.hasAttribute("disabled");

    this.textContent = "";
    this.classList.add("lht-input-mode-toggle");

    const group = document.createElement("div");
    group.className = "lht-input-mode-toggle__group";
    group.setAttribute("role", "radiogroup");
    group.setAttribute("aria-label", groupLabel);

    const fileOption = this.createOption({
      id: fileId,
      name: groupName,
      label: fileLabel,
      value: "file",
      checked: defaultMode !== "source",
      disabled
    });

    const sourceOption = this.createOption({
      id: sourceId,
      name: groupName,
      label: sourceLabel,
      value: "source",
      checked: defaultMode === "source",
      disabled
    });

    group.appendChild(fileOption.label);
    group.appendChild(sourceOption.label);
    this.appendChild(group);

    this._fileRadio = fileOption.input;
    this._sourceRadio = sourceOption.input;

    const onChange = () => this.applyModeUi();
    this._fileRadio.addEventListener("change", onChange);
    this._sourceRadio.addEventListener("change", onChange);

    this.applyModeUi();
  }

  createOption({ id, name, label, value, checked, disabled }) {
    const optionLabel = document.createElement("label");
    optionLabel.className = "lht-input-mode-toggle__option";

    const input = document.createElement("input");
    input.id = id;
    input.type = "radio";
    input.name = name;
    input.value = value;
    input.checked = !!checked;
    input.disabled = !!disabled;

    const text = document.createElement("span");
    text.textContent = label;

    optionLabel.appendChild(input);
    optionLabel.appendChild(text);
    return { label: optionLabel, input };
  }

  getMode() {
    return this._sourceRadio?.checked ? "source" : "file";
  }

  setMode(mode) {
    const normalized = (mode || "").trim().toLowerCase();
    const sourceMode = normalized === "source";
    if (this._sourceRadio) this._sourceRadio.checked = sourceMode;
    if (this._fileRadio) this._fileRadio.checked = !sourceMode;
    this.applyModeUi();
  }

  applyModeUi() {
    const sourceMode = this.getMode() === "source";
    const sourceTargetId = (this.getAttribute("source-target-id") || "").trim();
    const fileTargetId = (this.getAttribute("file-target-id") || "").trim();

    if (sourceTargetId) {
      const sourceTarget = document.getElementById(sourceTargetId);
      if (sourceTarget) sourceTarget.classList.toggle("md-hidden", !sourceMode);
    }
    if (fileTargetId) {
      const fileTarget = document.getElementById(fileTargetId);
      if (fileTarget) fileTarget.classList.toggle("md-hidden", sourceMode);
    }

    const onChangeFnName = (this.getAttribute("on-change") || "").trim();
    if (onChangeFnName) {
      const fn = window[onChangeFnName];
      if (typeof fn === "function") fn(this.getMode());
    }

    this.dispatchEvent(new CustomEvent("input-mode-change", {
      detail: { mode: this.getMode() },
      bubbles: true
    }));
  }
}

class LhtPreviewOutput extends HTMLElement {
  connectedCallback() {
    if (this.dataset.initialized === "true") return;
    this.dataset.initialized = "true";

    const previewId = (this.getAttribute("preview-id") || "previewText").trim();
    const copyButtonId = (this.getAttribute("copy-button-id") || "copyBtn").trim();
    const copyTargetId = (this.getAttribute("copy-target-id") || previewId).trim();
    const placeholder = this.getAttribute("placeholder") || "未変換";
    const copyLabel = (this.getAttribute("copy-label") || "コピー").trim();
    const copyAriaLabel = (this.getAttribute("copy-aria-label") || `${copyLabel}をコピー`).trim();
    const previewTag = (this.getAttribute("preview-tag") || "div").trim().toLowerCase();
    const showCopyButton = !this.hasAttribute("no-copy");

    this.textContent = "";
    this.classList.add("lht-preview-output");

    const root = document.createElement("div");
    root.className = "lht-preview-output__root";

    const preview = document.createElement(previewTag === "pre" ? "pre" : "div");
    preview.id = previewId;
    preview.className = "lht-preview-output__preview";
    preview.textContent = placeholder;
    root.appendChild(preview);
    this._previewNode = preview;

    if (showCopyButton) {
      const copyButton = document.createElement("button");
      copyButton.type = "button";
      copyButton.id = copyButtonId;
      copyButton.className = "lht-preview-output__copy-button";
      copyButton.setAttribute("aria-label", copyAriaLabel);
      copyButton.innerHTML = '<svg aria-hidden="true" viewBox="0 0 24 24" class="lht-preview-output__copy-icon" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="9" y="9" width="11" height="11" rx="2"></rect><rect x="4" y="4" width="11" height="11" rx="2"></rect></svg>';
      copyButton.addEventListener("click", () => this.copy(copyTargetId));
      root.appendChild(copyButton);
      this._copyButton = copyButton;
    }

    this.appendChild(root);
  }

  getText() {
    return (this._previewNode?.textContent || "").trim();
  }

  setText(text) {
    if (!this._previewNode) return;
    this._previewNode.textContent = text == null ? "" : String(text);
  }

  clear() {
    if (!this._previewNode) return;
    const placeholder = this.getAttribute("placeholder") || "";
    this._previewNode.textContent = placeholder;
  }

  async copy(targetId) {
    const target = document.getElementById(targetId);
    const text = (target?.textContent || "").trim();
    if (!text) return;
    try {
      if (navigator.clipboard && typeof navigator.clipboard.writeText === "function") {
        await navigator.clipboard.writeText(text);
      } else {
        const temp = document.createElement("textarea");
        temp.value = text;
        document.body.appendChild(temp);
        temp.select();
        document.execCommand("copy");
        document.body.removeChild(temp);
      }
      if (typeof window.showToast === "function") {
        window.showToast("コピーしました");
      }
    } catch (_) {
      // コピー不可環境では失敗を握りつぶす
    }
  }
}

class LhtFileSelect extends HTMLElement {
  connectedCallback() {
    if (this.dataset.initialized === "true") return;
    this.dataset.initialized = "true";

    const inputId = (this.getAttribute("input-id") || "fileInput").trim();
    const buttonId = (this.getAttribute("button-id") || "fileSelectBtn").trim();
    const fileNameId = (this.getAttribute("file-name-id") || "fileNameText").trim();
    const accept = (this.getAttribute("accept") || "").trim();
    const buttonLabel = (this.getAttribute("button-label") || "ファイルを選択").trim();
    const placeholder = (this.getAttribute("placeholder") || "未選択").trim();
    const showFileName = this.hasAttribute("show-file-name");
    const autoOpenValue = (this.getAttribute("auto-open") || "").trim().toLowerCase();
    const autoOpen = autoOpenValue !== "false";

    this.textContent = "";

    const root = document.createElement("div");
    root.className = "lht-file-select";

    const hasMdFilledButton = !!(window.customElements && window.customElements.get("md-filled-button"));
    const triggerButton = document.createElement(hasMdFilledButton ? "md-filled-button" : "button");
    if (!hasMdFilledButton) {
      triggerButton.type = "button";
    }
    triggerButton.id = buttonId;
    triggerButton.className = `lht-file-select__button${hasMdFilledButton ? "" : " lht-file-select__button--fallback"}`;

    const icon = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    icon.setAttribute("slot", "icon");
    icon.setAttribute("aria-hidden", "true");
    icon.setAttribute("viewBox", "0 0 24 24");
    icon.setAttribute("class", "lht-file-select__button-icon");
    icon.setAttribute("fill", "none");
    icon.setAttribute("stroke", "currentColor");
    icon.setAttribute("stroke-width", "1.9");
    icon.setAttribute("stroke-linecap", "round");
    icon.setAttribute("stroke-linejoin", "round");
    icon.innerHTML = '<path d="M4 7h7l2 2h7v8a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2z"></path><path d="M12 10v6"></path><path d="M9 13l3 3 3-3"></path>';

    const labelNode = document.createElement("span");
    labelNode.className = "lht-file-select__button-text";
    labelNode.textContent = buttonLabel;
    triggerButton.appendChild(icon);
    triggerButton.appendChild(labelNode);

    const input = document.createElement("input");
    input.id = inputId;
    input.type = "file";
    input.className = "md-file";
    input.hidden = true;
    if (accept) input.setAttribute("accept", accept);
    if (this.hasAttribute("multiple")) input.multiple = true;

    const fileName = document.createElement("span");
    fileName.id = fileNameId;
    fileName.className = "lht-file-select__file-name";
    fileName.textContent = placeholder;
    if (!showFileName) fileName.hidden = true;

    if (this.hasAttribute("disabled")) {
      input.disabled = true;
      triggerButton.disabled = true;
    }

    triggerButton.addEventListener("click", () => {
      const beforeOpenEvent = new CustomEvent("lht-file-select:before-open", {
        detail: {
          inputId,
          buttonId,
          input,
          triggerButton,
          autoOpen
        },
        bubbles: true,
        cancelable: true
      });
      const canAutoOpen = this.dispatchEvent(beforeOpenEvent);
      if (autoOpen && canAutoOpen) {
        input.click();
      }
    });
    input.addEventListener("change", () => {
      const names = Array.from(input.files || []).map((file) => file.name).filter(Boolean);
      fileName.textContent = names.length > 0 ? names.join(", ") : placeholder;
      this.dispatchEvent(new CustomEvent("lht-file-select:change", {
        detail: {
          files: Array.from(input.files || []),
          names,
          input,
          fileName
        },
        bubbles: true
      }));
    });

    root.appendChild(triggerButton);
    root.appendChild(fileName);
    this.appendChild(root);
    this.appendChild(input);
  }
}

class LhtSwitchHelp extends HTMLElement {
  connectedCallback() {
    if (this.dataset.initialized === "true") return;
    this.dataset.initialized = "true";

    const switchId = (this.getAttribute("switch-id") || "").trim();
    if (!switchId) return;
    const labelText = (this.getAttribute("label") || "").trim();
    const helpLabel = (this.getAttribute("help-label") || `${labelText}の説明`).trim();
    const helpContentHtml = this.innerHTML.trim();
    const onChangeFnName = (this.getAttribute("on-change") || "").trim();
    const isChecked = this.hasAttribute("checked");
    const isHelpWide = this.hasAttribute("help-wide");

    this.textContent = "";

    const label = document.createElement("label");
    label.className = "md-switch-label";

    const hasMdSwitch = !!(window.customElements && window.customElements.get("md-switch"));
    const switchControl = hasMdSwitch
      ? this._createMaterialSwitch(switchId, isChecked)
      : this._createFallbackSwitch(switchId, isChecked);

    if (onChangeFnName) {
      switchControl.control.addEventListener("change", () => {
        const fn = window[onChangeFnName];
        if (typeof fn === "function") {
          fn();
        }
      });
    }
    label.appendChild(switchControl.node);

    const labelSpan = document.createElement("span");
    labelSpan.textContent = labelText;
    label.appendChild(labelSpan);

    if (helpContentHtml) {
      const help = document.createElement("lht-help-tooltip");
      help.setAttribute("label", helpLabel);
      if (isHelpWide) {
        help.setAttribute("wide", "");
      }
      help.innerHTML = helpContentHtml;
      label.appendChild(help);
    }

    this.appendChild(label);
  }

  _createMaterialSwitch(switchId, isChecked) {
    const mdSwitch = document.createElement("md-switch");
    mdSwitch.id = switchId;
    Object.defineProperty(mdSwitch, "checked", {
      get() {
        return !!mdSwitch.selected;
      },
      set(value) {
        mdSwitch.selected = !!value;
      }
    });
    if (isChecked) {
      mdSwitch.selected = true;
      mdSwitch.setAttribute("selected", "");
    }
    return { node: mdSwitch, control: mdSwitch };
  }

  _createFallbackSwitch(switchId, isChecked) {
    const input = document.createElement("input");
    input.id = switchId;
    input.type = "checkbox";
    input.className = "md-switch-input";
    input.checked = isChecked;
    input.setAttribute("role", "switch");
    input.setAttribute("aria-checked", isChecked ? "true" : "false");

    input.addEventListener("change", () => {
      input.setAttribute("aria-checked", input.checked ? "true" : "false");
    });

    const visual = document.createElement("span");
    visual.className = "md-switch";
    visual.setAttribute("aria-hidden", "true");

    const fragment = document.createDocumentFragment();
    fragment.appendChild(input);
    fragment.appendChild(visual);
    return { node: fragment, control: input };
  }
}

class LhtCommandBlock extends HTMLElement {
  connectedCallback() {
    if (this.dataset.initialized === "true") return;
    this.dataset.initialized = "true";

    const commandId = (this.getAttribute("command-id") || "").trim();
    if (!commandId) return;
    const copyButtons = (this.getAttribute("copy-buttons") || "single").trim().toLowerCase();
    const isDual = copyButtons === "dual";

    this.textContent = "";

    const block = document.createElement("div");
    block.className = "md-code-block";

    const code = document.createElement("code");
    code.id = commandId;
    code.className = `md-code${isDual ? " md-code--dual-copy" : ""}`;
    block.appendChild(code);

    const topCopyButton = this.createCopyButton("コピー", () => this.copyFromCommand(commandId));
    block.appendChild(topCopyButton);

    if (isDual) {
      const bottomCopyButton = this.createCopyButton("コピー（右下）", () => this.copyFromCommand(commandId));
      bottomCopyButton.classList.add("md-copy-button--bottom-right");
      block.appendChild(bottomCopyButton);
    }

    this.appendChild(block);
  }

  createCopyButton(label, onClick) {
    const hasMdIconButton = !!(window.customElements && window.customElements.get("md-icon-button"));
    const button = document.createElement(hasMdIconButton ? "md-icon-button" : "button");
    button.className = `md-copy-button md-copy-button--surface${hasMdIconButton ? "" : " md-copy-button--fallback"}`;
    if (!hasMdIconButton) {
      button.type = "button";
    }
    button.setAttribute("aria-label", label);
    button.innerHTML = '<svg viewBox="0 0 24 24" class="md-icon-small" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><use href="#md-icon-copy" xlink:href="#md-icon-copy"></use></svg>';
    button.addEventListener("click", onClick);
    return button;
  }

  async copyFromCommand(commandId) {
    const commandElement = document.getElementById(commandId);
    if (!commandElement) return;
    const text = (commandElement.textContent || "").trim();
    if (!text) return;
    try {
      if (navigator.clipboard && typeof navigator.clipboard.writeText === "function") {
        await navigator.clipboard.writeText(text);
      } else {
        const temp = document.createElement("textarea");
        temp.value = text;
        document.body.appendChild(temp);
        temp.select();
        document.execCommand("copy");
        document.body.removeChild(temp);
      }
      if (typeof window.showToast === "function") {
        window.showToast("コピーしました");
      }
    } catch (_) {
      // Clipboard API 利用不可環境では失敗を無視
    }
  }
}

class LhtIndexCardLink extends HTMLElement {
  connectedCallback() {
    if (this.dataset.initialized === "true") return;
    this.dataset.initialized = "true";

    const href = (this.getAttribute("href") || "").trim();
    if (!href) return;

    const title = (this.getAttribute("title") || "").trim();
    const descAttr = (this.getAttribute("desc") || "").trim();
    const iconAttr = (this.getAttribute("icon") || "").trim();
    const target = (this.getAttribute("target") || "").trim();
    const relAttr = (this.getAttribute("rel") || "").trim();
    const variant = (this.getAttribute("variant") || "default").trim().toLowerCase();
    const arrowMode = (this.getAttribute("arrow") || "auto").trim().toLowerCase();
    const badgeText = (this.getAttribute("badge") || "").trim();
    const descLines = (this.getAttribute("desc-lines") || "").trim();

    if (!title || !descAttr) {
      const missing = [];
      if (!title) missing.push("title");
      if (!descAttr) missing.push("desc");
      // Fail fast for authoring mistakes in index cards.
      console.warn(`[lht-index-card-link] Missing required attribute(s): ${missing.join(", ")}`, this);
      return;
    }

    this.textContent = "";

    const link = document.createElement("a");
    link.href = href;
    link.className = "md-link-card";
    const isExternalHref = /^(https?:)?\/\//i.test(href);
    const isExternal = variant === "external" || isExternalHref || target === "_blank";
    const effectiveTarget = target || (isExternal ? "_blank" : "");
    if (effectiveTarget) link.target = effectiveTarget;
    if (effectiveTarget === "_blank") {
      link.rel = relAttr || "noopener noreferrer";
    } else if (relAttr) {
      link.rel = relAttr;
    }
    if (variant === "simple") link.classList.add("lht-index-card-link--simple");
    if (isExternal) link.classList.add("lht-index-card-link--external");

    const head = document.createElement("div");
    head.className = "md-card-head";

    const h3 = document.createElement("h3");
    h3.className = "md-card-title";
    if (iconAttr) {
      const iconContainer = document.createElement("span");
      iconContainer.className = "lht-index-card-link__icon";
      iconContainer.textContent = iconAttr;
      h3.appendChild(iconContainer);
    }
    const titleContainer = document.createElement("span");
    titleContainer.className = "lht-index-card-link__title";
    titleContainer.textContent = title;
    h3.appendChild(titleContainer);
    if (badgeText) {
      const badge = document.createElement("span");
      badge.className = "lht-index-card-link__badge";
      badge.textContent = badgeText;
      h3.appendChild(badge);
    }

    const arrow = document.createElement("span");
    arrow.className = "md-card-arrow";
    const showArrow = arrowMode === "auto" ? variant !== "simple" : arrowMode !== "none";
    arrow.textContent = isExternal ? "↗" : "→";
    if (!showArrow) arrow.hidden = true;

    const desc = document.createElement("p");
    desc.className = "md-card-desc";
    desc.textContent = descAttr;
    if (descLines && /^\d+$/.test(descLines)) {
      desc.classList.add("lht-index-card-link__desc--clamp");
      desc.style.setProperty("--lht-desc-lines", descLines);
    }

    head.appendChild(h3);
    head.appendChild(arrow);
    link.appendChild(head);
    link.appendChild(desc);
    this.appendChild(link);
  }
}

class LhtPageHero extends HTMLElement {
  connectedCallback() {
    if (this.dataset.initialized === "true") return;
    this.dataset.initialized = "true";

    const title = (this.getAttribute("title") || "").trim();
    if (!title) return;
    const subtitle = (this.getAttribute("subtitle") || "").trim();
    const icon = (this.getAttribute("icon") || "").trim();
    const helpLabel = (this.getAttribute("help-label") || "説明").trim();
    const homeHref = (this.getAttribute("menu-home-href") || "../index.html").trim();
    const homeLabel = (this.getAttribute("menu-home-label") || "トップへ戻る").trim();
    const useWideHelp = this.hasAttribute("help-wide");
    const showMenu = !this.hasAttribute("no-menu");
    const helpHtml = this.innerHTML.trim();
    const actionHref = (this.getAttribute("action-href") || "").trim();
    const actionLabel = (this.getAttribute("action-label") || "").trim();
    const actionAriaLabel = (this.getAttribute("action-aria-label") || actionLabel).trim();
    const actionIconId = (this.getAttribute("action-icon-id") || "").trim();

    this.textContent = "";
    this.classList.add("lht-page-hero");

    const topRow = document.createElement("div");
    topRow.className = "lht-page-hero__title-row";

    const titleMain = document.createElement("span");
    titleMain.className = "lht-page-hero__title-main";

    const heading = document.createElement("h1");
    heading.className = "lht-page-hero__title";
    if (icon) {
      const iconNode = document.createElement("span");
      iconNode.className = "lht-page-hero__icon";
      iconNode.setAttribute("aria-hidden", "true");
      iconNode.textContent = icon;
      heading.appendChild(iconNode);
    }
    const titleNode = document.createElement("span");
    titleNode.textContent = title;
    heading.appendChild(titleNode);
    titleMain.appendChild(heading);

    if (helpHtml) {
      const help = document.createElement("lht-help-tooltip");
      help.setAttribute("label", helpLabel);
      if (useWideHelp) {
        help.setAttribute("wide", "");
      }
      help.innerHTML = helpHtml;
      titleMain.appendChild(help);
    }

    if (actionHref && (actionLabel || actionIconId)) {
      const actionLink = document.createElement("a");
      actionLink.href = actionHref;
      actionLink.className = "ms-hero-link";
      if (!actionLabel) {
        actionLink.classList.add("ms-hero-link--icon-only");
      }
      actionLink.target = "_blank";
      actionLink.rel = "noopener noreferrer";
      if (actionAriaLabel) {
        actionLink.setAttribute("aria-label", actionAriaLabel);
      }
      if (actionIconId) {
        const iconSvg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
        iconSvg.setAttribute("aria-hidden", "true");
        iconSvg.setAttribute("viewBox", "0 0 16 16");
        iconSvg.setAttribute("class", "ms-btn-icon");
        iconSvg.setAttribute("fill", "currentColor");
        const useNode = document.createElementNS("http://www.w3.org/2000/svg", "use");
        useNode.setAttribute("href", `#${actionIconId}`);
        useNode.setAttributeNS("http://www.w3.org/1999/xlink", "xlink:href", `#${actionIconId}`);
        iconSvg.appendChild(useNode);
        actionLink.appendChild(iconSvg);
      }
      if (actionLabel) {
        const labelNode = document.createElement("span");
        labelNode.textContent = actionLabel;
        actionLink.appendChild(labelNode);
      }
      titleMain.appendChild(actionLink);
    }

    topRow.appendChild(titleMain);

    if (showMenu) {
      const actions = document.createElement("span");
      actions.className = "lht-page-hero__actions";
      const menu = document.createElement("lht-page-menu");
      menu.setAttribute("home-href", homeHref);
      menu.setAttribute("home-label", homeLabel);
      actions.appendChild(menu);
      topRow.appendChild(actions);
    }

    this.appendChild(topRow);

    if (subtitle) {
      const subtitleNode = document.createElement("div");
      subtitleNode.className = "lht-page-hero__subtitle";
      subtitleNode.textContent = subtitle;
      this.appendChild(subtitleNode);
    }
  }
}

class LhtPageMenu extends HTMLElement {
  connectedCallback() {
    if (this.dataset.initialized === "true") return;
    this.dataset.initialized = "true";

    const homeHref = (this.getAttribute("home-href") || "../index.html").trim();
    const homeLabel = (this.getAttribute("home-label") || "トップへ戻る").trim();

    this.textContent = "";
    this.classList.add("lht-page-menu");

    const button = document.createElement("button");
    button.type = "button";
    button.className = "md-menu-button md-icon-btn";
    button.setAttribute("aria-label", "メニュー");
    button.innerHTML = '<svg aria-hidden="true" viewBox="0 0 24 24" class="md-icon-20" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M4 6h16"></path><path d="M4 12h16"></path><path d="M4 18h16"></path></svg>';

    const panel = document.createElement("div");
    panel.className = "md-menu-panel md-hidden";

    const link = document.createElement("a");
    link.className = "md-menu-link";
    link.href = homeHref;
    link.textContent = homeLabel;
    panel.appendChild(link);

    button.addEventListener("click", () => {
      panel.classList.toggle("md-hidden");
    });

    document.addEventListener("pointerdown", (event) => {
      if (!this.contains(event.target)) {
        panel.classList.add("md-hidden");
      }
    });

    this.appendChild(button);
    this.appendChild(panel);
  }
}

if (!customElements.get("lht-help-tooltip")) {
  customElements.define("lht-help-tooltip", LhtHelpTooltip);
}
if (!customElements.get("lht-text-field-help")) {
  customElements.define("lht-text-field-help", LhtTextFieldHelp);
}
if (!customElements.get("lht-select-help")) {
  customElements.define("lht-select-help", LhtSelectHelp);
}
if (!customElements.get("lht-file-select")) {
  customElements.define("lht-file-select", LhtFileSelect);
}
if (!customElements.get("lht-loading-overlay")) {
  customElements.define("lht-loading-overlay", LhtLoadingOverlay);
}
if (!customElements.get("lht-toast")) {
  customElements.define("lht-toast", LhtToast);
}
if (!customElements.get("lht-error-alert")) {
  customElements.define("lht-error-alert", LhtErrorAlert);
}
if (!customElements.get("lht-input-mode-toggle")) {
  customElements.define("lht-input-mode-toggle", LhtInputModeToggle);
}
if (!customElements.get("lht-preview-output")) {
  customElements.define("lht-preview-output", LhtPreviewOutput);
}
if (!customElements.get("lht-switch-help")) {
  customElements.define("lht-switch-help", LhtSwitchHelp);
}
if (!customElements.get("lht-command-block")) {
  customElements.define("lht-command-block", LhtCommandBlock);
}
if (!customElements.get("lht-index-card-link")) {
  customElements.define("lht-index-card-link", LhtIndexCardLink);
}
if (!customElements.get("lht-page-hero")) {
  customElements.define("lht-page-hero", LhtPageHero);
}
if (!customElements.get("lht-page-menu")) {
  customElements.define("lht-page-menu", LhtPageMenu);
}

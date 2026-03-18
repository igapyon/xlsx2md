// @vitest-environment jsdom

import { readFileSync } from "node:fs";
import { resolve } from "node:path";

import { describe, expect, it, vi } from "vitest";

import "./js/components.js";

const componentsCss = readFileSync(resolve(process.cwd(), "lht-cmn/css/components.css"), "utf8");

function waitForMicrotask() {
  return new Promise((resolve) => queueMicrotask(resolve));
}

function defineMaterialTestDoubles() {
  const definitions = {
    "md-icon-button": class extends HTMLElement {},
    "md-filled-button": class extends HTMLElement {},
    "md-switch": class extends HTMLElement {},
    "md-outlined-text-field": class extends HTMLElement {},
    "md-outlined-select": class extends HTMLElement {},
    "md-select-option": class extends HTMLElement {}
  };

  for (const [tagName, ctor] of Object.entries(definitions)) {
    if (!customElements.get(tagName)) {
      customElements.define(tagName, ctor);
    }
  }
}

function renderCriticalComponentsMarkup() {
  document.body.innerHTML = `
    <lht-help-tooltip label="help">body</lht-help-tooltip>
    <lht-text-field-help field-id="textField" label="Text"></lht-text-field-help>
    <lht-select-help field-id="selectField"></lht-select-help>
    <lht-file-select input-id="fileInput" button-id="fileButton"></lht-file-select>
    <lht-switch-help switch-id="switchField" label="Switch">help</lht-switch-help>
    <lht-command-block command-id="cmdField"></lht-command-block>
  `;
}

function mountTooltipStyles() {
  document.head.innerHTML = `<style>${componentsCss}</style>`;
}

describe("lht-select-help declarative options", () => {
  it("does not enable child observer after consuming declarative JSON script", async () => {
    document.body.innerHTML = `
      <lht-select-help field-id="test-select">
        <script type="application/json" slot="options">[
          {"value":"a","label":"Alpha","selected":true}
        ]</script>
      </lht-select-help>
    `;

    const element = document.querySelector("lht-select-help");
    const field = element.querySelector("select");

    expect(field).not.toBeNull();
    expect(field.options).toHaveLength(1);
    expect(field.options[0].value).toBe("a");
    expect(field.options[0].textContent).toBe("Alpha");
    expect(element.querySelector("script[slot='options']")).toBeNull();
    expect(element._optionsObserver ?? null).toBeNull();

    const lateOption = document.createElement("option");
    lateOption.value = "b";
    lateOption.textContent = "Beta";
    element.appendChild(lateOption);
    await waitForMicrotask();

    expect(field.options).toHaveLength(1);
    expect(field.options[0].value).toBe("a");
  });

  it("supports setOptions and preserves selected value by default", () => {
    document.body.innerHTML = `
      <lht-select-help field-id="dynamic-select"></lht-select-help>
    `;

    const element = document.querySelector("lht-select-help");
    const field = element.querySelector("select");

    element.setOptions([
      { value: "a", label: "Alpha" },
      { value: "b", label: "Beta", selected: true }
    ]);
    expect(field.value).toBe("b");

    element.setOptions([
      { value: "b", label: "Beta 2" },
      { value: "c", label: "Gamma" }
    ]);

    expect(field.value).toBe("b");
    expect(field.options).toHaveLength(2);
    expect(field.options[0].textContent).toBe("Beta 2");
  });

  it("clears selected value when preserveValue is disabled or missing from next options", () => {
    document.body.innerHTML = `
      <lht-select-help field-id="dynamic-select-2"></lht-select-help>
    `;

    const element = document.querySelector("lht-select-help");
    const field = element.querySelector("select");

    element.setOptions([
      { value: "a", label: "Alpha", selected: true },
      { value: "b", label: "Beta" }
    ]);
    expect(field.value).toBe("a");

    element.setOptions([
      { value: "c", label: "Gamma" }
    ], { preserveValue: false });
    expect(field.value).toBe("c");

    element.setValue("c");
    expect(element.getValue()).toBe("c");

    element.setOptions([
      { value: "x", label: "Ex" }
    ]);
    expect(field.value).toBe("");
  });

  it("renders fallback supporting text and keeps title for native select", () => {
    document.body.innerHTML = `
      <lht-select-help
        field-id="test-select"
        label="Choice"
        help-text="Select one option"
      >
        <script type="application/json" slot="options">[
          {"value":"a","label":"Alpha","selected":true}
        ]</script>
      </lht-select-help>
    `;

    const field = document.querySelector("lht-select-help select");
    const supportingText = document.querySelector("lht-select-help .lht-select-help__supporting-text");

    expect(field).not.toBeNull();
    expect(field.title).toBe("Select one option");
    expect(field.getAttribute("aria-label")).toBe("Choice");
    expect(supportingText).not.toBeNull();
    expect(supportingText.textContent).toBe("Select one option");
    expect(supportingText.hidden).toBe(true);
  });

  it("shows fallback select help text on focus and hides it after blur delay", () => {
    vi.useFakeTimers();
    document.body.innerHTML = `
      <lht-select-help
        field-id="test-select"
        label="Choice"
        help-text="Select one option"
        hide-delay-ms="220"
      >
        <script type="application/json" slot="options">[
          {"value":"a","label":"Alpha","selected":true}
        ]</script>
      </lht-select-help>
    `;

    const field = document.querySelector("lht-select-help select");
    const supportingText = document.querySelector("lht-select-help .lht-select-help__supporting-text");

    field.dispatchEvent(new Event("focus"));
    expect(supportingText.hidden).toBe(false);
    expect(supportingText.getAttribute("aria-hidden")).toBe("false");

    field.dispatchEvent(new Event("blur"));
    vi.advanceTimersByTime(219);
    expect(supportingText.hidden).toBe(false);

    vi.advanceTimersByTime(1);
    expect(supportingText.hidden).toBe(true);
    expect(supportingText.getAttribute("aria-hidden")).toBe("true");
    vi.useRealTimers();
  });

  it("uses the default 160ms hide delay for fallback select when hide-delay-ms is omitted", () => {
    vi.useFakeTimers();
    document.body.innerHTML = `
      <lht-select-help
        field-id="test-select"
        label="Choice"
        help-text="Select one option"
      >
        <script type="application/json" slot="options">[
          {"value":"a","label":"Alpha","selected":true}
        ]</script>
      </lht-select-help>
    `;

    const field = document.querySelector("lht-select-help select");
    const supportingText = document.querySelector("lht-select-help .lht-select-help__supporting-text");

    field.dispatchEvent(new Event("focus"));
    field.dispatchEvent(new Event("blur"));

    vi.advanceTimersByTime(159);
    expect(supportingText.hidden).toBe(false);

    vi.advanceTimersByTime(1);
    expect(supportingText.hidden).toBe(true);
    vi.useRealTimers();
  });
});

describe("lht-help-tooltip fallback", () => {
  it("renders a native button when md-icon-button is unavailable", () => {
    document.body.innerHTML = `
      <lht-help-tooltip label="説明ラベル">
        <strong>help</strong>
      </lht-help-tooltip>
    `;

    const button = document.querySelector("lht-help-tooltip .md-help-icon-button--fallback");
    const tooltip = document.querySelector("lht-help-tooltip .md-tooltip-content");

    expect(button).not.toBeNull();
    expect(button.tagName).toBe("BUTTON");
    expect(button.getAttribute("aria-label")).toBe("説明ラベル");
    expect(tooltip).not.toBeNull();
    expect(tooltip.innerHTML).toContain("<strong>help</strong>");
  });

  it("bundles the self-contained CSS contract for anchoring and visibility", () => {
    expect(componentsCss).toContain("lht-help-tooltip {");
    expect(componentsCss).toContain("position: relative;");
    expect(componentsCss).toContain("overflow: visible;");
    expect(componentsCss).toContain("lht-help-tooltip .md-tooltip-group {");
    expect(componentsCss).toContain("lht-help-tooltip .md-tooltip-content {");
    expect(componentsCss).toContain("display: none;");
    expect(componentsCss).toContain("lht-help-tooltip .md-tooltip-group:hover .md-tooltip-content,");
    expect(componentsCss).toContain("lht-help-tooltip .md-tooltip-group:focus-within .md-tooltip-content {");
    expect(componentsCss).toContain("display: block;");
  });

  it("supports placement auto and clamps to the lower-overflow side", () => {
    document.body.innerHTML = `
      <lht-help-tooltip label="説明ラベル" placement="auto">
        help
      </lht-help-tooltip>
    `;

    const element = document.querySelector("lht-help-tooltip");
    const group = element.querySelector(".md-tooltip-group");
    const tooltip = element.querySelector(".md-tooltip-content");

    Object.defineProperty(window, "innerWidth", { configurable: true, value: 320 });
    Object.defineProperty(window, "innerHeight", { configurable: true, value: 640 });
    group.getBoundingClientRect = () => ({
      left: 260,
      right: 290,
      top: 100,
      bottom: 130,
      width: 30,
      height: 30
    });
    tooltip.getBoundingClientRect = () => ({
      width: 120,
      height: 60
    });

    element._applyTooltipPlacement();

    expect(tooltip.dataset.placement).toBe("right");
    expect(tooltip.style.left).toBe("-76px");
    expect(tooltip.style.top).toBe("-15px");
  });

  it("hides by default, becomes visible while active, and resets on leave", () => {
    mountTooltipStyles();
    document.body.innerHTML = `
      <lht-help-tooltip label="説明ラベル">
        help
      </lht-help-tooltip>
    `;

    const element = document.querySelector("lht-help-tooltip");
    const tooltip = element.querySelector(".md-tooltip-content");

    expect(window.getComputedStyle(tooltip).display).toBe("none");

    element._handleTooltipEnter();
    expect(window.getComputedStyle(tooltip).display).toBe("block");
    expect(tooltip.style.left).not.toBe("");
    expect(tooltip.style.top).not.toBe("");

    element._handleTooltipLeave();
    expect(window.getComputedStyle(tooltip).display).toBe("none");
    expect(tooltip.style.left).toBe("");
    expect(tooltip.style.top).toBe("");
  });

  it("supports Escape to force-hide the active tooltip", () => {
    document.body.innerHTML = `
      <lht-help-tooltip label="説明ラベル">
        help
      </lht-help-tooltip>
    `;

    const element = document.querySelector("lht-help-tooltip");
    const group = element.querySelector(".md-tooltip-group");
    const button = group.querySelector(".md-help-icon-button");

    element._handleTooltipEnter();
    button.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));

    expect(group.getAttribute("data-force-hidden")).toBe("true");
    expect(element._activeTooltip).toBe(false);

    element._handleTooltipEnter();
    expect(group.hasAttribute("data-force-hidden")).toBe(false);
  });
});

describe("lht-text-field-help fallback", () => {
  it("renders native input fallback when md-outlined-text-field is unavailable", () => {
    document.body.innerHTML = `
      <lht-text-field-help
        field-id="nameField"
        label="Name"
        placeholder="Your name"
        value="Alice"
        help-text="Enter your name"
      ></lht-text-field-help>
    `;

    const field = document.querySelector("lht-text-field-help input");
    const supportingText = document.querySelector("lht-text-field-help .lht-text-field-help__supporting-text");

    expect(field).not.toBeNull();
    expect(field.id).toBe("nameField");
    expect(field.value).toBe("Alice");
    expect(field.getAttribute("aria-label")).toBe("Name");
    expect(field.title).toBe("Enter your name");
    expect(supportingText).not.toBeNull();
    expect(supportingText.textContent).toBe("Enter your name");
    expect(supportingText.hidden).toBe(true);
  });

  it("renders native textarea fallback when rows is specified", () => {
    document.body.innerHTML = `
      <lht-text-field-help
        field-id="memoField"
        label="Memo"
        rows="4"
        value="hello"
      ></lht-text-field-help>
    `;

    const field = document.querySelector("lht-text-field-help textarea");

    expect(field).not.toBeNull();
    expect(field.id).toBe("memoField");
    expect(field.getAttribute("rows")).toBe("4");
    expect(field.value).toBe("hello");
  });

  it("shows help text on focus and hides it after blur delay", () => {
    vi.useFakeTimers();
    document.body.innerHTML = `
      <lht-text-field-help
        field-id="nameField"
        label="Name"
        help-text="Enter your name"
        hide-delay-ms="240"
      ></lht-text-field-help>
    `;

    const field = document.querySelector("lht-text-field-help input");
    const supportingText = document.querySelector("lht-text-field-help .lht-text-field-help__supporting-text");

    field.dispatchEvent(new Event("focus"));
    expect(supportingText.hidden).toBe(false);
    expect(supportingText.getAttribute("aria-hidden")).toBe("false");

    field.dispatchEvent(new Event("blur"));
    vi.advanceTimersByTime(239);
    expect(supportingText.hidden).toBe(false);

    vi.advanceTimersByTime(1);
    expect(supportingText.hidden).toBe(true);
    expect(supportingText.getAttribute("aria-hidden")).toBe("true");
    vi.useRealTimers();
  });

  it("uses the default 160ms hide delay when hide-delay-ms is omitted", () => {
    vi.useFakeTimers();
    document.body.innerHTML = `
      <lht-text-field-help
        field-id="nameField"
        label="Name"
        help-text="Enter your name"
      ></lht-text-field-help>
    `;

    const field = document.querySelector("lht-text-field-help input");
    const supportingText = document.querySelector("lht-text-field-help .lht-text-field-help__supporting-text");

    field.dispatchEvent(new Event("focus"));
    field.dispatchEvent(new Event("blur"));

    vi.advanceTimersByTime(159);
    expect(supportingText.hidden).toBe(false);

    vi.advanceTimersByTime(1);
    expect(supportingText.hidden).toBe(true);
    vi.useRealTimers();
  });

  it("supports clearable fallback text fields and dispatches input/change when cleared", async () => {
    document.body.innerHTML = `
      <lht-text-field-help
        field-id="nameField"
        label="Name"
        value="Alice"
        clearable
      ></lht-text-field-help>
    `;

    const field = document.querySelector("lht-text-field-help input");
    const clearButton = document.querySelector("lht-text-field-help .lht-text-field-help__clear-button");
    const inputListener = vi.fn();
    const changeListener = vi.fn();

    field.addEventListener("input", inputListener);
    field.addEventListener("change", changeListener);

    await waitForMicrotask();

    expect(clearButton).not.toBeNull();
    expect(clearButton.hidden).toBe(false);

    clearButton.click();

    expect(field.value).toBe("");
    expect(clearButton.hidden).toBe(true);
    expect(inputListener).toHaveBeenCalledTimes(1);
    expect(changeListener).toHaveBeenCalledTimes(1);
  });
});

describe("lht-file-select events", () => {
  it("dispatches before-open and auto-clicks input by default", () => {
    document.body.innerHTML = `
      <lht-file-select input-id="fileInput" button-id="fileSelectBtn"></lht-file-select>
    `;

    const element = document.querySelector("lht-file-select");
    const button = document.getElementById("fileSelectBtn");
    const input = document.getElementById("fileInput");
    const beforeOpen = vi.fn();
    const clickSpy = vi.fn();

    element.addEventListener("lht-file-select:before-open", beforeOpen);
    input.click = clickSpy;

    button.click();

    expect(beforeOpen).toHaveBeenCalledTimes(1);
    expect(clickSpy).toHaveBeenCalledTimes(1);
    expect(beforeOpen.mock.calls[0][0].detail.autoOpen).toBe(true);
  });

  it("supports host-owned open flow via auto-open=false and emits change event", () => {
    document.body.innerHTML = `
      <lht-file-select
        input-id="fileInput"
        button-id="fileSelectBtn"
        file-name-id="fileNameText"
        show-file-name
        auto-open="false"
      ></lht-file-select>
    `;

    const element = document.querySelector("lht-file-select");
    const button = document.getElementById("fileSelectBtn");
    const input = document.getElementById("fileInput");
    const fileName = document.getElementById("fileNameText");
    const beforeOpen = vi.fn((event) => {
      expect(event.detail.autoOpen).toBe(false);
    });
    const changeListener = vi.fn();
    const clickSpy = vi.fn();

    element.addEventListener("lht-file-select:before-open", beforeOpen);
    element.addEventListener("lht-file-select:change", changeListener);
    input.click = clickSpy;

    button.click();

    expect(beforeOpen).toHaveBeenCalledTimes(1);
    expect(clickSpy).not.toHaveBeenCalled();

    Object.defineProperty(input, "files", {
      configurable: true,
      value: [{ name: "score.musicxml" }]
    });
    input.dispatchEvent(new Event("change", { bubbles: true }));

    expect(fileName.textContent).toBe("score.musicxml");
    expect(changeListener).toHaveBeenCalledTimes(1);
    expect(changeListener.mock.calls[0][0].detail.names).toEqual(["score.musicxml"]);
  });
});

describe("lht-error-alert variants", () => {
  it("uses alert/assertive for error variant by default", () => {
    document.body.innerHTML = `<lht-error-alert text="boom"></lht-error-alert>`;

    const element = document.querySelector("lht-error-alert");

    expect(element.getAttribute("variant")).toBe("error");
    expect(element.getAttribute("role")).toBe("alert");
    expect(element.getAttribute("aria-live")).toBe("assertive");
  });

  it("uses status/polite for warning and info variants", () => {
    document.body.innerHTML = `
      <lht-error-alert variant="warning" text="warn"></lht-error-alert>
      <lht-error-alert variant="info" text="info"></lht-error-alert>
    `;

    const warning = document.querySelector('lht-error-alert[variant="warning"]');
    const info = document.querySelector('lht-error-alert[variant="info"]');

    expect(warning.getAttribute("role")).toBe("status");
    expect(warning.getAttribute("aria-live")).toBe("polite");
    expect(info.getAttribute("role")).toBe("status");
    expect(info.getAttribute("aria-live")).toBe("polite");
  });
});

describe("lht-command-block fallback", () => {
  it("renders native copy buttons when md-icon-button is unavailable", () => {
    document.body.innerHTML = `
      <lht-command-block command-id="cmd" copy-buttons="dual"></lht-command-block>
    `;

    const buttons = document.querySelectorAll("lht-command-block button.md-copy-button--fallback");
    const code = document.querySelector("lht-command-block code#cmd");

    expect(code).not.toBeNull();
    expect(buttons).toHaveLength(2);
  });
});

describe("lht-switch-help fallback", () => {
  it("renders supported fallback DOM when md-switch is unavailable", () => {
    const onChange = vi.fn();
    window.testSwitchChange = onChange;

    document.body.innerHTML = `
      <lht-switch-help switch-id="demo-switch" label="Demo" on-change="testSwitchChange" checked>
        help
      </lht-switch-help>
    `;

    const input = document.getElementById("demo-switch");
    const visual = document.querySelector("lht-switch-help .md-switch-input + .md-switch");

    expect(input).not.toBeNull();
    expect(input.tagName).toBe("INPUT");
    expect(input.checked).toBe(true);
    expect(visual).not.toBeNull();

    input.checked = false;
    input.dispatchEvent(new Event("change", { bubbles: true }));

    expect(input.getAttribute("aria-checked")).toBe("false");
    expect(onChange).toHaveBeenCalledTimes(1);

    delete window.testSwitchChange;
  });
});

describe("lht critical components mode matrix", () => {
  it("renders self-contained fallback DOM when material components are unavailable", () => {
    renderCriticalComponentsMarkup();

    expect(document.querySelector("lht-help-tooltip .md-help-icon-button--fallback")).not.toBeNull();
    expect(document.querySelector("lht-text-field-help input")).not.toBeNull();
    expect(document.querySelector("lht-select-help select")).not.toBeNull();
    expect(document.querySelector("lht-file-select button.lht-file-select__button--fallback")).not.toBeNull();
    expect(document.querySelector("lht-switch-help input.md-switch-input")).not.toBeNull();
    expect(document.querySelector("lht-command-block button.md-copy-button--fallback")).not.toBeNull();
  });

  it("renders md-* primitives when material components are registered", () => {
    defineMaterialTestDoubles();
    renderCriticalComponentsMarkup();

    expect(document.querySelector("lht-help-tooltip md-icon-button")).not.toBeNull();
    expect(document.querySelector("lht-text-field-help md-outlined-text-field")).not.toBeNull();
    expect(document.querySelector("lht-select-help md-outlined-select")).not.toBeNull();
    expect(document.querySelector("lht-file-select md-filled-button")).not.toBeNull();
    expect(document.querySelector("lht-switch-help md-switch")).not.toBeNull();
    expect(document.querySelector("lht-command-block md-icon-button")).not.toBeNull();
  });
});

describe("lht components in material-loaded mode", () => {
  it("uses md-* elements when the corresponding custom elements are registered", () => {
    defineMaterialTestDoubles();

    document.body.innerHTML = `
      <lht-help-tooltip label="help">body</lht-help-tooltip>
      <lht-text-field-help field-id="textField" label="Text"></lht-text-field-help>
      <lht-select-help field-id="selectField"></lht-select-help>
      <lht-file-select input-id="fileInputLoaded" button-id="fileButtonLoaded"></lht-file-select>
      <lht-switch-help switch-id="switchLoaded" label="Switch"></lht-switch-help>
      <lht-command-block command-id="cmdLoaded"></lht-command-block>
    `;

    expect(document.querySelector("lht-help-tooltip md-icon-button")).not.toBeNull();
    expect(document.querySelector("lht-text-field-help md-outlined-text-field")).not.toBeNull();
    expect(document.querySelector("lht-select-help md-outlined-select")).not.toBeNull();
    expect(document.querySelector("lht-file-select md-filled-button")).not.toBeNull();
    expect(document.querySelector("lht-switch-help md-switch")).not.toBeNull();
    expect(document.querySelector("lht-command-block md-icon-button")).not.toBeNull();
  });
});

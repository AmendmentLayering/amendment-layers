/*
 * Task pane and shared runtime: wires the keyboard shortcuts (shortcuts.json)
 * and the pane's buttons to Engine.
 */
(() => {
  "use strict";

  const ACTIONS = {
    MarkDelete: () => Engine.markDelete(),
    MarkInsert: () => Engine.markInsert(),
    Adopted: () => Engine.adopted(),
    Failed: () => Engine.failed(),
  };

  function showStatus(text, kind) {
    const el = document.getElementById("status");
    el.textContent = text;
    el.className = `status ${kind === "ok" ? "" : kind}`;
  }

  function showLevel(state) {
    for (const b of document.querySelectorAll(".level")) {
      b.setAttribute("aria-checked", String(Number(b.dataset.level) === state.activeLevel));
    }
  }

  // Keys as Word reports them; where Word reports none (for example, after a
  // key clash with another add-in was settled the other way), the key from
  // shortcuts.json.
  async function showShortcuts() {
    let keys = {};
    let defaults = {};
    try {
      keys = (await Office.actions.getShortcuts()) || {};
    } catch (e) {
      // getShortcuts needs KeyboardShortcuts 1.1; the buttons still work without it.
    }
    try {
      const file = await (await fetch("shortcuts.json", { cache: "no-store" })).json();
      for (const s of file.shortcuts) defaults[s.action] = s.key.default;
    } catch (e) {
      // No defaults; show what Word reported.
    }
    for (const kbd of document.querySelectorAll("kbd[data-key]")) {
      kbd.textContent = keys[kbd.dataset.key] || defaults[kbd.dataset.key] || "";
    }
  }

  function showArmed(on) {
    document.getElementById("all-places").setAttribute("aria-pressed", String(on));
  }

  function showPresenter(open) {
    document.getElementById("presenter").setAttribute("aria-pressed", String(open));
  }

  Office.onReady(() => {
    Engine.init({ onStatus: showStatus, onState: showLevel, onArm: showArmed, onChange: () => Presenter.changed() });
    Presenter.init({ serial: Engine.serial, onStatus: showStatus, onOpenChange: showPresenter });
    document.getElementById("all-places").addEventListener("click", () => Engine.toggleAllPlaces());
    document.getElementById("presenter").addEventListener("click", () => Presenter.toggle());

    for (const [id, fn] of Object.entries(ACTIONS)) Office.actions.associate(id, fn);

    for (const b of document.querySelectorAll(".level")) {
      b.addEventListener("click", () => Engine.setLevel(Number(b.dataset.level)));
    }
    for (const b of document.querySelectorAll("[data-action]")) {
      b.addEventListener("click", () => Engine[b.dataset.action]());
    }

    // Load the runtime (and so the shortcuts) whenever this document opens,
    // and reopen the pane with it.
    Office.addin.setStartupBehavior(Office.StartupBehavior.load).catch(() => {});
    if (!Office.context.document.settings.get("Office.AutoShowTaskpaneWithDocument")) {
      Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
      Office.context.document.settings.saveAsync();
    }

    showShortcuts();
  });
})();

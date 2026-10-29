/* DocMerge – front-end behaviour (vanilla JS, no dependencies) */
(() => {
  "use strict";

  const app = document.getElementById("app");
  const cfg = {
    minFiles: parseInt(app.dataset.minFiles, 10),
    maxFiles: parseInt(app.dataset.maxFiles, 10),
    maxSizeBytes: parseInt(app.dataset.maxSizeMb, 10) * 1024 * 1024,
    mergeUrl: app.dataset.mergeUrl,
  };

  const $ = (id) => document.getElementById(id);
  const els = {
    messages: $("messages"),
    sr: $("sr-status"),
    form: $("merge-form"),
    uploadCard: $("upload-card"),
    listCard: $("list-card"),
    dropzone: $("dropzone"),
    input: $("file-input"),
    browseBtn: $("browse-btn"),
    count: $("file-count"),
    clearBtn: $("clear-btn"),
    empty: $("empty-state"),
    list: $("file-list"),
    template: $("file-item-template"),
    mergeBtn: $("merge-btn"),
    mergeLabel: $("merge-btn").querySelector(".btn-label"),
    hint: $("merge-hint"),
    result: $("result"),
    resultFilename: $("result-filename"),
    resultSize: $("result-size"),
    resultWarnings: $("result-warnings"),
    warningList: $("warning-list"),
    downloadBtn: $("download-btn"),
    newMergeBtn: $("new-merge-btn"),
  };

  /** @type {{id:number, file:File}[]} */
  let files = [];
  let nextId = 1;
  let busy = false;
  let currentResult = null;
  let dragId = null;

  const csrfToken = () => els.form.querySelector("[name=csrfmiddlewaretoken]").value;

  function formatSize(bytes) {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  }

  const isDocx = (file) => file.name.toLowerCase().endsWith(".docx");
  const sameFile = (a, b) =>
    a.name === b.name && a.size === b.size && a.lastModified === b.lastModified;
  const indexOfId = (id) => files.findIndex((f) => String(f.id) === String(id));

  function announce(text) {
    els.sr.textContent = text;
  }

  function clearMessages() {
    els.messages.replaceChildren();
  }

  function showMessage(type, text, details = []) {
    const box = document.createElement("div");
    box.className = `alert alert-${type}`;
    box.setAttribute("role", type === "error" ? "alert" : "status");

    const body = document.createElement("div");
    body.className = "alert-body";
    const p = document.createElement("p");
    p.textContent = text;
    body.appendChild(p);

    if (details.length) {
      const ul = document.createElement("ul");
      details.forEach((d) => {
        const li = document.createElement("li");
        li.textContent = d;
        ul.appendChild(li);
      });
      body.appendChild(ul);
    }

    const close = document.createElement("button");
    close.type = "button";
    close.className = "alert-close";
    close.setAttribute("aria-label", "Dismiss message");
    close.textContent = "×";
    close.addEventListener("click", () => box.remove());

    box.append(body, close);
    els.messages.replaceChildren(box);
  }

  function addFiles(fileList) {
    clearMessages();
    const problems = [];
    let added = 0;
    let limitHit = false;

    for (const file of Array.from(fileList)) {
      if (!isDocx(file)) {
        problems.push(`'${file.name}' is not a .docx file. Only Word (.docx) documents are supported.`);
        continue;
      }
      if (file.size === 0) {
        problems.push(`'${file.name}' is empty.`);
        continue;
      }
      if (file.size > cfg.maxSizeBytes) {
        problems.push(`'${file.name}' is ${formatSize(file.size)}, which is over the ${formatSize(cfg.maxSizeBytes)} limit.`);
        continue;
      }
      if (files.some((f) => sameFile(f.file, file))) {
        problems.push(`'${file.name}' has already been added.`);
        continue;
      }
      if (files.length >= cfg.maxFiles) {
        limitHit = true;
        continue;
      }
      files.push({ id: nextId++, file });
      added += 1;
    }

    if (limitHit) {
      problems.push(`You can merge up to ${cfg.maxFiles} documents at a time. Extra files were not added.`);
    }
    if (problems.length) {
      showMessage(
        "error",
        added ? `${added} document${added === 1 ? "" : "s"} added, but some files were skipped:` : "No documents were added:",
        problems
      );
    }
    if (added) announce(`${added} document${added === 1 ? "" : "s"} added. ${files.length} in total.`);
    render();
  }

  function render() {
    els.list.replaceChildren();

    files.forEach((item, index) => {
      const li = els.template.content.firstElementChild.cloneNode(true);
      const name = item.file.name;
      li.dataset.id = item.id;
      li.querySelector(".file-order").textContent = String(index + 1);
      const nameEl = li.querySelector(".file-name");
      nameEl.textContent = name;
      nameEl.title = name;
      li.querySelector(".file-size").textContent = formatSize(item.file.size);

      const up = li.querySelector('[data-action="up"]');
      const down = li.querySelector('[data-action="down"]');
      const remove = li.querySelector('[data-action="remove"]');
      up.disabled = index === 0;
      down.disabled = index === files.length - 1;
      up.setAttribute("aria-label", `Move ${name} up`);
      down.setAttribute("aria-label", `Move ${name} down`);
      remove.setAttribute("aria-label", `Remove ${name}`);

      els.list.appendChild(li);
    });

    const count = files.length;
    els.count.textContent = String(count);
    els.empty.hidden = count > 0;
    els.list.hidden = count === 0;
    els.clearBtn.hidden = count === 0;
    updateMergeState();
  }

  function updateMergeState() {
    const ready = files.length >= cfg.minFiles;
    els.mergeBtn.disabled = !ready || busy;

    if (busy) {
      els.hint.textContent = "Merging your documents. This can take a few moments…";
    } else if (ready) {
      els.hint.textContent = `${files.length} documents will be merged in the order shown above.`;
    } else {
      const missing = cfg.minFiles - files.length;
      els.hint.textContent =
        files.length === 0
          ? `Add at least ${cfg.minFiles} documents to enable merging.`
          : `Add ${missing === 1 ? "one more document" : `${missing} more documents`} to enable merging.`;
    }
  }

  function focusControl(id, action) {
    const li = els.list.querySelector(`[data-id="${id}"]`);
    if (!li) return;
    let btn = li.querySelector(`[data-action="${action}"]`);
    if (btn && btn.disabled) {
      btn = li.querySelector(`[data-action="${action === "up" ? "down" : "up"}"]`);
    }
    if (btn && !btn.disabled) btn.focus();
  }

  function moveItem(from, to) {
    if (to < 0 || to >= files.length || from === to) return;
    const [moved] = files.splice(from, 1);
    files.splice(to, 0, moved);
    announce(`${moved.file.name} moved to position ${to + 1} of ${files.length}.`);
    render();
    return moved.id;
  }

  els.list.addEventListener("click", (event) => {
    const btn = event.target.closest("button[data-action]");
    if (!btn || busy) return;
    const li = btn.closest(".file-item");
    const index = indexOfId(li.dataset.id);
    if (index < 0) return;
    const action = btn.dataset.action;

    if (action === "up" || action === "down") {
      const id = moveItem(index, action === "up" ? index - 1 : index + 1);
      if (id !== undefined) focusControl(id, action);
    } else if (action === "remove") {
      const [removed] = files.splice(index, 1);
      const neighbour = files[index] || files[index - 1];
      announce(`${removed.file.name} removed. ${files.length} remaining.`);
      render();
      if (neighbour) focusControl(neighbour.id, "remove");
      else els.browseBtn.focus();
    }
  });

  function clearDropMarkers() {
    els.list.querySelectorAll(".drop-before, .drop-after").forEach((el) => el.classList.remove("drop-before", "drop-after"));
  }

  function endDrag() {
    dragId = null;
    clearDropMarkers();
    els.list.querySelectorAll(".is-dragging").forEach((el) => el.classList.remove("is-dragging"));
  }

  els.list.addEventListener("dragstart", (event) => {
    const li = event.target.closest && event.target.closest(".file-item");
    if (!li || busy) return;
    dragId = li.dataset.id;
    li.classList.add("is-dragging");
    event.dataTransfer.effectAllowed = "move";
    event.dataTransfer.setData("text/plain", dragId);
  });

  els.list.addEventListener("dragover", (event) => {
    if (dragId === null) return;
    event.preventDefault();
    clearDropMarkers();
    const li = event.target.closest(".file-item");
    if (!li || li.dataset.id === dragId) return;
    const rect = li.getBoundingClientRect();
    li.classList.add(event.clientY > rect.top + rect.height / 2 ? "drop-after" : "drop-before");
  });

  els.list.addEventListener("drop", (event) => {
    if (dragId === null) return;
    event.preventDefault();
    const li = event.target.closest(".file-item");
    if (li && li.dataset.id !== dragId) {
      const rect = li.getBoundingClientRect();
      const after = event.clientY > rect.top + rect.height / 2;
      const from = indexOfId(dragId);
      const [moved] = files.splice(from, 1);
      const target = indexOfId(li.dataset.id);
      files.splice(after ? target + 1 : target, 0, moved);
      announce(`${moved.file.name} moved to position ${indexOfId(moved.id) + 1} of ${files.length}.`);
      render();
    }
    endDrag();
  });

  els.list.addEventListener("dragend", endDrag);

  const hasFiles = (event) =>
    Array.from((event.dataTransfer && event.dataTransfer.types) || []).includes("Files");

  els.browseBtn.addEventListener("click", (event) => {
    event.stopPropagation();
    els.input.click();
  });

  els.dropzone.addEventListener("click", (event) => {
    if (event.target.closest("button")) return;
    els.input.click();
  });

  els.input.addEventListener("change", () => {
    if (els.input.files.length) addFiles(els.input.files);
    els.input.value = "";
  });

  ["dragenter", "dragover"].forEach((type) =>
    els.dropzone.addEventListener(type, (event) => {
      if (!hasFiles(event)) return;
      event.preventDefault();
      event.dataTransfer.dropEffect = "copy";
      els.dropzone.classList.add("is-dragover");
    })
  );

  els.dropzone.addEventListener("dragleave", (event) => {
    if (els.dropzone.contains(event.relatedTarget)) return;
    els.dropzone.classList.remove("is-dragover");
  });

  els.dropzone.addEventListener("drop", (event) => {
    if (!hasFiles(event)) return;
    event.preventDefault();
    els.dropzone.classList.remove("is-dragover");
    if (!busy) addFiles(event.dataTransfer.files);
  });

  ["dragover", "drop"].forEach((type) =>
    window.addEventListener(type, (event) => {
      if (hasFiles(event)) event.preventDefault();
    })
  );

  els.clearBtn.addEventListener("click", () => {
    if (busy) return;
    files = [];
    clearMessages();
    announce("All documents removed.");
    render();
    els.browseBtn.focus();
  });

  function setBusy(flag) {
    busy = flag;
    els.mergeBtn.classList.toggle("is-loading", flag);
    els.mergeBtn.setAttribute("aria-busy", String(flag));
    els.mergeLabel.textContent = flag ? "Merging…" : "Merge Documents";
    els.uploadCard.toggleAttribute("inert", flag);
    els.listCard.toggleAttribute("inert", flag);
    updateMergeState();
  }

  function errorForStatus(status) {
    if (status === 403) return "Your session has expired. Please refresh the page and try again.";
    if (status === 413) return "The files are too large for the server to accept.";
    return "Something went wrong while merging. Please try again.";
  }

  els.form.addEventListener("submit", async (event) => {
    event.preventDefault();
    if (busy || files.length < cfg.minFiles) return;
    clearMessages();
    setBusy(true);

    const data = new FormData();
    files.forEach(({ file }) => data.append("documents", file, file.name));

    try {
      const response = await fetch(cfg.mergeUrl, {
        method: "POST",
        body: data,
        credentials: "same-origin",
        headers: { "X-CSRFToken": csrfToken(), "X-Requested-With": "XMLHttpRequest" },
      });

      let payload = null;
      try {
        payload = await response.json();
      } catch (_) {
        /* server returned a non-JSON error page */
      }

      if (!response.ok || !payload || !payload.ok) {
        throw new Error((payload && payload.error) || errorForStatus(response.status));
      }
      showResult(payload);
    } catch (error) {
      showMessage(
        "error",
        error instanceof TypeError
          ? "Could not reach the server. Please check your connection and try again."
          : error.message
      );
    } finally {
      setBusy(false);
    }
  });

  function showResult(payload) {
    currentResult = payload;
    els.resultFilename.textContent = payload.filename;
    els.resultSize.textContent = payload.size ? `· ${formatSize(payload.size)}` : "";
    els.downloadBtn.href = payload.download_url;
    els.downloadBtn.setAttribute("download", payload.filename);

    els.warningList.replaceChildren();
    const warnings = payload.warnings || [];
    warnings.forEach((text) => {
      const li = document.createElement("li");
      li.textContent = text;
      els.warningList.appendChild(li);
    });
    els.resultWarnings.hidden = warnings.length === 0;

    els.form.hidden = true;
    els.result.hidden = false;
    announce("Your documents have been merged successfully.");
    els.result.scrollIntoView({ behavior: "smooth", block: "start" });
    els.result.focus({ preventScroll: true });
  }

  els.newMergeBtn.addEventListener("click", () => {
    if (currentResult && currentResult.discard_url) {
      fetch(currentResult.discard_url, {
        method: "POST",
        credentials: "same-origin",
        headers: { "X-CSRFToken": csrfToken() },
      }).catch(() => {});
    }
    currentResult = null;
    files = [];
    clearMessages();
    els.result.hidden = true;
    els.form.hidden = false;
    render();
    window.scrollTo({ top: 0, behavior: "smooth" });
    els.browseBtn.focus();
  });

  render();
})();

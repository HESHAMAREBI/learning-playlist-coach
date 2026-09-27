(() => {
  "use strict";

  const VERSION = "0.4.2";
  const PANEL_ID = "lpt-panel";
  const BUTTON_ID = "lpt-floating-button";
  const MODAL_ID = "lpt-coach-modal";
  const COPY_TEMP_ID = "lpt-copy-temp";
  const STORE = "lpt.playlist.v3.";
  const BACKUP = "lpt.pageBackup.v3.";
  const SETTINGS = "lpt.settings.v1";
  const DEFAULT_AGENT = "https://m365.cloud.microsoft/chat/?titleId=T_86dbe9bd-1f55-1a86-021d-985dd825f2b7&source=embedded-builder";

  let timer = null;
  let busy = false;
  let lastCount = -1;
  let pendingOpen = false;
  let isOrphaned = false;
  let domObserver = null;

  // 1. فحص شامل ودقيق لصلاحية سياق الإضافة
  function isExtensionContextValid() {
    if (isOrphaned) return false;
    try {
      return Boolean(
        typeof chrome !== "undefined" &&
        chrome &&
        chrome.runtime &&
        Boolean(chrome.runtime.id)
      );
    } catch {
      isOrphaned = true;
      return false;
    }
  }

  // 2. إيقاف الـ Observers والمؤقتات إذا فقد السكربت اتصاله بالإضافة (Orphan detection)
  function checkOrphaned() {
    if (isOrphaned) return true;
    if (!isExtensionContextValid()) {
      isOrphaned = true;
      if (domObserver) {
        try {
          domObserver.disconnect();
          domObserver = null;
        } catch {}
      }
      if (timer) {
        clearTimeout(timer);
        timer = null;
      }
      console.warn("LPT: Extension context invalidated; content script detached safely.");
      return true;
    }
    return false;
  }

  // 3. اعتراض أي رفض غير معالج يخص انتهاء سياق الإضافة على مستوى الصفحة
  window.addEventListener("unhandledrejection", event => {
    const msg = String(event?.reason?.message || event?.reason || "");
    if (msg.includes("Extension context invalidated") || msg.includes("context invalidated")) {
      event.preventDefault();
      checkOrphaned();
    }
  });

  // 4. أغلفة آمنة (Safe Wrappers) للتعامل مع Chrome APIs دون رمي أي استثناءات
  async function safeStorageGet(keys) {
    if (checkOrphaned()) return null;
    try {
      if (chrome?.storage?.local) {
        return await chrome.storage.local.get(keys);
      }
    } catch (error) {
      checkOrphaned();
      console.warn("LPT: storage.local.get safely suppressed", error);
    }
    return null;
  }

  async function safeStorageSet(data) {
    if (checkOrphaned()) return false;
    try {
      if (chrome?.storage?.local) {
        await chrome.storage.local.set(data);
        return true;
      }
    } catch (error) {
      checkOrphaned();
      console.warn("LPT: storage.local.set safely suppressed", error);
    }
    return false;
  }

  async function safeSendMessage(payload) {
    if (checkOrphaned()) return null;
    try {
      if (chrome?.runtime?.sendMessage) {
        return await chrome.runtime.sendMessage(payload);
      }
    } catch (error) {
      checkOrphaned();
      console.warn("LPT: runtime.sendMessage safely suppressed", error);
    }
    return null;
  }

  const pid = () => new URL(location.href).searchParams.get("list") || "";
  const sk = id => STORE + id;
  const bk = id => BACKUP + id;

  const defaults = id => ({
    schemaVersion: 3,
    playlistId: id,
    playlistTitle: document.title.replace(/\s*-\s*YouTube\s*$/i, "").trim(),
    watched: {},
    speed: 1,
    planEnabled: false,
    goalDays: 7,
    lastVideoId: "",
    lastVideoTitle: "",
    lastPosition: 0,
    lastVideoUrl: "",
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString()
  });

  function normalize(x, id) {
    const s = { ...defaults(id), ...(x || {}) };
    s.playlistId = id;
    s.watched = s.watched && typeof s.watched === "object" ? s.watched : {};
    s.speed = Number(s.speed) || 1;
    s.goalDays = Math.max(1, Number(s.goalDays) || 7);
    s.lastPosition = Math.max(0, Number(s.lastPosition) || 0);
    return s;
  }

  async function readState(id) {
    let a = null;
    let b = null;

    if (!checkOrphaned()) {
      const d = await safeStorageGet([sk(id), `lpt:${id}`]);
      if (d) {
        a = d[sk(id)] || d[`lpt:${id}`];
      }
    }

    try {
      b = JSON.parse(localStorage.getItem(bk(id)) || "null");
    } catch {}

    const list = [a, b].filter(Boolean).sort((x, y) => String(y.updatedAt || "").localeCompare(String(x.updatedAt || "")));
    const s = normalize(list[0], id);
    await writeState(id, s, false);
    return s;
  }

  async function writeState(id, s, touch = true) {
    const n = normalize(s, id);
    if (touch) n.updatedAt = new Date().toISOString();

    if (!checkOrphaned()) {
      await safeStorageSet({ [sk(id)]: n });
    }

    try {
      localStorage.setItem(bk(id), JSON.stringify(n));
    } catch {}

    Object.assign(s, n);
  }

  async function settings() {
    const fallback = { agentUrl: DEFAULT_AGENT, coachName: "Ask AI Coach", promptLanguage: "ar" };
    try {
      const cached = JSON.parse(localStorage.getItem(SETTINGS) || "null");
      if (cached) Object.assign(fallback, cached);
    } catch {}

    if (checkOrphaned()) {
      return fallback;
    }

    const response = await safeSendMessage({ type: "LPT_GET_SETTINGS" });
    if (response?.settings) {
      Object.assign(fallback, response.settings);
      try {
        localStorage.setItem(SETTINGS, JSON.stringify(fallback));
      } catch {}
    }
    return fallback;
  }

  // ========================
  // ميزة النسخ الاحتياطي والاستعادة
  // ========================
  async function exportData() {
    try {
      const cfg = await settings();
      const currentId = pid();
      const playlistsMap = {};

      // 1. تجميع بيانات قوائم التشغيل من chrome.storage.local
      if (!checkOrphaned()) {
        try {
          const allStorage = await safeStorageGet(null);
          if (allStorage && typeof allStorage === "object") {
            for (const [key, value] of Object.entries(allStorage)) {
              if (key.startsWith(STORE) || key.startsWith("lpt:")) {
                const playlistId = value?.playlistId || key.replace(STORE, "").replace("lpt:", "");
                if (playlistId && value && typeof value === "object") {
                  playlistsMap[playlistId] = normalize(value, playlistId);
                }
              }
            }
          }
        } catch (err) {
          console.warn("LPT: Error scanning chrome.storage for export", err);
        }
      }

      // 2. تجميع من localStorage (النسخ الاحتياطية على مستوى الصفحة)
      try {
        for (let i = 0; i < localStorage.length; i++) {
          const key = localStorage.key(i);
          if (!key) continue;
          if (key.startsWith(BACKUP) || key.startsWith(STORE) || key.startsWith("lpt:")) {
            try {
              const raw = JSON.parse(localStorage.getItem(key) || "null");
              if (raw && typeof raw === "object") {
                const playlistId = raw.playlistId || key.replace(BACKUP, "").replace(STORE, "").replace("lpt:", "");
                if (playlistId) {
                  const norm = normalize(raw, playlistId);
                  if (!playlistsMap[playlistId] || String(norm.updatedAt || "") > String(playlistsMap[playlistId].updatedAt || "")) {
                    playlistsMap[playlistId] = norm;
                  }
                }
              }
            } catch {}
          }
        }
      } catch (err) {
        console.warn("LPT: Error scanning localStorage for export", err);
      }

      // 3. التأكد من حفظ وتضمين القائمة الحالية المفتوحة بأحدث حالة
      if (currentId) {
        try {
          const currentState = await readState(currentId);
          if (currentState) {
            playlistsMap[currentId] = currentState;
          }
        } catch {}
      }

      // 4. بناء بنية ملف النسخة الاحتياطية المطلوبة
      const backupData = {
        version: VERSION,
        createdAt: new Date().toISOString(),
        settings: {
          agentUrl: cfg.agentUrl || DEFAULT_AGENT,
          coachName: cfg.coachName || "Ask AI Coach",
          promptLanguage: cfg.promptLanguage || "ar",
          ...cfg
        },
        playlists: playlistsMap
      };

      // 5. التنزيل المباشر كملف JSON محلي
      const jsonStr = JSON.stringify(backupData, null, 2);
      const blob = new Blob([jsonStr], { type: "application/json;charset=utf-8" });
      const url = URL.createObjectURL(blob);
      const dateStr = new Date().toISOString().split("T")[0];
      const a = document.createElement("a");
      a.href = url;
      a.download = `LearningPlaylistCoach_Backup_${dateStr}.json`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    } catch (error) {
      console.error("LPT: Export failed", error);
      alert("❌ حدث خطأ أثناء تصدير البيانات.");
    }
  }

  async function importData(file) {
    if (!file) return;

    let text = "";
    try {
      text = await file.text();
    } catch {
      alert("❌ ملف النسخة الاحتياطية غير صالح");
      return;
    }

    let json = null;
    try {
      json = JSON.parse(text);
    } catch {
      alert("❌ ملف النسخة الاحتياطية غير صالح");
      return;
    }

    // 6. التحقق من صحة الملف
    if (
      !json ||
      typeof json !== "object" ||
      !json.version ||
      !json.settings ||
      typeof json.settings !== "object" ||
      !json.playlists ||
      typeof json.playlists !== "object"
    ) {
      alert("❌ ملف النسخة الاحتياطية غير صالح");
      return;
    }

    // 8. عدم فقدان البيانات الحالية: تأكيد من المستخدم
    const confirmed = confirm("سيتم استبدال البيانات الحالية.\nهل تريد المتابعة؟");
    if (!confirmed) return;

    try {
      // 9. توافق الإصدارات (Migration Layer)
      const migratedSettings = {
        agentUrl: json.settings.agentUrl || DEFAULT_AGENT,
        coachName: json.settings.coachName || "Ask AI Coach",
        promptLanguage: json.settings.promptLanguage || "ar",
        ...json.settings
      };

      const storagePayload = {
        [SETTINGS]: migratedSettings
      };

      const playlistsRaw = json.playlists;
      const migratedPlaylists = {};

      if (Array.isArray(playlistsRaw)) {
        playlistsRaw.forEach(item => {
          const id = item?.playlistId || item?.id;
          if (id) {
            const norm = normalize(item, id);
            norm.schemaVersion = 3;
            migratedPlaylists[id] = norm;
            storagePayload[sk(id)] = norm;
          }
        });
      } else if (typeof playlistsRaw === "object" && playlistsRaw !== null) {
        for (const [id, item] of Object.entries(playlistsRaw)) {
          if (!item) continue;
          const norm = normalize(item, id);
          norm.schemaVersion = 3;
          migratedPlaylists[id] = norm;
          storagePayload[sk(id)] = norm;
        }
      }

      // 10. التخزين المحلي فقط في chrome.storage.local
      if (!checkOrphaned()) {
        await safeStorageSet(storagePayload);
      }

      // 10. التخزين المحلي الاحتياطي في localStorage
      try {
        localStorage.setItem(SETTINGS, JSON.stringify(migratedSettings));
        for (const [id, item] of Object.entries(migratedPlaylists)) {
          localStorage.setItem(bk(id), JSON.stringify(item));
          localStorage.setItem(sk(id), JSON.stringify(item));
        }
      } catch (err) {
        console.warn("LPT: localStorage import save error", err);
      }

      // 7. إعادة تحميل الواجهة وإظهار نجاح الاستيراد
      schedule(10, true);
      setTimeout(() => {
        alert("✅ تم استيراد النسخة الاحتياطية بنجاح");
      }, 120);
    } catch (error) {
      console.error("LPT: Import failed", error);
      alert("❌ حدث خطأ أثناء استيراد البيانات.");
    }
  }

  function parseDuration(t) {
    const c = String(t || "").replace(/[^0-9:]/g, "");
    if (!c.includes(":")) return 0;
    const p = c.split(":").map(Number);
    return p.some(Number.isNaN) ? 0 : p.reduce((a, v) => a * 60 + v, 0);
  }

  function fmt(v) {
    v = Math.max(0, Math.round(Number(v) || 0));
    const h = Math.floor(v / 3600);
    const m = Math.floor((v % 3600) / 60);
    const s = v % 60;
    return h ? `${h}:${String(m).padStart(2, "0")}:${String(s).padStart(2, "0")}` : `${m}:${String(s).padStart(2, "0")}`;
  }

  function esc(v = "") {
    return String(v).replace(/[&<>'"]/g, c => ({
      "&": "&amp;",
      "<": "&lt;",
      ">": "&gt;",
      "'": "&#39;",
      '"': "&quot;"
    }[c]));
  }

  function items() {
    const id = pid();
    const map = new Map();
    document.querySelectorAll("ytd-playlist-video-renderer,ytd-playlist-panel-video-renderer").forEach((n, i) => {
      const a = n.querySelector('a#video-title[href*="watch"],a[href*="watch?v="]');
      if (!a) return;
      const u = new URL(a.href, location.origin);
      const v = u.searchParams.get("v");
      if (!v || map.has(v) || (u.searchParams.get("list") && u.searchParams.get("list") !== id)) return;
      const d = n.querySelector("ytd-thumbnail-overlay-time-status-renderer #text,ytd-thumbnail-overlay-time-status-renderer span,.badge-shape-wiz__text");
      map.set(v, {
        id: v,
        title: (a.title || a.textContent || `Video ${i + 1}`).trim(),
        duration: parseDuration(d?.textContent),
        href: u.href
      });
    });

    if (!map.size) {
      document.querySelectorAll('a[href*="watch?v="][href*="list="]').forEach((a, i) => {
        const u = new URL(a.href, location.origin);
        const v = u.searchParams.get("v");
        if (u.searchParams.get("list") !== id || !v || map.has(v)) return;
        map.set(v, {
          id: v,
          title: (a.title || a.textContent || `Video ${i + 1}`).trim(),
          duration: 0,
          href: u.href
        });
      });
    }

    return [...map.values()];
  }

  function context() {
    const video = document.querySelector("video");
    if (!video) return null;
    const seconds = Math.floor(video.currentTime || 0);
    const url = new URL(location.href);
    url.searchParams.set("t", seconds + "s");
    return {
      title: document.querySelector("h1.ytd-watch-metadata yt-formatted-string")?.textContent?.trim() || document.title.replace(/\s*-\s*YouTube\s*$/i, "").trim(),
      seconds,
      time: fmt(seconds),
      url: url.toString()
    };
  }

  function prompt(c) {
    return `أنا أدرس حاليًا من فيديو تعليمي.\n\nعنوان الفيديو:\n${c.title}\n\nالتوقيت الحالي:\n${c.time}\n\nرابط الفيديو عند التوقيت:\n${c.url}\n\nاشرح المفهوم المعروض عند هذه النقطة باللغة العربية، مع إبقاء المصطلحات التقنية الأساسية باللغة الإنجليزية.\n\nاشرح بالترتيب التالي:\n1. الفكرة الأساسية.\n2. الهدف العملي.\n3. مثال مبسط.\n4. نقطة قد تسبب التباسًا.\n5. سؤال تدريبي قصير.\n\nسؤالي هو:`;
  }

  async function copyText(text) {
    try {
      await navigator.clipboard.writeText(text);
      return true;
    } catch {}
    let ta = document.getElementById(COPY_TEMP_ID);
    if (!ta) {
      ta = document.createElement("textarea");
      ta.id = COPY_TEMP_ID;
      ta.style.cssText = "position:fixed;left:-9999px;top:0;opacity:0;pointer-events:none";
      document.body.appendChild(ta);
    }
    ta.value = text;
    ta.focus();
    ta.select();
    let ok = false;
    try {
      ok = document.execCommand("copy");
    } catch {}
    return ok;
  }

  function captureFrame() {
    const v = document.querySelector("video");
    if (!v || !v.videoWidth) return null;
    try {
      const c = document.createElement("canvas");
      c.width = v.videoWidth;
      c.height = v.videoHeight;
      c.getContext("2d").drawImage(v, 0, 0, c.width, c.height);
      return c.toDataURL("image/png");
    } catch {
      return null;
    }
  }

  function showCoachModal(ctx, text, agentUrl, coachName) {
    let box = document.getElementById(MODAL_ID);
    if (!box) {
      const existing = document.querySelectorAll(`#${MODAL_ID}`);
      if (existing.length > 0) {
        box = existing[0];
        for (let i = 1; i < existing.length; i++) {
          existing[i].style.display = "none";
        }
      } else {
        box = document.createElement("div");
        box.id = MODAL_ID;
        document.body.appendChild(box);
      }
    }

    const shot = captureFrame();
    box.style.display = "flex";
    box.classList.remove("lpt-hidden");
    box.innerHTML = `<div class="lpt-modal-card"><div class="lpt-modal-head"><strong>${esc(coachName)}</strong><button id="lpt-modal-x">×</button></div>${shot ? `<img class="lpt-shot" src="${shot}" alt="Video frame">` : `<div class="lpt-shot-note">تعذر التقاط إطار الفيديو. يمكن متابعة السؤال بدون صورة.</div>`}<label>السؤال الجاهز<textarea id="lpt-prompt-text">${esc(text)}</textarea></label><div id="lpt-copy-status"></div><div class="lpt-modal-actions"><button id="lpt-copy">نسخ السؤال</button>${shot ? '<button id="lpt-save-shot">حفظ اللقطة</button>' : ''}<button id="lpt-open-agent">فتح الـ Agent</button></div><p>بعد فتح الـ Agent الصق السؤال باستخدام Ctrl + V. الصورة تُحفظ كملف منفصل لإرفاقها يدويًا.</p></div>`;

    const hideModal = () => {
      box.style.display = "none";
      box.classList.add("lpt-hidden");
    };

    box.querySelector("#lpt-modal-x").onclick = hideModal;
    box.onclick = e => {
      if (e.target === box) hideModal();
    };

    box.querySelector("#lpt-copy").onclick = async () => {
      const val = box.querySelector("#lpt-prompt-text").value;
      const ok = await copyText(val);
      const s = box.querySelector("#lpt-copy-status");
      s.textContent = ok ? "✓ تم نسخ السؤال" : "تعذر النسخ التلقائي. النص محدد ويمكن نسخه بـ Ctrl + C";
      s.className = ok ? "success" : "error";
      if (!ok) {
        const t = box.querySelector("#lpt-prompt-text");
        t.focus();
        t.select();
      }
    };

    box.querySelector("#lpt-save-shot")?.addEventListener("click", () => {
      const a = document.createElement("a");
      a.href = shot;
      a.download = `video-frame-${ctx.seconds}s.png`;
      a.click();
    });

    box.querySelector("#lpt-open-agent").onclick = () => window.open(agentUrl, "_blank", "noopener,noreferrer");
  }

  function hideUI() {
    const p = document.getElementById(PANEL_ID);
    if (p) p.classList.add("lpt-hidden");
    const b = document.getElementById(BUTTON_ID);
    if (b) {
      b.style.display = "none";
      b.classList.add("lpt-hidden");
    }
    const modal = document.getElementById(MODAL_ID);
    if (modal) {
      modal.style.display = "none";
      modal.classList.add("lpt-hidden");
    }
    lastCount = -1;
  }

  function openSidebar() {
    pendingOpen = true;
    let p = document.getElementById(PANEL_ID);
    if (p) {
      p.classList.remove("lpt-hidden");
    }
    floating();
    schedule(20, true);
  }

  function toggleSidebar() {
    let p = document.getElementById(PANEL_ID);
    if (!p) {
      openSidebar();
      return;
    }
    if (p.classList.contains("lpt-hidden")) {
      p.classList.remove("lpt-hidden");
      pendingOpen = true;
      schedule(20, true);
    } else {
      p.classList.add("lpt-hidden");
      pendingOpen = false;
    }
  }

  function floating() {
    let b = document.getElementById(BUTTON_ID);
    if (!b) {
      const existing = document.querySelectorAll(`#${BUTTON_ID}`);
      if (existing.length > 0) {
        b = existing[0];
        for (let i = 1; i < existing.length; i++) {
          existing[i].style.display = "none";
        }
      } else {
        b = document.createElement("button");
        b.id = BUTTON_ID;
        b.textContent = "✓";
        b.title = "فتح/إغلاق متابع القائمة";
        document.body.appendChild(b);
      }
    }
    b.onclick = () => toggleSidebar();
    b.style.display = "";
    b.classList.remove("lpt-hidden");
  }

  async function render(open = false) {
    if (open) pendingOpen = true;
    if (busy) return;

    const id = pid();
    if (!id) {
      hideUI();
      return;
    }

    busy = true;
    const shouldOpen = open || pendingOpen;
    pendingOpen = false;

    try {
      floating();
      const list = items();
      const state = await readState(id);
      const cfg = await settings();

      const total = list.reduce((a, v) => a + v.duration, 0);
      const watched = list.reduce((a, v) => a + (state.watched[v.id] ? v.duration : 0), 0);
      const done = list.filter(v => state.watched[v.id]).length;
      const progress = list.length ? Math.round((done / list.length) * 100) : 0;
      const remaining = Math.max(0, total - watched) / state.speed;

      let p = document.getElementById(PANEL_ID);
      const isNew = !p;
      const wasHidden = p ? p.classList.contains("lpt-hidden") : false;

      if (!p) {
        const existingSidebars = document.querySelectorAll(`#${PANEL_ID}`);
        if (existingSidebars.length > 0) {
          p = existingSidebars[0];
          for (let i = 1; i < existingSidebars.length; i++) {
            existingSidebars[i].style.display = "none";
          }
        } else {
          p = document.createElement("aside");
          p.id = PANEL_ID;
          document.body.appendChild(p);
        }
      }

      p.innerHTML = `<div class="lpt-header"><div><strong>متابع قائمة التشغيل</strong><small>Learning Playlist Coach v${VERSION}</small></div><button id="lpt-close" title="إغلاق">×</button></div><div class="lpt-progress"><span style="width:${progress}%"></span></div><div class="lpt-stats"><div><b>${progress}%</b><span>الإنجاز</span></div><div><b>${done}/${list.length}</b><span>الفيديوهات</span></div><div><b>${fmt(remaining)}</b><span>المتبقي</span></div></div><div class="lpt-save-status">✓ الحفظ التلقائي مفعل <small>${esc(id)}</small></div><div class="lpt-controls"><label>سرعة المشاهدة<select id="lpt-speed">${[.75, 1, 1.25, 1.5, 1.75, 2].map(x => `<option value="${x}" ${state.speed === x ? "selected" : ""}>${x}x</option>`).join("")}</select></label><label class="lpt-plan-switch"><input id="lpt-plan" type="checkbox" ${state.planEnabled ? "checked" : ""}> تفعيل الخطة الدراسية</label><label class="${state.planEnabled ? "" : "lpt-hidden-control"}">عدد الأيام<input id="lpt-days" type="number" min="1" max="365" value="${state.goalDays}"></label></div><div class="lpt-summary"><span>الإجمالي: ${fmt(total)}</span><span>المشاهد: ${fmt(watched)}</span><span>${state.planEnabled ? `الهدف اليومي: ${fmt(remaining / state.goalDays)}` : "الخطة اختيارية"}</span></div><div class="lpt-coach"><button id="lpt-ask"><span>🤖</span><span><strong>${esc(cfg.coachName)}</strong><small>إنشاء السؤال واللقطة قبل فتح الوكيل</small></span></button><button id="lpt-settings" title="إعدادات الوكيل">⚙</button></div><div class="lpt-backup"><div class="lpt-backup-title">💾 النسخ الاحتياطي</div><div class="lpt-backup-actions"><button id="lpt-export" type="button">📤 تصدير البيانات</button><button id="lpt-import" type="button">📥 استيراد البيانات</button><input type="file" id="lpt-import-file" accept=".json" style="display:none"></div></div><div class="lpt-list">${list.length ? list.map((v, i) => `<label class="lpt-item ${state.watched[v.id] ? "done" : ""}"><input type="checkbox" data-id="${v.id}" ${state.watched[v.id] ? "checked" : ""}><span>${i + 1}</span><a href="${v.href}">${esc(v.title)}</a><time>${fmt(v.duration)}</time></label>`).join("") : '<div class="lpt-empty">لم تُقرأ الفيديوهات بعد.<button id="lpt-refresh">تحديث</button></div>'}</div><div class="lpt-footer"><button id="lpt-all">تحديد الظاهر كمكتمل</button><button id="lpt-reset" class="danger">إعادة ضبط العلامات</button></div>`;

      if (shouldOpen) {
        p.classList.remove("lpt-hidden");
      } else if (isNew) {
        p.classList.remove("lpt-hidden");
      } else if (wasHidden) {
        p.classList.add("lpt-hidden");
      } else {
        p.classList.remove("lpt-hidden");
      }

      p.querySelector("#lpt-close").onclick = () => {
        p.classList.add("lpt-hidden");
        pendingOpen = false;
      };

      p.querySelector("#lpt-refresh")?.addEventListener("click", () => schedule(50, true));

      p.querySelector("#lpt-speed").onchange = async e => {
        try {
          state.speed = Number(e.target.value);
          await writeState(id, state);
          schedule(30, true);
        } catch (err) {
          console.warn("LPT: speed update error", err);
        }
      };

      p.querySelector("#lpt-plan").onchange = async e => {
        try {
          state.planEnabled = e.target.checked;
          await writeState(id, state);
          schedule(30, true);
        } catch (err) {
          console.warn("LPT: plan update error", err);
        }
      };

      p.querySelector("#lpt-days").onchange = async e => {
        try {
          state.goalDays = Math.max(1, Number(e.target.value) || 7);
          await writeState(id, state);
        } catch (err) {
          console.warn("LPT: goal days update error", err);
        }
      };

      p.querySelectorAll("input[data-id]").forEach(x => {
        x.onchange = async e => {
          try {
            if (e.target.checked) {
              state.watched[e.target.dataset.id] = { completedAt: new Date().toISOString() };
            } else {
              delete state.watched[e.target.dataset.id];
            }
            await writeState(id, state);
            schedule(30, true);
          } catch (err) {
            console.warn("LPT: item toggle error", err);
          }
        };
      });

      p.querySelector("#lpt-all").onclick = async () => {
        try {
          list.forEach(v => {
            state.watched[v.id] = state.watched[v.id] || { completedAt: new Date().toISOString() };
          });
          await writeState(id, state);
          schedule(30, true);
        } catch (err) {
          console.warn("LPT: mark all error", err);
        }
      };

      p.querySelector("#lpt-reset").onclick = async () => {
        try {
          state.watched = {};
          await writeState(id, state);
          schedule(30, true);
        } catch (err) {
          console.warn("LPT: reset error", err);
        }
      };

      p.querySelector("#lpt-settings").onclick = async () => {
        if (checkOrphaned()) {
          alert("تم تحديث الإضافة بينما صفحة YouTube ما زالت مفتوحة. أعد تحميل صفحة YouTube ثم افتح الإعدادات من أيقونة الإضافة.");
          return;
        }
        const response = await safeSendMessage({ type: "LPT_OPEN_OPTIONS" });
        if (!response?.ok) {
          alert("تم تحديث الإضافة بينما صفحة YouTube ما زالت مفتوحة. أعد تحميل صفحة YouTube ثم افتح الإعدادات من أيقونة الإضافة.");
        }
      };

      p.querySelector("#lpt-ask").onclick = () => {
        const c = context();
        if (!c) {
          alert("افتح فيديو من قائمة التشغيل أولًا.");
          return;
        }
        document.querySelector("video")?.pause();
        showCoachModal(c, prompt(c), cfg.agentUrl, cfg.coachName);
      };

      // ربط أزرار النسخ الاحتياطي والاستعادة
      p.querySelector("#lpt-export").onclick = async () => {
        await exportData();
      };

      const fileInput = p.querySelector("#lpt-import-file");
      p.querySelector("#lpt-import").onclick = () => {
        fileInput.value = "";
        fileInput.click();
      };

      fileInput.onchange = async e => {
        const file = e.target.files?.[0];
        if (file) {
          await importData(file);
        }
      };

      lastCount = list.length;
    } catch (renderError) {
      console.warn("LPT: render error handled safely", renderError);
    } finally {
      busy = false;
      if (pendingOpen) {
        schedule(20, true);
      }
    }
  }

  function schedule(ms = 500, open = false) {
    if (isOrphaned) return;
    if (open) pendingOpen = true;
    clearTimeout(timer);
    timer = setTimeout(() => {
      const shouldOpen = pendingOpen;
      if (shouldOpen) {
        pendingOpen = false;
      }
      render(shouldOpen).catch(err => {
        console.warn("LPT: scheduled render error", err);
      });
    }, ms);
  }

  // إنشاء الـ MutationObserver وحفظ مرجعه لفصله إذا انتهت صلاحية السكربت
  domObserver = new MutationObserver(ms => {
    if (isOrphaned) return;
    const meaningful = ms.some(m => {
      const t = m.target instanceof Element ? m.target : m.target?.parentElement;
      return (
        t &&
        !t.closest?.(`#${PANEL_ID}`) &&
        !t.closest?.(`#${MODAL_ID}`) &&
        !t.closest?.(`#${BUTTON_ID}`)
      );
    });
    if (meaningful && items().length !== lastCount) schedule(700, false);
  });

  domObserver.observe(document.documentElement, { childList: true, subtree: true });

  window.addEventListener("yt-navigate-finish", () => {
    if (!isOrphaned) schedule(500, false);
  });
  window.addEventListener("popstate", () => {
    if (!isOrphaned) schedule(500, false);
  });
  document.addEventListener("yt-page-data-updated", () => {
    if (!isOrphaned) schedule(500, false);
  });

  try {
    if (isExtensionContextValid() && chrome?.runtime?.onMessage) {
      chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
        if (checkOrphaned()) return false;
        const type = message?.type || message?.action;
        if (type === "LPT_OPEN" || type === "open" || type === "OPEN") {
          openSidebar();
          if (typeof sendResponse === "function") {
            sendResponse({ ok: true, state: "opened" });
          }
          return true;
        }
        if (type === "LPT_TOGGLE" || type === "toggle" || type === "TOGGLE") {
          toggleSidebar();
          const p = document.getElementById(PANEL_ID);
          const isHidden = !p || p.classList.contains("lpt-hidden");
          if (typeof sendResponse === "function") {
            sendResponse({ ok: true, state: isHidden ? "hidden" : "opened" });
          }
          return true;
        }
      });
    }
  } catch (error) {
    console.warn("LPT: stale content script; refresh YouTube", error);
  }

  schedule(700, false);
})();
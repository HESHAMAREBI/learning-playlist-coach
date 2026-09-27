const SETTINGS_KEY = "lpt.settings.v1";
const DEFAULT_SETTINGS = {
  agentUrl: "https://m365.cloud.microsoft/chat/?titleId=T_86dbe9bd-1f55-1a86-021d-985dd825f2b7&source=embedded-builder",
  coachName: "Ask AI Coach"
};

chrome.runtime.onInstalled.addListener(() => {
  console.log("Learning Playlist Coach v0.4.1 installed");
});

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message?.type === "LPT_OPEN_OPTIONS") {
    chrome.runtime.openOptionsPage()
      .then(() => sendResponse({ ok: true }))
      .catch(error => sendResponse({ ok: false, error: String(error) }));
    return true;
  }

  if (message?.type === "LPT_GET_SETTINGS") {
    chrome.storage.local.get(SETTINGS_KEY)
      .then(data => sendResponse({ ok: true, settings: { ...DEFAULT_SETTINGS, ...(data[SETTINGS_KEY] || {}) } }))
      .catch(error => sendResponse({ ok: false, settings: DEFAULT_SETTINGS, error: String(error) }));
    return true;
  }
});

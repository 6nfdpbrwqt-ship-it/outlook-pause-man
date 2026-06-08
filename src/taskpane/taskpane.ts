/* global document, Office */

import { loadSettings, saveSettings, DEFAULT_SETTINGS, Settings } from "../shared/settings";

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    populate(loadSettings());
    document.getElementById("save")!.addEventListener("click", onSave);
  }
});

function populate(s: Settings) {
  (document.getElementById("chk-subject") as HTMLInputElement).checked = s.checkEmptySubject;
  (document.getElementById("chk-body") as HTMLInputElement).checked = s.checkEmptyBody;
  (document.getElementById("chk-external") as HTMLInputElement).checked = s.checkExternalDomain;
  (document.getElementById("chk-recipcount") as HTMLInputElement).checked = s.checkRecipientCount;
  (document.getElementById("recip-threshold") as HTMLInputElement).value = String(s.recipientCountThreshold);
  (document.getElementById("internal-domains") as HTMLTextAreaElement).value = s.internalDomains.join("\n");
  (document.getElementById("forbidden-words") as HTMLTextAreaElement).value = s.forbiddenWords.join("\n");
}

function readForm(): Settings {
  const lines = (id: string) =>
    (document.getElementById(id) as HTMLTextAreaElement).value
      .split(/\r?\n/)
      .map((s) => s.trim())
      .filter((s) => s.length > 0);

  return {
    checkEmptySubject: (document.getElementById("chk-subject") as HTMLInputElement).checked,
    checkEmptyBody: (document.getElementById("chk-body") as HTMLInputElement).checked,
    checkExternalDomain: (document.getElementById("chk-external") as HTMLInputElement).checked,
    checkRecipientCount: (document.getElementById("chk-recipcount") as HTMLInputElement).checked,
    recipientCountThreshold: parseInt((document.getElementById("recip-threshold") as HTMLInputElement).value, 10) || DEFAULT_SETTINGS.recipientCountThreshold,
    internalDomains: lines("internal-domains"),
    forbiddenWords: lines("forbidden-words"),
  };
}

async function onSave() {
  const status = document.getElementById("status")!;
  status.textContent = "保存中...";
  status.className = "pm-status";
  try {
    await saveSettings(readForm());
    status.textContent = "✅ 保存しました";
    status.className = "pm-status pm-status-ok";
    setTimeout(() => (status.textContent = ""), 2500);
  } catch (e: any) {
    status.textContent = `❌ 保存失敗: ${e?.message ?? e}`;
    status.className = "pm-status pm-status-err";
  }
}

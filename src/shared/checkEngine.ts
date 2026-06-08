import { Settings } from "./settings";

export type CheckLevel = "ok" | "info" | "warning" | "error";

export interface CheckResult {
  level: CheckLevel;
  title: string;
  detail: string;
}

interface MailData {
  subject: string;
  body: string;
  to: Office.EmailAddressDetails[];
  cc: Office.EmailAddressDetails[];
  bcc: Office.EmailAddressDetails[];
}

function isExternal(email: string, internalDomains: string[]): boolean {
  if (!email) return false;
  const at = email.indexOf("@");
  if (at < 0) return false;
  const domain = email.slice(at + 1).toLowerCase();
  if (internalDomains.length === 0) return false;
  return !internalDomains.some(
    (d) => domain === d.toLowerCase() || domain.endsWith("." + d.toLowerCase())
  );
}

/**
 * メール内容を検査して結果リストを返す。
 * Outlook Classic 版 (CheckEngine.cs) から移植。判定ルールは同等。
 */
export function runChecks(data: MailData, settings: Settings): CheckResult[] {
  const results: CheckResult[] = [];
  const all = [...data.to, ...data.cc, ...data.bcc];

  // 1. 宛先
  if (all.length === 0) {
    results.push({ level: "error", title: "宛先未入力", detail: "To / CC / BCC がすべて空です。" });
  }

  // 2. 件名
  if (settings.checkEmptySubject && !data.subject.trim()) {
    results.push({ level: "warning", title: "件名なし", detail: "件名が入力されていません。" });
  }

  // 3. 本文
  if (settings.checkEmptyBody && !data.body.trim()) {
    results.push({ level: "warning", title: "本文なし", detail: "本文が入力されていません。" });
  }

  // 4. 外部ドメイン
  if (settings.checkExternalDomain && settings.internalDomains.length > 0) {
    const externals = all.filter((r) => isExternal(r.emailAddress, settings.internalDomains));
    if (externals.length > 0) {
      results.push({
        level: "warning",
        title: `外部ドメイン ${externals.length} 名`,
        detail: externals.map((r) => r.emailAddress).join(", "),
      });
    }
  }

  // 5. 宛先人数
  if (settings.checkRecipientCount && all.length >= settings.recipientCountThreshold) {
    results.push({
      level: "warning",
      title: `宛先多数 (${all.length} 名)`,
      detail: `閾値 ${settings.recipientCountThreshold} 名以上です。`,
    });
  }

  // 6. 禁止ワード
  if (settings.forbiddenWords.length > 0) {
    const full = (data.subject || "") + " " + (data.body || "");
    const found = settings.forbiddenWords.filter((w) => full.includes(w));
    if (found.length > 0) {
      results.push({
        level: "error",
        title: "禁止ワード検出",
        detail: found.join(", "),
      });
    }
  }

  return results;
}

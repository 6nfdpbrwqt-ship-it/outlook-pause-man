/* global Office */

import { loadSettings, Settings } from "../shared/settings";
import { runChecks, CheckResult } from "../shared/checkEngine";

Office.onReady(() => {
  // Office.js ready
});

async function gatherMailData(item: Office.MessageCompose) {
  const get = <T>(fn: (cb: (r: Office.AsyncResult<T>) => void) => void) =>
    new Promise<T>((resolve) => fn((r) => resolve(r.value)));

  const [subject, body, to, cc, bcc] = await Promise.all([
    get<string>((cb) => item.subject.getAsync(cb)),
    get<string>((cb) => item.body.getAsync(Office.CoercionType.Text, cb)),
    get<Office.EmailAddressDetails[]>((cb) => item.to.getAsync(cb)),
    get<Office.EmailAddressDetails[]>((cb) => item.cc.getAsync(cb)),
    get<Office.EmailAddressDetails[]>((cb) => item.bcc.getAsync(cb)),
  ]);

  return { subject: subject || "", body: body || "", to: to || [], cc: cc || [], bcc: bcc || [] };
}

/**
 * OnMessageSend ハンドラ。
 * SendMode="PromptUser" なので、allowEvent: false + errorMessage を返すと
 * Outlook 側で「キャンセル / 続行」ダイアログが出る。
 */
async function onMessageSendHandler(event: Office.AddinCommands.Event) {
  try {
    const item = Office.context.mailbox.item as unknown as Office.MessageCompose;
    const settings: Settings = loadSettings();
    const data = await gatherMailData(item);
    const results: CheckResult[] = runChecks(data, settings);

    const errors = results.filter((r) => r.level === "error");
    const warnings = results.filter((r) => r.level === "warning");

    if (errors.length > 0) {
      // 完全ブロック(allowEvent: false で送信不可)
      const msg = errors.map((e) => `🚫 ${e.title}: ${e.detail}`).join("\n");
      event.completed({
        allowEvent: false,
        errorMessage: msg,
      } as any);
      return;
    }

    if (warnings.length > 0) {
      // 警告: 一旦ブロック → ダイアログを閉じてもう一度送信すれば送れる
      // (PromptUser モードでは再 send 時にハンドラが再実行される)
      const msg = "⚠ 確認してください\n\n"
        + warnings.map((w) => `・${w.title}: ${w.detail}`).join("\n")
        + "\n\n問題なければダイアログを閉じてもう一度送信ボタンを押してください。";
      event.completed({
        allowEvent: false,
        errorMessage: msg,
      } as any);
      return;
    }

    // すべて OK
    event.completed({ allowEvent: true });
  } catch (e) {
    // 失敗時は送信を許可してしまう(誤って送信ブロックを残さないため)
    console.error("OutlookPauseMan onMessageSend failed:", e);
    event.completed({ allowEvent: true });
  }
}

// Office に関数を登録(manifest の FunctionName と一致させる)
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);

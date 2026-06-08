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
  console.log("[OutlookPauseMan] onMessageSendHandler started");
  // 安全タイムアウト: 7 秒以内に event.completed を呼べなかったら強制送信許可
  const safety = setTimeout(() => {
    console.warn("[OutlookPauseMan] safety timeout fired, allowing send");
    event.completed({ allowEvent: true });
  }, 7000);
  const done = (opts: any) => {
    clearTimeout(safety);
    event.completed(opts);
  };

  try {
    const item = Office.context.mailbox.item as unknown as Office.MessageCompose;
    const settings: Settings = loadSettings();
    console.log("[OutlookPauseMan] gathering mail data...");
    const data = await gatherMailData(item);
    console.log("[OutlookPauseMan] gathered:", { subject: data.subject, bodyLen: data.body.length, to: data.to.length, cc: data.cc.length, bcc: data.bcc.length });
    const results: CheckResult[] = runChecks(data, settings);
    console.log("[OutlookPauseMan] results:", results);

    const errors = results.filter((r) => r.level === "error");
    const warnings = results.filter((r) => r.level === "warning");

    if (errors.length > 0) {
      // 完全ブロック(allowEvent: false で送信不可)
      const msg = errors.map((e) => `🚫 ${e.title}: ${e.detail}`).join("\n");
      done({
        allowEvent: false,
        errorMessage: msg,
      });
      return;
    }

    if (warnings.length > 0) {
      // 警告: 一旦ブロック → ダイアログを閉じてもう一度送信すれば送れる
      // (PromptUser モードでは再 send 時にハンドラが再実行される)
      const msg = "⚠ 確認してください\n\n"
        + warnings.map((w) => `・${w.title}: ${w.detail}`).join("\n")
        + "\n\n問題なければダイアログを閉じてもう一度送信ボタンを押してください。";
      done({
        allowEvent: false,
        errorMessage: msg,
      });
      return;
    }

    // すべて OK
    console.log("[OutlookPauseMan] all checks passed, allowing send");
    done({ allowEvent: true });
  } catch (e) {
    console.error("[OutlookPauseMan] onMessageSend failed:", e);
    done({ allowEvent: true });
  }
}

// Office に関数を登録(manifest の FunctionName と一致させる)
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);

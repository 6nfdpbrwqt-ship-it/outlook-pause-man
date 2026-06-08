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
  // デバッグ用: ハンドラの実行ステップを蓄積してダイアログに出す
  const trace: string[] = [];
  const t0 = Date.now();
  const log = (msg: string) => trace.push(`+${Date.now() - t0}ms ${msg}`);

  log("handler entered");

  // 安全タイムアウト: 6 秒以内に event.completed を呼べなかったら強制送信許可
  const safety = setTimeout(() => {
    log("SAFETY TIMEOUT (6s)");
    event.completed({
      allowEvent: false,
      errorMessage: "OutlookPauseMan: タイムアウト\n\n" + trace.join("\n"),
    } as any);
  }, 6000);
  const done = (opts: any) => {
    clearTimeout(safety);
    event.completed(opts);
  };

  try {
    const item = Office.context.mailbox.item as unknown as Office.MessageCompose;
    log("got item");
    const settings: Settings = loadSettings();
    log(`loaded settings (domains:${settings.internalDomains.length} words:${settings.forbiddenWords.length})`);
    log("calling gatherMailData...");
    const data = await gatherMailData(item);
    log(`gathered (subj:${data.subject.length}c body:${data.body.length}c to:${data.to.length} cc:${data.cc.length} bcc:${data.bcc.length})`);
    const results: CheckResult[] = runChecks(data, settings);
    log(`ran checks, ${results.length} results`);

    const errors = results.filter((r) => r.level === "error");
    const warnings = results.filter((r) => r.level === "warning");

    if (errors.length > 0) {
      log("blocking (errors)");
      const msg = errors.map((e) => `🚫 ${e.title}: ${e.detail}`).join("\n")
        + "\n\n---debug---\n" + trace.join("\n");
      done({ allowEvent: false, errorMessage: msg });
      return;
    }

    if (warnings.length > 0) {
      log("warning");
      const msg = "⚠ 確認してください\n\n"
        + warnings.map((w) => `・${w.title}: ${w.detail}`).join("\n")
        + "\n\n問題なければダイアログを閉じてもう一度送信ボタンを押してください。"
        + "\n\n---debug---\n" + trace.join("\n");
      done({ allowEvent: false, errorMessage: msg });
      return;
    }

    log("all ok, allowing");
    done({ allowEvent: true });
  } catch (e: any) {
    log(`EXCEPTION: ${e?.message || e}`);
    done({
      allowEvent: false,
      errorMessage: "OutlookPauseMan エラー\n\n" + trace.join("\n"),
    });
  }
}

// Office に関数を登録(manifest の FunctionName と一致させる)
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);

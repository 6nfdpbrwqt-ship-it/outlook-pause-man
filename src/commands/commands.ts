/* global Office */

// v1.0.5: associate を Office.onReady 内で確実に呼ぶ
const BUILD_TAG = "v1.0.5-onReady-2026-06-08";

function onMessageSendHandler(event: Office.AddinCommands.Event) {
  event.completed({
    allowEvent: false,
    errorMessage: `${BUILD_TAG}\n\nhandler ran at ${new Date().toISOString()}\n\nこれが見えれば新版が動いてる。`,
  } as any);
}

Office.onReady(() => {
  // Office.actions が確実に存在する状態で登録
  if (Office?.actions?.associate) {
    Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
  }
});

// onReady を待たずに念のためフォールバック登録(両方やる)
try {
  if ((globalThis as any).Office?.actions?.associate) {
    (globalThis as any).Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
  }
} catch {
  /* Office.actions 未準備なら onReady 側で登録される */
}

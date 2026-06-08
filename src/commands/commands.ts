/* global Office */

// ============================================================
// v1.0.4 トライアル: 完全に何もしない極小ハンドラ
// 即座に errorMessage を返すだけ。遅延ゼロのはず。
// これでも「予想以上に時間がかかっています」が出る場合は
// Office.actions.associate の登録が成立していない=ハンドラ自体が呼ばれていない。
// ============================================================

const BUILD_TAG = "v1.0.4-MINIMAL-2026-06-08-20:45";

Office.onReady(() => {});

function onMessageSendHandler(event: Office.AddinCommands.Event) {
  const t = Date.now();
  event.completed({
    allowEvent: false,
    errorMessage: `${BUILD_TAG}\n\nhandler ran at: ${new Date(t).toISOString()}\n\nこのメッセージが出たら新ハンドラ動作OK。`,
  } as any);
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);

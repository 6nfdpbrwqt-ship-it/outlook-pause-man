/* global Office */

export interface Settings {
  checkEmptySubject: boolean;
  checkEmptyBody: boolean;
  checkExternalDomain: boolean;
  checkRecipientCount: boolean;
  recipientCountThreshold: number;
  internalDomains: string[];
  forbiddenWords: string[];
}

export const DEFAULT_SETTINGS: Settings = {
  checkEmptySubject: true,
  checkEmptyBody: true,
  checkExternalDomain: true,
  checkRecipientCount: false,
  recipientCountThreshold: 10,
  internalDomains: ["wspartners.co.jp"],
  forbiddenWords: [],
};

const KEY = "outlookpauseman_settings_v1";

export function loadSettings(): Settings {
  try {
    const raw = Office.context.roamingSettings?.get(KEY);
    if (raw && typeof raw === "object") {
      return { ...DEFAULT_SETTINGS, ...raw };
    }
  } catch {
    // 起動順や roamingSettings 未対応の場面でも落とさない
  }
  return DEFAULT_SETTINGS;
}

export function saveSettings(settings: Settings): Promise<void> {
  return new Promise((resolve, reject) => {
    try {
      Office.context.roamingSettings.set(KEY, settings);
      Office.context.roamingSettings.saveAsync((r) => {
        if (r.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(new Error(r.error?.message || "save failed"));
      });
    } catch (e) {
      reject(e);
    }
  });
}

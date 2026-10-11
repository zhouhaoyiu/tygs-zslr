import { ref } from "vue";
import { registerSW } from "virtual:pwa-register";

export const offlineReady = ref(false);
export const offline = ref(!navigator.onLine);
export const pwaError = ref("");
export const updateAvailable = ref(false);
export const installAvailable = ref(false);
let installEvent;
window.addEventListener("online", () => {
  offline.value = false;
});
window.addEventListener("offline", () => {
  offline.value = true;
});
window.addEventListener("beforeinstallprompt", (event) => {
  event.preventDefault();
  installEvent = event;
  installAvailable.value = true;
});
window.addEventListener("appinstalled", () => {
  installAvailable.value = false;
});
export async function installApp() {
  if (!installEvent) return;
  await installEvent.prompt();
  await installEvent.userChoice;
  installAvailable.value = false;
  installEvent = null;
}
export const updateApp = registerSW({
  immediate: true,
  onOfflineReady() {
    offlineReady.value = true;
  },
  onNeedRefresh() {
    updateAvailable.value = true;
  },
  onRegisteredSW(_, registration) {
    if (registration?.active) offlineReady.value = true;
  },
  onRegisterError() {
    pwaError.value = "离线资源准备失败";
  },
});

// Сводки АГЗС — фоновый помощник: показывает уведомления, даже когда приложение закрыто.
const params = new URL(self.location).searchParams;
const API = params.get("api");
const DEV = params.get("dev");

self.addEventListener("install", () => self.skipWaiting());
self.addEventListener("activate", e => e.waitUntil(self.clients.claim()));

self.addEventListener("push", e => {
  e.waitUntil((async () => {
    let msgs = [];
    try {
      const r = await fetch(API, { method: "POST", body: JSON.stringify({ a: "inbox", dev: DEV }) });
      const j = await r.json();
      if (j && j.ok && Array.isArray(j.data)) msgs = j.data;
    } catch (err) {}
    if (!msgs.length) {
      const open = await self.registration.getNotifications();
      if (open.length) return;
      msgs = [{ id: "x", title: "Сводки АГЗС", body: "Есть новости. Откройте приложение." }];
    }
    for (const m of msgs) {
      await self.registration.showNotification(m.title, {
        body: m.body,
        tag: "agzs-" + m.id,
        icon: "icon.png",
        vibrate: m.urgent ? [300, 150, 300, 150, 300] : [200],
        requireInteraction: !!m.urgent,
        data: { url: self.registration.scope },
      });
    }
  })());
});

self.addEventListener("notificationclick", e => {
  e.notification.close();
  const url = (e.notification.data && e.notification.data.url) || self.registration.scope;
  e.waitUntil((async () => {
    const list = await self.clients.matchAll({ type: "window", includeUncontrolled: true });
    for (const c of list) {
      if (c.url.startsWith(self.registration.scope) && "focus" in c) { c.postMessage("refresh"); return c.focus(); }
    }
    return self.clients.openWindow(url);
  })());
});

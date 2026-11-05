// Сводки АГЗС — фоновый помощник: показывает уведомления, даже когда приложение закрыто.
const params = new URL(self.location).searchParams;
const API = params.get("api");
const DEV = params.get("dev");
const TOK = params.get("tok");

self.addEventListener("install", () => self.skipWaiting());
self.addEventListener("activate", e => e.waitUntil(self.clients.claim()));

self.addEventListener("push", e => {
  e.waitUntil((async () => {
    // Сразу показываем уведомление: Safari отключает уведомления сайту, если push пришёл, а показа не было.
    await self.registration.showNotification("Сводки АГЗС", { body: "Загружаю…", tag: "agzs-wait", icon: "icon.png", silent: true });
    let msgs = [];
    try {
      const ctl = new AbortController(); setTimeout(() => ctl.abort(), 15000);
      const r = await fetch(API, { method: "POST", body: JSON.stringify({ a: "inbox", dev: DEV, tok: TOK }), signal: ctl.signal });
      const j = await r.json();
      if (j && j.ok && Array.isArray(j.data)) msgs = j.data;
    } catch (err) {}
    const waiting = await self.registration.getNotifications({ tag: "agzs-wait" });
    if (!msgs.length) {
      const others = (await self.registration.getNotifications()).filter(n => n.tag !== "agzs-wait");
      if (others.length) { waiting.forEach(n => n.close()); return; }
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
    waiting.forEach(n => n.close());
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

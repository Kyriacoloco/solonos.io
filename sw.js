// Solonos service worker: shows push notifications and opens the right page when tapped. Version 1.0
self.addEventListener("install", () => self.skipWaiting());
self.addEventListener("activate", (e) => e.waitUntil(self.clients.claim()));
self.addEventListener("push", (e) => {
  let d = { title: "Solonos", body: "", url: "/app/tasks", tag: "solonos" };
  try { d = { ...d, ...e.data.json() }; } catch (_) { d.body = e.data ? e.data.text() : ""; }
  e.waitUntil(self.registration.showNotification(d.title, { body: d.body, tag: d.tag, icon: "/icon-192.png", badge: "/icon-192.png", data: { url: d.url }, renotify: false }));
});
self.addEventListener("notificationclick", (e) => {
  e.notification.close();
  const url = new URL((e.notification.data && e.notification.data.url) || "/app/tasks", self.location.origin).href;
  e.waitUntil(self.clients.matchAll({ type: "window", includeUncontrolled: true }).then((list) => {
    for (const c of list) { if ("focus" in c) { c.navigate(url); return c.focus(); } }
    return self.clients.openWindow(url);
  }));
});

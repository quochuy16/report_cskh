/* Chuyển giao diện sáng / tối.
   - Chạy trong <head> để đặt theme trước khi vẽ trang (không bị nháy trắng).
   - Chưa chọn lần nào thì theo cài đặt của máy; bấm nút thì ghi nhớ lựa chọn trong trình duyệt. */
(function () {
  "use strict";
  var KEY = "cskh-theme", root = document.documentElement;
  var mq = window.matchMedia ? window.matchMedia("(prefers-color-scheme: dark)") : null;
  function saved() { try { return localStorage.getItem(KEY); } catch (e) { return null; } }
  function current() { return saved() || (mq && mq.matches ? "dark" : "light"); }
  function apply(t) {
    root.setAttribute("data-theme", t);
    var b = document.getElementById("btnTheme");
    if (b) {
      var dark = t === "dark";
      b.textContent = dark ? "☀" : "☾";
      b.title = dark ? "Chuyển sang giao diện sáng" : "Chuyển sang giao diện tối";
      b.setAttribute("aria-label", b.title);
      b.setAttribute("aria-pressed", String(dark));
    }
  }
  apply(current());
  if (mq) {
    var onChange = function () { if (!saved()) apply(current()); };
    if (mq.addEventListener) mq.addEventListener("change", onChange); else if (mq.addListener) mq.addListener(onChange);
  }
  document.addEventListener("DOMContentLoaded", function () {
    var b = document.getElementById("btnTheme");
    if (!b) return;
    apply(current());
    b.addEventListener("click", function () {
      var t = root.getAttribute("data-theme") === "dark" ? "light" : "dark";
      try { localStorage.setItem(KEY, t); } catch (e) {}
      apply(t);
    });
  });
})();

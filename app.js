let count = 0;

function renderCount() {
  document.getElementById("count").innerText = count;
}

function increase() {
  count++;
  renderCount();
}

function decrease() {
  count--;
  renderCount();
}

function saveName() {
  const input = document.getElementById("nameInput");
  const name = input.value.trim();

  if (!name) return;

  localStorage.setItem("savedName", name);
  document.getElementById("savedName").innerText = "Saved name: " + name;
  input.value = "";
}

window.onload = function () {
  // Restore saved name
  const saved = localStorage.getItem("savedName");
  if (saved) {
    document.getElementById("savedName").innerText = "Saved name: " + saved;
  }

  // Small install hint for Android users
  const hint = document.getElementById("installHint");
  hint.innerText =
    "On Android Chrome: ⋮ menu → Install app / Add to Home screen (once hosted on HTTPS).";
};

// ✅ PWA: Register the service worker (offline + caching)
if ("serviceWorker" in navigator) {
  navigator.serviceWorker.register("./service-worker.js");
}
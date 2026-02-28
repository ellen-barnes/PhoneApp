const STORAGE_KEY = "accounts";

let accounts = [];

function escapeHtml(value) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function formatMoney(value) {
  return new Intl.NumberFormat("en-GB", {
    style: "currency",
    currency: "GBP"
  }).format(value);
}

function getSafeAmount(rawValue) {
  const parsed = Number.parseFloat(rawValue);
  return Number.isFinite(parsed) ? parsed : 0;
}

function saveAccounts() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(accounts));
}

function loadAccounts() {
  const saved = localStorage.getItem(STORAGE_KEY);
  if (!saved) {
    accounts = [];
    return;
  }

  try {
    const parsed = JSON.parse(saved);
    if (!Array.isArray(parsed)) {
      accounts = [];
      return;
    }

    accounts = parsed.map((account, index) => ({
      id: typeof account.id === "string" ? account.id : String(index),
      name: typeof account.name === "string" ? account.name : "Unnamed account",
      balance: getSafeAmount(account.balance)
    }));
  } catch {
    accounts = [];
  }
}

function calculateTotal() {
  return accounts.reduce((sum, account) => sum + account.balance, 0);
}

function renderTotal() {
  document.getElementById("totalBalance").innerText = formatMoney(calculateTotal());
}

function renderAccounts() {
  const list = document.getElementById("accountsList");
  list.innerHTML = "";

  if (accounts.length === 0) {
    list.innerHTML = '<p class="saved">No accounts yet. Add your first account above.</p>';
    return;
  }

  accounts.forEach((account) => {
    const safeName = escapeHtml(account.name);
    const card = document.createElement("div");
    card.className = "account-item";
    card.innerHTML = `
      <div class="account-header">
        <h3>${safeName}</h3>
        <p class="account-balance">${formatMoney(account.balance)}</p>
      </div>
      <div class="row">
        <input id="amount-${account.id}" class="input amount-input" type="number" step="0.01" placeholder="Amount" />
        <button class="btn btn-secondary" onclick="withdraw('${account.id}')">Remove</button>
        <button class="btn" onclick="deposit('${account.id}')">Add</button>
      </div>
    `;
    list.appendChild(card);
  });
}

function render() {
  renderTotal();
  renderAccounts();
}

function addAccount() {
  const nameInput = document.getElementById("accountNameInput");
  const balanceInput = document.getElementById("initialBalanceInput");

  const name = nameInput.value.trim();
  if (!name) {
    return;
  }

  const initialBalance = getSafeAmount(balanceInput.value);
  accounts.push({
    id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
    name,
    balance: initialBalance
  });

  nameInput.value = "";
  balanceInput.value = "";
  saveAccounts();
  render();
}

function getAccountAmount(accountId) {
  const amountInput = document.getElementById(`amount-${accountId}`);
  if (!amountInput) {
    return 0;
  }

  const amount = getSafeAmount(amountInput.value);
  amountInput.value = "";
  return amount > 0 ? amount : 0;
}

function deposit(accountId) {
  const amount = getAccountAmount(accountId);
  if (!amount) {
    return;
  }

  const account = accounts.find((item) => item.id === accountId);
  if (!account) {
    return;
  }

  account.balance += amount;
  saveAccounts();
  render();
}

function withdraw(accountId) {
  const amount = getAccountAmount(accountId);
  if (!amount) {
    return;
  }

  const account = accounts.find((item) => item.id === accountId);
  if (!account) {
    return;
  }

  account.balance -= amount;
  saveAccounts();
  render();
}

window.onload = function () {
  loadAccounts();
  render();

  const hint = document.getElementById("installHint");
  hint.innerText =
    "On Android Chrome: menu -> Install app / Add to Home screen (once hosted on HTTPS).";
};

if ("serviceWorker" in navigator) {
  navigator.serviceWorker.register("./service-worker.js");
}

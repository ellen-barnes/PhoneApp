const ACCOUNTS_STORAGE_KEY = "accounts";
const TRANSACTIONS_STORAGE_KEY = "transactions";
const CATEGORY_STORAGE_KEY = "transaction-category-options";
const EXPORT_META_STORAGE_KEY = "data-last-export-meta";
const DEFAULT_CATEGORY_OPTIONS = {
  expense: [
    "Audits",
    "Bills",
    "Charity",
    "Shopping",
    "Travel",
    "House",
    "Gym",
    "Sports",
    "Hobby",
    "Drinks",
    "Alcohol",
    "Pets",
    "Kids",
    "Social",
    "Clothing",
    "Dental",
    "Hair/Nail",
    "Medical",
    "Essential",
    "Unknown"
  ],
  income: [
    "Salary",
    "Bonus",
    "Investment",
    "Gifts",
    "Others",
    "Reimbursement",
    "Unknown"
  ]
};
const UNKNOWN_CATEGORY = "Unknown";
const ACCOUNT_ADJUSTMENT_CATEGORY = "Account Adjustment";

let accounts = [];
let transactions = [];
let categoryOptions = {
  expense: [...DEFAULT_CATEGORY_OPTIONS.expense],
  income: [...DEFAULT_CATEGORY_OPTIONS.income]
};
let selectedCategoryByType = { expense: "", income: "" };
let selectedGraphAccountId = "total";
let selectedCategoryBreakdownType = "expense";
let selectedCategoryBreakdownMonth = "";
let activeTab = "dashboard";
let transactionHistoryShowAll = false;
let showArchivedAccounts = false;
let expandedAccountId = "";
let activeTransactionModalId = "";
let transactionModalEditMode = false;
let transactionModalContext = null;
let graphShowAllHistory = true;
let graphShowTrendLine = true;
let graphPointTargets = [];
let activeGraphPointTransactionId = "";
let activeDataTab = "import";
let showImportFormatDetails = false;
let selectedImportFile = null;
let lastExportMeta = null;
let accountTransactionsShowAllById = {};
let selectedCategoryBreakdownMonthByType = { expense: "", income: "" };
let selectedCategoryBreakdownCategory = "";
let categoryBreakdownTransactionsShowAll = false;
const DAY_MS = 24 * 60 * 60 * 1000;
const CATEGORY_BREAKDOWN_COLORS = [
  "#1f77b4",
  "#ff7f0e",
  "#2ca02c",
  "#d62728",
  "#9467bd",
  "#8c564b",
  "#e377c2",
  "#7f7f7f",
  "#bcbd22",
  "#17becf",
  "#4e79a7",
  "#f28e2b"
];

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function normalizeMoney(value) {
  return Math.round(value * 100) / 100;
}

function getSafeAmount(rawValue) {
  const parsed = Number.parseFloat(rawValue);
  return Number.isFinite(parsed) ? normalizeMoney(parsed) : 0;
}

function formatMoney(value) {
  return new Intl.NumberFormat("en-GB", {
    style: "currency",
    currency: "GBP"
  }).format(value);
}

function formatSignedMoney(value) {
  if (value === 0) return formatMoney(0);
  return `${value > 0 ? "+" : "-"}${formatMoney(Math.abs(value))}`;
}

function formatAxisMoney(value) {
  if (Math.abs(value) >= 1000) {
    return new Intl.NumberFormat("en-GB", {
      style: "currency",
      currency: "GBP",
      notation: "compact",
      maximumFractionDigits: 1
    }).format(value);
  }
  return formatMoney(value);
}

function formatFileSize(bytes) {
  const size = Number.parseFloat(bytes);
  if (!Number.isFinite(size) || size <= 0) return "0 B";
  if (size < 1024) return `${Math.round(size)} B`;
  if (size < 1024 * 1024) return `${(size / 1024).toFixed(1)} KB`;
  return `${(size / (1024 * 1024)).toFixed(2)} MB`;
}

function formatDateTimeForDisplay(value) {
  const timestamp = Date.parse(value);
  if (!Number.isFinite(timestamp)) return "Never";
  return new Date(timestamp).toLocaleString("en-GB", {
    day: "2-digit",
    month: "short",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit"
  });
}

function normalizeHeader(value) {
  return String(value).trim().toLowerCase().replace(/[^a-z0-9]+/g, "");
}

function normalizeTransactionType(rawType) {
  const value = String(rawType || "").trim().toLowerCase();
  if (value.startsWith("exp")) return "expense";
  if (value.startsWith("inc")) return "income";
  if (value.startsWith("trans")) return "transfer";
  return "";
}

function toTrimmedString(value) {
  return typeof value === "string" ? value.trim() : String(value ?? "").trim();
}

function normalizeCategoryName(value) {
  return toTrimmedString(value).replace(/\s+/g, " ");
}

function cloneDefaultCategoryOptions() {
  return {
    expense: [...DEFAULT_CATEGORY_OPTIONS.expense],
    income: [...DEFAULT_CATEGORY_OPTIONS.income]
  };
}

function mergeCategoryList(base, extra) {
  const next = [];
  const seen = new Set();

  [...base, ...extra].forEach((name) => {
    const clean = normalizeCategoryName(name);
    if (!clean) return;
    const key = clean.toLowerCase();
    if (seen.has(key)) return;
    seen.add(key);
    next.push(clean);
  });

  return next;
}

function sanitizeCategoryOptions(raw) {
  const fallback = cloneDefaultCategoryOptions();
  if (!raw || typeof raw !== "object") {
    return fallback;
  }

  return {
    expense: mergeCategoryList(fallback.expense, Array.isArray(raw.expense) ? raw.expense : []),
    income: mergeCategoryList(fallback.income, Array.isArray(raw.income) ? raw.income : [])
  };
}

function saveCategoryOptions() {
  localStorage.setItem(CATEGORY_STORAGE_KEY, JSON.stringify(categoryOptions));
}

function loadCategoryOptions() {
  const saved = localStorage.getItem(CATEGORY_STORAGE_KEY);
  if (!saved) {
    categoryOptions = cloneDefaultCategoryOptions();
    return;
  }

  try {
    const parsed = JSON.parse(saved);
    categoryOptions = sanitizeCategoryOptions(parsed);
  } catch {
    categoryOptions = cloneDefaultCategoryOptions();
  }
}

function loadLastExportMeta() {
  const saved = localStorage.getItem(EXPORT_META_STORAGE_KEY);
  if (!saved) {
    lastExportMeta = null;
    return;
  }

  try {
    const parsed = JSON.parse(saved);
    if (!parsed || typeof parsed !== "object") {
      lastExportMeta = null;
      return;
    }
    lastExportMeta = {
      date: toTrimmedString(parsed.date),
      transactions: Number.parseInt(parsed.transactions, 10) || 0,
      accounts: Number.parseInt(parsed.accounts, 10) || 0
    };
  } catch {
    lastExportMeta = null;
  }
}

function saveLastExportMeta() {
  if (!lastExportMeta) {
    localStorage.removeItem(EXPORT_META_STORAGE_KEY);
    return;
  }

  localStorage.setItem(EXPORT_META_STORAGE_KEY, JSON.stringify(lastExportMeta));
}

function getCategoryListByType(type) {
  if (type !== "expense" && type !== "income") return [];
  if (!Array.isArray(categoryOptions[type])) {
    categoryOptions[type] = [];
  }
  return categoryOptions[type];
}

function ensureCategoryOption(type, rawCategory, shouldSave = true) {
  if (type !== "expense" && type !== "income") return "";

  const clean = normalizeCategoryName(rawCategory);
  if (!clean) return "";

  const list = getCategoryListByType(type);
  const existing = list.find((item) => item.toLowerCase() === clean.toLowerCase());
  if (existing) return existing;

  list.push(clean);
  if (shouldSave) {
    saveCategoryOptions();
  }

  return clean;
}

function getTransactionCategoryOrUnknown(type, rawCategory, shouldSave = false) {
  if (type !== "expense" && type !== "income") return "";
  const clean = normalizeCategoryName(rawCategory);
  return ensureCategoryOption(type, clean || UNKNOWN_CATEGORY, shouldSave);
}

function syncCategoriesFromTransactions() {
  let changed = false;
  transactions.forEach((tx) => {
    if (tx.type !== "expense" && tx.type !== "income") return;
    const beforeLength = getCategoryListByType(tx.type).length;
    ensureCategoryOption(tx.type, tx.category, false);
    if (getCategoryListByType(tx.type).length !== beforeLength) {
      changed = true;
    }
  });

  if (changed) {
    saveCategoryOptions();
  }
}

function isCategoryUsedByTransactions(type, categoryName) {
  if (type !== "expense" && type !== "income") return false;
  const clean = normalizeCategoryName(categoryName);
  if (!clean) return false;
  const key = clean.toLowerCase();

  return transactions.some((tx) => tx.type === type && normalizeCategoryName(tx.category).toLowerCase() === key);
}

function isValidDateInput(value) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(value))) return false;
  return Number.isFinite(Date.parse(`${value}T12:00:00`));
}

function getTodayDateInputValue() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const day = String(now.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function formatDateInputFromDate(date, useUtc = false) {
  const year = useUtc ? date.getUTCFullYear() : date.getFullYear();
  const month = String((useUtc ? date.getUTCMonth() : date.getMonth()) + 1).padStart(2, "0");
  const day = String(useUtc ? date.getUTCDate() : date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function normalizeImportedDate(rawDate) {
  if (rawDate instanceof Date && Number.isFinite(rawDate.getTime())) {
    return formatDateInputFromDate(rawDate);
  }

  if (typeof rawDate === "number" && Number.isFinite(rawDate)) {
    const excelEpochOffset = 25569;
    const msPerDay = 86400 * 1000;
    const date = new Date((rawDate - excelEpochOffset) * msPerDay);
    const formatted = formatDateInputFromDate(date, true);
    return isValidDateInput(formatted) ? formatted : getTodayDateInputValue();
  }

  const text = toTrimmedString(rawDate);
  if (!text) return getTodayDateInputValue();
  if (isValidDateInput(text)) return text;

  const isoLikeMatch = text.match(/^(\d{4})[\/.\-](\d{1,2})[\/.\-](\d{1,2})$/);
  if (isoLikeMatch) {
    const year = Number.parseInt(isoLikeMatch[1], 10);
    const month = Number.parseInt(isoLikeMatch[2], 10);
    const day = Number.parseInt(isoLikeMatch[3], 10);
    const normalized = `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
    if (isValidDateInput(normalized)) return normalized;
  }

  const ukMatch = text.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})$/);
  if (ukMatch) {
    const day = Number.parseInt(ukMatch[1], 10);
    const month = Number.parseInt(ukMatch[2], 10);
    let year = Number.parseInt(ukMatch[3], 10);
    if (year < 100) year += year < 70 ? 2000 : 1900;
    const normalized = `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
    if (isValidDateInput(normalized)) return normalized;
  }

  const parsed = Date.parse(text);
  if (Number.isFinite(parsed)) return formatDateInputFromDate(new Date(parsed));
  return getTodayDateInputValue();
}

function dateToTimestamp(dateValue) {
  const fallback = Date.parse(`${getTodayDateInputValue()}T12:00:00`);
  const parsed = Date.parse(`${dateValue}T12:00:00`);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function formatDateForDisplay(dateValue) {
  if (!isValidDateInput(dateValue)) return "Unknown date";
  const date = new Date(`${dateValue}T12:00:00`);
  return date.toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" });
}

function normalizeBreakdownType(rawType) {
  return rawType === "income" ? "income" : "expense";
}

function getMonthKeyFromDateValue(dateValue) {
  if (!isValidDateInput(dateValue)) return "";
  return String(dateValue).slice(0, 7);
}

function formatMonthKeyForDisplay(monthKey) {
  if (!/^\d{4}-\d{2}$/.test(monthKey)) return monthKey;
  const parsed = Date.parse(`${monthKey}-01T12:00:00`);
  if (!Number.isFinite(parsed)) return monthKey;
  return new Date(parsed).toLocaleDateString("en-GB", { month: "long", year: "numeric" });
}

function formatTimeTick(timestampMs, minX, maxX) {
  const spanMs = maxX - minX;
  const date = new Date(timestampMs);

  if (spanMs <= 36 * 60 * 60 * 1000) {
    return date.toLocaleTimeString("en-GB", { hour: "2-digit", minute: "2-digit" });
  }
  if (spanMs <= 120 * 24 * 60 * 60 * 1000) {
    return date.toLocaleDateString("en-GB", { day: "2-digit", month: "short" });
  }
  return date.toLocaleDateString("en-GB", { month: "short", year: "numeric" });
}

function generateId() {
  return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function saveAccounts() {
  localStorage.setItem(ACCOUNTS_STORAGE_KEY, JSON.stringify(accounts));
}

function saveTransactions() {
  localStorage.setItem(TRANSACTIONS_STORAGE_KEY, JSON.stringify(transactions));
}

function persistAll() {
  saveAccounts();
  saveTransactions();
  saveCategoryOptions();
}

function loadAccounts() {
  const saved = localStorage.getItem(ACCOUNTS_STORAGE_KEY);
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

    accounts = parsed.map((account, index) => {
      const id = typeof account?.id === "string" ? account.id : String(index);
      const name = toTrimmedString(account?.name) || "Unnamed account";
      const initialBalance = getSafeAmount(account?.initialBalance ?? account?.balance ?? 0);
      const createdDate = isValidDateInput(account?.createdDate) ? account.createdDate : getTodayDateInputValue();
      const archived = Boolean(account?.archived);
      return { id, name, initialBalance, createdDate, archived };
    });
  } catch {
    accounts = [];
  }
}

function getActiveAccounts() {
  return accounts.filter((account) => !account.archived);
}

function getArchivedAccounts() {
  return accounts.filter((account) => account.archived);
}

function loadTransactions() {
  const saved = localStorage.getItem(TRANSACTIONS_STORAGE_KEY);
  if (!saved) {
    transactions = [];
    return;
  }

  try {
    const parsed = JSON.parse(saved);
    if (!Array.isArray(parsed)) {
      transactions = [];
      return;
    }

    transactions = parsed
      .map((tx, index) => {
        const type = normalizeTransactionType(tx?.type);
        if (!type) return null;

        const amount = getSafeAmount(tx?.amount);
        if (amount <= 0) return null;

        const date = normalizeImportedDate(tx?.date);
        const parsedCreatedAt = Date.parse(tx?.createdAt);
        const createdAt = Number.isFinite(parsedCreatedAt) ? new Date(parsedCreatedAt).toISOString() : new Date().toISOString();

        const note = toTrimmedString(tx?.note);
        const category = normalizeCategoryName(tx?.category);
        const currency = toTrimmedString(tx?.currency);

        if (type === "transfer") {
          const fromAccountId = toTrimmedString(tx?.fromAccountId);
          const toAccountId = toTrimmedString(tx?.toAccountId);
          if (!fromAccountId || !toAccountId || fromAccountId === toAccountId) return null;

          return {
            id: typeof tx?.id === "string" ? tx.id : `legacy-${index}`,
            type,
            date,
            amount,
            accountId: "",
            fromAccountId,
            toAccountId,
            note,
            category: "",
            currency,
            createdAt
          };
        }

        const legacyAccount = type === "expense" ? tx?.fromAccountId : tx?.toAccountId;
        const accountId = toTrimmedString(tx?.accountId || legacyAccount);
        if (!accountId) return null;

        return {
          id: typeof tx?.id === "string" ? tx.id : `legacy-${index}`,
          type,
          date,
          amount,
          accountId,
          fromAccountId: "",
          toAccountId: "",
          note,
          category: getTransactionCategoryOrUnknown(type, category, false),
          currency,
          createdAt
        };
      })
      .filter((tx) => tx !== null);
  } catch {
    transactions = [];
  }
}

function sortTransactionsAscending() {
  return [...transactions].sort((a, b) => {
    const byDate = dateToTimestamp(a.date) - dateToTimestamp(b.date);
    if (byDate !== 0) return byDate;

    const byCreatedAt = Date.parse(a.createdAt) - Date.parse(b.createdAt);
    if (byCreatedAt !== 0) return byCreatedAt;
    return a.id.localeCompare(b.id);
  });
}

function sortTransactionsDescending() {
  return [...sortTransactionsAscending()].reverse();
}

function getAccountNameById(accountId) {
  if (accountId === "total") return "All Accounts";
  const account = accounts.find((item) => item.id === accountId);
  return account ? account.name : "Deleted account";
}

function getBalanceMap() {
  const balances = {};
  accounts.forEach((account) => {
    balances[account.id] = account.initialBalance;
  });

  const accountIds = new Set(accounts.map((account) => account.id));
  sortTransactionsAscending().forEach((tx) => {
    if (tx.type === "expense" && accountIds.has(tx.accountId)) {
      balances[tx.accountId] = normalizeMoney((balances[tx.accountId] || 0) - tx.amount);
    }

    if (tx.type === "income" && accountIds.has(tx.accountId)) {
      balances[tx.accountId] = normalizeMoney((balances[tx.accountId] || 0) + tx.amount);
    }

    if (tx.type === "transfer") {
      if (accountIds.has(tx.fromAccountId)) {
        balances[tx.fromAccountId] = normalizeMoney((balances[tx.fromAccountId] || 0) - tx.amount);
      }
      if (accountIds.has(tx.toAccountId)) {
        balances[tx.toAccountId] = normalizeMoney((balances[tx.toAccountId] || 0) + tx.amount);
      }
    }
  });

  return balances;
}

function calculateTotalFromBalances(balances) {
  return normalizeMoney(Object.values(balances).reduce((sum, value) => sum + value, 0));
}

function getGraphTransactionChanges(tx, accountIds) {
  const changes = {};

  if (tx.type === "expense" && accountIds.has(tx.accountId)) {
    changes[tx.accountId] = -tx.amount;
  }

  if (tx.type === "income" && accountIds.has(tx.accountId)) {
    changes[tx.accountId] = tx.amount;
  }

  if (tx.type === "transfer") {
    if (accountIds.has(tx.fromAccountId)) {
      changes[tx.fromAccountId] = (changes[tx.fromAccountId] || 0) - tx.amount;
    }
    if (accountIds.has(tx.toAccountId)) {
      changes[tx.toAccountId] = (changes[tx.toAccountId] || 0) + tx.amount;
    }
  }

  return changes;
}

function getGraphInitialBalance(accountId) {
  if (accountId === "total") {
    return normalizeMoney(accounts.reduce((sum, account) => sum + getSafeAmount(account.initialBalance), 0));
  }

  const account = accounts.find((item) => item.id === accountId);
  return getSafeAmount(account ? account.initialBalance : 0);
}

function buildGraphAllPoints(accountId) {
  const points = [];
  const accountIds = new Set(accounts.map((account) => account.id));
  let runningBalance = getGraphInitialBalance(accountId);

  sortTransactionsAscending().forEach((tx) => {
    const changes = getGraphTransactionChanges(tx, accountIds);
    const includePoint = accountId === "total"
      ? Object.keys(changes).length > 0
      : Object.prototype.hasOwnProperty.call(changes, accountId);

    if (!includePoint) return;

    const delta = accountId === "total"
      ? Object.values(changes).reduce((sum, value) => sum + value, 0)
      : (changes[accountId] || 0);

    runningBalance = normalizeMoney(runningBalance + delta);
    points.push({
      x: dateToTimestamp(tx.date),
      y: runningBalance,
      source: "transaction",
      txId: tx.id
    });
  });

  // Keep x coordinates unique so same-day transactions are all clickable.
  let previousTimestamp = -Infinity;
  points.forEach((point) => {
    if (point.x <= previousTimestamp) {
      point.x = previousTimestamp + 1;
    }
    previousTimestamp = point.x;
  });

  return points;
}

function getGraphBalanceBeforeTimestamp(points, initialBalance, timestamp) {
  let balance = initialBalance;
  for (const point of points) {
    if (point.x >= timestamp) break;
    balance = point.y;
  }
  return balance;
}

function getGraphBalanceAtOrBeforeTimestamp(points, initialBalance, timestamp) {
  let balance = initialBalance;
  for (const point of points) {
    if (point.x > timestamp) break;
    balance = point.y;
  }
  return balance;
}

function renderTotal() {
  const balances = getBalanceMap();
  const activeIds = new Set(getActiveAccounts().map((account) => account.id));
  const activeTotal = normalizeMoney(
    Object.entries(balances).reduce((sum, [accountId, amount]) => {
      if (!activeIds.has(accountId)) return sum;
      return sum + amount;
    }, 0)
  );
  document.getElementById("totalBalance").innerText = formatMoney(activeTotal);
}

function renderAccounts() {
  const list = document.getElementById("accountsList");
  const archivedList = document.getElementById("archivedAccountsList");
  const archivedToggleButton = document.getElementById("archivedAccountsToggleBtn");
  list.innerHTML = "";
  if (archivedList) archivedList.innerHTML = "";

  const activeAccounts = getActiveAccounts();
  const archivedAccounts = getArchivedAccounts();
  const balances = getBalanceMap();
  const activeAccountIds = new Set(activeAccounts.map((account) => account.id));

  if (!activeAccountIds.has(expandedAccountId)) {
    expandedAccountId = "";
  }

  const sortedActiveAccounts = [...activeAccounts].sort((a, b) => {
    const balanceA = getSafeAmount(balances[a.id] || 0);
    const balanceB = getSafeAmount(balances[b.id] || 0);
    if (balanceA !== balanceB) {
      return balanceB - balanceA;
    }
    return a.name.localeCompare(b.name, undefined, { sensitivity: "base" });
  });

  if (sortedActiveAccounts.length === 0) {
    list.innerHTML = archivedAccounts.length === 0
      ? '<p class="saved">No accounts yet. Add your first account above.</p>'
      : '<p class="saved">No active accounts. Add a new account to continue.</p>';
  }

  sortedActiveAccounts.forEach((account) => {
    const safeName = escapeHtml(account.name);
    const accountBalance = getSafeAmount(balances[account.id] || 0);
    const canArchive = Math.abs(accountBalance) < 0.0001;
    const isExpanded = expandedAccountId === account.id;
    const accountTransactions = getTransactionsForAccount(account.id);
    const showAllTransactions = Boolean(accountTransactionsShowAllById[account.id]);
    const visibleTransactions = showAllTransactions ? accountTransactions : accountTransactions.slice(0, 5);
    const accountTransactionsMarkup = accountTransactions.length === 0
      ? '<p class="saved account-transactions-empty">No transactions for this account yet.</p>'
      : `
        <div class="account-transactions-list">
          ${visibleTransactions.map((tx) => `
            <button type="button" class="transaction-line-btn account-transaction-line" onclick="event.stopPropagation(); openTransactionModal('${tx.id}')">
              <span class="account-transaction-left">
                <span class="account-transaction-date">${escapeHtml(formatDateForDisplay(tx.date))}</span>
                <span class="account-transaction-label">${escapeHtml(getTransactionCompactLabel(tx))}</span>
              </span>
              <span class="transaction-line-amount ${getAccountTransactionAmountClass(tx, account.id)}">${escapeHtml(getAccountTransactionAmountText(tx, account.id))}</span>
            </button>
          `).join("")}
        </div>
        ${accountTransactions.length > 5
          ? `<div class="row account-transactions-toggle-row"><button class="btn btn-secondary btn-small" type="button" onclick="event.stopPropagation(); toggleAccountTransactionsShowAll('${account.id}')">${showAllTransactions ? "Show Last 5" : "Show All"}</button></div>`
          : ""}
      `;
    const deleteButton = canArchive
      ? `<button class="btn btn-danger btn-small" onclick="event.stopPropagation(); deleteAccount('${account.id}')">Delete</button>`
      : "";
    const card = document.createElement("div");
    card.className = "account-item";
    card.innerHTML = `
      <div
        class="account-header account-summary ${isExpanded ? "expanded" : ""}"
        role="button"
        tabindex="0"
        aria-expanded="${isExpanded}"
        onclick="toggleAccountAdjustPanel('${account.id}')"
        onkeydown="onAccountHeaderKeyDown(event, '${account.id}')"
      >
        <div>
          <h3>${safeName}</h3>
        </div>
        <div class="account-actions">
          <p class="account-balance">${formatMoney(accountBalance)}</p>
          ${deleteButton}
        </div>
      </div>
      <div class="row account-adjust-row ${isExpanded ? "" : "hidden"}" onclick="event.stopPropagation()">
        <input id="amount-${account.id}" class="input amount-input" type="number" step="0.01" placeholder="Amount (+/-)" />
        <button class="btn btn-small" onclick="adjustAccount('${account.id}')">Adjust</button>
        <button class="btn btn-secondary btn-small" onclick="silentAdjust('${account.id}')">Silent Adjust</button>
      </div>
      <div class="account-transactions ${isExpanded ? "" : "hidden"}" onclick="event.stopPropagation()">
        ${accountTransactionsMarkup}
      </div>
    `;
    list.appendChild(card);
  });

  if (archivedToggleButton) {
    const toggleLabel = showArchivedAccounts ? "Hide" : "Show";
    archivedToggleButton.innerText = `Archived Accounts (${archivedAccounts.length}) - ${toggleLabel}`;
  }

  if (!archivedList) return;
  archivedList.classList.toggle("hidden", !showArchivedAccounts);
  if (!showArchivedAccounts) return;

  if (archivedAccounts.length === 0) {
    archivedList.innerHTML = '<p class="saved">No archived accounts.</p>';
    return;
  }

  archivedAccounts.forEach((account) => {
    const card = document.createElement("div");
    card.className = "account-item";
    card.innerHTML = `
      <div class="account-header">
        <div>
          <h3>${escapeHtml(account.name)}</h3>
          <p class="saved">Archived</p>
        </div>
        <div class="account-actions">
          <button class="btn btn-small" onclick="unarchiveAccount('${account.id}')">Unarchive</button>
        </div>
      </div>
    `;
    archivedList.appendChild(card);
  });
}

function populateAccountSelect(selectId, selectedValue = "") {
  const select = document.getElementById(selectId);
  const previousValue = selectedValue || select.value;
  select.innerHTML = "";

  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = "Select account";
  select.appendChild(placeholder);

  getActiveAccounts().forEach((account) => {
    const option = document.createElement("option");
    option.value = account.id;
    option.textContent = account.name;
    select.appendChild(option);
  });

  select.value = Array.from(select.options).some((option) => option.value === previousValue)
    ? previousValue
    : "";
}

function closeCategoryDropdown() {
  const menu = document.getElementById("transactionCategoryMenu");
  const trigger = document.getElementById("transactionCategoryTrigger");
  if (menu) menu.classList.add("hidden");
  if (trigger) trigger.classList.remove("open");
}

function setCategoryTriggerLabel(type, selectedValue = "") {
  const label = document.getElementById("transactionCategoryLabel");
  if (!label) return;

  if (type !== "expense" && type !== "income") {
    label.innerText = "Select category";
    return;
  }

  const clean = normalizeCategoryName(selectedValue || selectedCategoryByType[type] || "");
  label.innerText = clean || "Select category";
}

function deleteCategoryFromType(type, categoryName) {
  if (type !== "expense" && type !== "income") return;

  const selected = normalizeCategoryName(categoryName);
  if (!selected) return;
  if (selected.toLowerCase() === UNKNOWN_CATEGORY.toLowerCase()) {
    window.alert('"Unknown" is a protected category and cannot be deleted.');
    return;
  }

  if (isCategoryUsedByTransactions(type, selected)) {
    window.alert("You cannot delete this category because it is already used by at least one transaction.");
    return;
  }

  const confirmed = window.confirm(`Delete category "${selected}" from ${type} categories?`);
  if (!confirmed) return;

  const key = selected.toLowerCase();
  categoryOptions[type] = getCategoryListByType(type).filter((name) => name.toLowerCase() !== key);
  saveCategoryOptions();

  if (normalizeCategoryName(selectedCategoryByType[type]).toLowerCase() === key) {
    selectedCategoryByType[type] = "";
  }

  populateCategorySelect(type, selectedCategoryByType[type] || "");
}

function promptAddCategory(type) {
  if (type !== "expense" && type !== "income") return;

  const entered = window.prompt(`New ${type} category name:`);
  const newCategory = normalizeCategoryName(entered);
  if (!newCategory) return;

  const savedCategory = ensureCategoryOption(type, newCategory, true);
  selectedCategoryByType[type] = savedCategory;
  populateCategorySelect(type, savedCategory);
  closeCategoryDropdown();
}

function selectCategory(type, categoryName) {
  if (type !== "expense" && type !== "income") return;
  selectedCategoryByType[type] = normalizeCategoryName(categoryName);
  populateCategorySelect(type, selectedCategoryByType[type]);
  closeCategoryDropdown();
}

function populateCategorySelect(type, selectedValue = "") {
  const trigger = document.getElementById("transactionCategoryTrigger");
  const menu = document.getElementById("transactionCategoryMenu");
  if (!trigger || !menu) return;

  if (type !== "expense" && type !== "income") {
    trigger.disabled = true;
    menu.innerHTML = "";
    setCategoryTriggerLabel(type, "");
    closeCategoryDropdown();
    return;
  }

  const categories = getCategoryListByType(type);
  const selected = normalizeCategoryName(selectedValue || selectedCategoryByType[type] || "");
  selectedCategoryByType[type] = selected;

  menu.innerHTML = "";

  categories.forEach((category) => {
    const row = document.createElement("div");
    row.className = "category-option-row";

    const optionButton = document.createElement("button");
    optionButton.type = "button";
    optionButton.className = "category-option-btn";
    optionButton.innerText = category;
    if (selected && category.toLowerCase() === selected.toLowerCase()) {
      optionButton.classList.add("active");
    }
    optionButton.addEventListener("click", () => selectCategory(type, category));

    const deleteButton = document.createElement("button");
    deleteButton.type = "button";
    deleteButton.className = "category-delete-btn";
    deleteButton.innerText = "X";
    deleteButton.title = `Delete ${category}`;
    deleteButton.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      deleteCategoryFromType(type, category);
    });

    row.appendChild(optionButton);
    row.appendChild(deleteButton);
    menu.appendChild(row);
  });

  const addRow = document.createElement("button");
  addRow.type = "button";
  addRow.className = "category-add-btn";
  addRow.innerText = "+ Add new category...";
  addRow.addEventListener("click", () => promptAddCategory(type));
  menu.appendChild(addRow);

  setCategoryTriggerLabel(type, selected);
  trigger.disabled = getActiveAccounts().length === 0;
}

function toggleCategoryDropdown() {
  const type = normalizeTransactionType(document.getElementById("transactionTypeInput").value);
  if (type !== "expense" && type !== "income") return;
  if (getActiveAccounts().length === 0) return;

  const menu = document.getElementById("transactionCategoryMenu");
  const trigger = document.getElementById("transactionCategoryTrigger");
  if (!menu || !trigger) return;

  const isHidden = menu.classList.contains("hidden");
  if (!isHidden) {
    closeCategoryDropdown();
    return;
  }

  populateCategorySelect(type, selectedCategoryByType[type] || "");
  menu.classList.remove("hidden");
  trigger.classList.add("open");
}

function renderTransactionForm() {
  populateAccountSelect("transactionAccountInput");
  populateAccountSelect("transactionFromAccountInput");
  populateAccountSelect("transactionToAccountInput");

  const dateInput = document.getElementById("transactionDateInput");
  const addButton = document.getElementById("addTransactionBtn");
  const typeSelect = document.getElementById("transactionTypeInput");
  const categoryTrigger = document.getElementById("transactionCategoryTrigger");
  const openTransactionModalButton = document.getElementById("openAddTransactionBtn");

  if (!dateInput.value) {
    dateInput.value = getTodayDateInputValue();
  }

  const disabled = getActiveAccounts().length === 0;
  if (openTransactionModalButton) {
    openTransactionModalButton.disabled = disabled;
  }
  addButton.disabled = disabled;
  typeSelect.disabled = disabled;
  dateInput.disabled = disabled;
  document.getElementById("transactionAmountInput").disabled = disabled;
  document.getElementById("transactionNoteInput").disabled = disabled;
  document.getElementById("transactionAccountInput").disabled = disabled;
  document.getElementById("transactionFromAccountInput").disabled = disabled;
  document.getElementById("transactionToAccountInput").disabled = disabled;
  if (categoryTrigger) {
    categoryTrigger.disabled = disabled;
  }

  onTransactionTypeChange();
}

function openAddAccountModal() {
  const modal = document.getElementById("addAccountModal");
  if (!modal) return;
  modal.classList.remove("hidden");
}

function closeAddAccountModal() {
  const modal = document.getElementById("addAccountModal");
  if (modal) modal.classList.add("hidden");
}

function onAddAccountModalBackdropClick(event) {
  if (event.target && event.target.id === "addAccountModal") {
    closeAddAccountModal();
  }
}

function openAddTransactionModal() {
  if (getActiveAccounts().length === 0) {
    window.alert("Add an account first.");
    return;
  }

  renderTransactionForm();
  const modal = document.getElementById("addTransactionModal");
  if (modal) modal.classList.remove("hidden");
}

function closeAddTransactionModal() {
  closeCategoryDropdown();
  const modal = document.getElementById("addTransactionModal");
  if (modal) modal.classList.add("hidden");
}

function onAddTransactionModalBackdropClick(event) {
  if (event.target && event.target.id === "addTransactionModal") {
    closeAddTransactionModal();
  }
}

function getTransactionCompactLabel(tx) {
  if (tx.type === "transfer") {
    return `${getAccountNameById(tx.fromAccountId)} -> ${getAccountNameById(tx.toAccountId)}`;
  }

  return normalizeCategoryName(tx.category) || UNKNOWN_CATEGORY;
}

function getTransactionCompactAmountText(tx) {
  const amount = formatMoney(tx.amount);
  if (tx.type === "income") return `+ ${amount}`;
  if (tx.type === "expense") return `- ${amount}`;
  return `-> ${amount}`;
}

function getTransactionCompactAmountClass(tx) {
  if (tx.type === "income") return "transaction-line-income";
  if (tx.type === "expense") return "transaction-line-expense";
  return "transaction-line-transfer";
}

function getAccountTransactionAmountText(tx, accountId) {
  if (tx.type === "transfer") {
    if (tx.toAccountId === accountId) return `+ ${formatMoney(tx.amount)}`;
    if (tx.fromAccountId === accountId) return `- ${formatMoney(tx.amount)}`;
  }

  return getTransactionCompactAmountText(tx);
}

function getAccountTransactionAmountClass(tx, accountId) {
  if (tx.type === "transfer") {
    if (tx.toAccountId === accountId) return "transaction-line-income";
    if (tx.fromAccountId === accountId) return "transaction-line-expense";
  }

  return getTransactionCompactAmountClass(tx);
}

function getTransactionsForAccount(accountId) {
  return sortTransactionsDescending().filter((tx) => {
    if (tx.type === "transfer") {
      return tx.fromAccountId === accountId || tx.toAccountId === accountId;
    }
    return tx.accountId === accountId;
  });
}

function getTransactionNetImpact(tx) {
  if (tx.type === "income") return tx.amount;
  if (tx.type === "expense") return -tx.amount;
  return 0;
}

function shouldShowTransactionInRecentHistory(tx) {
  const todayStart = Date.parse(`${getTodayDateInputValue()}T00:00:00`);
  const recentWindowStart = todayStart - (29 * DAY_MS);
  return dateToTimestamp(tx.date) >= recentWindowStart;
}

function updateTransactionHistoryRangeButton() {
  const button = document.getElementById("transactionHistoryRangeBtn");
  if (!button) return;

  button.innerText = transactionHistoryShowAll ? "Show Last 30 Days" : "Show All";
  button.disabled = transactions.length === 0;
}

function closeTransactionModal() {
  const closingTransactionId = activeTransactionModalId;
  activeTransactionModalId = "";
  transactionModalEditMode = false;
  transactionModalContext = null;
  if (closingTransactionId && closingTransactionId === activeGraphPointTransactionId) {
    activeGraphPointTransactionId = "";
  }
  const modal = document.getElementById("transactionModal");
  const modalBody = document.getElementById("transactionModalBody");
  const deleteButton = document.getElementById("transactionModalDeleteBtn");
  const saveButton = document.getElementById("transactionModalSaveBtn");
  const title = document.getElementById("transactionModalTitle");
  if (modal) modal.classList.add("hidden");
  if (modalBody) modalBody.innerHTML = "";
  if (deleteButton) deleteButton.disabled = true;
  if (saveButton) {
    saveButton.disabled = true;
    saveButton.innerText = "Edit Transaction";
  }
  if (title) title.innerText = "Transaction Details";
}

function onTransactionModalBackdropClick(event) {
  if (event.target && event.target.id === "transactionModal") {
    closeTransactionModal();
  }
}

function populateTransactionEditAccountSelect(selectId, selectedValue = "") {
  const select = document.getElementById(selectId);
  if (!select) return;

  const previousValue = toTrimmedString(selectedValue);
  const sortedAccounts = [...accounts].sort((a, b) => {
    if (a.archived !== b.archived) return a.archived ? 1 : -1;
    return a.name.localeCompare(b.name, undefined, { sensitivity: "base" });
  });

  select.innerHTML = "";

  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = "Select account";
  select.appendChild(placeholder);

  sortedAccounts.forEach((account) => {
    const option = document.createElement("option");
    option.value = account.id;
    option.textContent = account.archived ? `${account.name} (Archived)` : account.name;
    select.appendChild(option);
  });

  const hasMatch = Array.from(select.options).some((option) => option.value === previousValue);
  if (!hasMatch && previousValue) {
    const fallback = document.createElement("option");
    fallback.value = previousValue;
    fallback.textContent = "Deleted account";
    select.appendChild(fallback);
  }

  select.value = previousValue;
}

function populateTransactionEditCategorySelect(type, selectedValue = "") {
  const categorySelect = document.getElementById("transactionEditCategoryInput");
  if (!categorySelect) return;

  if (type !== "expense" && type !== "income") {
    categorySelect.innerHTML = "";
    categorySelect.disabled = true;
    return;
  }

  const unknownCategory = ensureCategoryOption(type, UNKNOWN_CATEGORY, false);
  const list = getCategoryListByType(type);
  const selected = normalizeCategoryName(selectedValue);
  categorySelect.innerHTML = "";

  list.forEach((category) => {
    const option = document.createElement("option");
    option.value = category;
    option.textContent = category;
    categorySelect.appendChild(option);
  });

  if (selected && !list.some((item) => item.toLowerCase() === selected.toLowerCase())) {
    const customOption = document.createElement("option");
    customOption.value = selected;
    customOption.textContent = selected;
    categorySelect.appendChild(customOption);
  }

  categorySelect.disabled = false;
  categorySelect.value = selected || unknownCategory;
}

function syncTransactionEditTypeUI(preferredCategory = "") {
  const typeInput = document.getElementById("transactionEditTypeInput");
  const accountGroup = document.getElementById("transactionEditAccountGroup");
  const transferGroup = document.getElementById("transactionEditTransferGroup");
  const categoryGroup = document.getElementById("transactionEditCategoryGroup");
  const accountInput = document.getElementById("transactionEditAccountInput");
  const fromInput = document.getElementById("transactionEditFromAccountInput");
  const toInput = document.getElementById("transactionEditToAccountInput");
  if (!typeInput || !accountGroup || !transferGroup || !categoryGroup) return;

  const type = normalizeTransactionType(typeInput.value);
  const isTransfer = type === "transfer";

  accountGroup.classList.toggle("hidden", isTransfer);
  transferGroup.classList.toggle("hidden", !isTransfer);
  categoryGroup.classList.toggle("hidden", isTransfer);

  if (accountInput) accountInput.disabled = isTransfer;
  if (fromInput) fromInput.disabled = !isTransfer;
  if (toInput) toInput.disabled = !isTransfer;

  populateTransactionEditCategorySelect(type, preferredCategory);
}

function onTransactionEditTypeChange() {
  syncTransactionEditTypeUI("");
}

function getTransactionTypeLabel(type) {
  if (type === "expense") return "Expense";
  if (type === "income") return "Income";
  if (type === "transfer") return "Transfer";
  return "Unknown";
}

function getTransactionModalRows(tx) {
  const rows = [
    { label: "Type", value: getTransactionTypeLabel(tx.type) },
    { label: "Date", value: formatDateForDisplay(tx.date) },
    { label: "Amount", value: getTransactionCompactAmountText(tx) }
  ];

  if (transactionModalContext && Number.isFinite(transactionModalContext.balanceAfter)) {
    rows.unshift({
      label: transactionModalContext.balanceLabel || "Balance after",
      value: formatMoney(transactionModalContext.balanceAfter)
    });
  }

  if (tx.type === "transfer") {
    rows.push({ label: "From account", value: getAccountNameById(tx.fromAccountId) });
    rows.push({ label: "To account", value: getAccountNameById(tx.toAccountId) });
  } else {
    rows.push({ label: "Account", value: getAccountNameById(tx.accountId) });
    rows.push({ label: "Category", value: normalizeCategoryName(tx.category) || UNKNOWN_CATEGORY });
  }

  rows.push({ label: "Currency", value: tx.currency ? tx.currency.toUpperCase() : "GBP" });
  rows.push({ label: "Note", value: tx.note ? tx.note : "No note" });
  return rows;
}

function renderTransactionModalContent() {
  if (!activeTransactionModalId) return;

  const tx = transactions.find((item) => item.id === activeTransactionModalId);
  if (!tx) {
    closeTransactionModal();
    return;
  }

  const modalBody = document.getElementById("transactionModalBody");
  const deleteButton = document.getElementById("transactionModalDeleteBtn");
  const saveButton = document.getElementById("transactionModalSaveBtn");
  const title = document.getElementById("transactionModalTitle");
  if (!modalBody || !deleteButton || !saveButton) return;

  deleteButton.disabled = false;
  saveButton.disabled = false;
  saveButton.innerText = transactionModalEditMode ? "Save Changes" : "Edit Transaction";
  if (title) title.innerText = transactionModalEditMode ? "Edit Transaction" : "Transaction Details";

  if (!transactionModalEditMode) {
    const rows = getTransactionModalRows(tx);
    modalBody.innerHTML = rows
      .map((row) => `<p class="transaction-modal-line"><span class="transaction-modal-label">${escapeHtml(row.label)}</span><span>${escapeHtml(row.value)}</span></p>`)
      .join("");
    return;
  }

  modalBody.innerHTML = `
    <div class="stack">
      <label class="label" for="transactionEditTypeInput">Type</label>
      <select id="transactionEditTypeInput" class="input" onchange="onTransactionEditTypeChange()">
        <option value="expense">Expense</option>
        <option value="income">Income</option>
        <option value="transfer">Transfer</option>
      </select>

      <label class="label" for="transactionEditDateInput">Date</label>
      <input id="transactionEditDateInput" class="input" type="date" />

      <label class="label" for="transactionEditAmountInput">Amount</label>
      <input id="transactionEditAmountInput" class="input" type="number" step="0.01" min="0.01" />

      <div id="transactionEditAccountGroup" class="stack">
        <label class="label" for="transactionEditAccountInput">Account</label>
        <select id="transactionEditAccountInput" class="input"></select>
      </div>

      <div id="transactionEditTransferGroup" class="stack hidden">
        <label class="label" for="transactionEditFromAccountInput">From account</label>
        <select id="transactionEditFromAccountInput" class="input"></select>
        <label class="label" for="transactionEditToAccountInput">To account</label>
        <select id="transactionEditToAccountInput" class="input"></select>
      </div>

      <div id="transactionEditCategoryGroup" class="stack">
        <label class="label" for="transactionEditCategoryInput">Category</label>
        <select id="transactionEditCategoryInput" class="input"></select>
      </div>

      <label class="label" for="transactionEditCurrencyInput">Currency</label>
      <input id="transactionEditCurrencyInput" class="input" placeholder="GBP" />

      <label class="label" for="transactionEditNoteInput">Note</label>
      <input id="transactionEditNoteInput" class="input" placeholder="Optional note" />
    </div>
  `;

  populateTransactionEditAccountSelect("transactionEditAccountInput", tx.accountId);
  populateTransactionEditAccountSelect("transactionEditFromAccountInput", tx.fromAccountId);
  populateTransactionEditAccountSelect("transactionEditToAccountInput", tx.toAccountId);

  const typeInput = document.getElementById("transactionEditTypeInput");
  const dateInput = document.getElementById("transactionEditDateInput");
  const amountInput = document.getElementById("transactionEditAmountInput");
  const noteInput = document.getElementById("transactionEditNoteInput");
  const currencyInput = document.getElementById("transactionEditCurrencyInput");
  const accountInput = document.getElementById("transactionEditAccountInput");
  const fromInput = document.getElementById("transactionEditFromAccountInput");
  const toInput = document.getElementById("transactionEditToAccountInput");

  if (typeInput) typeInput.value = tx.type;
  if (dateInput) dateInput.value = normalizeImportedDate(tx.date);
  if (amountInput) amountInput.value = String(tx.amount);
  if (noteInput) noteInput.value = tx.note || "";
  if (currencyInput) currencyInput.value = (tx.currency || "GBP").toUpperCase();
  if (accountInput) accountInput.value = tx.accountId || "";
  if (fromInput) fromInput.value = tx.fromAccountId || "";
  if (toInput) toInput.value = tx.toAccountId || "";
  syncTransactionEditTypeUI(tx.category || "");
}

function onTransactionModalEditOrSave() {
  if (!activeTransactionModalId) return;

  if (!transactionModalEditMode) {
    transactionModalEditMode = true;
    renderTransactionModalContent();
    return;
  }

  saveTransactionFromModal();
}

function openTransactionModal(transactionId, context = null) {
  const tx = transactions.find((item) => item.id === transactionId);
  if (!tx) {
    closeTransactionModal();
    return;
  }

  const modal = document.getElementById("transactionModal");
  if (!modal) return;

  activeTransactionModalId = tx.id;
  transactionModalEditMode = false;
  transactionModalContext = context && typeof context === "object"
    ? {
        balanceAfter: Number(context.balanceAfter),
        balanceLabel: toTrimmedString(context.balanceLabel)
      }
    : null;
  if (!transactionModalContext) {
    activeGraphPointTransactionId = "";
  }
  renderTransactionModalContent();
  modal.classList.remove("hidden");
}

function saveTransactionFromModal() {
  if (!activeTransactionModalId) return;

  const txIndex = transactions.findIndex((item) => item.id === activeTransactionModalId);
  if (txIndex < 0) {
    closeTransactionModal();
    return;
  }

  const typeInput = document.getElementById("transactionEditTypeInput");
  const dateInput = document.getElementById("transactionEditDateInput");
  const amountInput = document.getElementById("transactionEditAmountInput");
  const noteInput = document.getElementById("transactionEditNoteInput");
  const currencyInput = document.getElementById("transactionEditCurrencyInput");
  const accountInput = document.getElementById("transactionEditAccountInput");
  const fromInput = document.getElementById("transactionEditFromAccountInput");
  const toInput = document.getElementById("transactionEditToAccountInput");
  const categoryInput = document.getElementById("transactionEditCategoryInput");
  if (!typeInput || !dateInput || !amountInput || !noteInput || !currencyInput || !accountInput || !fromInput || !toInput || !categoryInput) {
    return;
  }

  const type = normalizeTransactionType(typeInput.value);
  const amount = getSafeAmount(amountInput.value);
  if (!type) {
    window.alert("Choose a transaction type.");
    return;
  }
  if (amount <= 0) {
    window.alert("Amount must be greater than 0.");
    return;
  }

  const date = normalizeImportedDate(dateInput.value);
  const note = toTrimmedString(noteInput.value);
  const currency = toTrimmedString(currencyInput.value).toUpperCase() || "GBP";
  const accountIds = new Set(accounts.map((account) => account.id));
  const tx = transactions[txIndex];

  if (type === "transfer") {
    const fromAccountId = toTrimmedString(fromInput.value);
    const toAccountId = toTrimmedString(toInput.value);
    if (!accountIds.has(fromAccountId) || !accountIds.has(toAccountId) || fromAccountId === toAccountId) {
      window.alert("Choose different valid from/to accounts.");
      return;
    }

    tx.type = type;
    tx.date = date;
    tx.amount = amount;
    tx.accountId = "";
    tx.fromAccountId = fromAccountId;
    tx.toAccountId = toAccountId;
    tx.category = "";
    tx.note = note;
    tx.currency = currency;
  } else {
    const accountId = toTrimmedString(accountInput.value);
    if (!accountIds.has(accountId)) {
      window.alert("Choose a valid account.");
      return;
    }

    const categoryValue = normalizeCategoryName(categoryInput.value);
    if (!categoryValue) {
      window.alert('Choose a category. Use "Unknown" if needed.');
      return;
    }

    const category = ensureCategoryOption(type, categoryValue, true);
    tx.type = type;
    tx.date = date;
    tx.amount = amount;
    tx.accountId = accountId;
    tx.fromAccountId = "";
    tx.toAccountId = "";
    tx.category = category;
    tx.note = note;
    tx.currency = currency;
  }

  saveTransactions();
  render();
  closeTransactionModal();
}

function deleteTransactionFromModal() {
  if (!activeTransactionModalId) return;
  const transactionId = activeTransactionModalId;
  closeTransactionModal();
  deleteTransaction(transactionId);
}

function renderTransactions() {
  const list = document.getElementById("transactionsList");
  list.innerHTML = "";
  updateTransactionHistoryRangeButton();

  if (activeTransactionModalId && !transactions.some((tx) => tx.id === activeTransactionModalId)) {
    closeTransactionModal();
  }

  if (transactions.length === 0) {
    list.innerHTML = '<p class="saved">No transactions yet.</p>';
    return;
  }

  const historyTransactions = sortTransactionsDescending();
  const visibleTransactions = transactionHistoryShowAll
    ? historyTransactions
    : historyTransactions.filter((tx) => shouldShowTransactionInRecentHistory(tx));

  if (visibleTransactions.length === 0) {
    list.innerHTML = '<p class="saved">No transactions in the last 30 days. Tap "Show All" to view full history.</p>';
    return;
  }

  const groupedByDate = new Map();
  visibleTransactions.forEach((tx) => {
    if (!groupedByDate.has(tx.date)) {
      groupedByDate.set(tx.date, []);
    }
    groupedByDate.get(tx.date).push(tx);
  });

  groupedByDate.forEach((items, dateValue) => {
    const group = document.createElement("div");
    group.className = "transaction-date-group";

    const dayNet = normalizeMoney(items.reduce((sum, tx) => sum + getTransactionNetImpact(tx), 0));

    const heading = document.createElement("div");
    heading.className = "transaction-date-heading";
    const headingDate = document.createElement("span");
    headingDate.className = "transaction-date-title";
    headingDate.innerText = formatDateForDisplay(dateValue);

    const headingNet = document.createElement("span");
    headingNet.className = "transaction-date-net";
    headingNet.innerText = formatSignedMoney(dayNet);

    heading.appendChild(headingDate);
    heading.appendChild(headingNet);
    group.appendChild(heading);

    const lines = document.createElement("div");
    lines.className = "transaction-lines";

    items.forEach((tx) => {
      const lineButton = document.createElement("button");
      lineButton.type = "button";
      lineButton.className = "transaction-line-btn";
      lineButton.addEventListener("click", () => openTransactionModal(tx.id));

      const label = document.createElement("span");
      label.className = "transaction-line-label";
      label.innerText = getTransactionCompactLabel(tx);

      const amount = document.createElement("span");
      amount.className = `transaction-line-amount ${getTransactionCompactAmountClass(tx)}`;
      amount.innerText = getTransactionCompactAmountText(tx);

      lineButton.appendChild(label);
      lineButton.appendChild(amount);
      lines.appendChild(lineButton);
    });

    group.appendChild(lines);
    list.appendChild(group);
  });
}

function getCategoryBreakdownAvailableMonths(type) {
  const months = new Set();
  sortTransactionsDescending().forEach((tx) => {
    if (tx.type !== type) return;
    const monthKey = getMonthKeyFromDateValue(tx.date);
    if (!monthKey) return;
    months.add(monthKey);
  });
  return Array.from(months);
}

function getCategoryBreakdownData(type, monthKey) {
  const totals = new Map();
  transactions.forEach((tx) => {
    if (tx.type !== type) return;
    if (getMonthKeyFromDateValue(tx.date) !== monthKey) return;

    const category = normalizeCategoryName(tx.category) || UNKNOWN_CATEGORY;
    totals.set(category, normalizeMoney((totals.get(category) || 0) + tx.amount));
  });

  return Array.from(totals.entries())
    .map(([category, total]) => ({ category, total }))
    .sort((a, b) => {
      if (b.total !== a.total) return b.total - a.total;
      return a.category.localeCompare(b.category);
    });
}

function getCategoryBreakdownTransactions(type, monthKey, categoryName) {
  const targetCategory = normalizeCategoryName(categoryName) || UNKNOWN_CATEGORY;
  return sortTransactionsDescending().filter((tx) => {
    if (tx.type !== type) return false;
    if (getMonthKeyFromDateValue(tx.date) !== monthKey) return false;
    const txCategory = normalizeCategoryName(tx.category) || UNKNOWN_CATEGORY;
    return txCategory.toLowerCase() === targetCategory.toLowerCase();
  });
}

function onCategoryBreakdownLegendClick(categoryName) {
  const normalized = normalizeCategoryName(categoryName) || UNKNOWN_CATEGORY;
  if (!normalized) return;

  if (selectedCategoryBreakdownCategory === normalized) {
    selectedCategoryBreakdownCategory = "";
    categoryBreakdownTransactionsShowAll = false;
  } else {
    selectedCategoryBreakdownCategory = normalized;
    categoryBreakdownTransactionsShowAll = false;
  }

  renderCategoryBreakdown();
}

function onCategoryBreakdownLegendContainerClick(event) {
  const legend = document.getElementById("categoryBreakdownLegend");
  if (!legend) return;

  const targetButton = event.target instanceof Element
    ? event.target.closest(".category-breakdown-legend-btn")
    : null;
  if (!targetButton || !legend.contains(targetButton)) return;

  const categoryName = normalizeCategoryName(targetButton.dataset.category || "");
  if (!categoryName) return;
  onCategoryBreakdownLegendClick(categoryName);
}

function toggleCategoryBreakdownTransactionsShowAll() {
  categoryBreakdownTransactionsShowAll = !categoryBreakdownTransactionsShowAll;
  renderCategoryBreakdown();
}

function getCategoryBreakdownColor(index) {
  return CATEGORY_BREAKDOWN_COLORS[index % CATEGORY_BREAKDOWN_COLORS.length];
}

function drawCategoryBreakdownEmptyChart(context, width, height, message) {
  context.clearRect(0, 0, width, height);
  context.fillStyle = "#6b6b6b";
  context.font = "15px Segoe UI, Arial, sans-serif";
  context.textAlign = "center";
  context.textBaseline = "middle";
  context.fillText(message, width / 2, height / 2);
}

function renderCategoryBreakdown() {
  const typeSelect = document.getElementById("categoryBreakdownTypeInput");
  const monthSelect = document.getElementById("categoryBreakdownMonthInput");
  const canvas = document.getElementById("categoryBreakdownChart");
  const summary = document.getElementById("categoryBreakdownSummary");
  const legend = document.getElementById("categoryBreakdownLegend");
  if (!typeSelect || !monthSelect || !canvas || !summary || !legend) return;

  selectedCategoryBreakdownType = normalizeBreakdownType(selectedCategoryBreakdownType);
  typeSelect.value = selectedCategoryBreakdownType;
  if (selectedCategoryBreakdownMonthByType[selectedCategoryBreakdownType]) {
    selectedCategoryBreakdownMonth = selectedCategoryBreakdownMonthByType[selectedCategoryBreakdownType];
  }

  const availableMonths = getCategoryBreakdownAvailableMonths(selectedCategoryBreakdownType);
  monthSelect.innerHTML = "";
  if (availableMonths.length === 0) {
    selectedCategoryBreakdownMonth = "";
    selectedCategoryBreakdownMonthByType[selectedCategoryBreakdownType] = "";
    selectedCategoryBreakdownCategory = "";
    categoryBreakdownTransactionsShowAll = false;
    const emptyOption = document.createElement("option");
    emptyOption.value = "";
    emptyOption.textContent = "No months available";
    monthSelect.appendChild(emptyOption);
    monthSelect.disabled = true;

    const context = canvas.getContext("2d");
    const cssWidth = Math.max(canvas.clientWidth, 320);
    const cssHeight = 280;
    const pixelRatio = window.devicePixelRatio || 1;
    canvas.width = Math.floor(cssWidth * pixelRatio);
    canvas.height = Math.floor(cssHeight * pixelRatio);
    context.setTransform(pixelRatio, 0, 0, pixelRatio, 0, 0);
    drawCategoryBreakdownEmptyChart(context, cssWidth, cssHeight, "No transactions yet for this type.");
    summary.innerText = `No ${selectedCategoryBreakdownType} transactions available yet.`;
    legend.innerHTML = "";
    return;
  }

  monthSelect.disabled = false;
  availableMonths.forEach((monthKey) => {
    const option = document.createElement("option");
    option.value = monthKey;
    option.textContent = formatMonthKeyForDisplay(monthKey);
    monthSelect.appendChild(option);
  });

  if (!availableMonths.includes(selectedCategoryBreakdownMonth)) {
    const savedForType = selectedCategoryBreakdownMonthByType[selectedCategoryBreakdownType];
    selectedCategoryBreakdownMonth = availableMonths.includes(savedForType) ? savedForType : availableMonths[0];
  }
  selectedCategoryBreakdownMonthByType[selectedCategoryBreakdownType] = selectedCategoryBreakdownMonth;
  monthSelect.value = selectedCategoryBreakdownMonth;

  const data = getCategoryBreakdownData(selectedCategoryBreakdownType, selectedCategoryBreakdownMonth);
  const total = normalizeMoney(data.reduce((sum, item) => sum + item.total, 0));
  const monthLabel = formatMonthKeyForDisplay(selectedCategoryBreakdownMonth);

  const context = canvas.getContext("2d");
  const cssWidth = Math.max(canvas.clientWidth, 320);
  const cssHeight = 280;
  const pixelRatio = window.devicePixelRatio || 1;
  canvas.width = Math.floor(cssWidth * pixelRatio);
  canvas.height = Math.floor(cssHeight * pixelRatio);
  context.setTransform(pixelRatio, 0, 0, pixelRatio, 0, 0);
  context.clearRect(0, 0, cssWidth, cssHeight);

  if (data.length === 0 || total <= 0) {
    drawCategoryBreakdownEmptyChart(context, cssWidth, cssHeight, "No category data for this month.");
    summary.innerText = `${selectedCategoryBreakdownType === "expense" ? "Expenses" : "Income"} in ${monthLabel}: ${formatMoney(0)} across 0 categories.`;
    legend.innerHTML = "";
    selectedCategoryBreakdownCategory = "";
    categoryBreakdownTransactionsShowAll = false;
    return;
  }

  const centerX = cssWidth / 2;
  const centerY = cssHeight / 2;
  const radius = Math.min(cssWidth, cssHeight) * 0.34;
  const ringWidth = Math.max(26, radius * 0.45);
  let startAngle = -Math.PI / 2;
  let consumedAngle = 0;

  data.forEach((item, index) => {
    const sweepAngle = index === data.length - 1
      ? (Math.PI * 2 - consumedAngle)
      : ((item.total / total) * Math.PI * 2);
    const endAngle = startAngle + sweepAngle;

    context.strokeStyle = getCategoryBreakdownColor(index);
    context.lineWidth = ringWidth;
    context.lineCap = "butt";
    context.beginPath();
    context.arc(centerX, centerY, radius, startAngle, endAngle);
    context.stroke();

    consumedAngle += sweepAngle;
    startAngle = endAngle;
  });

  context.fillStyle = "#111";
  context.font = "700 16px Segoe UI, Arial, sans-serif";
  context.textAlign = "center";
  context.textBaseline = "middle";
  context.fillText(formatMoney(total), centerX, centerY - 8);
  context.fillStyle = "#666";
  context.font = "12px Segoe UI, Arial, sans-serif";
  context.fillText(monthLabel, centerX, centerY + 12);

  const typeLabel = selectedCategoryBreakdownType === "expense" ? "Expenses" : "Income";
  summary.innerText = `${typeLabel} in ${monthLabel}: ${formatMoney(total)} across ${data.length} categor${data.length === 1 ? "y" : "ies"}.`;

  const availableCategoryKeys = new Set(
    data.map((item) => (normalizeCategoryName(item.category) || UNKNOWN_CATEGORY).toLowerCase())
  );
  if (!availableCategoryKeys.has((normalizeCategoryName(selectedCategoryBreakdownCategory) || "").toLowerCase())) {
    selectedCategoryBreakdownCategory = "";
    categoryBreakdownTransactionsShowAll = false;
  }

  legend.innerHTML = "";
  data.forEach((item, index) => {
    const percentage = total > 0 ? ((item.total / total) * 100).toFixed(1) : "0.0";
    const valueText = `${formatMoney(item.total)} (${percentage}%)`;
    const normalizedCategory = normalizeCategoryName(item.category) || UNKNOWN_CATEGORY;
    const isActive = selectedCategoryBreakdownCategory === normalizedCategory;
    const wrapper = document.createElement("div");
    wrapper.className = `category-breakdown-legend-item${isActive ? " active" : ""}`;

    const rowButton = document.createElement("button");
    rowButton.type = "button";
    rowButton.className = `category-breakdown-legend-row category-breakdown-legend-btn${isActive ? " active" : ""}`;
    rowButton.dataset.category = normalizedCategory;

    const left = document.createElement("div");
    left.className = "category-breakdown-legend-left";

    const color = document.createElement("span");
    color.className = "category-breakdown-color";
    color.style.background = getCategoryBreakdownColor(index);
    const name = document.createElement("span");
    name.className = "category-breakdown-name";
    name.innerText = normalizedCategory;
    left.appendChild(color);
    left.appendChild(name);

    const value = document.createElement("span");
    value.className = "category-breakdown-value";
    value.innerText = valueText;

    rowButton.appendChild(left);
    rowButton.appendChild(value);
    wrapper.appendChild(rowButton);

    if (isActive) {
      const categoryTransactions = getCategoryBreakdownTransactions(
        selectedCategoryBreakdownType,
        selectedCategoryBreakdownMonth,
        selectedCategoryBreakdownCategory
      );
      const visibleTransactions = categoryBreakdownTransactionsShowAll
        ? categoryTransactions
        : categoryTransactions.slice(0, 5);

      const details = document.createElement("div");
      details.className = "category-breakdown-legend-details";

      const count = document.createElement("p");
      count.className = "saved category-breakdown-transactions-count";
      count.innerText = `${categoryTransactions.length} transaction${categoryTransactions.length === 1 ? "" : "s"}`;
      details.appendChild(count);

      if (visibleTransactions.length === 0) {
        const empty = document.createElement("p");
        empty.className = "saved account-transactions-empty";
        empty.innerText = "No transactions in this category.";
        details.appendChild(empty);
      } else {
        const txList = document.createElement("div");
        txList.className = "account-transactions-list";
        visibleTransactions.forEach((tx) => {
          const txButton = document.createElement("button");
          txButton.type = "button";
          txButton.className = "transaction-line-btn account-transaction-line";
          txButton.addEventListener("click", () => openTransactionModal(tx.id));

          const leftWrap = document.createElement("span");
          leftWrap.className = "account-transaction-left";

          const date = document.createElement("span");
          date.className = "account-transaction-date";
          date.innerText = formatDateForDisplay(tx.date);
          const label = document.createElement("span");
          label.className = "account-transaction-label";
          label.innerText = getAccountNameById(tx.accountId);
          leftWrap.appendChild(date);
          leftWrap.appendChild(label);

          const amount = document.createElement("span");
          amount.className = `transaction-line-amount ${getTransactionCompactAmountClass(tx)}`;
          amount.innerText = getTransactionCompactAmountText(tx);

          txButton.appendChild(leftWrap);
          txButton.appendChild(amount);
          txList.appendChild(txButton);
        });
        details.appendChild(txList);
      }

      if (categoryTransactions.length > 5) {
        const toggleRow = document.createElement("div");
        toggleRow.className = "row account-transactions-toggle-row";
        const toggleBtn = document.createElement("button");
        toggleBtn.className = "btn btn-secondary btn-small";
        toggleBtn.type = "button";
        toggleBtn.innerText = categoryBreakdownTransactionsShowAll ? "Show Last 5" : "Show All";
        toggleBtn.addEventListener("click", () => toggleCategoryBreakdownTransactionsShowAll());
        toggleRow.appendChild(toggleBtn);
        details.appendChild(toggleRow);
      }

      wrapper.appendChild(details);
    }

    legend.appendChild(wrapper);
  });
}

function renderGraphFilter() {
  const select = document.getElementById("graphAccountFilter");
  const previousSelection = selectedGraphAccountId;
  select.innerHTML = "";

  const totalOption = document.createElement("option");
  totalOption.value = "total";
  totalOption.textContent = "All Accounts (Total)";
  select.appendChild(totalOption);

  accounts.forEach((account) => {
    const option = document.createElement("option");
    option.value = account.id;
    option.textContent = account.name;
    select.appendChild(option);
  });

  const available = Array.from(select.options).map((option) => option.value);
  selectedGraphAccountId = available.includes(previousSelection) ? previousSelection : "total";
  select.value = selectedGraphAccountId;
  syncGraphControlsFromState();
}

function getGraphDateBoundsTimestamps(accountId = selectedGraphAccountId) {
  const points = buildGraphAllPoints(accountId);
  if (points.length === 0) {
    const now = dateToTimestamp(getTodayDateInputValue());
    return { start: now, end: now };
  }

  return {
    start: points[0].x,
    end: points[points.length - 1].x
  };
}

function getGraphDateBounds(accountId = selectedGraphAccountId) {
  const bounds = getGraphDateBoundsTimestamps(accountId);
  return {
    start: formatDateInputFromDate(new Date(bounds.start)),
    end: formatDateInputFromDate(new Date(bounds.end))
  };
}

function timestampToDayStart(timestamp) {
  const date = new Date(timestamp);
  date.setHours(0, 0, 0, 0);
  return date.getTime();
}

function getGraphBoundsDayWindow(accountId = selectedGraphAccountId) {
  const bounds = getGraphDateBoundsTimestamps(accountId);
  return {
    startDay: timestampToDayStart(bounds.start),
    endDay: timestampToDayStart(bounds.end)
  };
}

function getGraphRangeDaysFromInputs() {
  const startInput = document.getElementById("graphDateStart");
  const endInput = document.getElementById("graphDateEnd");
  if (!startInput || !endInput) return null;
  if (!isValidDateInput(startInput.value) || !isValidDateInput(endInput.value)) return null;

  const startDay = Date.parse(`${startInput.value}T00:00:00`);
  const endDay = Date.parse(`${endInput.value}T00:00:00`);
  if (!Number.isFinite(startDay) || !Number.isFinite(endDay) || endDay < startDay) return null;

  const spanDays = Math.floor((endDay - startDay) / DAY_MS) + 1;
  return { startDay, endDay, spanDays: Math.max(1, spanDays) };
}

function clampGraphDateInputsToBounds() {
  const startInput = document.getElementById("graphDateStart");
  const endInput = document.getElementById("graphDateEnd");
  if (!startInput || !endInput) return;

  const boundsDays = getGraphBoundsDayWindow(selectedGraphAccountId);
  const totalDays = Math.max(1, Math.floor((boundsDays.endDay - boundsDays.startDay) / DAY_MS) + 1);

  const range = getGraphRangeDaysFromInputs();
  if (!range) {
    setGraphRangeInputsFromTimestamps(boundsDays.startDay, boundsDays.endDay);
    return;
  }

  const spanDays = Math.min(Math.max(1, range.spanDays), totalDays);
  let startDay = range.startDay;
  if (startDay < boundsDays.startDay) {
    startDay = boundsDays.startDay;
  }

  const latestStart = boundsDays.endDay - (spanDays - 1) * DAY_MS;
  if (startDay > latestStart) {
    startDay = latestStart;
  }
  if (startDay < boundsDays.startDay) {
    startDay = boundsDays.startDay;
  }

  const endDay = startDay + (spanDays - 1) * DAY_MS;
  setGraphRangeInputsFromTimestamps(startDay, endDay);
}

function setGraphRangeInputsFromTimestamps(startTimestamp, endTimestamp) {
  const startInput = document.getElementById("graphDateStart");
  const endInput = document.getElementById("graphDateEnd");
  if (!startInput || !endInput) return;

  startInput.value = formatDateInputFromDate(new Date(startTimestamp));
  endInput.value = formatDateInputFromDate(new Date(endTimestamp));
}

function syncGraphControlsFromState() {
  const checkbox = document.getElementById("graphAllHistoryCheckbox");
  if (checkbox) {
    checkbox.checked = graphShowAllHistory;
  }
  const trendCheckbox = document.getElementById("graphTrendLineCheckbox");
  if (trendCheckbox) {
    trendCheckbox.checked = graphShowTrendLine;
  }

  const startInput = document.getElementById("graphDateStart");
  const endInput = document.getElementById("graphDateEnd");
  const dateRangeGroup = document.getElementById("graphDateRangeGroup");
  const scrollRow = document.getElementById("graphScrollRow");
  if (dateRangeGroup) dateRangeGroup.classList.toggle("hidden", graphShowAllHistory);
  if (scrollRow) scrollRow.classList.toggle("hidden", graphShowAllHistory);
  if (!startInput || !endInput) return;

  const bounds = getGraphDateBounds(selectedGraphAccountId);
  if (graphShowAllHistory) {
    startInput.value = bounds.start;
    endInput.value = bounds.end;
    syncGraphRangeScrollFromCurrentRange();
    return;
  }

  if (!isValidDateInput(startInput.value)) startInput.value = bounds.start;
  if (!isValidDateInput(endInput.value)) endInput.value = bounds.end;
  clampGraphDateInputsToBounds();
  syncGraphRangeScrollFromCurrentRange();
}

function onGraphAllHistoryChange() {
  const checkbox = document.getElementById("graphAllHistoryCheckbox");
  graphShowAllHistory = checkbox ? checkbox.checked : true;
  closeGraphPointModal();
  syncGraphControlsFromState();
  renderGraph();
}

function onGraphTrendLineChange() {
  const checkbox = document.getElementById("graphTrendLineCheckbox");
  graphShowTrendLine = checkbox ? checkbox.checked : true;
  renderGraph();
}

function onGraphDateRangeChange() {
  graphShowAllHistory = false;
  const checkbox = document.getElementById("graphAllHistoryCheckbox");
  if (checkbox) checkbox.checked = false;
  syncGraphControlsFromState();
  closeGraphPointModal();
  renderGraph();
}

function getGraphRangeFilter() {
  const bounds = getGraphDateBoundsTimestamps(selectedGraphAccountId);
  const startInput = document.getElementById("graphDateStart");
  const endInput = document.getElementById("graphDateEnd");
  const startValue = startInput ? startInput.value : "";
  const endValue = endInput ? endInput.value : "";

  if (graphShowAllHistory) {
    return {
      mode: "all",
      invalid: false,
      start: bounds.start,
      end: bounds.end,
      startValue: formatDateInputFromDate(new Date(bounds.start)),
      endValue: formatDateInputFromDate(new Date(bounds.end))
    };
  }

  if (!isValidDateInput(startValue) || !isValidDateInput(endValue)) {
    return { mode: "range", invalid: true, reason: "Choose both start and end dates." };
  }

  const start = Date.parse(`${startValue}T00:00:00`);
  const end = Date.parse(`${endValue}T23:59:59.999`);
  if (!Number.isFinite(start) || !Number.isFinite(end) || end < start) {
    return { mode: "range", invalid: true, reason: "End date must be on or after start date." };
  }

  return { mode: "range", invalid: false, start, end, startValue, endValue };
}

function clampGraphRangeToBounds(start, end, bounds) {
  let nextStart = start;
  let nextEnd = end;
  const maxSpan = Math.max(bounds.end - bounds.start, 24 * 60 * 60 * 1000);
  let span = nextEnd - nextStart;
  if (!Number.isFinite(span) || span <= 0) {
    span = 24 * 60 * 60 * 1000;
    nextEnd = nextStart + span;
  }
  if (span > maxSpan) span = maxSpan;

  if (nextStart < bounds.start) {
    nextStart = bounds.start;
    nextEnd = nextStart + span;
  }
  if (nextEnd > bounds.end) {
    nextEnd = bounds.end;
    nextStart = nextEnd - span;
  }

  if (nextStart < bounds.start) nextStart = bounds.start;
  if (nextEnd > bounds.end) nextEnd = bounds.end;
  if (nextEnd <= nextStart) nextEnd = nextStart + 24 * 60 * 60 * 1000;
  if (nextEnd > bounds.end) nextEnd = bounds.end;
  if (nextStart < bounds.start) nextStart = bounds.start;
  if (nextEnd < nextStart) nextEnd = nextStart;

  return { start: nextStart, end: nextEnd };
}

function getGraphCurrentRangeFromControls() {
  const bounds = getGraphDateBoundsTimestamps(selectedGraphAccountId);
  const rangeFilter = getGraphRangeFilter();
  if (rangeFilter.invalid || graphShowAllHistory) {
    return { start: bounds.start, end: bounds.end, bounds };
  }
  return { start: rangeFilter.start, end: rangeFilter.end, bounds };
}

function syncGraphRangeScrollFromCurrentRange() {
  const slider = document.getElementById("graphRangeScroll");
  if (!slider) return;
  if (graphShowAllHistory) {
    slider.min = "0";
    slider.max = "0";
    slider.step = "1";
    slider.value = "0";
    slider.disabled = true;
    return;
  }

  const range = getGraphRangeDaysFromInputs();
  if (!range) {
    slider.min = "0";
    slider.max = "0";
    slider.step = "1";
    slider.value = "0";
    slider.disabled = true;
    return;
  }

  const boundsDays = getGraphBoundsDayWindow(selectedGraphAccountId);
  const totalDays = Math.max(1, Math.floor((boundsDays.endDay - boundsDays.startDay) / DAY_MS) + 1);
  const spanDays = Math.min(range.spanDays, totalDays);
  const maxOffset = Math.max(0, totalDays - spanDays);
  const currentOffsetRaw = Math.round((range.startDay - boundsDays.startDay) / DAY_MS);
  const currentOffset = Math.min(maxOffset, Math.max(0, currentOffsetRaw));

  slider.min = "0";
  slider.max = String(maxOffset);
  slider.step = "1";
  slider.value = String(currentOffset);
  slider.disabled = maxOffset === 0;
}

function applyGraphRangeFromTimestamps(startTimestamp, endTimestamp) {
  const bounds = getGraphDateBoundsTimestamps(selectedGraphAccountId);
  const clamped = clampGraphRangeToBounds(startTimestamp, endTimestamp, bounds);
  graphShowAllHistory = false;
  const checkbox = document.getElementById("graphAllHistoryCheckbox");
  if (checkbox) checkbox.checked = false;
  setGraphRangeInputsFromTimestamps(clamped.start, clamped.end);
  syncGraphControlsFromState();
  closeGraphPointModal();
  renderGraph();
}

function zoomGraphRange(zoomIn) {
  const { start, end, bounds } = getGraphCurrentRangeFromControls();
  const minSpan = 24 * 60 * 60 * 1000;
  const maxSpan = Math.max(bounds.end - bounds.start, minSpan);
  const span = Math.max(minSpan, end - start);
  const factor = zoomIn ? 0.7 : (1 / 0.7);
  const nextSpan = Math.max(minSpan, Math.min(maxSpan, span * factor));
  if (!zoomIn && graphShowAllHistory && Math.abs(nextSpan - span) < 1) {
    return;
  }

  // Keep the right-hand date fixed while zooming.
  const nextEnd = end;
  const nextStart = nextEnd - nextSpan;
  applyGraphRangeFromTimestamps(nextStart, nextEnd);
}

function onGraphZoomInClick() {
  zoomGraphRange(true);
}

function onGraphZoomOutClick() {
  zoomGraphRange(false);
}

function onGraphRangeScrollChange() {
  if (graphShowAllHistory) return;

  const slider = document.getElementById("graphRangeScroll");
  if (!slider) return;

  const range = getGraphRangeDaysFromInputs();
  if (!range) return;

  const boundsDays = getGraphBoundsDayWindow(selectedGraphAccountId);
  const totalDays = Math.max(1, Math.floor((boundsDays.endDay - boundsDays.startDay) / DAY_MS) + 1);
  const spanDays = Math.min(range.spanDays, totalDays);
  const maxOffset = Math.max(0, totalDays - spanDays);
  if (maxOffset <= 0) return;

  const rawOffset = Number.parseInt(slider.value, 10);
  const clampedOffset = Number.isFinite(rawOffset) ? Math.min(maxOffset, Math.max(0, rawOffset)) : 0;
  const nextStartDay = boundsDays.startDay + clampedOffset * DAY_MS;
  const nextEndDay = nextStartDay + (spanDays - 1) * DAY_MS;
  applyGraphRangeFromTimestamps(nextStartDay, nextEndDay);
}

function closeGraphPointModal() {
  const graphTransactionId = activeGraphPointTransactionId;
  activeGraphPointTransactionId = "";

  if (graphTransactionId && activeTransactionModalId === graphTransactionId) {
    closeTransactionModal();
  }

  const modal = document.getElementById("graphPointModal");
  const body = document.getElementById("graphPointModalBody");
  if (modal) modal.classList.add("hidden");
  if (body) body.innerHTML = "";
}

function onGraphPointModalBackdropClick(event) {
  if (event.target && event.target.id === "graphPointModal") {
    closeGraphPointModal();
  }
}

function showGraphPointModal(point) {
  if (!point || !point.txId) return;

  const tx = transactions.find((item) => item.id === point.txId);
  if (!tx) return;
  activeGraphPointTransactionId = tx.id;

  const balanceLabel = selectedGraphAccountId === "total"
    ? "All accounts balance after"
    : `${getAccountNameById(selectedGraphAccountId)} balance after`;

  openTransactionModal(tx.id, {
    balanceAfter: point.y,
    balanceLabel
  });
}

function onGraphCanvasClick(event) {
  if (activeTab !== "graph") return;
  if (graphPointTargets.length === 0) return;

  const canvas = document.getElementById("balanceChart");
  if (!canvas) return;

  const rect = canvas.getBoundingClientRect();
  const clickX = event.clientX - rect.left;
  const clickY = event.clientY - rect.top;
  const hitRadius = 10;
  const hitRadiusSq = hitRadius * hitRadius;

  let closest = null;
  graphPointTargets.forEach((target) => {
    const dx = target.x - clickX;
    const dy = target.y - clickY;
    const distanceSq = dx * dx + dy * dy;
    if (distanceSq > hitRadiusSq) return;
    if (!closest || distanceSq < closest.distanceSq) {
      closest = { point: target.point, distanceSq };
    }
  });

  if (!closest) return;
  showGraphPointModal(closest.point);
}

function drawEmptyChart(context, width, height) {
  context.clearRect(0, 0, width, height);
  context.fillStyle = "#6b6b6b";
  context.font = "15px Segoe UI, Arial, sans-serif";
  context.textAlign = "center";
  context.textBaseline = "middle";
  context.fillText("Not enough history yet. Add transactions to build a trend.", width / 2, height / 2);
}

function getTrendLineForSeries(series, minX, maxX) {
  if (!Array.isArray(series) || series.length < 2) return null;

  const count = series.length;
  const xMean = series.reduce((sum, point) => sum + point.x, 0) / count;
  const yMean = series.reduce((sum, point) => sum + point.y, 0) / count;

  let numerator = 0;
  let denominator = 0;
  series.forEach((point) => {
    const dx = point.x - xMean;
    numerator += dx * (point.y - yMean);
    denominator += dx * dx;
  });

  if (!Number.isFinite(denominator) || Math.abs(denominator) < 1e-9) return null;
  const slope = numerator / denominator;
  const intercept = yMean - slope * xMean;
  if (!Number.isFinite(slope) || !Number.isFinite(intercept)) return null;

  return {
    yAtMin: slope * minX + intercept,
    yAtMax: slope * maxX + intercept
  };
}

function renderGraph() {
  if (activeTab !== "graph") return;

  const canvas = document.getElementById("balanceChart");
  const summary = document.getElementById("graphSummary");
  const context = canvas.getContext("2d");
  graphPointTargets = [];

  if (activeGraphPointTransactionId && !transactions.some((tx) => tx.id === activeGraphPointTransactionId)) {
    closeGraphPointModal();
  }

  const cssWidth = Math.max(canvas.clientWidth, 320);
  const cssHeight = 280;
  const pixelRatio = window.devicePixelRatio || 1;

  canvas.width = Math.floor(cssWidth * pixelRatio);
  canvas.height = Math.floor(cssHeight * pixelRatio);
  context.setTransform(pixelRatio, 0, 0, pixelRatio, 0, 0);
  context.clearRect(0, 0, cssWidth, cssHeight);

  const rangeFilter = getGraphRangeFilter();
  if (rangeFilter.invalid) {
    drawEmptyChart(context, cssWidth, cssHeight);
    summary.innerText = rangeFilter.reason;
    return;
  }

  const allSeries = buildGraphAllPoints(selectedGraphAccountId);
  const series = rangeFilter.mode === "range"
    ? allSeries.filter((point) => point.x >= rangeFilter.start && point.x <= rangeFilter.end)
    : allSeries;
  const currentBalances = getBalanceMap();
  const currentValue = selectedGraphAccountId === "total"
    ? calculateTotalFromBalances(currentBalances)
    : getSafeAmount(currentBalances[selectedGraphAccountId]);
  const graphInitialBalance = getGraphInitialBalance(selectedGraphAccountId);
  const shownStartTimestamp = rangeFilter.mode === "range"
    ? rangeFilter.start
    : (allSeries.length > 0 ? allSeries[0].x : dateToTimestamp(getTodayDateInputValue()));
  const shownEndTimestamp = rangeFilter.mode === "range"
    ? rangeFilter.end
    : (allSeries.length > 0 ? allSeries[allSeries.length - 1].x : shownStartTimestamp);
  const periodStartBalance = getGraphBalanceBeforeTimestamp(allSeries, graphInitialBalance, shownStartTimestamp);
  const periodEndBalance = getGraphBalanceAtOrBeforeTimestamp(allSeries, graphInitialBalance, shownEndTimestamp);
  const periodChange = normalizeMoney(periodEndBalance - periodStartBalance);

  if (series.length === 0) {
    drawEmptyChart(context, cssWidth, cssHeight);
    const shownLine = rangeFilter.mode === "range"
      ? `Shown: ${formatDateForDisplay(rangeFilter.startValue)} - ${formatDateForDisplay(rangeFilter.endValue)}.`
      : "Shown: all history.";
    const detailLine = rangeFilter.mode === "range"
      ? "No transaction points in this period."
      : "No transaction history to plot yet.";
    summary.innerText = `${getAccountNameById(selectedGraphAccountId)} current: ${formatMoney(currentValue)}.\nChange over shown period: ${formatSignedMoney(periodChange)}.\n${shownLine}\n${detailLine}`;
    return;
  }

  const padding = { top: 14, right: 14, bottom: 40, left: 62 };
  const chartWidth = cssWidth - padding.left - padding.right;
  const chartHeight = cssHeight - padding.top - padding.bottom;
  if (chartWidth <= 0 || chartHeight <= 0) return;

  const minX = Math.min(...series.map((point) => point.x));
  let maxX = Math.max(...series.map((point) => point.x));
  if (minX === maxX) maxX += 1;

  let minY = Math.min(...series.map((point) => point.y));
  let maxY = Math.max(...series.map((point) => point.y));
  minY = Math.min(minY, 0);
  maxY = Math.max(maxY, 0);
  if (minY === maxY) {
    minY -= 1;
    maxY += 1;
  }

  const toX = (value) => padding.left + ((value - minX) / (maxX - minX)) * chartWidth;
  const toY = (value) => padding.top + ((maxY - value) / (maxY - minY)) * chartHeight;

  context.strokeStyle = "#ececec";
  context.lineWidth = 1;
  const yTicks = 4;
  for (let tick = 0; tick <= yTicks; tick += 1) {
    const ratio = tick / yTicks;
    const y = padding.top + ratio * chartHeight;
    const value = maxY - ratio * (maxY - minY);

    context.beginPath();
    context.moveTo(padding.left, y);
    context.lineTo(padding.left + chartWidth, y);
    context.stroke();

    context.fillStyle = "#666";
    context.font = "12px Segoe UI, Arial, sans-serif";
    context.textAlign = "left";
    context.textBaseline = "middle";
    context.fillText(formatAxisMoney(value), 8, y);
  }

  context.strokeStyle = "#b5b5b5";
  context.beginPath();
  context.moveTo(padding.left, padding.top);
  context.lineTo(padding.left, padding.top + chartHeight);
  context.lineTo(padding.left + chartWidth, padding.top + chartHeight);
  context.stroke();

  const xTicks = 3;
  for (let tick = 0; tick <= xTicks; tick += 1) {
    const ratio = tick / xTicks;
    const x = padding.left + ratio * chartWidth;
    const time = minX + ratio * (maxX - minX);

    context.fillStyle = "#666";
    context.font = "12px Segoe UI, Arial, sans-serif";
    context.textAlign = tick === 0 ? "left" : tick === xTicks ? "right" : "center";
    context.textBaseline = "top";
    context.fillText(formatTimeTick(time, minX, maxX), x, padding.top + chartHeight + 10);
  }

  context.strokeStyle = "#111";
  context.lineWidth = 2;
  context.beginPath();
  series.forEach((point, index) => {
    const x = toX(point.x);
    const y = toY(point.y);
    if (index === 0) context.moveTo(x, y);
    else context.lineTo(x, y);
  });
  context.stroke();

  const trendLine = getTrendLineForSeries(series, minX, maxX);
  if (graphShowTrendLine && trendLine) {
    context.save();
    context.beginPath();
    context.rect(padding.left, padding.top, chartWidth, chartHeight);
    context.clip();
    context.strokeStyle = "#c24e2a";
    context.lineWidth = 1.5;
    context.setLineDash([6, 4]);
    context.beginPath();
    context.moveTo(toX(minX), toY(trendLine.yAtMin));
    context.lineTo(toX(maxX), toY(trendLine.yAtMax));
    context.stroke();
    context.restore();
    context.setLineDash([]);
  }

  series.forEach((point) => {
    const x = toX(point.x);
    const y = toY(point.y);
    context.fillStyle = "#111";
    context.beginPath();
    context.arc(x, y, 2.6, 0, Math.PI * 2);
    context.fill();

    graphPointTargets.push({ x, y, point });
  });

  const rangeSuffix = rangeFilter.mode === "range"
    ? `Shown: ${formatDateForDisplay(rangeFilter.startValue)} - ${formatDateForDisplay(rangeFilter.endValue)}.`
    : "Shown: all history.";
  summary.innerText = `${getAccountNameById(selectedGraphAccountId)} current: ${formatMoney(currentValue)}.\nChange over shown period: ${formatSignedMoney(periodChange)}.\n${rangeSuffix}\nTap a transaction dot for details.`;
}

function render() {
  renderTotal();
  renderAccounts();
  renderTransactionForm();
  renderTransactions();
  renderDataSection();
  renderGraphFilter();
  renderGraph();
  renderCategoryBreakdown();
}

function addAccount() {
  const nameInput = document.getElementById("accountNameInput");
  const balanceInput = document.getElementById("initialBalanceInput");

  const name = toTrimmedString(nameInput.value);
  if (!name) return;

  accounts.push({
    id: generateId(),
    name,
    initialBalance: getSafeAmount(balanceInput.value),
    createdDate: getTodayDateInputValue(),
    archived: false
  });

  nameInput.value = "";
  balanceInput.value = "";
  saveAccounts();
  render();
  closeAddAccountModal();
}

function getSignedAccountAmount(accountId) {
  const amountInput = document.getElementById(`amount-${accountId}`);
  if (!amountInput) return 0;

  const amount = getSafeAmount(amountInput.value);
  amountInput.value = "";
  return amount;
}

function adjustAccount(accountId) {
  const delta = getSignedAccountAmount(accountId);
  if (!delta) return;
  const type = delta > 0 ? "income" : "expense";
  const adjustmentCategory = ensureCategoryOption(type, ACCOUNT_ADJUSTMENT_CATEGORY, true);

  transactions.push({
    id: generateId(),
    type,
    date: getTodayDateInputValue(),
    amount: Math.abs(delta),
    accountId,
    fromAccountId: "",
    toAccountId: "",
    note: "Quick adjust",
    category: adjustmentCategory,
    currency: "GBP",
    createdAt: new Date().toISOString()
  });
  saveTransactions();
  render();
}

function silentlyAdjustAccount(accountId, delta) {
  const account = accounts.find((item) => item.id === accountId);
  if (!account) return;

  // Keep the adjustment on the opening-balance event so graph history shifts from inception.
  account.createdDate = isValidDateInput(account.createdDate) ? account.createdDate : getTodayDateInputValue();
  account.initialBalance = normalizeMoney(account.initialBalance + delta);
  saveAccounts();
  render();
}

function silentAdjust(accountId) {
  const delta = getSignedAccountAmount(accountId);
  if (!delta) return;
  silentlyAdjustAccount(accountId, delta);
}

function deleteAccount(accountId) {
  const account = accounts.find((item) => item.id === accountId);
  if (!account) return;
  if (account.archived) return;

  const balances = getBalanceMap();
  const currentBalance = getSafeAmount(balances[accountId] || 0);
  if (Math.abs(currentBalance) >= 0.0001) {
    window.alert("Only accounts with a zero balance can be archived.");
    return;
  }

  const confirmed = window.confirm(`Archive account "${account.name}"?`);
  if (!confirmed) return;

  account.archived = true;
  saveAccounts();
  render();
}

function unarchiveAccount(accountId) {
  const account = accounts.find((item) => item.id === accountId);
  if (!account) return;
  if (!account.archived) return;

  account.archived = false;
  saveAccounts();
  render();
}

function toggleArchivedAccountsVisibility() {
  showArchivedAccounts = !showArchivedAccounts;
  renderAccounts();
}

function toggleAccountAdjustPanel(accountId) {
  expandedAccountId = expandedAccountId === accountId ? "" : accountId;
  renderAccounts();
}

function toggleAccountTransactionsShowAll(accountId) {
  accountTransactionsShowAllById[accountId] = !Boolean(accountTransactionsShowAllById[accountId]);
  renderAccounts();
}

function onAccountHeaderKeyDown(event, accountId) {
  if (event.key !== "Enter" && event.key !== " " && event.key !== "Spacebar") return;
  event.preventDefault();
  toggleAccountAdjustPanel(accountId);
}

function onTransactionTypeChange() {
  const type = normalizeTransactionType(document.getElementById("transactionTypeInput").value);
  const accountGroup = document.getElementById("transactionAccountGroup");
  const transferGroup = document.getElementById("transactionTransferGroup");
  const categoryGroup = document.getElementById("transactionCategoryGroup");
  const categoryTrigger = document.getElementById("transactionCategoryTrigger");

  closeCategoryDropdown();

  if (type === "transfer") {
    accountGroup.classList.add("hidden");
    transferGroup.classList.remove("hidden");
    if (categoryGroup) categoryGroup.classList.add("hidden");
    if (categoryTrigger) categoryTrigger.disabled = true;
    setCategoryTriggerLabel(type, "");
    return;
  }

  accountGroup.classList.remove("hidden");
  transferGroup.classList.add("hidden");
  if (categoryGroup) categoryGroup.classList.remove("hidden");
  if (categoryTrigger) categoryTrigger.disabled = getActiveAccounts().length === 0;
  populateCategorySelect(type, selectedCategoryByType[type] || "");
}

function addTransaction() {
  if (getActiveAccounts().length === 0) return;

  const type = normalizeTransactionType(document.getElementById("transactionTypeInput").value);
  const dateInput = document.getElementById("transactionDateInput");
  const amountInput = document.getElementById("transactionAmountInput");
  const noteInput = document.getElementById("transactionNoteInput");
  const accountInput = document.getElementById("transactionAccountInput");
  const fromInput = document.getElementById("transactionFromAccountInput");
  const toInput = document.getElementById("transactionToAccountInput");

  if (!type) return;

  const amount = getSafeAmount(amountInput.value);
  if (amount <= 0) return;

  const date = normalizeImportedDate(dateInput.value);
  const note = toTrimmedString(noteInput.value);
  const accountIds = new Set(getActiveAccounts().map((account) => account.id));

  if (type === "transfer") {
    const fromAccountId = toTrimmedString(fromInput.value);
    const toAccountId = toTrimmedString(toInput.value);
    if (!accountIds.has(fromAccountId) || !accountIds.has(toAccountId) || fromAccountId === toAccountId) return;

    transactions.push({
      id: generateId(),
      type,
      date,
      amount,
      accountId: "",
      fromAccountId,
      toAccountId,
      note,
      category: "",
      currency: "GBP",
      createdAt: new Date().toISOString()
    });
  } else {
    const accountId = toTrimmedString(accountInput.value);
    if (!accountIds.has(accountId)) return;
    const selectedCategory = normalizeCategoryName(selectedCategoryByType[type] || "");
    if (!selectedCategory) {
      window.alert('Choose a category. Use "Unknown" if needed.');
      return;
    }
    const category = ensureCategoryOption(type, selectedCategory, true);
    selectedCategoryByType[type] = category;

    transactions.push({
      id: generateId(),
      type,
      date,
      amount,
      accountId,
      fromAccountId: "",
      toAccountId: "",
      note,
      category,
      currency: "GBP",
      createdAt: new Date().toISOString()
    });
  }

  amountInput.value = "";
  noteInput.value = "";
  saveTransactions();
  render();
  closeAddTransactionModal();
}

function deleteTransaction(transactionId) {
  const transaction = transactions.find((tx) => tx.id === transactionId);
  if (!transaction) return;

  const confirmed = window.confirm("Delete this transaction?");
  if (!confirmed) return;

  if (activeTransactionModalId === transactionId) {
    closeTransactionModal();
  }
  if (activeGraphPointTransactionId === transactionId) {
    closeGraphPointModal();
  }

  transactions = transactions.filter((tx) => tx.id !== transactionId);
  saveTransactions();
  render();
}

function parseCsv(text) {
  const rows = [];
  let currentCell = "";
  let currentRow = [];
  let inQuotes = false;

  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];
    const next = text[i + 1];

    if (char === '"') {
      if (inQuotes && next === '"') {
        currentCell += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (char === "," && !inQuotes) {
      currentRow.push(currentCell);
      currentCell = "";
      continue;
    }

    if ((char === "\n" || char === "\r") && !inQuotes) {
      if (char === "\r" && next === "\n") i += 1;
      currentRow.push(currentCell);
      currentCell = "";

      if (currentRow.some((cell) => toTrimmedString(cell) !== "")) {
        rows.push(currentRow);
      }
      currentRow = [];
      continue;
    }

    currentCell += char;
  }

  currentRow.push(currentCell);
  if (currentRow.some((cell) => toTrimmedString(cell) !== "")) {
    rows.push(currentRow);
  }

  if (rows.length === 0) return [];

  const headers = rows[0].map((header) => toTrimmedString(header));
  return rows.slice(1).map((row) => {
    const obj = {};
    headers.forEach((header, index) => {
      obj[header] = row[index] ?? "";
    });
    return obj;
  });
}

async function readImportRows(file) {
  const fileName = file.name.toLowerCase();

  if (fileName.endsWith(".csv")) {
    const text = await file.text();
    return { rows: parseCsv(text), rowNumberOffset: 2 };
  }

  if (fileName.endsWith(".json")) {
    const text = (await file.text()).replace(/^\uFEFF/, "");
    let parsed;
    try {
      parsed = JSON.parse(text);
    } catch {
      throw new Error("JSON file is invalid.");
    }

    if (Array.isArray(parsed)) {
      return { rows: parsed, rowNumberOffset: 1 };
    }

    if (parsed && typeof parsed === "object") {
      if (Array.isArray(parsed.transactions)) {
        return { rows: parsed.transactions, rowNumberOffset: 1 };
      }
      if (Array.isArray(parsed.data)) {
        return { rows: parsed.data, rowNumberOffset: 1 };
      }
    }

    throw new Error("JSON must be an array of transactions (or an object with a transactions/data array).");
  }

  if (!window.XLSX) {
    throw new Error("Excel reader failed to load. Check internet access and retry, or import as CSV.");
  }

  const buffer = await file.arrayBuffer();
  const workbook = window.XLSX.read(buffer, {
    type: "array",
    cellDates: true,
    raw: true
  });

  if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
    return { rows: [], rowNumberOffset: 2 };
  }

  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
  return {
    rows: window.XLSX.utils.sheet_to_json(firstSheet, { defval: "", raw: true }),
    rowNumberOffset: 2
  };
}

function getNormalizedRow(row) {
  const normalized = {};
  Object.entries(row).forEach(([key, value]) => {
    normalized[normalizeHeader(key)] = value;
  });
  return normalized;
}

function getFirstValue(row, keys) {
  for (const key of keys) {
    if (Object.prototype.hasOwnProperty.call(row, key)) {
      return row[key];
    }
  }
  return "";
}

function parseImportedAmount(rawAmount) {
  if (typeof rawAmount === "number") {
    return Number.isFinite(rawAmount) ? normalizeMoney(Math.abs(rawAmount)) : 0;
  }

  const text = toTrimmedString(rawAmount);
  if (!text) return 0;

  const match = text.match(/[-+]?\d[\d,]*(?:\.\d+)?/);
  if (!match) return 0;

  const numeric = Number.parseFloat(match[0].replaceAll(",", ""));
  return Number.isFinite(numeric) ? normalizeMoney(Math.abs(numeric)) : 0;
}

function parseTransferAccountsFromField(rawAccountField) {
  const text = toTrimmedString(rawAccountField);
  if (!text) return { from: "", to: "" };

  const parts = text
    .split(/\s*(?:\u279B|\u2192|\u27A1|\u27F6|\u2B95|->|=>)\s*/)
    .map((part) => toTrimmedString(part))
    .filter(Boolean);
  if (parts.length < 2) return { from: "", to: "" };
  return { from: parts[0], to: parts[parts.length - 1] };
}

function getOrCreateAccountIdByName(accountName, createdDate, counters) {
  const cleanName = toTrimmedString(accountName);
  if (!cleanName) return "";

  const existingActive = accounts.find((account) => !account.archived && account.name.toLowerCase() === cleanName.toLowerCase());
  if (existingActive) return existingActive.id;

  const existingArchived = accounts.find((account) => account.archived && account.name.toLowerCase() === cleanName.toLowerCase());
  if (existingArchived) {
    existingArchived.archived = false;
    return existingArchived.id;
  }

  const newAccount = {
    id: generateId(),
    name: cleanName,
    initialBalance: 0,
    createdDate: isValidDateInput(createdDate) ? createdDate : getTodayDateInputValue(),
    archived: false
  };
  accounts.push(newAccount);
  counters.createdAccounts += 1;
  return newAccount.id;
}

function importTransactionsFromRows(rows, rowNumberOffset = 2) {
  const counters = { added: 0, skipped: 0, createdAccounts: 0, failedRows: [] };

  rows.forEach((rawRow, index) => {
    const row = getNormalizedRow(rawRow);
    const rowNumber = index + rowNumberOffset;
    const accountSinglePreview = toTrimmedString(getFirstValue(row, ["account", "accountname"]));
    let accountFr = toTrimmedString(getFirstValue(row, ["accountfr", "accountfrom", "fromaccount"]));
    let accountTo = toTrimmedString(getFirstValue(row, ["accountto", "toaccount"]));

    const recordFailure = (reason) => {
      const typeValue = toTrimmedString(getFirstValue(row, ["type", "transactiontype"])) || "(blank)";
      const amountValue = toTrimmedString(getFirstValue(row, ["amount", "value"])) || "(blank)";
      counters.failedRows.push(
        `Row ${rowNumber}: ${reason} | type=${typeValue}, amount=${amountValue}, account=${accountSinglePreview || "(blank)"}, account fr=${accountFr || "(blank)"}, account to=${accountTo || "(blank)"}`
      );
      counters.skipped += 1;
    };

    const type = normalizeTransactionType(getFirstValue(row, ["type", "transactiontype"]));
    if (!type) {
      recordFailure("Invalid or missing type");
      return;
    }

    if (type === "transfer" && (!accountFr || !accountTo) && accountSinglePreview) {
      const parsedAccounts = parseTransferAccountsFromField(accountSinglePreview);
      accountFr = accountFr || parsedAccounts.from;
      accountTo = accountTo || parsedAccounts.to;
    }

    let amount = parseImportedAmount(getFirstValue(row, ["amount", "value"]));
    if (type === "transfer" && amount <= 0) {
      amount = parseImportedAmount(getFirstValue(row, ["currency", "curr"]));
    }
    if (amount <= 0) {
      recordFailure("Amount must be greater than 0");
      return;
    }

    const date = normalizeImportedDate(getFirstValue(row, ["date", "transactiondate"]));
    const importedCategory = normalizeCategoryName(getFirstValue(row, ["category", "cat"]));
    const remark = toTrimmedString(getFirstValue(row, ["remark", "note", "memo", "description"]));
    const currency = toTrimmedString(getFirstValue(row, ["currency", "curr"])).toUpperCase();

    const noteParts = [];
    if (remark) noteParts.push(remark);
    if (currency && currency !== "GBP") noteParts.push(`Currency: ${currency}`);
    const note = noteParts.join(" | ");

    const createdAt = new Date(dateToTimestamp(date) + index).toISOString();

    if (type === "transfer") {
      const fromAccountId = getOrCreateAccountIdByName(accountFr, date, counters);
      const toAccountId = getOrCreateAccountIdByName(accountTo, date, counters);
      if (!fromAccountId || !toAccountId || fromAccountId === toAccountId) {
        recordFailure("Transfer requires different from/to accounts");
        return;
      }

      transactions.push({
        id: generateId(),
        type,
        date,
        amount,
        accountId: "",
        fromAccountId,
        toAccountId,
        note,
        category: "",
        currency: currency || "GBP",
        createdAt
      });
      counters.added += 1;
      return;
    }

    const singleAccountName = accountSinglePreview || (type === "expense" ? (accountFr || accountTo) : (accountTo || accountFr));
    const accountId = getOrCreateAccountIdByName(singleAccountName, date, counters);
    if (!accountId) {
      recordFailure("Missing account name for income/expense");
      return;
    }
    if (!importedCategory) {
      recordFailure('Missing category for income/expense (use "Unknown" if needed)');
      return;
    }
    const category = ensureCategoryOption(type, importedCategory, false);

    transactions.push({
      id: generateId(),
      type,
      date,
      amount,
      accountId,
      fromAccountId: "",
      toAccountId: "",
      note,
      category,
      currency: currency || "GBP",
      createdAt
    });
    counters.added += 1;
  });

  return counters;
}

function showImportFailuresPopup(failedRows) {
  if (!failedRows || failedRows.length === 0) {
    return;
  }

  const maxRowsInPopup = 120;
  const shownRows = failedRows.slice(0, maxRowsInPopup);
  const remaining = failedRows.length - shownRows.length;
  const footer = remaining > 0
    ? `\n...and ${remaining} more skipped row(s).`
    : "";

  window.alert(
    `Skipped rows during import:\n\n${shownRows.join("\n")}${footer}`
  );
}

function setStatusPanelMessage(elementId, message, state = "info") {
  const panel = document.getElementById(elementId);
  if (!panel) return;

  const text = toTrimmedString(message);
  if (!text) {
    panel.innerText = "";
    panel.className = "status-panel hidden";
    return;
  }

  panel.innerText = text;
  panel.className = "status-panel";
  if (state === "success") panel.classList.add("status-success");
  if (state === "error") panel.classList.add("status-error");
}

function getCurrentImportFile() {
  if (selectedImportFile) return selectedImportFile;
  const input = document.getElementById("importFileInput");
  return input && input.files && input.files[0] ? input.files[0] : null;
}

function updateImportFileSelectionState() {
  const info = document.getElementById("importSelectedFileInfo");
  const importButton = document.getElementById("importTransactionsBtn");
  const file = getCurrentImportFile();

  if (importButton) {
    importButton.disabled = !file;
  }

  if (!info) return;
  if (!file) {
    info.innerText = "";
    info.classList.add("hidden");
    return;
  }

  info.innerText = `Selected: ${file.name} (${formatFileSize(file.size)})`;
  info.classList.remove("hidden");
}

function triggerImportFilePicker() {
  const input = document.getElementById("importFileInput");
  if (input) input.click();
}

function onImportFileInputChange() {
  const input = document.getElementById("importFileInput");
  selectedImportFile = input && input.files && input.files[0] ? input.files[0] : null;
  updateImportFileSelectionState();
}

function onImportDropzoneKeyDown(event) {
  if (event.key !== "Enter" && event.key !== " " && event.key !== "Spacebar") return;
  event.preventDefault();
  triggerImportFilePicker();
}

function setImportDropzoneDragState(isDragOver) {
  const dropzone = document.getElementById("importDropzone");
  if (!dropzone) return;
  dropzone.classList.toggle("drag-over", isDragOver);
}

function initImportDropzone() {
  const dropzone = document.getElementById("importDropzone");
  if (!dropzone) return;
  if (dropzone.dataset.bound === "1") return;
  dropzone.dataset.bound = "1";

  const onDragActive = (event) => {
    event.preventDefault();
    event.stopPropagation();
    setImportDropzoneDragState(true);
  };

  const onDragInactive = (event) => {
    event.preventDefault();
    event.stopPropagation();
    setImportDropzoneDragState(false);
  };

  dropzone.addEventListener("dragenter", onDragActive);
  dropzone.addEventListener("dragover", onDragActive);
  dropzone.addEventListener("dragleave", onDragInactive);
  dropzone.addEventListener("dragend", onDragInactive);
  dropzone.addEventListener("drop", (event) => {
    event.preventDefault();
    event.stopPropagation();
    setImportDropzoneDragState(false);

    const file = event.dataTransfer && event.dataTransfer.files && event.dataTransfer.files[0]
      ? event.dataTransfer.files[0]
      : null;
    if (!file) return;

    selectedImportFile = file;
    const input = document.getElementById("importFileInput");
    if (input) {
      try {
        const transfer = new DataTransfer();
        transfer.items.add(file);
        input.files = transfer.files;
      } catch {
        input.value = "";
      }
    }
    updateImportFileSelectionState();
  });
}

function renderImportFormatPanel() {
  const panel = document.getElementById("importFormatPanel");
  const toggleButton = document.getElementById("importFormatToggleBtn");
  if (panel) {
    panel.classList.toggle("hidden", !showImportFormatDetails);
  }
  if (toggleButton) {
    toggleButton.innerText = showImportFormatDetails ? "Hide required format <-" : "View required format ->";
    toggleButton.setAttribute("aria-expanded", showImportFormatDetails ? "true" : "false");
  }
}

function toggleImportFormatPanel() {
  showImportFormatDetails = !showImportFormatDetails;
  renderImportFormatPanel();
}

function switchDataTab(tabName) {
  activeDataTab = tabName === "export" ? "export" : "import";

  const importTab = document.getElementById("dataTabImport");
  const exportTab = document.getElementById("dataTabExport");
  const importPanel = document.getElementById("dataImportPanel");
  const exportPanel = document.getElementById("dataExportPanel");
  if (!importTab || !exportTab || !importPanel || !exportPanel) return;

  const importActive = activeDataTab === "import";
  importTab.classList.toggle("active", importActive);
  exportTab.classList.toggle("active", !importActive);
  importTab.setAttribute("aria-selected", importActive ? "true" : "false");
  exportTab.setAttribute("aria-selected", importActive ? "false" : "true");
  importTab.tabIndex = importActive ? 0 : -1;
  exportTab.tabIndex = importActive ? -1 : 0;
  importPanel.classList.toggle("hidden", !importActive);
  exportPanel.classList.toggle("hidden", importActive);
}

function onDataTabKeyDown(event, currentTab) {
  const keys = ["ArrowLeft", "ArrowRight", "Home", "End"];
  if (!keys.includes(event.key)) return;
  event.preventDefault();

  if (event.key === "Home") {
    switchDataTab("import");
    const importTab = document.getElementById("dataTabImport");
    if (importTab) importTab.focus();
    return;
  }

  if (event.key === "End") {
    switchDataTab("export");
    const exportTab = document.getElementById("dataTabExport");
    if (exportTab) exportTab.focus();
    return;
  }

  const isImport = currentTab === "import";
  const nextTab = event.key === "ArrowRight"
    ? (isImport ? "export" : "import")
    : (isImport ? "export" : "import");
  switchDataTab(nextTab);
  const nextTabElement = document.getElementById(nextTab === "import" ? "dataTabImport" : "dataTabExport");
  if (nextTabElement) nextTabElement.focus();
}

function renderDataExportMeta() {
  const meta = document.getElementById("dataExportMeta");
  if (!meta) return;

  const dateText = lastExportMeta && lastExportMeta.date
    ? formatDateTimeForDisplay(lastExportMeta.date)
    : "Never";
  const transactionCount = lastExportMeta && Number.isInteger(lastExportMeta.transactions)
    ? lastExportMeta.transactions
    : transactions.length;
  const accountCount = lastExportMeta && Number.isInteger(lastExportMeta.accounts)
    ? lastExportMeta.accounts
    : accounts.length;

  meta.innerText = `Last export: ${dateText} | ${transactionCount} transactions | ${accountCount} accounts`;
}

function renderDataSection() {
  switchDataTab(activeDataTab);
  renderImportFormatPanel();
  updateImportFileSelectionState();
  renderDataExportMeta();
}

async function importTransactionsFile() {
  const fileInput = document.getElementById("importFileInput");
  const file = getCurrentImportFile();

  if (!file) {
    setStatusPanelMessage("dataImportStatus", "Choose a file first.", "error");
    return;
  }

  setStatusPanelMessage("dataImportStatus", "Importing...", "info");

  try {
    const importPayload = await readImportRows(file);
    const rows = Array.isArray(importPayload) ? importPayload : importPayload.rows;
    const rowNumberOffset = Number.isInteger(importPayload?.rowNumberOffset) ? importPayload.rowNumberOffset : 2;
    if (!rows || rows.length === 0) {
      setStatusPanelMessage("dataImportStatus", "No rows found in the file.", "error");
      return;
    }

    const counters = importTransactionsFromRows(rows, rowNumberOffset);
    persistAll();
    render();
    selectedImportFile = null;
    if (fileInput) fileInput.value = "";
    updateImportFileSelectionState();
    showImportFailuresPopup(counters.failedRows);

    if (counters.added === 0) {
      setStatusPanelMessage("dataImportStatus", `No transactions imported. Skipped ${counters.skipped} row(s).`, "error");
      return;
    }

    setStatusPanelMessage(
      "dataImportStatus",
      `Imported ${counters.added} transactions.`,
      "success"
    );
  } catch (error) {
    setStatusPanelMessage("dataImportStatus", `Import failed: ${error instanceof Error ? error.message : "Unknown error"}`, "error");
  }
}

function buildAccountExportRows() {
  const balances = getBalanceMap();
  return [...accounts]
    .sort((a, b) => {
      const balanceA = getSafeAmount(balances[a.id] || 0);
      const balanceB = getSafeAmount(balances[b.id] || 0);
      if (balanceA !== balanceB) return balanceB - balanceA;
      return a.name.localeCompare(b.name, undefined, { sensitivity: "base" });
    })
    .map((account) => ({
      id: account.id,
      account: account.name,
      current_balance: getSafeAmount(balances[account.id] || 0),
      opening_balance: getSafeAmount(account.initialBalance),
      archived: account.archived ? "Yes" : "No",
      created_date: account.createdDate
    }));
}

function buildTransactionExportRows() {
  return sortTransactionsAscending().map((tx) => {
    const accountName = tx.type === "transfer" ? "" : getAccountNameById(tx.accountId);
    const fromAccountName = tx.type === "transfer"
      ? getAccountNameById(tx.fromAccountId)
      : (tx.type === "expense" ? accountName : "");
    const toAccountName = tx.type === "transfer"
      ? getAccountNameById(tx.toAccountId)
      : (tx.type === "income" ? accountName : "");

    return {
      id: tx.id,
      type: tx.type,
      amount: getSafeAmount(tx.amount),
      account: accountName,
      account_fr: fromAccountName,
      account_to: toAccountName,
      category: tx.type === "transfer" ? "" : (normalizeCategoryName(tx.category) || UNKNOWN_CATEGORY),
      remark: tx.note || "",
      currency: (tx.currency || "GBP").toUpperCase(),
      date: tx.date,
      created_at: tx.createdAt || ""
    };
  });
}

function exportAllData() {
  if (!window.XLSX) {
    setStatusPanelMessage("dataExportStatus", "Excel export is unavailable right now. Check internet access and retry.", "error");
    return;
  }

  const accountRows = buildAccountExportRows();
  const transactionRows = buildTransactionExportRows();

  const workbook = window.XLSX.utils.book_new();
  const accountSheet = window.XLSX.utils.json_to_sheet(accountRows.length > 0 ? accountRows : [{
    id: "",
    account: "",
    current_balance: 0,
    opening_balance: 0,
    archived: "",
    created_date: ""
  }]);
  const transactionSheet = window.XLSX.utils.json_to_sheet(transactionRows.length > 0 ? transactionRows : [{
    id: "",
    type: "",
    amount: 0,
    account: "",
    account_fr: "",
    account_to: "",
    category: "",
    remark: "",
    currency: "",
    date: "",
    created_at: ""
  }]);

  window.XLSX.utils.book_append_sheet(workbook, accountSheet, "Accounts");
  window.XLSX.utils.book_append_sheet(workbook, transactionSheet, "Transactions");

  const fileName = `account-tracker-export-${getTodayDateInputValue()}.xlsx`;
  window.XLSX.writeFile(workbook, fileName);

  lastExportMeta = {
    date: new Date().toISOString(),
    transactions: transactionRows.length,
    accounts: accountRows.length
  };
  saveLastExportMeta();
  renderDataExportMeta();
  setStatusPanelMessage(
    "dataExportStatus",
    `Exported ${accountRows.length} account(s) and ${transactionRows.length} transaction(s) to ${fileName}.`,
    "success"
  );
}

function onGraphFilterChange() {
  const select = document.getElementById("graphAccountFilter");
  selectedGraphAccountId = select.value;
  closeGraphPointModal();
  syncGraphControlsFromState();
  renderGraph();
}

function onCategoryBreakdownTypeChange() {
  const typeSelect = document.getElementById("categoryBreakdownTypeInput");
  const previousType = normalizeBreakdownType(selectedCategoryBreakdownType);
  if (previousType === "expense" || previousType === "income") {
    selectedCategoryBreakdownMonthByType[previousType] = selectedCategoryBreakdownMonth;
  }

  selectedCategoryBreakdownType = normalizeBreakdownType(typeSelect ? typeSelect.value : "expense");
  selectedCategoryBreakdownMonth = selectedCategoryBreakdownMonthByType[selectedCategoryBreakdownType] || selectedCategoryBreakdownMonth;
  selectedCategoryBreakdownCategory = "";
  categoryBreakdownTransactionsShowAll = false;
  renderCategoryBreakdown();
}

function onCategoryBreakdownMonthChange() {
  const monthSelect = document.getElementById("categoryBreakdownMonthInput");
  selectedCategoryBreakdownMonth = toTrimmedString(monthSelect ? monthSelect.value : "");
  selectedCategoryBreakdownMonthByType[selectedCategoryBreakdownType] = selectedCategoryBreakdownMonth;
  selectedCategoryBreakdownCategory = "";
  categoryBreakdownTransactionsShowAll = false;
  renderCategoryBreakdown();
}

function toggleTransactionHistoryRange() {
  if (transactions.length === 0) return;
  transactionHistoryShowAll = !transactionHistoryShowAll;
  renderTransactions();
}

function clearAllTransactions() {
  if (transactions.length === 0) {
    window.alert("No transaction history to clear.");
    return;
  }

  const confirmed = window.confirm("Clear all transaction history? This cannot be undone.");
  if (!confirmed) {
    return;
  }

  transactions = [];
  transactionHistoryShowAll = false;
  closeTransactionModal();
  closeGraphPointModal();
  saveTransactions();
  render();

  setStatusPanelMessage("dataImportStatus", "All transaction history was cleared.", "info");
}

function showTab(tabName) {
  activeTab = ["dashboard", "transactions", "graph", "categories"].includes(tabName) ? tabName : "dashboard";

  if (activeTab !== "transactions") {
    closeCategoryDropdown();
    closeAddTransactionModal();
  }
  if (activeTab !== "dashboard") {
    closeAddAccountModal();
  }
  if (activeTab !== "graph") {
    closeGraphPointModal();
  }

  document.getElementById("dashboardTab").classList.toggle("hidden", activeTab !== "dashboard");
  document.getElementById("transactionsTab").classList.toggle("hidden", activeTab !== "transactions");
  document.getElementById("graphTab").classList.toggle("hidden", activeTab !== "graph");
  document.getElementById("categoriesTab").classList.toggle("hidden", activeTab !== "categories");

  document.getElementById("tabDashboardBtn").classList.toggle("active", activeTab === "dashboard");
  document.getElementById("tabTransactionsBtn").classList.toggle("active", activeTab === "transactions");
  document.getElementById("tabGraphBtn").classList.toggle("active", activeTab === "graph");
  document.getElementById("tabCategoriesBtn").classList.toggle("active", activeTab === "categories");

  if (activeTab === "graph") {
    renderGraph();
  }
  if (activeTab === "categories") {
    renderCategoryBreakdown();
  }
}

function mountSharedModalsToBody() {
  ["transactionModal"].forEach((id) => {
    const modal = document.getElementById(id);
    if (!modal) return;
    if (modal.parentElement !== document.body) {
      document.body.appendChild(modal);
    }
  });
}

let promptedServiceWorkerScriptUrl = "";

function promptForServiceWorkerUpdate(registration) {
  if (!registration || !registration.waiting) return;

  const waiting = registration.waiting;
  const waitingKey = toTrimmedString(waiting.scriptURL || "waiting");
  if (waitingKey && waitingKey === promptedServiceWorkerScriptUrl) return;
  if (waitingKey) promptedServiceWorkerScriptUrl = waitingKey;

  const shouldUpdate = window.confirm("A new version of Ledger is available. Update now?");
  if (!shouldUpdate) return;

  waiting.postMessage({ type: "SKIP_WAITING" });
}

function registerServiceWorkerWithUpdatePrompt() {
  if (!("serviceWorker" in navigator)) return;

  let reloadingForUpdate = false;
  navigator.serviceWorker.addEventListener("controllerchange", () => {
    if (reloadingForUpdate) return;
    reloadingForUpdate = true;
    window.location.reload();
  });

  navigator.serviceWorker.register("./service-worker.js")
    .then((registration) => {
      if (!registration) return;

      promptForServiceWorkerUpdate(registration);

      registration.addEventListener("updatefound", () => {
        const installing = registration.installing;
        if (!installing) return;

        installing.addEventListener("statechange", () => {
          if (installing.state === "installed" && navigator.serviceWorker.controller) {
            promptForServiceWorkerUpdate(registration);
          }
        });
      });

      registration.update().catch(() => {});
    })
    .catch(() => {});
}

window.onload = function () {
  mountSharedModalsToBody();
  registerServiceWorkerWithUpdatePrompt();
  loadAccounts();
  loadCategoryOptions();
  loadTransactions();
  loadLastExportMeta();
  syncCategoriesFromTransactions();
  render();
  showTab("dashboard");
  initImportDropzone();

  const graphCanvas = document.getElementById("balanceChart");
  if (graphCanvas) {
    graphCanvas.addEventListener("click", onGraphCanvasClick);
  }
  const categoryLegend = document.getElementById("categoryBreakdownLegend");
  if (categoryLegend) {
    categoryLegend.addEventListener("click", onCategoryBreakdownLegendContainerClick);
  }

  const hint = document.getElementById("installHint");
  if (hint) hint.innerText = "";
};

window.addEventListener("resize", () => {
  if (activeTab === "graph") {
    renderGraph();
  }
  if (activeTab === "categories") {
    renderCategoryBreakdown();
  }
});

document.addEventListener("click", (event) => {
  const dropdown = document.getElementById("transactionCategoryDropdown");
  if (!dropdown) return;

  if (!dropdown.contains(event.target)) {
    closeCategoryDropdown();
  }
});



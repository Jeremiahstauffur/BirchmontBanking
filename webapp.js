(function () {
  const seed = window.WORKBOOK_SEED;
  const STORAGE_KEY = "jere-anna-budget-clone-v2";
  const PAYPERIOD_INCOME_ACCOUNTS = new Set(["FNB - PERSONAL", "WALLET"]);
  const PAYPERIOD_EXCLUDED_CATEGORIES = new Set(["Added", "Budget Loans", "Withdrawals"]);
  const DEFAULT_CATEGORY_NAME = "General";
  const NO_SUBCATEGORY_LABEL = "No subcategory";
  const emptyTwinLine = () => ({ budget: "", category: "", subcategory: "", amount: "", personalNotes: "" });

  if (!seed) {
    document.body.innerHTML = "<p style='padding:24px;font-family:sans-serif'>Starting data could not be loaded.</p>";
    return;
  }

  const today = todayIso();
  const state = loadState();
  normalizeState();
  state.twinLines = [emptyTwinLine(), emptyTwinLine(), emptyTwinLine()];

  document.addEventListener("DOMContentLoaded", init);

  function init() {
    bindStaticEvents();
    render();
    primeForms();
  }

  function loadState() {
    const defaults = {
      currentView: "dashboard",
      selfEmploymentSubview: "income",
      selectedTransactionId: "",
      pendingReconcileIds: [],
      budgetProgressMainGroup: "",
      budgetProgressCategory: "",
      editing: {
        register: "",
        budgetEntries: "",
      },
      filters: {
        bankingSearch: "",
        bankingAccount: "All accounts",
        budgetSearch: "",
        budgetBudget: "All budgets",
        budgetCategory: "All categories",
        budgetSubcategory: "All subcategories",
        transactionIdSearch: "",
      },
      settings: {
        planningFrequency: seed.defaults.paycheckFrequency || "Semi-Monthly (24/yr)",
        payperiodStart: seed.defaults.payperiodStart || today,
        payperiodCycleDays: seed.defaults.payperiodCycleDays || 14,
        bankOverviewStepDays: seed.defaults.bankOverviewStepDays || 2,
        bankOverviewStart: addDays(today, -30 * (seed.defaults.bankOverviewStepDays || 2)),
        budgetOverviewStepDays: seed.defaults.budgetOverviewStepDays || 1,
        budgetOverviewStart: addDays(today, -30 * (seed.defaults.budgetOverviewStepDays || 1)),
      },
      budgetHierarchy: buildInitialBudgetHierarchy(seed.tables.budgetEntries, seed.tables.biWeeklyExpenses),
      tables: clone(seed.tables),
    };

    try {
      const saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || "null");
      if (!saved) return defaults;
      return {
        currentView: saved.currentView || defaults.currentView,
        selfEmploymentSubview: saved.selfEmploymentSubview || defaults.selfEmploymentSubview,
        selectedTransactionId: saved.selectedTransactionId || defaults.selectedTransactionId,
        pendingReconcileIds: saved.pendingReconcileIds || defaults.pendingReconcileIds,
        budgetProgressMainGroup: saved.budgetProgressMainGroup || defaults.budgetProgressMainGroup,
        budgetProgressCategory: saved.budgetProgressCategory || defaults.budgetProgressCategory,
        filters: { ...defaults.filters, ...(saved.filters || {}) },
        settings: { ...defaults.settings, ...(saved.settings || {}) },
        budgetHierarchy: saved.budgetHierarchy || defaults.budgetHierarchy,
        tables: {
          register: saved.tables?.register || defaults.tables.register,
          budgetEntries: saved.tables?.budgetEntries || defaults.tables.budgetEntries,
          selfEmploymentIncome: saved.tables?.selfEmploymentIncome || defaults.tables.selfEmploymentIncome,
          selfEmploymentExpenses: saved.tables?.selfEmploymentExpenses || defaults.tables.selfEmploymentExpenses,
          mileageTracker: saved.tables?.mileageTracker || defaults.tables.mileageTracker,
          biWeeklyExpenses: saved.tables?.biWeeklyExpenses || defaults.tables.biWeeklyExpenses,
          carCostCalculator: saved.tables?.carCostCalculator || defaults.tables.carCostCalculator,
        },
      };
    } catch (error) {
      return defaults;
    }
  }

  function normalizeState() {
    state.tables.register = state.tables.register.map((row, index) => ({
      ...row,
      transactionId: row.transactionId || `BANK-${String(row.sourceRow || index + 1).padStart(6, "0")}`,
      reconciled: Boolean(row.reconciled),
    }));
    state.tables.budgetEntries = state.tables.budgetEntries.map((row, index) => normalizeBudgetEntryRow(row, index));
    state.tables.biWeeklyExpenses = state.tables.biWeeklyExpenses.map((row, index) => normalizePlanningRow(row, index));
    state.tables.carCostCalculator = state.tables.carCostCalculator.map((row, index) => {
      const perMile = safeNumber(row.perMile);
      const perPaycheck = safeNumber(row.perPaycheck);
      const basisMilesPerPaycheck =
        safeNumber(row.basisMilesPerPaycheck) ||
        (perMile ? round2(perPaycheck / perMile) : inferMileageBasisFromName(row.name));
      return {
        ...row,
        id: row.id || `carcost-${index + 1}`,
        basisMilesPerPaycheck,
      };
    });
    state.editing = { register: state.editing?.register || "", budgetEntries: state.editing?.budgetEntries || "" };
    state.editDrafts = {
      register: null,
      budgetEntries: null,
    };
    state.budgetHierarchy = normalizeBudgetHierarchy(state.budgetHierarchy);
    ensureBudgetHierarchyCoverage();
    normalizeHierarchySelectionState();
  }

  function persist() {
    localStorage.setItem(
      STORAGE_KEY,
      JSON.stringify({
        currentView: state.currentView,
        selfEmploymentSubview: state.selfEmploymentSubview,
        selectedTransactionId: state.selectedTransactionId,
        pendingReconcileIds: state.pendingReconcileIds,
        budgetProgressMainGroup: state.budgetProgressMainGroup,
        budgetProgressCategory: state.budgetProgressCategory,
        filters: state.filters,
        settings: state.settings,
        budgetHierarchy: state.budgetHierarchy,
        tables: state.tables,
      }),
    );
  }

  function render() {
    normalizeHierarchySelectionState();
    renderMeta();
    renderLookups();
    renderNav();
    renderDashboard();
    renderBanking();
    renderBudgeting();
    renderSetup();
    renderTransactionLookup();
    renderPlanning();
    renderPayperiod();
    renderSelfEmployment();
    syncControls();
    persist();
  }

  function renderMeta() {
    setText("meta-workbook", "Local starting data");
    setText("meta-modified", formatDateTime(seed.meta.workbookLastModified));
    setText("meta-extracted", formatDateTime(seed.meta.extractedAt));
  }

  function renderLookups() {
    const lookups = getDynamicLookups();
    renderDatalist("accounts-list", lookups.accounts);
    renderDatalist("categories-list", lookups.categories);
    renderDatalist("subcategories-list", lookups.subcategories);
    renderDatalist("budget-notes-list", lookups.budgetNotes);
    fillSelect("banking-account-filter", ["All accounts", ...lookups.accounts], state.filters.bankingAccount);
    fillSelect("budget-budget-filter", ["All budgets", ...lookups.budgets], state.filters.budgetBudget);
    fillSelect(
      "budget-category-filter",
      ["All categories", ...getCategoryOptionsForSelection(state.filters.budgetBudget)],
      state.filters.budgetCategory,
    );
    fillSelect(
      "budget-subcategory-filter",
      ["All subcategories", ...getSubcategoryOptionsForSelection(state.filters.budgetBudget, state.filters.budgetCategory)],
      state.filters.budgetSubcategory,
    );
    renderBudgetEntrySelectors();
    fillSelect(
      "planning-frequency",
      seed.lookups.paycheckFrequencies && seed.lookups.paycheckFrequencies.length
        ? seed.lookups.paycheckFrequencies
        : [state.settings.planningFrequency],
      state.settings.planningFrequency,
    );
  }

  function renderNav() {
    document.querySelectorAll(".view-tab").forEach((button) => {
      button.classList.toggle("is-active", button.dataset.view === state.currentView);
    });
    document.querySelectorAll(".view-section").forEach((panel) => {
      panel.classList.toggle("is-active", panel.dataset.viewPanel === state.currentView);
    });
    document.querySelectorAll(".subview-tab").forEach((button) => {
      button.classList.toggle("is-active", button.dataset.seView === state.selfEmploymentSubview);
    });
    document.querySelectorAll(".se-subview").forEach((panel) => {
      panel.classList.toggle("is-active", panel.dataset.sePanel === state.selfEmploymentSubview);
    });
  }

  function renderDashboard() {
    const registerRows = getRegisterRows();
    const budgetRows = getBudgetRows();
    const planningRows = getPlanningRows();
    const latestPaycheck = getLatestPaycheckAmount(registerRows);
    const unassigned = round2(sum(registerRows, "amount") - sum(budgetRows, "amount"));
    const cards = [
      {
        label: "Current Bank Balance",
        value: formatMoney(sum(registerRows, "amount")),
        detail: `${registerRows.length.toLocaleString()} register rows`,
      },
      {
        label: "Unassigned Cash",
        value: formatMoney(unassigned),
        detail: "Register total minus budget ledger total",
      },
      {
        label: "Planned Per Paycheck",
        value: formatMoney(sum(planningRows, "amount")),
        detail: state.settings.planningFrequency,
      },
      {
        label: "Latest Paycheck Margin",
        value: formatMoney(round2(latestPaycheck - sum(planningRows, "amount"))),
        detail: `Latest FNB personal deposit ${formatMoney(latestPaycheck)}`,
      },
    ];

    byId("dashboard-summary").innerHTML = cards.map(renderStatCard).join("");
    renderBarSummary("account-summary", getAccountSummary(registerRows), "Account", "Current balance");
    renderBarSummary("category-summary", getCategorySummary(budgetRows, unassigned), "Budget", "Current balance");
    renderOverviewTable("bank-overview", buildBankOverview(registerRows));
    renderOverviewTable("budget-overview", buildBudgetOverview(registerRows, budgetRows));
  }

  function renderBanking() {
    const rows = getRegisterRows();
    const search = state.filters.bankingSearch.trim().toLowerCase();
    const filtered = rows.filter((row) => {
      const matchesAccount =
        state.filters.bankingAccount === "All accounts" || row.accountName === state.filters.bankingAccount;
      const haystack = [row.transactionId, row.accountName, row.checkNumber, row.purchaseDate, row.clearedDate, formatMoney(row.amount)]
        .join(" ")
        .toLowerCase();
      return matchesAccount && (!search || haystack.includes(search));
    });

    const currentReconciledBalance = round2(
      filtered.reduce((total, row) => total + (row.reconciled ? safeNumber(row.amount) : 0), 0),
    );
    const pendingRows = filtered.filter((row) => state.pendingReconcileIds.includes(row.id) && !row.reconciled);
    const pendingAmount = round2(sum(pendingRows, "amount"));
    const previewReconciledBalance = round2(currentReconciledBalance + pendingAmount);
    const unreconciledFilteredRows = filtered.filter((row) => !row.reconciled);
    const allFilteredUnreconciledSelected =
      unreconciledFilteredRows.length > 0 &&
      unreconciledFilteredRows.every((row) => state.pendingReconcileIds.includes(row.id));

    byId("reconcile-summary").innerHTML = `
      <div class="summary-chip">
        <span>Filtered rows</span>
        <strong>${filtered.length.toLocaleString()}</strong>
      </div>
      <div class="summary-chip">
        <span>Current reconciled balance</span>
        <strong>${formatMoney(currentReconciledBalance)}</strong>
      </div>
      <div class="summary-chip">
        <span>Checked to reconcile</span>
        <strong>${formatMoney(pendingAmount)}</strong>
      </div>
      <div class="summary-chip">
        <span>Preview reconciled balance</span>
        <strong>${formatMoney(previewReconciledBalance)}</strong>
      </div>
    `;
    byId("reconcile-selected").disabled = pendingRows.length === 0;
    byId("banking-select-all").checked = allFilteredUnreconciledSelected;
    byId("banking-select-all").indeterminate =
      !allFilteredUnreconciledSelected && pendingRows.length > 0 && unreconciledFilteredRows.length > 0;

    renderTable(
      "banking-table",
      [
        {
          label: "Reconcile",
          render: (row) =>
            row.reconciled
              ? `<span class="pill">Reconciled</span>`
              : `<input type="checkbox" data-reconcile-id="${escapeAttr(row.id)}" ${state.pendingReconcileIds.includes(row.id) ? "checked" : ""} />`,
        },
        { label: "Transaction ID", render: (row) => transactionLinkCell(row.transactionId) },
        {
          label: "Cleared",
          render: (row) =>
            isEditingRow("register", row.id)
              ? `<input name="clearedDate" type="date" data-inline-edit-field="register" value="${escapeAttr(getRowDraft("register", row).clearedDate || "")}" />`
              : dateCell(row.clearedDate),
        },
        {
          label: "Account",
          render: (row) =>
            isEditingRow("register", row.id)
              ? `<input name="accountName" list="accounts-list" data-inline-edit-field="register" value="${escapeAttr(getRowDraft("register", row).accountName || "")}" />`
              : escapeHtml(row.accountName || ""),
        },
        {
          label: "Check #",
          render: (row) =>
            isEditingRow("register", row.id)
              ? `<input name="checkNumber" data-inline-edit-field="register" value="${escapeAttr(getRowDraft("register", row).checkNumber || "")}" />`
              : mono(row.checkNumber || "-"),
        },
        {
          label: "Amount",
          render: (row) =>
            isEditingRow("register", row.id)
              ? `<input name="amount" type="text" inputmode="decimal" data-money-input data-inline-edit-field="register" value="${escapeAttr(formatMoneyInputValue(getRowDraft("register", row).amount))}" />`
              : moneyCell(row.amount),
        },
        {
          label: "Purchase",
          render: (row) =>
            isEditingRow("register", row.id)
              ? `<input name="purchaseDate" type="date" data-inline-edit-field="register" value="${escapeAttr(getRowDraft("register", row).purchaseDate || "")}" />`
              : dateCell(row.purchaseDate),
        },
        { label: "Balance", render: (row) => moneyCell(row.computedBalance) },
        { label: "Actions", render: (row) => editActionButtons("register", row.id) },
      ],
      filtered.slice(0, 300),
      filtered.length > 300 ? `Showing 300 of ${filtered.length.toLocaleString()} matching rows.` : `${filtered.length.toLocaleString()} matching rows.`,
    );

    renderTwinLines();
  }

  function renderBudgeting() {
    const rows = getBudgetRows();
    const search = state.filters.budgetSearch.trim().toLowerCase();
    const filtered = rows.filter((row) => {
      const matchesBudget =
        state.filters.budgetBudget === "All budgets" || row.budget === state.filters.budgetBudget;
      const matchesCategory =
        state.filters.budgetCategory === "All categories" || row.category === state.filters.budgetCategory;
      const matchesSubcategory =
        state.filters.budgetSubcategory === "All subcategories" || row.subcategory === state.filters.budgetSubcategory;
      const haystack = [row.transactionId, row.budget, row.category, row.subcategory, row.personalNotes, row.date].join(" ").toLowerCase();
      return matchesBudget && matchesCategory && matchesSubcategory && (!search || haystack.includes(search));
    });

    byId("budget-subtotal").innerHTML = `
      <span>Filtered subtotal</span>
      <strong>${formatMoney(sum(filtered, "amount"))}</strong>
    `;

    renderTable(
      "budgeting-table",
      [
        { label: "Transaction ID", render: (row) => transactionLinkCell(row.transactionId) },
        {
          label: "Date",
          render: (row) =>
            isEditingRow("budgetEntries", row.id)
              ? `<input name="date" type="date" data-inline-edit-field="budgetEntries" value="${escapeAttr(getRowDraft("budgetEntries", row).date || "")}" />`
              : dateCell(row.date),
        },
        {
          label: "Budget",
          render: (row) =>
            isEditingRow("budgetEntries", row.id)
              ? renderSelectMarkup(
                  "budget",
                  getDynamicLookups().budgets,
                  getRowDraft("budgetEntries", row).budget,
                  "Choose budget",
                  `data-inline-edit-field="budgetEntries"`,
                )
              : escapeHtml(row.budget || ""),
        },
        {
          label: "Category",
          render: (row) =>
            isEditingRow("budgetEntries", row.id)
              ? renderSelectMarkup(
                  "category",
                  getCategoryOptionsForSelection(getRowDraft("budgetEntries", row).budget),
                  getRowDraft("budgetEntries", row).category,
                  "Choose category",
                  `data-inline-edit-field="budgetEntries"`,
                )
              : escapeHtml(row.category || ""),
        },
        {
          label: "Subcategory",
          render: (row) =>
            isEditingRow("budgetEntries", row.id)
              ? renderSelectMarkup(
                  "subcategory",
                  getSubcategoryOptionsForSelection(getRowDraft("budgetEntries", row).budget, getRowDraft("budgetEntries", row).category),
                  getRowDraft("budgetEntries", row).subcategory,
                  NO_SUBCATEGORY_LABEL,
                  `data-inline-edit-field="budgetEntries"`,
                )
              : escapeHtml(row.subcategory || NO_SUBCATEGORY_LABEL),
        },
        {
          label: "Amount",
          render: (row) =>
            isEditingRow("budgetEntries", row.id)
              ? `<input name="amount" type="text" inputmode="decimal" data-money-input data-inline-edit-field="budgetEntries" value="${escapeAttr(formatMoneyInputValue(getRowDraft("budgetEntries", row).amount))}" />`
              : moneyCell(row.amount),
        },
        { label: "Running", render: (row) => moneyCell(row.computedBudget) },
        {
          label: "Notes",
          render: (row) =>
            isEditingRow("budgetEntries", row.id)
              ? `<input name="personalNotes" list="budget-notes-list" data-inline-edit-field="budgetEntries" value="${escapeAttr(getRowDraft("budgetEntries", row).personalNotes || "")}" />`
              : escapeHtml(row.personalNotes || ""),
        },
        { label: "Address", render: (row) => mono(row.address) },
        { label: "Actions", render: (row) => editActionButtons("budgetEntries", row.id) },
      ],
      filtered.slice(0, 300),
      filtered.length > 300 ? `Showing 300 of ${filtered.length.toLocaleString()} matching rows.` : `${filtered.length.toLocaleString()} matching rows.`,
    );
  }

  function renderTransactionLookup() {
    const search = (state.filters.transactionIdSearch || state.selectedTransactionId || "").trim().toLowerCase();
    const groups = getTransactionLookupGroups().filter((group) => !search || group.transactionId.toLowerCase().includes(search));
    const summaryRows = groups.slice(0, 200);

    renderTable(
      "transaction-id-summary",
      [
        { label: "Transaction ID", render: (row) => `<button class="table-link" type="button" data-select-txn="${escapeAttr(row.transactionId)}">${escapeHtml(row.transactionId)}</button>` },
        { label: "Bank Rows", render: (row) => numberCell(row.bankRows.length) },
        { label: "Budget Rows", render: (row) => numberCell(row.budgetRows.length) },
        { label: "Bank Total", render: (row) => moneyCell(sum(row.bankRows, "amount")) },
        { label: "Budget Total", render: (row) => moneyCell(sum(row.budgetRows, "amount")) },
      ],
      summaryRows,
      summaryRows.length ? `${summaryRows.length.toLocaleString()} matching transaction IDs.` : "No transaction IDs matched.",
    );

    const selected = groups.find((group) => group.transactionId === state.selectedTransactionId) || groups[0];
    if (selected && !state.selectedTransactionId) {
      state.selectedTransactionId = selected.transactionId;
    }

    renderTable(
      "transaction-id-bank-results",
      [
        { label: "Transaction ID", render: (row) => mono(row.transactionId) },
        { label: "Cleared", render: (row) => dateCell(row.clearedDate) },
        { label: "Account", render: (row) => escapeHtml(row.accountName || "") },
        { label: "Amount", render: (row) => moneyCell(row.amount) },
        { label: "Purchase", render: (row) => dateCell(row.purchaseDate) },
      ],
      selected ? selected.bankRows : [],
      selected ? `Bank rows linked to ${selected.transactionId}.` : "Choose a transaction ID to inspect.",
    );

    renderTable(
      "transaction-id-budget-results",
      [
        { label: "Transaction ID", render: (row) => mono(row.transactionId) },
        { label: "Date", render: (row) => dateCell(row.date) },
        { label: "Budget", render: (row) => escapeHtml(row.budget || "") },
        { label: "Category", render: (row) => escapeHtml(row.category || "") },
        { label: "Subcategory", render: (row) => escapeHtml(row.subcategory || NO_SUBCATEGORY_LABEL) },
        { label: "Amount", render: (row) => moneyCell(row.amount) },
        { label: "Notes", render: (row) => escapeHtml(row.personalNotes || "") },
      ],
      selected ? selected.budgetRows : [],
      selected ? `Budget rows linked to ${selected.transactionId}.` : "Choose a transaction ID to inspect.",
    );
  }

  function renderSetup() {
    const budgets = state.budgetHierarchy;
    const categoryCount = budgets.reduce((total, budget) => total + budget.categories.length, 0);
    const subcategoryCount = budgets.reduce(
      (total, budget) => total + budget.categories.reduce((innerTotal, category) => innerTotal + category.subcategories.length, 0),
      0,
    );

    byId("setup-summary").innerHTML = `
      <div class="summary-chip">
        <span>Main budgets</span>
        <strong>${budgets.length.toLocaleString()}</strong>
      </div>
      <div class="summary-chip">
        <span>Categories</span>
        <strong>${categoryCount.toLocaleString()}</strong>
      </div>
      <div class="summary-chip">
        <span>Subcategories</span>
        <strong>${subcategoryCount.toLocaleString()}</strong>
      </div>
    `;

    byId("setup-board").innerHTML = budgets.length
      ? budgets
          .map(
            (budget) => `
            <section class="setup-budget-card" data-drop-target="budget" data-budget-name="${escapeAttr(budget.name)}">
              <div class="setup-budget-head">
                <div>
                  <span class="setup-kicker">Budget</span>
                  <h3>${escapeHtml(budget.name)}</h3>
                </div>
                <span class="pill">${budget.categories.length} categories</span>
              </div>
              <div class="setup-drop-note">Drop a category anywhere on this card to move it here.</div>
              <div class="setup-category-stack">
                ${budget.categories
                  .map(
                    (category) => `
                    <article
                      class="setup-category-card"
                      draggable="true"
                      data-drag-type="category"
                      data-drop-target="category"
                      data-budget-name="${escapeAttr(budget.name)}"
                      data-category-name="${escapeAttr(category.name)}"
                    >
                      <div class="setup-category-head">
                        <div>
                          <span class="setup-kicker">Category</span>
                          <strong>${escapeHtml(category.name)}</strong>
                        </div>
                        <span class="pill">${category.subcategories.length} subcategories</span>
                      </div>
                      <div class="setup-drop-note">Drop a subcategory anywhere in this card to move it here.</div>
                      <div class="setup-subcategory-wrap">
                        ${
                          category.subcategories.length
                            ? category.subcategories
                                .map(
                                  (subcategory) => `
                                  <button
                                    class="setup-subcategory-chip"
                                    type="button"
                                    draggable="true"
                                    data-drag-type="subcategory"
                                    data-budget-name="${escapeAttr(budget.name)}"
                                    data-category-name="${escapeAttr(category.name)}"
                                    data-subcategory-name="${escapeAttr(subcategory)}"
                                  >${escapeHtml(subcategory)}</button>`,
                                )
                                .join("")
                            : `<span class="empty-state compact">No subcategories yet.</span>`
                        }
                      </div>
                      <form class="setup-inline-form" data-setup-add="subcategory" data-budget-name="${escapeAttr(budget.name)}" data-category-name="${escapeAttr(category.name)}">
                        <label>
                          New subcategory
                          <input name="name" placeholder="Add a subcategory" required />
                        </label>
                        <button class="button button-secondary" type="submit">Add Subcategory</button>
                      </form>
                    </article>`,
                  )
                  .join("")}
              </div>
              <form class="setup-inline-form" data-setup-add="category" data-budget-name="${escapeAttr(budget.name)}">
                <label>
                  New category
                  <input name="name" placeholder="Add a category" required />
                </label>
                <button class="button button-secondary" type="submit">Add Category</button>
              </form>
            </section>`,
          )
          .join("")
      : `<div class="empty-state">Add a budget to start building your setup.</div>`;
  }

  function renderPlanning() {
    const planningRows = getPlanningRows();
    const annualTotal = sum(planningRows, "amountPerYear");
    const perPaycheckTotal = sum(planningRows, "amount");
    const latestPaycheck = getLatestPaycheckAmount(getRegisterRows());
    const summary = [
      { label: "Annual Planned Spend", value: formatMoney(annualTotal), detail: `${planningRows.length} planning rows` },
      { label: "Per Paycheck", value: formatMoney(perPaycheckTotal), detail: state.settings.planningFrequency },
      { label: "Latest Paycheck", value: formatMoney(latestPaycheck), detail: "Latest positive FNB personal deposit" },
      { label: "Paycheck Margin", value: formatMoney(round2(latestPaycheck - perPaycheckTotal)), detail: "Latest paycheck minus planned total" },
    ];
    byId("planning-summary").innerHTML = summary.map(renderStatCard).join("");
    renderEditableBiweeklyExpenses(planningRows);
    renderEditableCarCosts(getCarCostRows());
    renderBudgetProgress();
  }

  function renderPayperiod() {
    const report = getPayperiodReport();
    const summaryCards = [
      { label: "Opening Balance", value: formatMoney(report.openingBalance), detail: formatDate(report.startDate) },
      { label: "Total Income", value: formatMoney(report.totalIncome), detail: `${report.incomeRows.length} rows` },
      { label: "Total Expenses", value: formatMoney(report.totalExpenses), detail: `${report.expenseRows.length} rows` },
      { label: "Closing Balance", value: formatMoney(report.closingBalance), detail: formatDate(report.endDate) },
      { label: "Net", value: formatMoney(report.net), detail: "Closing minus opening" },
    ];
    byId("payperiod-summary").innerHTML = summaryCards.map(renderStatCard).join("");

    renderTable(
      "payperiod-income",
      [
        { label: "Transaction ID", render: (row) => transactionLinkCell(row.transactionId) },
        { label: "Date", render: (row) => dateCell(row.purchaseDate) },
        { label: "Account", render: (row) => escapeHtml(row.accountName) },
        { label: "Check #", render: (row) => mono(row.checkNumber || "-") },
        { label: "Amount", render: (row) => moneyCell(row.amount) },
      ],
      report.incomeRows,
      `${report.incomeRows.length} income rows in the pay period.`,
    );

    const expenseColumns = [
      { label: "Transaction ID", render: (row) => transactionLinkCell(row.transactionId) },
      { label: "Date", render: (row) => dateCell(row.date) },
      { label: "Budget", render: (row) => escapeHtml(row.budget || "") },
      { label: "Category", render: (row) => escapeHtml(row.category) },
      { label: "Subcategory", render: (row) => escapeHtml(row.subcategory || NO_SUBCATEGORY_LABEL) },
      { label: "Amount", render: (row) => moneyCell(row.amount) },
    ];
    renderTable("payperiod-expenses-left", expenseColumns, report.expenseRows.slice(0, 26), "First 26 expense rows.");
    renderTable(
      "payperiod-expenses-right",
      expenseColumns,
      report.expenseRows.slice(26),
      report.expenseRows.length > 26 ? "Overflow rows after the first 26." : "No overflow rows.",
    );
  }

  function renderSelfEmployment() {
    renderNav();
    renderTable(
      "se-income-table",
      [
        { label: "Date", render: (row) => dateCell(row.date) },
        { label: "Label", render: (row) => escapeHtml(row.label || "") },
        { label: "Amount", render: (row) => moneyCell(row.amount) },
        { label: "Sales Tax", render: (row) => moneyCell(row.salesTax) },
        { label: "SE Tax", render: (row) => moneyCell(row.selfEmploymentTax) },
        { label: "Tithe", render: (row) => moneyCell(row.tithe) },
        { label: "Left", render: (row) => moneyCell(row.amountLeft) },
        { label: "", render: (row) => deleteButton("selfEmploymentIncome", row.id) },
      ],
      getSelfEmploymentIncomeRows(),
      "Formulas use the same percentages as your current budgeting workflow.",
    );
    renderTable(
      "se-expenses-table",
      [
        { label: "Date", render: (row) => dateCell(row.date) },
        { label: "Tax Category", render: (row) => escapeHtml(row.taxCategory || "") },
        { label: "Amount", render: (row) => moneyCell(row.amount) },
        { label: "Rate", render: (row) => percentCell(row.taxDeductibleRate) },
        { label: "Deduction", render: (row) => moneyCell(row.taxableIncomeDeduction) },
        { label: "Yearly Total", render: (row) => moneyCell(row.yearlyDeductible) },
        { label: "", render: (row) => deleteButton("selfEmploymentExpenses", row.id) },
      ],
      getSelfEmploymentExpenseRows(),
      "Yearly totals are grouped by category and year.",
    );
    renderTable(
      "mileage-table",
      [
        { label: "Date", render: (row) => dateCell(row.date) },
        { label: "Vehicle", render: (row) => escapeHtml(row.vehicle || "") },
        { label: "Purpose", render: (row) => escapeHtml(row.purpose || "") },
        { label: "Miles", render: (row) => numberCell(row.miles) },
        { label: "Yearly Miles", render: (row) => numberCell(row.yearlyMileage) },
        { label: "Deduction", render: (row) => moneyCell(row.taxableIncomeDeduction) },
        { label: "", render: (row) => deleteButton("mileageTracker", row.id) },
      ],
      getMileageRows(),
      "Mileage deduction is calculated at $0.65 per mile.",
    );
  }

  function getRegisterRows() {
    const effectiveDate = todayIso();
    const grouped = new Map();
    for (const row of state.tables.register) {
      const account = row.accountName || "Unknown";
      const clearedDate = row.clearedDate || effectiveDate;
      if (!grouped.has(account)) grouped.set(account, new Map());
      const bucket = grouped.get(account);
      bucket.set(clearedDate, round2((bucket.get(clearedDate) || 0) + safeNumber(row.amount)));
    }

    const cumulative = new Map();
    grouped.forEach((dateMap, account) => {
      let running = 0;
      const output = new Map();
      [...dateMap.entries()].sort(([left], [right]) => compareDate(left, right)).forEach(([dateValue, amount]) => {
        running = round2(running + amount);
        output.set(dateValue, running);
      });
      cumulative.set(account, output);
    });

    return [...state.tables.register]
      .map((row) => {
        const clearedDate = row.clearedDate || effectiveDate;
        const account = row.accountName || "Unknown";
        return { ...row, computedBalance: cumulative.get(account)?.get(clearedDate) || 0 };
      })
      .sort((left, right) => (right.sourceRow || 0) - (left.sourceRow || 0));
  }

  function getBudgetRows() {
    const running = new Map();
    const computed = new Map();
    [...state.tables.budgetEntries]
      .sort((left, right) => compareDate(left.date, right.date) || safeNumber(left.address) - safeNumber(right.address))
      .forEach((row) => {
        const key = `${row.budget || ""}::${row.category || ""}::${row.subcategory || ""}`;
        const next = round2((running.get(key) || 0) + safeNumber(row.amount));
        running.set(key, next);
        computed.set(row.id, next);
      });

    return [...state.tables.budgetEntries]
      .map((row) => ({ ...row, computedBudget: computed.get(row.id) || 0 }))
      .sort((left, right) => (right.sourceRow || 0) - (left.sourceRow || 0));
  }

  function getPlanningRows() {
    const cyclesPerYear = parseCyclesPerYear(state.settings.planningFrequency);
    const annualTotal = sum(state.tables.biWeeklyExpenses, "amountPerYear");
    return [...state.tables.biWeeklyExpenses]
      .map((row) => ({
        ...row,
        amountPerYear: safeNumber(row.amountPerYear),
        amount: round2(safeNumber(row.amountPerYear) / cyclesPerYear),
        shareOfYearlyBudget: annualTotal ? safeNumber(row.amountPerYear) / annualTotal : 0,
      }))
      .sort(
        (left, right) =>
          compareText(left.budget, right.budget) ||
          compareText(left.category, right.category) ||
          compareText(left.subcategory, right.subcategory),
      );
  }

  function getPlanningCopyRows() {
    return getPlanningRows()
      .filter((row) => row.budget && row.category && safeNumber(row.amount) !== 0)
      .map((row) => ({
        budget: row.budget,
        category: row.category,
        subcategory: row.subcategory,
        amount: row.amount,
        date: byId("planning-post-date")?.value || todayIso(),
      }));
  }

  function getCarCostRows() {
    return [...state.tables.carCostCalculator].map((row) => {
      const milesPerCycle = safeNumber(row.milesPerCycle);
      const amount = safeNumber(row.amount);
      const perMile = milesPerCycle ? round2(amount / milesPerCycle) : 0;
      const basisMilesPerPaycheck = safeNumber(row.basisMilesPerPaycheck) || inferMileageBasisFromName(row.name);
      return {
        ...row,
        amount,
        milesPerCycle,
        basisMilesPerPaycheck,
        perMile,
        perPaycheck: round2(perMile * basisMilesPerPaycheck),
      };
    });
  }

  function getBudgetProgressRows() {
    const payperiod = getPayperiodReport();
    const planningBudgets = getPlanningRows().map((row) => row.budget).filter(Boolean);
    const mainGroups = new Map();
    const categoryGroups = new Map();
    const subcategoryGroups = new Map();

    for (const row of getBudgetRows()) {
      if (planningBudgets.length && !planningBudgets.includes(row.budget)) continue;
      const budgetKey = row.budget || "Unbudgeted";
      const categoryKey = `${budgetKey}::${row.category || DEFAULT_CATEGORY_NAME}`;
      const subcategoryKey = `${categoryKey}::${row.subcategory || ""}`;

      accumulateBudgetProgress(mainGroups, budgetKey, { key: budgetKey, label: budgetKey }, row, payperiod);
      accumulateBudgetProgress(
        categoryGroups,
        categoryKey,
        {
          key: categoryKey,
          label: row.category || DEFAULT_CATEGORY_NAME,
          parentKey: budgetKey,
        },
        row,
        payperiod,
      );
      accumulateBudgetProgress(
        subcategoryGroups,
        subcategoryKey,
        {
          key: subcategoryKey,
          label: row.subcategory || NO_SUBCATEGORY_LABEL,
          parentKey: categoryKey,
        },
        row,
        payperiod,
      );
    }

    const mainRows = finalizeBudgetProgress([...mainGroups.values()]).sort((left, right) => compareText(left.label, right.label));
    const categoryRows = finalizeBudgetProgress(
      [...categoryGroups.values()].filter((row) => !state.budgetProgressMainGroup || row.parentKey === state.budgetProgressMainGroup),
    ).sort((left, right) => compareText(left.label, right.label));
    const subcategoryRows = finalizeBudgetProgress(
      [...subcategoryGroups.values()].filter((row) => !state.budgetProgressCategory || row.parentKey === state.budgetProgressCategory),
    ).sort((left, right) => compareText(left.label, right.label));

    return { mainRows, categoryRows, subcategoryRows };
  }

  function accumulateBudgetProgress(store, key, seedRow, row, payperiod) {
    if (!store.has(key)) {
      store.set(key, {
        ...seedRow,
        currentTotal: 0,
        addedThisPayperiod: 0,
        spentThisPayperiod: 0,
      });
    }
    const group = store.get(key);
    group.currentTotal = round2(group.currentTotal + safeNumber(row.amount));
    if (row.date && compareDate(row.date, payperiod.startDate) >= 0 && compareDate(row.date, payperiod.endDate) < 0) {
      if (safeNumber(row.amount) > 0) group.addedThisPayperiod = round2(group.addedThisPayperiod + safeNumber(row.amount));
      if (safeNumber(row.amount) < 0) group.spentThisPayperiod = round2(group.spentThisPayperiod + Math.abs(safeNumber(row.amount)));
    }
  }

  function finalizeBudgetProgress(rows) {
    return rows
      .map((row) => ({
        ...row,
        spentBasis: round2(Math.max(row.addedThisPayperiod, row.spentThisPayperiod)),
        overspentAmount: round2(Math.max(0, row.spentThisPayperiod - row.addedThisPayperiod)),
        spentRatio:
          Math.max(row.addedThisPayperiod, row.spentThisPayperiod) > 0
            ? clamp01(Math.min(row.spentThisPayperiod, row.addedThisPayperiod) / Math.max(row.addedThisPayperiod, row.spentThisPayperiod))
            : 0,
        overspentRatio:
          Math.max(row.addedThisPayperiod, row.spentThisPayperiod) > 0
            ? clamp01(Math.max(0, row.spentThisPayperiod - row.addedThisPayperiod) / Math.max(row.addedThisPayperiod, row.spentThisPayperiod))
            : 0,
      }))
      .filter((row) => row.label && (row.currentTotal !== 0 || row.addedThisPayperiod !== 0 || row.spentThisPayperiod !== 0));
  }

  function getSelfEmploymentIncomeRows() {
    return [...state.tables.selfEmploymentIncome]
      .map((row) => {
        const amount = safeNumber(row.amount);
        const salesTaxable = Boolean(row.salesTaxable);
        const selfEmploymentTaxEnabled = Boolean(row.selfEmploymentTaxEnabled);
        const salesTax = round2(salesTaxable ? amount * 0.07875 : 0);
        const selfEmploymentTax = round2(0.2 * (amount - salesTax));
        const tithe = round2((amount - selfEmploymentTax * (selfEmploymentTaxEnabled ? 1 : 0) - salesTax) * 0.13);
        const amountLeft = round2(amount - selfEmploymentTax * (selfEmploymentTaxEnabled ? 1 : 0) - tithe - salesTax * (salesTaxable ? 1 : 0));
        return { ...row, salesTax, selfEmploymentTax, tithe, amountLeft };
      })
      .sort((left, right) => compareDate(right.date, left.date) || (right.sourceRow || 0) - (left.sourceRow || 0));
  }

  function getSelfEmploymentExpenseRows() {
    const running = new Map();
    const computed = new Map();
    [...state.tables.selfEmploymentExpenses]
      .sort((left, right) => compareDate(left.date, right.date) || (left.sourceRow || 0) - (right.sourceRow || 0))
      .forEach((row) => {
        const year = (row.date || "").slice(0, 4) || "unknown";
        const key = `${year}::${row.taxCategory || ""}`;
        const deduction = round2(safeNumber(row.amount) * safeNumber(row.taxDeductibleRate));
        const next = round2((running.get(key) || 0) + deduction);
        running.set(key, next);
        computed.set(row.id, { deduction, yearlyDeductible: next });
      });

    return [...state.tables.selfEmploymentExpenses]
      .map((row) => ({
        ...row,
        taxableIncomeDeduction: computed.get(row.id)?.deduction || 0,
        yearlyDeductible: computed.get(row.id)?.yearlyDeductible || 0,
      }))
      .sort((left, right) => compareDate(right.date, left.date) || (right.sourceRow || 0) - (left.sourceRow || 0));
  }

  function getMileageRows() {
    const running = new Map();
    const computed = new Map();
    [...state.tables.mileageTracker]
      .sort((left, right) => compareDate(left.date, right.date) || (left.sourceRow || 0) - (right.sourceRow || 0))
      .forEach((row) => {
        const year = (row.date || "").slice(0, 4) || "unknown";
        const key = `${year}::${row.vehicle || ""}::${row.purpose || ""}`;
        const deduction = round2(0.65 * safeNumber(row.miles));
        const next = round2((running.get(key) || 0) + safeNumber(row.miles));
        running.set(key, next);
        computed.set(row.id, { deduction, yearlyMileage: next });
      });

    return [...state.tables.mileageTracker]
      .map((row) => ({
        ...row,
        taxableIncomeDeduction: computed.get(row.id)?.deduction || 0,
        yearlyMileage: computed.get(row.id)?.yearlyMileage || 0,
      }))
      .sort((left, right) => compareDate(right.date, left.date) || (right.sourceRow || 0) - (left.sourceRow || 0));
  }

  function getAccountSummary(registerRows) {
    const grouped = new Map();
    for (const row of registerRows) {
      const account = row.accountName || "Unknown";
      grouped.set(account, round2((grouped.get(account) || 0) + safeNumber(row.amount)));
    }
    return [...grouped.entries()].map(([label, value]) => ({ label, value })).sort((left, right) => Math.abs(right.value) - Math.abs(left.value));
  }

  function getCategorySummary(budgetRows, unassigned) {
    const grouped = new Map();
    for (const row of budgetRows) {
      const budget = row.budget || "Unknown";
      grouped.set(budget, round2((grouped.get(budget) || 0) + safeNumber(row.amount)));
    }
    return [{ label: "Unassigned", value: unassigned }, ...[...grouped.entries()].map(([label, value]) => ({ label, value }))].sort(
      (left, right) => Math.abs(right.value) - Math.abs(left.value),
    );
  }

  function buildBankOverview(registerRows) {
    const step = Math.max(1, safeNumber(state.settings.bankOverviewStepDays) || 1);
    const startDate = state.settings.bankOverviewStart || addDays(todayIso(), -30 * step);
    const dates = Array.from({ length: 31 }, (_, index) => addDays(startDate, index * step));
    const accounts = getDynamicLookups().accounts;
    const rows = dates.map((dateValue) => {
      const values = {};
      let overall = 0;
      for (const row of registerRows) {
        const clearedDate = row.clearedDate || todayIso();
        if (compareDate(clearedDate, dateValue) <= 0) {
          overall = round2(overall + safeNumber(row.amount));
          values[row.accountName] = round2((values[row.accountName] || 0) + safeNumber(row.amount));
        }
      }
      return { date: dateValue, overall, values };
    });
    return { columns: ["Date", "Overall Balance", ...accounts], rows: rows.map((row) => [row.date, row.overall, ...accounts.map((account) => row.values[account] || 0)]) };
  }

  function buildBudgetOverview(registerRows, budgetRows) {
    const step = Math.max(1, safeNumber(state.settings.budgetOverviewStepDays) || 1);
    const startDate = state.settings.budgetOverviewStart || addDays(todayIso(), -30 * step);
    const dates = Array.from({ length: 31 }, (_, index) => addDays(startDate, index * step));
    const budgets = getDynamicLookups().budgets;
    const rows = dates.map((dateValue) => {
      const values = {};
      let registerTotal = 0;
      let budgetTotal = 0;
      for (const row of registerRows) {
        if (row.purchaseDate && compareDate(row.purchaseDate, dateValue) <= 0) registerTotal = round2(registerTotal + safeNumber(row.amount));
      }
      for (const row of budgetRows) {
        if (row.date && compareDate(row.date, dateValue) <= 0) {
          budgetTotal = round2(budgetTotal + safeNumber(row.amount));
          values[row.budget] = round2((values[row.budget] || 0) + safeNumber(row.amount));
        }
      }
      return { date: dateValue, unassigned: round2(registerTotal - budgetTotal), values };
    });
    return { columns: ["Date", "Unassigned", ...budgets], rows: rows.map((row) => [row.date, row.unassigned, ...budgets.map((budget) => row.values[budget] || 0)]) };
  }

  function getPayperiodReport() {
    const registerRows = getRegisterRows();
    const budgetRows = getBudgetRows();
    const startDate = state.settings.payperiodStart || todayIso();
    const endDate = addDays(startDate, safeNumber(state.settings.payperiodCycleDays) || 14);
    const openingBalance = round2(registerRows.filter((row) => row.purchaseDate && compareDate(row.purchaseDate, startDate) < 0).reduce((total, row) => total + safeNumber(row.amount), 0));
    const closingBalance = round2(registerRows.filter((row) => row.purchaseDate && compareDate(row.purchaseDate, endDate) < 0).reduce((total, row) => total + safeNumber(row.amount), 0));
    const incomeRows = registerRows
      .filter((row) => row.purchaseDate && compareDate(row.purchaseDate, startDate) >= 0 && compareDate(row.purchaseDate, endDate) < 0 && safeNumber(row.amount) > 0 && PAYPERIOD_INCOME_ACCOUNTS.has(row.accountName))
      .sort((left, right) => compareDate(left.purchaseDate, right.purchaseDate) || (left.sourceRow || 0) - (right.sourceRow || 0));
    const expenseRows = budgetRows
      .filter(
        (row) =>
          row.date &&
          compareDate(row.date, startDate) >= 0 &&
          compareDate(row.date, endDate) < 0 &&
          safeNumber(row.amount) < 0 &&
          !PAYPERIOD_EXCLUDED_CATEGORIES.has(row.category) &&
          !PAYPERIOD_EXCLUDED_CATEGORIES.has(row.subcategory),
      )
      .sort((left, right) => compareDate(left.date, right.date) || (left.sourceRow || 0) - (right.sourceRow || 0));
    return {
      startDate,
      endDate,
      openingBalance,
      closingBalance,
      incomeRows,
      totalIncome: round2(sum(incomeRows, "amount")),
      expenseRows,
      totalExpenses: round2(closingBalance - openingBalance - round2(sum(incomeRows, "amount"))),
      net: round2(closingBalance - openingBalance),
    };
  }

  function getLatestPaycheckAmount(registerRows) {
    const candidates = registerRows
      .filter((row) => row.accountName === "FNB - PERSONAL" && safeNumber(row.amount) > 0)
      .sort((left, right) => compareDate(right.purchaseDate, left.purchaseDate) || (right.sourceRow || 0) - (left.sourceRow || 0));
    return safeNumber(candidates[0]?.amount);
  }

  function getDynamicLookups() {
    const budgets = state.budgetHierarchy.map((row) => row.name);
    return {
      accounts: unique(state.tables.register.map((row) => row.accountName)),
      budgets,
      categories: unique(
        state.budgetHierarchy.flatMap((budget) => budget.categories.map((category) => category.name)),
      ),
      subcategories: unique(
        state.budgetHierarchy.flatMap((budget) =>
          budget.categories.flatMap((category) => category.subcategories),
        ),
      ),
      budgetNotes: unique(state.tables.budgetEntries.map((row) => row.personalNotes)),
    };
  }

  function getTransactionLookupGroups() {
    const groups = new Map();
    for (const row of getRegisterRows()) {
      const id = row.transactionId || "";
      if (!groups.has(id)) groups.set(id, { transactionId: id, bankRows: [], budgetRows: [] });
      groups.get(id).bankRows.push(row);
    }
    for (const row of getBudgetRows()) {
      const id = row.transactionId || "";
      if (!groups.has(id)) groups.set(id, { transactionId: id, bankRows: [], budgetRows: [] });
      groups.get(id).budgetRows.push(row);
    }
    return [...groups.values()].filter((group) => group.transactionId).sort((left, right) => compareText(left.transactionId, right.transactionId));
  }

  function bindStaticEvents() {
    document.querySelectorAll(".view-tab").forEach((button) => {
      button.addEventListener("click", () => {
        state.currentView = button.dataset.view;
        renderNav();
        persist();
      });
    });
    document.querySelectorAll(".subview-tab").forEach((button) => {
      button.addEventListener("click", () => {
        state.selfEmploymentSubview = button.dataset.seView;
        renderNav();
        persist();
      });
    });
    byId("export-state").addEventListener("click", exportState);
    byId("reset-state").addEventListener("click", resetState);
    byId("banking-form").addEventListener("submit", handleBankingSubmit);
    byId("budget-form").addEventListener("submit", handleBudgetSubmit);
    byId("budget-form").addEventListener("change", handleBudgetFormChange);
    byId("twin-form").addEventListener("submit", handleTwinSubmit);
    byId("setup-budget-form").addEventListener("submit", handleSetupBudgetSubmit);
    byId("planning-post-form").addEventListener("submit", handlePlanningPost);
    byId("planning-frequency-form").addEventListener("change", (event) => {
      state.settings.planningFrequency = event.target.value;
      renderPlanning();
      persist();
    });
    byId("se-income-form").addEventListener("submit", handleSelfEmploymentIncomeSubmit);
    byId("se-expense-form").addEventListener("submit", handleSelfEmploymentExpenseSubmit);
    byId("mileage-form").addEventListener("submit", handleMileageSubmit);
    byId("payperiod-form").addEventListener("change", handlePayperiodChange);
    byId("bank-overview-form").addEventListener("change", handleBankOverviewChange);
    byId("budget-overview-form").addEventListener("change", handleBudgetOverviewChange);
    byId("banking-search").addEventListener("input", (event) => { state.filters.bankingSearch = event.target.value; renderBanking(); persist(); });
    byId("banking-account-filter").addEventListener("change", (event) => { state.filters.bankingAccount = event.target.value; renderBanking(); persist(); });
    byId("banking-select-all").addEventListener("change", handleBankingSelectAllToggle);
    byId("reconcile-selected").addEventListener("click", handleReconcileSelected);
    byId("budget-search").addEventListener("input", (event) => { state.filters.budgetSearch = event.target.value; renderBudgeting(); persist(); });
    byId("budget-budget-filter").addEventListener("change", handleBudgetFilterChange);
    byId("budget-category-filter").addEventListener("change", handleBudgetCategoryFilterChange);
    byId("budget-subcategory-filter").addEventListener("change", handleBudgetSubcategoryFilterChange);
    byId("transaction-id-search").addEventListener("input", (event) => {
      state.filters.transactionIdSearch = event.target.value;
      state.selectedTransactionId = event.target.value.trim();
      renderTransactionLookup();
      persist();
    });
    byId("planning-post-date").addEventListener("change", renderPlanning);
    byId("add-twin-line").addEventListener("click", () => { state.twinLines.push(emptyTwinLine()); renderTwinLines(); });
    byId("add-biweekly-row").addEventListener("click", () => {
      state.tables.biWeeklyExpenses.push({ id: makeId("biweekly"), budget: "", category: "", subcategory: "", amountPerYear: 0 });
      renderPlanning();
      persist();
    });
    byId("add-car-cost-row").addEventListener("click", () => {
      state.tables.carCostCalculator.push({ id: makeId("carcost"), name: "", amount: 0, milesPerCycle: 0, basisMilesPerPaycheck: inferMileageBasisFromName("") });
      renderPlanning();
      persist();
    });
    byId("twin-lines").addEventListener("input", handleTwinLineInput);
    byId("twin-lines").addEventListener("change", handleTwinLineInput);
    byId("twin-lines").addEventListener("click", handleTwinLineClick);
    byId("planning-table").addEventListener("input", handlePlanningTableInput);
    byId("planning-table").addEventListener("change", handlePlanningTableInput);
    byId("planning-table").addEventListener("click", handlePlanningTableClick);
    byId("car-cost-table").addEventListener("input", handleCarCostTableInput);
    byId("car-cost-table").addEventListener("click", handleCarCostTableClick);
    byId("banking-table").addEventListener("input", handleBankingTableInput);
    byId("banking-table").addEventListener("change", handleBankingTableInput);
    byId("budgeting-table").addEventListener("input", handleBudgetingTableInput);
    byId("budgeting-table").addEventListener("change", handleBudgetingTableInput);
    byId("setup-board").addEventListener("submit", handleSetupBoardSubmit);
    byId("setup-board").addEventListener("dragstart", handleSetupDragStart);
    byId("setup-board").addEventListener("dragend", handleSetupDragEnd);
    byId("setup-board").addEventListener("dragover", handleSetupDragOver);
    byId("setup-board").addEventListener("dragleave", handleSetupDragLeave);
    byId("setup-board").addEventListener("drop", handleSetupDrop);
    document.body.addEventListener("focusin", handleMoneyInputFocus);
    document.body.addEventListener("focusout", handleMoneyInputBlur);
    document.body.addEventListener("click", handleGlobalClick);
  }

  function primeForms() {
    setInputValue("planning-post-date", todayIso());
    byId("banking-form").elements.clearedDate.value = todayIso();
    byId("banking-form").elements.purchaseDate.value = todayIso();
    byId("budget-form").elements.date.value = todayIso();
    byId("se-income-form").elements.date.value = todayIso();
    byId("se-expense-form").elements.date.value = todayIso();
    byId("mileage-form").elements.date.value = todayIso();
    const twinForm = byId("twin-form");
    twinForm.elements.clearedDate.value = seed.defaults.twinEntry.clearedDate || todayIso();
    twinForm.elements.accountName.value = seed.defaults.twinEntry.accountName || "";
    twinForm.elements.checkNumber.value = seed.defaults.twinEntry.checkNumber || "";
    twinForm.elements.purchaseDate.value = seed.defaults.twinEntry.purchaseDate || todayIso();
    syncBudgetEntryFormSelection();
    syncControls();
  }

  function syncControls() {
    setInputValue("banking-search", state.filters.bankingSearch);
    setInputValue("budget-search", state.filters.budgetSearch);
    setInputValue("budget-budget-filter", state.filters.budgetBudget);
    setInputValue("transaction-id-search", state.filters.transactionIdSearch || state.selectedTransactionId);
    setInputValue("payperiod-start", state.settings.payperiodStart);
    setInputValue("payperiod-days", state.settings.payperiodCycleDays);
    setInputValue("bank-overview-start", state.settings.bankOverviewStart);
    setInputValue("bank-overview-step", state.settings.bankOverviewStepDays);
    setInputValue("budget-overview-start", state.settings.budgetOverviewStart);
    setInputValue("budget-overview-step", state.settings.budgetOverviewStepDays);
    formatMoneyInputs(document);
  }

  function handleBankingSubmit(event) {
    event.preventDefault();
    const form = event.currentTarget;
    state.tables.register.unshift({
      id: makeId("register"),
      transactionId: generateTransactionId(),
      sourceRow: nextSourceRow(state.tables.register),
      balance: null,
      clearedDate: form.elements.clearedDate.value,
      accountName: form.elements.accountName.value.trim(),
      checkNumber: form.elements.checkNumber.value.trim(),
      amount: safeMoney(form.elements.amount.value),
      purchaseDate: form.elements.purchaseDate.value,
    });
    form.reset();
    primeForms();
    render();
  }

  function handleBankingTableInput(event) {
    const checkbox = event.target.closest("[data-reconcile-id]");
    if (checkbox) {
      const id = checkbox.dataset.reconcileId;
      if (checkbox.checked) {
        if (!state.pendingReconcileIds.includes(id)) state.pendingReconcileIds.push(id);
      } else {
        state.pendingReconcileIds = state.pendingReconcileIds.filter((item) => item !== id);
      }
      renderBanking();
      persist();
      return;
    }
    const field = event.target.closest('[data-inline-edit-field="register"]');
    if (!field || !state.editDrafts.register) return;
    state.editDrafts.register[field.name] = field.name === "amount" ? safeMoney(field.value) : field.value;
    if (field.matches("[data-money-input]") && event.type === "change") {
      formatMoneyInput(field);
    }
    persist();
  }

  function handleBudgetingTableInput(event) {
    const field = event.target.closest('[data-inline-edit-field="budgetEntries"]');
    if (!field || !state.editDrafts.budgetEntries) return;
    if (field.name === "amount") {
      state.editDrafts.budgetEntries.amount = safeMoney(field.value);
    } else if (field.name === "budget") {
      state.editDrafts.budgetEntries.budget = field.value;
      state.editDrafts.budgetEntries.category = "";
      state.editDrafts.budgetEntries.subcategory = "";
      renderBudgeting();
      persist();
      return;
    } else if (field.name === "category") {
      state.editDrafts.budgetEntries.category = field.value;
      state.editDrafts.budgetEntries.subcategory = "";
      renderBudgeting();
      persist();
      return;
    } else {
      state.editDrafts.budgetEntries[field.name] = field.value;
    }
    if (field.matches("[data-money-input]") && event.type === "change") {
      formatMoneyInput(field);
    }
    persist();
  }

  function handleBankingSelectAllToggle(event) {
    const checked = event.target.checked;
    const rows = getRegisterRows().filter((row) => {
      const matchesAccount =
        state.filters.bankingAccount === "All accounts" || row.accountName === state.filters.bankingAccount;
      const haystack = [row.transactionId, row.accountName, row.checkNumber, row.purchaseDate, row.clearedDate, formatMoney(row.amount)]
        .join(" ")
        .toLowerCase();
      const matchesSearch = !state.filters.bankingSearch.trim() || haystack.includes(state.filters.bankingSearch.trim().toLowerCase());
      return matchesAccount && matchesSearch && !row.reconciled;
    });

    if (checked) {
      const ids = new Set([...state.pendingReconcileIds, ...rows.map((row) => row.id)]);
      state.pendingReconcileIds = [...ids];
    } else {
      const idsToRemove = new Set(rows.map((row) => row.id));
      state.pendingReconcileIds = state.pendingReconcileIds.filter((id) => !idsToRemove.has(id));
    }

    renderBanking();
    persist();
  }

  function handleReconcileSelected() {
    if (!state.pendingReconcileIds.length) return;
    const selected = new Set(state.pendingReconcileIds);
    state.tables.register = state.tables.register.map((row) =>
      selected.has(row.id) ? { ...row, reconciled: true } : row,
    );
    state.pendingReconcileIds = [];
    renderBanking();
    persist();
  }

  function handleBudgetFormChange(event) {
    if (event.target.name === "budget") {
      byId("budget-form").elements.category.value = "";
      byId("budget-form").elements.subcategory.value = "";
    }
    if (event.target.name === "category") {
      byId("budget-form").elements.subcategory.value = "";
    }
    renderLookups();
    persist();
  }

  function handleBudgetFilterChange(event) {
    state.filters.budgetBudget = event.target.value;
    state.filters.budgetCategory = "All categories";
    state.filters.budgetSubcategory = "All subcategories";
    renderLookups();
    renderBudgeting();
    persist();
  }

  function handleBudgetCategoryFilterChange(event) {
    state.filters.budgetCategory = event.target.value;
    state.filters.budgetSubcategory = "All subcategories";
    renderLookups();
    renderBudgeting();
    persist();
  }

  function handleBudgetSubcategoryFilterChange(event) {
    state.filters.budgetSubcategory = event.target.value;
    renderBudgeting();
    persist();
  }

  function handleSetupBudgetSubmit(event) {
    event.preventDefault();
    const name = normalizeHierarchyName(event.currentTarget.elements.name.value);
    if (!name) return;
    addBudgetNode(name);
    event.currentTarget.reset();
    render();
  }

  function handleBudgetSubmit(event) {
    event.preventDefault();
    const form = event.currentTarget;
    const nextAddress = Math.max(0, ...state.tables.budgetEntries.map((row) => safeNumber(row.address))) + 1;
    state.tables.budgetEntries.unshift({
      id: makeId("budgetEntries"),
      transactionId: generateTransactionId(),
      sourceRow: nextAddress,
      budget: form.elements.budget.value.trim(),
      category: form.elements.category.value.trim(),
      subcategory: form.elements.subcategory.value.trim(),
      amount: safeMoney(form.elements.amount.value),
      date: form.elements.date.value,
      personalNotes: form.elements.personalNotes.value.trim(),
      address: nextAddress,
    });
    form.reset();
    primeForms();
    render();
  }

  function handleTwinSubmit(event) {
    event.preventDefault();
    const form = event.currentTarget;
    const splits = state.twinLines
      .map((line) => ({
        budget: line.budget.trim(),
        category: line.category.trim(),
        subcategory: line.subcategory.trim(),
        amount: safeMoney(line.amount),
        personalNotes: line.personalNotes.trim(),
      }))
      .filter((line) => line.budget && line.category && safeMoney(line.amount) !== 0);
    if (!splits.length) return;

    const transactionId = generateTransactionId();
    const totalAmount = round2(splits.reduce((total, line) => total + safeMoney(line.amount), 0));
    state.tables.register.unshift({
      id: makeId("register"),
      transactionId,
      sourceRow: nextSourceRow(state.tables.register),
      balance: null,
      clearedDate: form.elements.clearedDate.value,
      accountName: form.elements.accountName.value.trim(),
      checkNumber: form.elements.checkNumber.value.trim(),
      amount: totalAmount,
      purchaseDate: form.elements.purchaseDate.value,
    });

    let nextAddress = Math.max(0, ...state.tables.budgetEntries.map((row) => safeNumber(row.address)));
    for (const split of splits) {
      nextAddress += 1;
      state.tables.budgetEntries.unshift({
        id: makeId("budgetEntries"),
        transactionId,
        sourceRow: nextAddress,
        budget: split.budget,
        category: split.category,
        subcategory: split.subcategory,
        amount: split.amount,
        date: form.elements.purchaseDate.value,
        personalNotes: split.personalNotes,
        address: nextAddress,
      });
    }
    state.twinLines = [emptyTwinLine(), emptyTwinLine(), emptyTwinLine()];
    primeForms();
    render();
  }

  function handlePlanningPost(event) {
    event.preventDefault();
    const postDate = byId("planning-post-date").value || todayIso();
    let nextAddress = Math.max(0, ...state.tables.budgetEntries.map((row) => safeNumber(row.address)));
    for (const row of getPlanningCopyRows()) {
      nextAddress += 1;
      state.tables.budgetEntries.unshift({
        id: makeId("budgetEntries"),
        transactionId: generateTransactionId(),
        sourceRow: nextAddress,
        budget: row.budget,
        category: row.category,
        subcategory: row.subcategory,
        amount: row.amount,
        date: postDate,
        personalNotes: "",
        address: nextAddress,
      });
    }
    render();
  }

  function handleSelfEmploymentIncomeSubmit(event) {
    event.preventDefault();
    const form = event.currentTarget;
    state.tables.selfEmploymentIncome.unshift({ id: makeId("se-income"), sourceRow: nextSourceRow(state.tables.selfEmploymentIncome), date: form.elements.date.value, amount: safeMoney(form.elements.amount.value), label: form.elements.label.value.trim(), notes: form.elements.notes.value.trim(), salesTaxable: form.elements.salesTaxable.checked, selfEmploymentTaxEnabled: form.elements.selfEmploymentTaxEnabled.checked });
    form.reset();
    primeForms();
    renderSelfEmployment();
    persist();
  }

  function handleSelfEmploymentExpenseSubmit(event) {
    event.preventDefault();
    const form = event.currentTarget;
    state.tables.selfEmploymentExpenses.unshift({ id: makeId("se-expense"), sourceRow: nextSourceRow(state.tables.selfEmploymentExpenses), date: form.elements.date.value, amount: safeMoney(form.elements.amount.value), taxCategory: form.elements.taxCategory.value.trim(), taxDeductibleRate: safeNumber(form.elements.taxDeductibleRate.value), notes: form.elements.notes.value.trim() });
    form.reset();
    primeForms();
    renderSelfEmployment();
    persist();
  }

  function handleMileageSubmit(event) {
    event.preventDefault();
    const form = event.currentTarget;
    state.tables.mileageTracker.unshift({ id: makeId("mileage"), sourceRow: nextSourceRow(state.tables.mileageTracker), date: form.elements.date.value, miles: safeNumber(form.elements.miles.value), vehicle: form.elements.vehicle.value.trim(), purpose: form.elements.purpose.value.trim(), notes: form.elements.notes.value.trim() });
    form.reset();
    primeForms();
    renderSelfEmployment();
    persist();
  }

  function handlePayperiodChange() {
    state.settings.payperiodStart = byId("payperiod-start").value;
    state.settings.payperiodCycleDays = safeNumber(byId("payperiod-days").value);
    renderPayperiod();
    renderPlanning();
    persist();
  }

  function handleBankOverviewChange() {
    state.settings.bankOverviewStart = byId("bank-overview-start").value;
    state.settings.bankOverviewStepDays = safeNumber(byId("bank-overview-step").value);
    renderDashboard();
    persist();
  }

  function handleBudgetOverviewChange() {
    state.settings.budgetOverviewStart = byId("budget-overview-start").value;
    state.settings.budgetOverviewStepDays = safeNumber(byId("budget-overview-step").value);
    renderDashboard();
    persist();
  }

  function handleTwinLineInput(event) {
    const line = event.target.closest(".twin-line");
    if (!line) return;
    const index = safeNumber(line.dataset.index);
    state.twinLines[index][event.target.name] = event.target.value;
    if (event.target.name === "budget") {
      state.twinLines[index].category = "";
      state.twinLines[index].subcategory = "";
      renderTwinLines();
      persist();
      return;
    }
    if (event.target.name === "category") {
      state.twinLines[index].subcategory = "";
      renderTwinLines();
      persist();
      return;
    }
    if (event.target.matches("[data-money-input]")) {
      formatMoneyInput(event.target);
    }
    updateTwinTotal();
    persist();
  }

  function handleTwinLineClick(event) {
    const button = event.target.closest("[data-remove-twin-line]");
    if (!button) return;
    state.twinLines.splice(safeNumber(button.dataset.removeTwinLine), 1);
    if (!state.twinLines.length) state.twinLines.push(emptyTwinLine());
    renderTwinLines();
  }

  function handlePlanningTableInput(event) {
    const row = event.target.closest("[data-biweekly-id]");
    if (!row) return;
    const id = row.dataset.biweeklyId;
    const target = state.tables.biWeeklyExpenses.find((item) => item.id === id);
    if (!target) return;
    if (event.target.name === "amountPerYear") {
      target.amountPerYear = safeMoney(event.target.value);
    } else if (event.target.name === "budget") {
      target.budget = event.target.value;
      target.category = "";
      target.subcategory = "";
    } else if (event.target.name === "category") {
      target.category = event.target.value;
      target.subcategory = "";
    } else {
      target[event.target.name] = event.target.value;
    }
    renderPlanning();
    persist();
  }

  function handlePlanningTableClick(event) {
    const button = event.target.closest("[data-delete-biweekly]");
    if (!button) return;
    state.tables.biWeeklyExpenses = state.tables.biWeeklyExpenses.filter((row) => row.id !== button.dataset.deleteBiweekly);
    renderPlanning();
    persist();
  }

  function handleCarCostTableInput(event) {
    const row = event.target.closest("[data-carcost-id]");
    if (!row) return;
    const id = row.dataset.carcostId;
    const target = state.tables.carCostCalculator.find((item) => item.id === id);
    if (!target) return;
    target[event.target.name] = event.target.name === "name"
      ? event.target.value
      : event.target.name === "amount"
        ? safeMoney(event.target.value)
        : safeNumber(event.target.value);
    if (event.target.name === "name" && !safeNumber(target.basisMilesPerPaycheck)) {
      target.basisMilesPerPaycheck = inferMileageBasisFromName(event.target.value);
    }
    renderPlanning();
    persist();
  }

  function handleCarCostTableClick(event) {
    const button = event.target.closest("[data-delete-carcost]");
    if (!button) return;
    state.tables.carCostCalculator = state.tables.carCostCalculator.filter((row) => row.id !== button.dataset.deleteCarcost);
    renderPlanning();
    persist();
  }

  function handleSetupBoardSubmit(event) {
    const form = event.target.closest("[data-setup-add]");
    if (!form) return;
    event.preventDefault();
    const name = normalizeHierarchyName(form.elements.name.value);
    if (!name) return;

    if (form.dataset.setupAdd === "category") {
      addCategoryNode(form.dataset.budgetName, name);
    }

    if (form.dataset.setupAdd === "subcategory") {
      addSubcategoryNode(form.dataset.budgetName, form.dataset.categoryName, name);
    }

    form.reset();
    render();
  }

  function handleSetupDragStart(event) {
    const node = event.target.closest("[data-drag-type]");
    if (!node) return;
    const payload =
      node.dataset.dragType === "category"
        ? {
            type: "category",
            budgetName: node.dataset.budgetName,
            categoryName: node.dataset.categoryName,
          }
        : {
            type: "subcategory",
            budgetName: node.dataset.budgetName,
            categoryName: node.dataset.categoryName,
            subcategoryName: node.dataset.subcategoryName,
          };
    state.dragPayload = payload;
    event.dataTransfer.effectAllowed = "move";
    event.dataTransfer.setData("text/plain", JSON.stringify(payload));
  }

  function handleSetupDragEnd() {
    clearSetupDropState();
    delete state.dragPayload;
  }

  function handleSetupDragOver(event) {
    const payload = state.dragPayload;
    if (!payload) return;
    const target = getSetupDropTarget(event.target, payload.type);
    if (!target) return;
    event.preventDefault();
    clearSetupDropState();
    target.classList.add("is-drop-target");
  }

  function handleSetupDragLeave(event) {
    const target = event.target.closest(".is-drop-target");
    if (target && !target.contains(event.relatedTarget)) {
      target.classList.remove("is-drop-target");
    }
  }

  function handleSetupDrop(event) {
    const payload = state.dragPayload;
    if (!payload) return;
    const target = getSetupDropTarget(event.target, payload.type);
    if (!target) return;
    event.preventDefault();
    clearSetupDropState();

    if (payload.type === "category") {
      moveCategoryToBudget(payload.budgetName, payload.categoryName, target.dataset.budgetName);
    }

    if (payload.type === "subcategory") {
      moveSubcategoryToCategory(
        payload.budgetName,
        payload.categoryName,
        payload.subcategoryName,
        target.dataset.budgetName,
        target.dataset.categoryName,
      );
    }

    delete state.dragPayload;
    render();
  }

  function handleMoneyInputFocus(event) {
    const input = event.target.closest("[data-money-input]");
    if (!input) return;
    window.requestAnimationFrame(() => input.select());
  }

  function handleMoneyInputBlur(event) {
    const input = event.target.closest("[data-money-input]");
    if (!input) return;
    formatMoneyInput(input);
  }

  function handleGlobalClick(event) {
    const startInlineEditButton = event.target.closest("[data-start-inline-edit]");
    if (startInlineEditButton) {
      startInlineEdit(startInlineEditButton.dataset.startInlineEdit, startInlineEditButton.dataset.editId);
      return;
    }
    const saveInlineEditButton = event.target.closest("[data-save-inline-edit]");
    if (saveInlineEditButton) {
      saveInlineEdit(saveInlineEditButton.dataset.saveInlineEdit, saveInlineEditButton.dataset.editId);
      return;
    }
    const cancelInlineEditButton = event.target.closest("[data-cancel-inline-edit]");
    if (cancelInlineEditButton) {
      cancelInlineEdit(cancelInlineEditButton.dataset.cancelInlineEdit);
      return;
    }
    const deleteButtonNode = event.target.closest("[data-delete-table]");
    if (deleteButtonNode) {
      const tableName = deleteButtonNode.dataset.deleteTable;
      const id = deleteButtonNode.dataset.deleteId;
      state.tables[tableName] = state.tables[tableName].filter((row) => row.id !== id);
      state.pendingReconcileIds = state.pendingReconcileIds.filter((item) => item !== id);
      if (state.editing[tableName] === id) cancelInlineEdit(tableName);
      render();
      return;
    }
    const selectTxn = event.target.closest("[data-select-txn]");
    if (selectTxn) {
      state.selectedTransactionId = selectTxn.dataset.selectTxn;
      state.filters.transactionIdSearch = state.selectedTransactionId;
      state.currentView = "transaction-lookup";
      render();
      return;
    }
    const selectBudgetMain = event.target.closest("[data-budget-progress-main]");
    if (selectBudgetMain) {
      state.budgetProgressMainGroup = selectBudgetMain.dataset.budgetProgressMain;
      state.budgetProgressCategory = "";
      renderPlanning();
      persist();
      return;
    }
    const selectBudgetCategory = event.target.closest("[data-budget-progress-category]");
    if (selectBudgetCategory) {
      state.budgetProgressCategory = selectBudgetCategory.dataset.budgetProgressCategory;
      renderPlanning();
      persist();
      return;
    }
    const resetBudgetProgress = event.target.closest("[data-budget-progress-reset]");
    if (resetBudgetProgress) {
      if (resetBudgetProgress.dataset.budgetProgressReset === "all") {
        state.budgetProgressMainGroup = "";
        state.budgetProgressCategory = "";
      } else if (resetBudgetProgress.dataset.budgetProgressReset === "main") {
        state.budgetProgressCategory = "";
      }
      renderPlanning();
      persist();
    }
  }

  function renderTwinLines() {
    byId("twin-lines").innerHTML = state.twinLines.map((line, index) => `
      <div class="twin-line" data-index="${index}">
        <label>Budget${renderSelectMarkup("budget", getDynamicLookups().budgets, line.budget, "Choose budget")}</label>
        <label>Category${renderSelectMarkup("category", getCategoryOptionsForSelection(line.budget), line.category, "Choose category")}</label>
        <label>Subcategory${renderSelectMarkup("subcategory", getSubcategoryOptionsForSelection(line.budget, line.category), line.subcategory, NO_SUBCATEGORY_LABEL)}</label>
        <label>Amount<input name="amount" type="text" inputmode="decimal" data-money-input value="${escapeAttr(formatMoneyInputValue(line.amount))}" /></label>
        <label>Notes<input name="personalNotes" list="budget-notes-list" value="${escapeAttr(line.personalNotes)}" /></label>
        <button class="button button-danger" type="button" data-remove-twin-line="${index}">Remove</button>
      </div>`).join("");
    updateTwinTotal();
  }

  function renderEditableBiweeklyExpenses(rows) {
    byId("planning-table").innerHTML = `
      <table>
        <thead><tr><th>Budget</th><th>Category</th><th>Subcategory</th><th>Amount / Year</th><th>Amount / Paycheck</th><th>Share</th><th></th></tr></thead>
        <tbody>
          ${rows.map((row) => `
            <tr data-biweekly-id="${row.id}">
              <td>${renderSelectMarkup("budget", getDynamicLookups().budgets, row.budget, "Choose budget")}</td>
              <td>${renderSelectMarkup("category", getCategoryOptionsForSelection(row.budget), row.category, "Choose category")}</td>
              <td>${renderSelectMarkup("subcategory", getSubcategoryOptionsForSelection(row.budget, row.category), row.subcategory, NO_SUBCATEGORY_LABEL)}</td>
              <td><input name="amountPerYear" type="text" inputmode="decimal" data-money-input value="${escapeAttr(formatMoneyInputValue(row.amountPerYear))}" /></td>
              <td>${moneyCell(row.amount)}</td>
              <td>${percentCell(row.shareOfYearlyBudget)}</td>
              <td><button class="button button-secondary" type="button" data-delete-biweekly="${row.id}">Delete</button></td>
            </tr>`).join("")}
        </tbody>
      </table>
      <div class="table-note">Per-paycheck values update immediately from the selected frequency.</div>`;
  }

  function renderEditableCarCosts(rows) {
    byId("car-cost-table").innerHTML = `
      <table>
        <thead><tr><th>Name</th><th>Amount</th><th>Miles / Cycle</th><th>Per Mile</th><th>Per Paycheck</th><th></th></tr></thead>
        <tbody>
          ${rows.map((row) => `
            <tr data-carcost-id="${row.id}">
              <td><input name="name" value="${escapeAttr(row.name || "")}" /></td>
              <td><input name="amount" type="text" inputmode="decimal" data-money-input value="${escapeAttr(formatMoneyInputValue(row.amount))}" /></td>
              <td><input name="milesPerCycle" type="number" step="0.01" value="${escapeAttr(row.milesPerCycle)}" /></td>
              <td>${moneyCell(row.perMile)}</td>
              <td>${moneyCell(row.perPaycheck)}</td>
              <td><button class="button button-secondary" type="button" data-delete-carcost="${row.id}">Delete</button></td>
            </tr>`).join("")}
        </tbody>
      </table>
      <div class="table-note">Per-paycheck values use the row’s inferred mileage basis from your current data.</div>`;
  }

  function renderBudgetProgress() {
    const { mainRows, categoryRows, subcategoryRows } = getBudgetProgressRows();
    let level = "main";
    let rows = mainRows;

    if (state.budgetProgressCategory) {
      level = "subcategory";
      rows = subcategoryRows;
    } else if (state.budgetProgressMainGroup) {
      level = "category";
      rows = categoryRows;
    }

    const crumbs = [
      `<button class="table-link" type="button" data-budget-progress-reset="all">Main Budgets</button>`,
    ];
    if (state.budgetProgressMainGroup) {
      crumbs.push(
        `<button class="table-link" type="button" data-budget-progress-reset="main">${escapeHtml(state.budgetProgressMainGroup)}</button>`,
      );
    }
    if (state.budgetProgressCategory) {
      crumbs.push(`<span class="mono">${escapeHtml(getBudgetProgressCategoryLabel(state.budgetProgressCategory))}</span>`);
    }

    byId("budget-progress-table").innerHTML = rows.length ? `
      <div class="table-stack">
        <div class="progress-breadcrumbs">${crumbs.join(`<span>/</span>`)}</div>
        ${rows.map((row) => `
          <div class="progress-card">
            <div class="progress-head">
              <strong>
                ${
                  level === "main"
                    ? `<button class="table-link progress-link" type="button" data-budget-progress-main="${escapeAttr(row.key)}">${escapeHtml(row.label)}</button>`
                    : level === "category"
                      ? `<button class="table-link progress-link" type="button" data-budget-progress-category="${escapeAttr(row.key)}">${escapeHtml(row.label)}</button>`
                      : escapeHtml(row.label)
                }
              </strong>
              <span>Current total ${formatMoney(row.currentTotal)}</span>
            </div>
            <div class="progress-meta">
              <span>Added ${formatMoney(row.addedThisPayperiod)}</span>
              <span>Spent ${formatMoney(row.spentThisPayperiod)} of ${formatMoney(row.spentBasis)}</span>
              ${row.overspentAmount > 0 ? `<span class="negative">Overspent ${formatMoney(row.overspentAmount)}</span>` : ``}
            </div>
            <div class="bar-track">
              <div class="progress-stack">
                <div class="bar-fill spent" style="width:${row.spentRatio * 100}%"></div>
                <div class="bar-fill overspend" style="width:${row.overspentRatio * 100}%"></div>
              </div>
            </div>
          </div>`).join("")}
      </div>` : `<div class="empty-state">No budget activity matched the current payperiod.</div>`;
  }

  function resetState() {
    if (!window.confirm("Reset the site back to the starting local data?")) return;
    localStorage.removeItem(STORAGE_KEY);
    window.location.reload();
  }

  function exportState() {
    const blob = new Blob([JSON.stringify({ exportedAt: new Date().toISOString(), settings: state.settings, budgetHierarchy: state.budgetHierarchy, tables: state.tables }, null, 2)], { type: "application/json" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "jere-anna-budget-web-export.json";
    link.click();
    URL.revokeObjectURL(link.href);
  }

  function normalizeBudgetEntryRow(row, index) {
    const legacyBudget = typeof row.budget !== "string";
    const budget = normalizeHierarchyName(legacyBudget ? row.category : row.budget);
    const category = normalizeHierarchyName(legacyBudget ? row.subcategory : row.category) || DEFAULT_CATEGORY_NAME;
    const subcategory = normalizeHierarchyName(legacyBudget ? "" : row.subcategory);
    return {
      ...row,
      transactionId: row.transactionId || `BUDGET-${String(row.address || row.sourceRow || index + 1).padStart(6, "0")}`,
      budget,
      category,
      subcategory,
    };
  }

  function normalizePlanningRow(row, index) {
    const legacyBudget = typeof row.budget !== "string";
    return {
      ...row,
      id: row.id || `biweekly-${index + 1}`,
      budget: normalizeHierarchyName(legacyBudget ? row.category : row.budget),
      category: normalizeHierarchyName(legacyBudget ? row.subcategory : row.category) || DEFAULT_CATEGORY_NAME,
      subcategory: normalizeHierarchyName(legacyBudget ? "" : row.subcategory),
    };
  }

  function buildInitialBudgetHierarchy(budgetEntries, planningRows) {
    const rows = [
      ...budgetEntries.map((row, index) => normalizeBudgetEntryRow(row, index)),
      ...planningRows.map((row, index) => normalizePlanningRow(row, index)),
    ];
    return buildBudgetHierarchyFromRows(rows);
  }

  function buildBudgetHierarchyFromRows(rows) {
    const hierarchy = [];
    for (const row of rows) {
      if (!row.budget || !row.category) continue;
      addBudgetNode(row.budget, hierarchy);
      addCategoryNode(row.budget, row.category, hierarchy);
      if (row.subcategory) addSubcategoryNode(row.budget, row.category, row.subcategory, hierarchy);
    }
    return normalizeBudgetHierarchy(hierarchy);
  }

  function normalizeBudgetHierarchy(hierarchy) {
    const budgets = [];
    for (const budget of hierarchy || []) {
      const budgetName = normalizeHierarchyName(budget.name);
      if (!budgetName || budgets.some((item) => item.name === budgetName)) continue;
      const categories = [];
      for (const category of budget.categories || []) {
        const categoryName = normalizeHierarchyName(category.name) || DEFAULT_CATEGORY_NAME;
        if (categories.some((item) => item.name === categoryName)) continue;
        const subcategories = [];
        for (const subcategory of category.subcategories || []) {
          const subcategoryName = normalizeHierarchyName(subcategory);
          if (!subcategoryName || subcategories.includes(subcategoryName)) continue;
          subcategories.push(subcategoryName);
        }
        categories.push({ name: categoryName, subcategories });
      }
      budgets.push({ name: budgetName, categories });
    }
    return budgets;
  }

  function ensureBudgetHierarchyCoverage() {
    const rows = [...state.tables.budgetEntries, ...state.tables.biWeeklyExpenses];
    for (const row of rows) {
      if (!row.budget || !row.category) continue;
      addBudgetNode(row.budget);
      addCategoryNode(row.budget, row.category);
      if (row.subcategory) addSubcategoryNode(row.budget, row.category, row.subcategory);
    }
  }

  function normalizeHierarchySelectionState() {
    const lookups = getDynamicLookups();
    if (!["All budgets", ...lookups.budgets].includes(state.filters.budgetBudget)) {
      state.filters.budgetBudget = "All budgets";
    }
    if (!["All categories", ...getCategoryOptionsForSelection(state.filters.budgetBudget)].includes(state.filters.budgetCategory)) {
      state.filters.budgetCategory = "All categories";
    }
    if (
      !["All subcategories", ...getSubcategoryOptionsForSelection(state.filters.budgetBudget, state.filters.budgetCategory)].includes(
        state.filters.budgetSubcategory,
      )
    ) {
      state.filters.budgetSubcategory = "All subcategories";
    }
    if (state.budgetProgressMainGroup && !lookups.budgets.includes(state.budgetProgressMainGroup)) {
      state.budgetProgressMainGroup = "";
    }
    if (state.budgetProgressCategory && !findCategoryByProgressKey(state.budgetProgressCategory)) {
      state.budgetProgressCategory = "";
    }
  }

  function renderBudgetEntrySelectors() {
    const form = byId("budget-form");
    if (!form) return;
    const budgets = getDynamicLookups().budgets;
    const currentBudget = form.elements.budget?.value || budgets[0] || "";
    fillSelect("budget-form-budget", budgets, currentBudget);
    const finalBudget = form.elements.budget.value || budgets[0] || "";
    fillSelect(
      "budget-form-category",
      getCategoryOptionsForSelection(finalBudget).map((value) => ({ label: value, value })),
      form.elements.category?.value || "",
    );
    const finalCategory = form.elements.category.value || getCategoryOptionsForSelection(finalBudget)[0] || "";
    fillSelect(
      "budget-form-subcategory",
      [{ label: NO_SUBCATEGORY_LABEL, value: "" }, ...getSubcategoryOptionsForSelection(finalBudget, finalCategory).map((value) => ({ label: value, value }))],
      form.elements.subcategory?.value || "",
    );
  }

  function syncBudgetEntryFormSelection() {
    renderBudgetEntrySelectors();
    const form = byId("budget-form");
    if (!form) return;
    if (!form.elements.budget.value) form.elements.budget.value = getDynamicLookups().budgets[0] || "";
    renderBudgetEntrySelectors();
  }

  function getCategoryOptionsForSelection(budgetName) {
    if (!budgetName || budgetName === "All budgets") {
      return unique(state.budgetHierarchy.flatMap((budget) => budget.categories.map((category) => category.name)));
    }
    return getBudgetNode(budgetName)?.categories.map((category) => category.name) || [];
  }

  function getSubcategoryOptionsForSelection(budgetName, categoryName) {
    if (!categoryName || categoryName === "All categories") {
      if (!budgetName || budgetName === "All budgets") {
        return unique(
          state.budgetHierarchy.flatMap((budget) => budget.categories.flatMap((category) => category.subcategories)),
        );
      }
      return unique((getBudgetNode(budgetName)?.categories || []).flatMap((category) => category.subcategories));
    }
    if (!budgetName || budgetName === "All budgets") {
      return unique(
        state.budgetHierarchy.flatMap((budget) =>
          budget.categories
            .filter((category) => category.name === categoryName)
            .flatMap((category) => category.subcategories),
        ),
      );
    }
    return getCategoryNode(budgetName, categoryName)?.subcategories || [];
  }

  function addBudgetNode(name, hierarchy = state.budgetHierarchy) {
    const budgetName = normalizeHierarchyName(name);
    if (!budgetName) return null;
    let node = hierarchy.find((budget) => budget.name === budgetName);
    if (!node) {
      node = { name: budgetName, categories: [] };
      hierarchy.push(node);
    }
    return node;
  }

  function addCategoryNode(budgetName, categoryName, hierarchy = state.budgetHierarchy) {
    const budget = addBudgetNode(budgetName, hierarchy);
    const normalizedCategory = normalizeHierarchyName(categoryName) || DEFAULT_CATEGORY_NAME;
    if (!budget) return null;
    let node = budget.categories.find((category) => category.name === normalizedCategory);
    if (!node) {
      node = { name: normalizedCategory, subcategories: [] };
      budget.categories.push(node);
    }
    return node;
  }

  function addSubcategoryNode(budgetName, categoryName, subcategoryName, hierarchy = state.budgetHierarchy) {
    const category = addCategoryNode(budgetName, categoryName, hierarchy);
    const normalizedSubcategory = normalizeHierarchyName(subcategoryName);
    if (!category || !normalizedSubcategory) return null;
    if (!category.subcategories.includes(normalizedSubcategory)) {
      category.subcategories.push(normalizedSubcategory);
    }
    return normalizedSubcategory;
  }

  function getBudgetNode(name) {
    return state.budgetHierarchy.find((budget) => budget.name === name) || null;
  }

  function getCategoryNode(budgetName, categoryName) {
    return getBudgetNode(budgetName)?.categories.find((category) => category.name === categoryName) || null;
  }

  function findCategoryByProgressKey(key) {
    const [budgetName, categoryName] = String(key || "").split("::");
    if (!budgetName || !categoryName) return null;
    return { budgetName, categoryName };
  }

  function getBudgetProgressCategoryLabel(key) {
    return findCategoryByProgressKey(key)?.categoryName || key;
  }

  function moveCategoryToBudget(fromBudgetName, categoryName, toBudgetName) {
    if (!fromBudgetName || !categoryName || !toBudgetName || fromBudgetName === toBudgetName) return;
    const fromBudget = getBudgetNode(fromBudgetName);
    const toBudget = addBudgetNode(toBudgetName);
    if (!fromBudget || !toBudget) return;
    const index = fromBudget.categories.findIndex((category) => category.name === categoryName);
    if (index === -1) return;
    const [movingCategory] = fromBudget.categories.splice(index, 1);
    const existing = toBudget.categories.find((category) => category.name === categoryName);
    if (existing) {
      for (const subcategory of movingCategory.subcategories) {
        if (!existing.subcategories.includes(subcategory)) existing.subcategories.push(subcategory);
      }
    } else {
      toBudget.categories.push(movingCategory);
    }
    reassignRowsForCategoryMove(fromBudgetName, categoryName, toBudgetName);
  }

  function moveSubcategoryToCategory(fromBudgetName, fromCategoryName, subcategoryName, toBudgetName, toCategoryName) {
    if (!fromBudgetName || !fromCategoryName || !subcategoryName || !toBudgetName || !toCategoryName) return;
    if (fromBudgetName === toBudgetName && fromCategoryName === toCategoryName) return;
    const fromCategory = getCategoryNode(fromBudgetName, fromCategoryName);
    const toCategory = addCategoryNode(toBudgetName, toCategoryName);
    if (!fromCategory || !toCategory) return;
    fromCategory.subcategories = fromCategory.subcategories.filter((item) => item !== subcategoryName);
    if (!toCategory.subcategories.includes(subcategoryName)) {
      toCategory.subcategories.push(subcategoryName);
    }
    reassignRowsForSubcategoryMove(fromBudgetName, fromCategoryName, subcategoryName, toBudgetName, toCategoryName);
  }

  function reassignRowsForCategoryMove(fromBudgetName, categoryName, toBudgetName) {
    state.tables.budgetEntries = state.tables.budgetEntries.map((row) =>
      row.budget === fromBudgetName && row.category === categoryName ? { ...row, budget: toBudgetName } : row,
    );
    state.tables.biWeeklyExpenses = state.tables.biWeeklyExpenses.map((row) =>
      row.budget === fromBudgetName && row.category === categoryName ? { ...row, budget: toBudgetName } : row,
    );
    if (state.budgetProgressMainGroup === fromBudgetName) state.budgetProgressMainGroup = toBudgetName;
    if (state.budgetProgressCategory === `${fromBudgetName}::${categoryName}`) {
      state.budgetProgressCategory = `${toBudgetName}::${categoryName}`;
    }
  }

  function reassignRowsForSubcategoryMove(fromBudgetName, fromCategoryName, subcategoryName, toBudgetName, toCategoryName) {
    state.tables.budgetEntries = state.tables.budgetEntries.map((row) =>
      row.budget === fromBudgetName && row.category === fromCategoryName && row.subcategory === subcategoryName
        ? { ...row, budget: toBudgetName, category: toCategoryName }
        : row,
    );
    state.tables.biWeeklyExpenses = state.tables.biWeeklyExpenses.map((row) =>
      row.budget === fromBudgetName && row.category === fromCategoryName && row.subcategory === subcategoryName
        ? { ...row, budget: toBudgetName, category: toCategoryName }
        : row,
    );
    if (state.budgetProgressCategory === `${fromBudgetName}::${fromCategoryName}`) {
      state.budgetProgressCategory = `${toBudgetName}::${toCategoryName}`;
    }
  }

  function startInlineEdit(tableName, id) {
    const row = state.tables[tableName]?.find((item) => item.id === id);
    if (!row) return;
    state.editing[tableName] = id;
    state.editDrafts[tableName] = clone(row);
    render();
  }

  function cancelInlineEdit(tableName) {
    state.editing[tableName] = "";
    state.editDrafts[tableName] = null;
    render();
  }

  function saveInlineEdit(tableName, id) {
    if (state.editing[tableName] !== id || !state.editDrafts[tableName]) return;
    const draft = clone(state.editDrafts[tableName]);
    if (tableName === "register") {
      draft.accountName = String(draft.accountName || "").trim();
      draft.checkNumber = String(draft.checkNumber || "").trim();
      draft.amount = safeMoney(draft.amount);
    }
    if (tableName === "budgetEntries") {
      draft.budget = normalizeHierarchyName(draft.budget);
      draft.category = normalizeHierarchyName(draft.category) || DEFAULT_CATEGORY_NAME;
      draft.subcategory = normalizeHierarchyName(draft.subcategory);
      draft.personalNotes = String(draft.personalNotes || "").trim();
      draft.amount = safeMoney(draft.amount);
      addBudgetNode(draft.budget);
      addCategoryNode(draft.budget, draft.category);
      if (draft.subcategory) addSubcategoryNode(draft.budget, draft.category, draft.subcategory);
    }
    state.tables[tableName] = state.tables[tableName].map((row) => (row.id === id ? { ...row, ...draft } : row));
    state.editing[tableName] = "";
    state.editDrafts[tableName] = null;
    render();
  }

  function isEditingRow(tableName, id) {
    return state.editing?.[tableName] === id;
  }

  function getRowDraft(tableName, row) {
    return isEditingRow(tableName, row.id) && state.editDrafts?.[tableName] ? state.editDrafts[tableName] : row;
  }

  function getSetupDropTarget(target, dragType) {
    if (dragType === "category") return target.closest('[data-drop-target="budget"]');
    return target.closest('[data-drop-target="category"]');
  }

  function clearSetupDropState() {
    document.querySelectorAll(".is-drop-target").forEach((node) => node.classList.remove("is-drop-target"));
  }

  function renderSelectMarkup(name, values, selectedValue, emptyLabel, attrs = "") {
    const options = [
      `<option value="" ${!selectedValue ? "selected" : ""}>${escapeHtml(emptyLabel || "Choose one")}</option>`,
      ...values.map((value) => `<option value="${escapeAttr(value)}" ${value === selectedValue ? "selected" : ""}>${escapeHtml(value)}</option>`),
    ];
    return `<select name="${escapeAttr(name)}" ${attrs}>${options.join("")}</select>`;
  }

  function normalizeHierarchyName(value) {
    return String(value || "").trim();
  }

  function renderStatCard(card) {
    return `<article class="stat-card"><small>${escapeHtml(card.label)}</small><strong>${escapeHtml(card.value)}</strong><span>${escapeHtml(card.detail)}</span></article>`;
  }
  function renderBarSummary(targetId, rows, labelTitle, valueTitle) {
    const max = Math.max(...rows.map((row) => Math.abs(row.value)), 1);
    byId(targetId).innerHTML = `<div class="table-stack"><div class="table-caption">${escapeHtml(labelTitle)} and ${escapeHtml(valueTitle)}</div>${rows.map((row) => `<div class="bar-row"><div>${escapeHtml(row.label)}</div><div class="${classForMoney(row.value)}">${escapeHtml(formatMoney(row.value))}</div><div class="bar-track"><div class="bar-fill" style="width:${(Math.abs(row.value) / max) * 100}%"></div></div></div>`).join("")}</div>`;
  }
  function renderOverviewTable(targetId, overview) {
    byId(targetId).innerHTML = `<table><thead><tr>${overview.columns.map((column) => `<th>${escapeHtml(column)}</th>`).join("")}</tr></thead><tbody>${overview.rows.map((row) => `<tr>${row.map((value, index) => index === 0 ? `<td>${dateCell(value)}</td>` : `<td>${moneyCell(value)}</td>`).join("")}</tr>`).join("")}</tbody></table>`;
  }
  function renderTable(targetId, columns, rows, note) {
    byId(targetId).innerHTML = rows.length ? `<table><thead><tr>${columns.map((column) => `<th>${escapeHtml(column.label)}</th>`).join("")}</tr></thead><tbody>${rows.map((row) => `<tr data-row-id="${escapeAttr(row.id || "")}">${columns.map((column) => `<td>${column.render(row)}</td>`).join("")}</tr>`).join("")}</tbody></table>${note ? `<div class="table-note">${escapeHtml(note)}</div>` : ""}` : `<div class="empty-state">Nothing to show here yet.</div>${note ? `<div class="table-note">${escapeHtml(note)}</div>` : ""}`;
  }
  function fillSelect(id, values, selectedValue) {
    byId(id).innerHTML = values
      .map((value) => {
        const option = typeof value === "string" ? { label: value, value } : value;
        return `<option value="${escapeAttr(option.value)}" ${option.value === selectedValue ? "selected" : ""}>${escapeHtml(option.label)}</option>`;
      })
      .join("");
  }
  function renderDatalist(id, values) { byId(id).innerHTML = values.map((value) => `<option value="${escapeAttr(value)}"></option>`).join(""); }
  function setInputValue(id, value) { const element = byId(id); if (element) element.value = value ?? ""; }
  function deleteButton(tableName, id) { return `<button class="button button-secondary" type="button" data-delete-table="${escapeAttr(tableName)}" data-delete-id="${escapeAttr(id)}">Delete</button>`; }
  function editActionButtons(tableName, id) {
    if (isEditingRow(tableName, id)) {
      return `<div class="action-group"><button class="button button-primary" type="button" data-save-inline-edit="${escapeAttr(tableName)}" data-edit-id="${escapeAttr(id)}">Save</button><button class="button button-secondary" type="button" data-cancel-inline-edit="${escapeAttr(tableName)}">Cancel</button></div>`;
    }
    return `<div class="action-group"><button class="button button-secondary" type="button" data-start-inline-edit="${escapeAttr(tableName)}" data-edit-id="${escapeAttr(id)}">Edit</button>${deleteButton(tableName, id)}</div>`;
  }
  function transactionLinkCell(value) { return `<button class="table-link mono" type="button" data-select-txn="${escapeAttr(value)}">${escapeHtml(value)}</button>`; }
  function moneyCell(value) { return `<span class="money ${classForMoney(value)}">${escapeHtml(formatMoney(value))}</span>`; }
  function percentCell(value) { return `<span class="number">${escapeHtml(formatPercent(value))}</span>`; }
  function numberCell(value) { return `<span class="number">${escapeHtml(formatNumber(value))}</span>`; }
  function dateCell(value) { return `<span class="date-cell">${escapeHtml(formatDate(value))}</span>`; }
  function mono(value) { return `<span class="mono">${escapeHtml(String(value ?? ""))}</span>`; }
  function updateTwinTotal() { setText("twin-total", formatMoney(state.twinLines.reduce((total, line) => total + safeNumber(line.amount), 0))); }
  function formatMoney(value) { return new Intl.NumberFormat("en-US", { style: "currency", currency: "USD", minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(safeNumber(value)); }
  function formatMoneyInputValue(value) { return value === "" || value === null || value === undefined ? "" : formatMoney(safeMoney(value)); }
  function formatPercent(value) { return `${(safeNumber(value) * 100).toFixed(2)}%`; }
  function formatNumber(value) { const amount = safeNumber(value); return Number.isInteger(amount) ? amount.toLocaleString("en-US") : amount.toLocaleString("en-US", { maximumFractionDigits: 2 }); }
  function formatDate(value) { if (!value) return "-"; return new Intl.DateTimeFormat("en-US", { month: "short", day: "numeric", year: "numeric" }).format(new Date(`${value}T00:00:00`)); }
  function formatDateTime(value) { if (!value) return "-"; return new Intl.DateTimeFormat("en-US", { month: "short", day: "numeric", year: "numeric", hour: "numeric", minute: "2-digit" }).format(new Date(value)); }
  function classForMoney(value) { return safeNumber(value) >= 0 ? "positive" : "negative"; }
  function sum(rows, key) { return round2(rows.reduce((total, row) => total + safeNumber(row[key]), 0)); }
  function parseNumeric(value) {
    if (typeof value === "number") return Number.isFinite(value) ? value : 0;
    const text = String(value ?? "").trim();
    if (!text) return 0;
    const cleaned = text.replace(/[$,\s]/g, "").replace(/[^\d.-]/g, "");
    const amount = Number(cleaned);
    return Number.isFinite(amount) ? amount : 0;
  }
  function safeNumber(value) { return parseNumeric(value); }
  function safeMoney(value) { return round2(parseNumeric(value)); }
  function round2(value) { return Math.round((parseNumeric(value) + Number.EPSILON) * 100) / 100; }
  function clamp01(value) { return Math.max(0, Math.min(1, safeNumber(value))); }
  function addDays(value, days) { const base = new Date(`${value}T00:00:00`); base.setDate(base.getDate() + safeNumber(days)); return base.toISOString().slice(0, 10); }
  function compareDate(left, right) { return new Date(`${left || "1900-01-01"}T00:00:00`) - new Date(`${right || "1900-01-01"}T00:00:00`); }
  function compareText(left, right) { return String(left || "").localeCompare(String(right || "")); }
  function parseCyclesPerYear(label) { const match = String(label || "").match(/\((\d+)\/yr\)/i); return match ? Number(match[1]) : 24; }
  function makeId(prefix) { return `${prefix}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`; }
  function generateTransactionId() { return `TXN-${Date.now().toString(36).toUpperCase()}-${Math.random().toString(36).slice(2, 6).toUpperCase()}`; }
  function nextSourceRow(rows) { return Math.max(0, ...rows.map((row) => safeNumber(row.sourceRow))) + 1; }
  function unique(values) { return [...new Set(values.filter(Boolean))].sort((left, right) => String(left).localeCompare(String(right))); }
  function inferMileageBasisFromName(name) {
    const text = String(name || "").toLowerCase();
    if (text.includes("truck") || text.includes("chevy")) return 400;
    if (text.includes("car") || text.includes("buick") || text.includes("park avenue")) return 1500;
    return 1900;
  }
  function byId(id) { return document.getElementById(id); }
  function setText(id, value) { byId(id).textContent = value; }
  function formatMoneyInput(input) {
    if (!input) return;
    const raw = String(input.value ?? "").trim();
    input.value = raw ? formatMoney(safeMoney(raw)) : "";
  }
  function formatMoneyInputs(root) {
    (root || document).querySelectorAll("[data-money-input]").forEach((input) => formatMoneyInput(input));
  }
  function clone(value) { return JSON.parse(JSON.stringify(value)); }
  function escapeHtml(value) { return String(value ?? "").replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll('"', "&quot;").replaceAll("'", "&#39;"); }
  function escapeAttr(value) { return escapeHtml(value); }
  function todayIso() { const now = new Date(); const offset = now.getTimezoneOffset(); return new Date(now.getTime() - offset * 60000).toISOString().slice(0, 10); }
})();

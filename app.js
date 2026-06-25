const fmtKrw = new Intl.NumberFormat("ko-KR", {
  style: "currency",
  currency: "KRW",
  maximumFractionDigits: 0,
});

const fmtCny = new Intl.NumberFormat("en-US", {
  style: "currency",
  currency: "CNY",
  currencyDisplay: "narrowSymbol",
  minimumFractionDigits: 2,
  maximumFractionDigits: 2,
});

const state = {
  meta: null,
  invoiceMeta: null,
  items: [],
  shipping: [],
  costs: {
    tax: [
      { label: "관세", amount: 0 },
      { label: "부가세", amount: 0 },
    ],
    logistics: [
      { label: "원산지 발급비용", amount: 0 },
      { label: "통관수수료", amount: 0 },
      { label: "서류발급비용", amount: 0 },
      { label: "해상운임비", amount: 0 },
    ],
    other: [],
  },
  rates: {
    cny: 220.44,
    usd: 1523.64,
    commission: 0,
  },
};

const $ = (selector) => document.querySelector(selector);
const $$ = (selector) => [...document.querySelectorAll(selector)];

function n(value) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

function uid() {
  return `item-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function itemBaseCny(item) {
  return n(item.quantity) * n(item.unitCny) + n(item.inlandCny) + n(item.extraCny);
}

function costTotal() {
  return Object.values(state.costs)
    .flat()
    .reduce((sum, cost) => sum + n(cost.amount), 0);
}

function calc() {
  const baseCny = state.items.reduce((sum, item) => sum + itemBaseCny(item), 0);
  const chinaKrw = baseCny * state.rates.cny;
  const commissionKrw = chinaKrw * (state.rates.commission / 100);
  const koreaKrw = costTotal();
  const grand = chinaKrw + commissionKrw + koreaKrw;

  return {
    baseCny,
    chinaKrw,
    commissionKrw,
    koreaKrw,
    grand,
    items: state.items.map((item) => {
      const cny = itemBaseCny(item);
      const ratio = baseCny > 0 ? cny / baseCny : 0;
      const itemChinaKrw = cny * state.rates.cny;
      const itemCommission = itemChinaKrw * (state.rates.commission / 100);
      const allocated = koreaKrw * ratio;
      const total = itemChinaKrw + itemCommission + allocated;
      const unit = n(item.shippedQty) > 0 ? total / n(item.shippedQty) : 0;
      return { ...item, cny, itemChinaKrw, itemCommission, allocated, total, unit };
    }),
  };
}

function renderItems() {
  const list = $("#itemList");
  if (!state.items.length) {
    list.innerHTML = '<div class="empty">엑셀을 업로드하면 품목이 여기에 표시됩니다.</div>';
    $("#itemCount").textContent = "0개";
    return;
  }

  list.innerHTML = state.items
    .map(
      (item, index) => `
        <article class="item-card" data-id="${item.id}">
          <div class="item-main">
            <label class="field">
              <span>품명</span>
              <input data-key="name" value="${escapeHtml(item.name)}" />
            </label>
            <label class="field">
              <span>옵션/사이즈</span>
              <input data-key="option" value="${escapeHtml(item.option || "")}" />
            </label>
            <label class="field">
              <span>MARK</span>
              <input data-key="mark" value="${escapeHtml(item.mark || "")}" />
            </label>
          </div>
          <div class="item-sub">
            <label class="field">
              <span>수량</span>
              <input data-key="quantity" type="number" step="1" value="${n(item.quantity)}" />
            </label>
            <label class="field">
              <span>중국 단가</span>
              <input data-key="unitCny" type="number" step="0.01" value="${n(item.unitCny)}" />
            </label>
            <label class="field">
              <span>내륙 운송비</span>
              <input data-key="inlandCny" type="number" step="0.01" value="${n(item.inlandCny)}" />
            </label>
            <label class="field">
              <span>출고수량</span>
              <input data-key="shippedQty" type="number" step="1" value="${n(item.shippedQty)}" />
            </label>
            <button class="icon-button remove-item" title="품목 삭제">×</button>
          </div>
          <div class="item-total" id="item-total-${index}"></div>
        </article>
      `,
    )
    .join("");

  $("#itemCount").textContent = `${state.items.length}개`;
}

function renderCosts() {
  for (const group of ["tax", "logistics", "other"]) {
    const host = $(`#${group}Costs`);
    const costs = state.costs[group];
    if (!costs.length) {
      host.innerHTML = '<div class="empty">등록된 항목이 없습니다.</div>';
      continue;
    }
    host.innerHTML = costs
      .map(
        (cost, index) => `
          <div class="cost-row" data-group="${group}" data-index="${index}">
            <label class="field">
              <span>항목</span>
              <input data-cost-key="label" value="${escapeHtml(cost.label)}" />
            </label>
            <label class="field">
              <span>금액</span>
              <input data-cost-key="amount" type="number" step="1" value="${n(cost.amount)}" />
            </label>
            <button class="icon-button remove-cost" title="비용 삭제">×</button>
          </div>
        `,
      )
      .join("");
  }
}

function renderTotals() {
  const result = calc();
  $("#totalCnyLabel").textContent = `${fmtCny.format(result.baseCny)} 포함`;
  $("#chinaTotalLabel").textContent = fmtKrw.format(result.chinaKrw + result.commissionKrw);
  $("#koreaTotalLabel").textContent = fmtKrw.format(result.koreaKrw);
  $("#bundleKrw").textContent = fmtKrw.format(result.chinaKrw + result.commissionKrw);
  $("#bundleCny").textContent = fmtCny.format(result.baseCny);
  $("#grandTotal").textContent = fmtKrw.format(result.grand);

  result.items.forEach((item, index) => {
    const el = $(`#item-total-${index}`);
    if (el) {
      const extra = n(item.extraCny) ? ` · 부대비용 ${fmtCny.format(item.extraCny)} 포함` : "";
      el.textContent = `중국 원가 ${fmtCny.format(item.cny)} · 원화 ${fmtKrw.format(item.itemChinaKrw)}${extra}`;
    }
  });

  const finalHost = $("#finalItems");
  if (!result.items.length) {
    finalHost.innerHTML = '<div class="empty">계산할 품목이 없습니다.</div>';
    return;
  }

  finalHost.innerHTML = result.items
    .map(
      (item) => `
        <div class="final-item">
          <div>
            <b>${escapeHtml(item.name || "품목")}</b>
            <span>${escapeHtml(item.option || item.mark || "")} · 배분비용 ${fmtKrw.format(item.allocated)}</span>
          </div>
          <strong>${fmtKrw.format(item.unit)}</strong>
        </div>
      `,
    )
    .join("");
}

function renderAll() {
  $("#cnyRate").value = state.rates.cny;
  $("#usdRate").value = state.rates.usd;
  $("#commissionRate").value = state.rates.commission;
  renderItems();
  renderCosts();
  renderTotals();
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

async function importFile(file) {
  const form = new FormData();
  form.append("file", file);
  $("#fileName").textContent = "엑셀 파일을 읽는 중...";
  const response = await fetch("/api/import", { method: "POST", body: form });
  const payload = await response.json();
  if (!response.ok || payload.error) {
    throw new Error(payload.error || "엑셀 파일을 읽지 못했습니다.");
  }
  state.meta = payload.meta;
  state.items = payload.items.map((item) => ({ ...item, id: item.id || uid() }));
  state.shipping = payload.shipping || [];
  const extraText = payload.meta.additionalCostCount
    ? ` · 부대비용 ${payload.meta.additionalCostCount}개 배분`
    : "";
  $("#fileName").textContent = `${payload.meta.filename} · ${state.items.length}개 품목 인식${extraText}`;
  renderAll();
}

async function importInvoice(file) {
  const form = new FormData();
  form.append("file", file);
  $("#invoiceName").textContent = "인보이스 PDF를 읽는 중...";
  const response = await fetch("/api/import-invoice", { method: "POST", body: form });
  const payload = await response.json();
  if (!response.ok || payload.error) {
    throw new Error(payload.error || "인보이스 PDF를 읽지 못했습니다.");
  }

  state.invoiceMeta = payload.meta;
  state.costs.tax = payload.costs.tax || [];
  state.costs.logistics = payload.costs.logistics || [];
  state.costs.other = payload.costs.other || [];

  const total = payload.totals?.grand || costTotal();
  $("#invoiceName").textContent = `${payload.meta.filename} · ${fmtKrw.format(total)} 인식`;
  renderAll();
}

function applyInputs() {
  state.rates.cny = n($("#cnyRate").value);
  state.rates.usd = n($("#usdRate").value);
  state.rates.commission = n($("#commissionRate").value);
}

function download(filename, text) {
  const blob = new Blob([text], { type: "application/json;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

function savedName() {
  const raw = $("#saveName").value.trim() || "수입원가계산";
  return raw.replace(/[\\/:*?"<>|]/g, "_");
}

function loadDemo() {
  state.meta = { filename: "예시 데이터" };
  state.items = [
    {
      id: uid(),
      name: "가방",
      option: "샘플 옵션",
      mark: "GDH-001~340",
      quantity: 340,
      unitCny: 25,
      inlandCny: 260,
      shippedQty: 340,
    },
  ];
  state.costs.tax = [
    { label: "관세", amount: 0 },
    { label: "부가세", amount: 0 },
  ];
  state.costs.logistics = [
    { label: "원산지 발급비용", amount: 0 },
    { label: "통관수수료", amount: 0 },
    { label: "서류발급비용", amount: 0 },
    { label: "해상운임비", amount: 0 },
  ];
  state.costs.other = [];
  $("#fileName").textContent = "예시 데이터";
  renderAll();
}

$("#fileInput").addEventListener("change", async (event) => {
  const file = event.target.files?.[0];
  if (!file) return;
  try {
    await importFile(file);
  } catch (error) {
    $("#fileName").textContent = "파일 인식 실패";
    alert(error.message);
  }
});

$("#invoiceInput").addEventListener("change", async (event) => {
  const file = event.target.files?.[0];
  if (!file) return;
  try {
    await importInvoice(file);
  } catch (error) {
    $("#invoiceName").textContent = "PDF 인식 실패";
    alert(error.message);
  }
});

$("#dropZone").addEventListener("dragover", (event) => {
  event.preventDefault();
  event.currentTarget.classList.add("dragging");
});

$("#dropZone").addEventListener("dragleave", (event) => {
  event.currentTarget.classList.remove("dragging");
});

$("#dropZone").addEventListener("drop", async (event) => {
  event.preventDefault();
  event.currentTarget.classList.remove("dragging");
  const file = event.dataTransfer.files?.[0];
  if (!file) return;
  try {
    await importFile(file);
  } catch (error) {
    alert(error.message);
  }
});

document.body.addEventListener("input", (event) => {
  const target = event.target;
  if (!(target instanceof HTMLInputElement)) return;

  if (["cnyRate", "usdRate", "commissionRate"].includes(target.id)) {
    applyInputs();
    renderTotals();
    return;
  }

  const itemCard = target.closest(".item-card");
  if (itemCard && target.dataset.key) {
    const item = state.items.find((entry) => entry.id === itemCard.dataset.id);
    if (!item) return;
    item[target.dataset.key] = target.type === "number" ? n(target.value) : target.value;
    renderTotals();
    return;
  }

  const costRow = target.closest(".cost-row");
  if (costRow && target.dataset.costKey) {
    const cost = state.costs[costRow.dataset.group][Number(costRow.dataset.index)];
    cost[target.dataset.costKey] = target.type === "number" ? n(target.value) : target.value;
    renderTotals();
  }
});

document.body.addEventListener("click", (event) => {
  const target = event.target;
  if (!(target instanceof HTMLElement)) return;

  if (target.matches(".remove-item")) {
    const card = target.closest(".item-card");
    state.items = state.items.filter((item) => item.id !== card.dataset.id);
    renderAll();
  }

  if (target.matches(".remove-cost")) {
    const row = target.closest(".cost-row");
    state.costs[row.dataset.group].splice(Number(row.dataset.index), 1);
    renderAll();
  }

  if (target.dataset.addCost) {
    state.costs[target.dataset.addCost].push({ label: "추가 비용", amount: 0 });
    renderAll();
  }
});

$("#addItem").addEventListener("click", () => {
  state.items.push({
    id: uid(),
    name: "새 품목",
    option: "",
    mark: "",
    quantity: 1,
    unitCny: 0,
    inlandCny: 0,
    shippedQty: 1,
  });
  renderAll();
});

$("#clearItems").addEventListener("click", () => {
  state.items = [];
  renderAll();
});

$("#refreshRates").addEventListener("click", () => {
  applyInputs();
  renderTotals();
});

$("#resetAll").addEventListener("click", () => {
  localStorage.removeItem("import-cost-calculator");
  state.meta = null;
  state.invoiceMeta = null;
  state.items = [];
  state.shipping = [];
  state.costs.tax = [
    { label: "관세", amount: 0 },
    { label: "부가세", amount: 0 },
  ];
  state.costs.logistics = [
    { label: "원산지 발급비용", amount: 0 },
    { label: "통관수수료", amount: 0 },
    { label: "서류발급비용", amount: 0 },
    { label: "해상운임비", amount: 0 },
  ];
  state.costs.other = [];
  $("#fileName").textContent = "shipping list 엑셀 파일을 선택하세요";
  $("#invoiceName").textContent = "프렌드 해운항공 INV PDF를 선택하세요";
  renderAll();
});

$("#saveLocal").addEventListener("click", () => {
  applyInputs();
  localStorage.setItem("import-cost-calculator", JSON.stringify(state));
  alert("현재 계산 내용이 이 브라우저에 저장되었습니다.");
});

$("#downloadJson").addEventListener("click", () => {
  applyInputs();
  const result = calc();
  const payload = { savedAt: new Date().toISOString(), state, result };
  download(`${savedName()}.json`, JSON.stringify(payload, null, 2));
});

$("#loadDemo").addEventListener("click", loadDemo);

const saved = localStorage.getItem("import-cost-calculator");
if (saved) {
  try {
    Object.assign(state, JSON.parse(saved));
    $("#fileName").textContent = state.meta?.filename || "저장된 계산 불러옴";
    if (state.invoiceMeta?.filename) {
      $("#invoiceName").textContent = `${state.invoiceMeta.filename} · 저장된 비용`;
    }
  } catch {
    localStorage.removeItem("import-cost-calculator");
  }
}

renderAll();

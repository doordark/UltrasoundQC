/* 超声报告数据纠错系统 — 前端交互逻辑（原生 JS，零外部依赖） */
"use strict";

const API = ""; // 与前端同源；分离部署时改为后端地址，如 http://server:8000

const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => Array.from(document.querySelectorAll(sel));

function esc(s) {
  return String(s ?? "").replace(
    /[&<>"']/g,
    (c) =>
      ({
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        '"': "&quot;",
        "'": "&#39;",
      })[c],
  );
}

/* ================= 提示 ================= */
function toast(msg, type = "") {
  const el = document.createElement("div");
  el.className = `toast ${type}`;
  el.textContent = msg;
  $("#toastWrap").appendChild(el);
  setTimeout(() => el.remove(), 3800);
}

async function fetchJSON(url, options = {}) {
  let resp;
  try {
    resp = await fetch(API + url, options);
  } catch (e) {
    throw new Error("无法连接后端服务，请确认服务已启动");
  }
  let body = null;
  try {
    body = await resp.json();
  } catch (e) {
    /* 非 JSON 响应 */
  }
  if (!resp.ok) {
    const detail = body && (body.detail || body.message);
    throw new Error(detail || `请求失败（HTTP ${resp.status}）`);
  }
  return body;
}

/* ================= 页签 ================= */
$$(".tab-btn").forEach((btn) => {
  btn.addEventListener("click", () => {
    $$(".tab-btn").forEach((b) => b.classList.remove("active"));
    $$(".tab-panel").forEach((p) => p.classList.remove("active"));
    btn.classList.add("active");
    $(`#tab-${btn.dataset.tab}`).classList.add("active");
  });
});

/* ================= 健康状态 ================= */
async function loadHealth() {
  try {
    const h = await fetchJSON("/api/health");
    $("#healthPill").className = "pill pill-ok";
    $("#healthText").textContent = "服务正常";
    $("#engineVersion").textContent =
      `引擎 v${h.engine_version} · ${h.rules_source}`;
  } catch (e) {
    $("#healthPill").className = "pill pill-error";
    $("#healthText").textContent = "服务不可用";
  }
}

/* ================= 步骤条 ================= */
function renderStepper(container, steps) {
  container.innerHTML = steps
    .map((s, i) => {
      const cls = s.state === "pending" ? "" : s.state;
      const mark = s.state === "done" ? "✓" : s.state === "error" ? "!" : i + 1;
      const line =
        i < steps.length - 1
          ? `<span class="st-line ${steps[i + 1].state === "pending" ? "" : "done"}"></span>`
          : "";
      return `<span class="st ${cls}"><span class="st-bullet">${mark}</span>${esc(s.label)}</span>${line}`;
    })
    .join("");
}

function makeSteps(labels, activeIdx = -1, doneAll = false, error = false) {
  return labels.map((label, i) => {
    let state = "pending";
    if (doneAll) state = "done";
    else if (error && i === activeIdx) state = "error";
    else if (i === activeIdx) state = "active";
    else if (i < activeIdx) state = "done";
    return { label, state };
  });
}

/* ================= 上传流程 ================= */
const UPLOAD_STEPS = ["上传文件", "解析数据", "纠错分析", "完成"];
const uploadState = {
  filename: "",
  files: [],
  reports: [],
  warnings: [],
  format: "",
};

function setUploadStepper(activeIdx, doneAll = false, error = false) {
  renderStepper(
    $("#uploadStepper"),
    makeSteps(UPLOAD_STEPS, activeIdx, doneAll, error),
  );
}

const dropzone = $("#dropzone");
const fileInput = $("#fileInput");

dropzone.addEventListener("click", () => fileInput.click());
dropzone.addEventListener("keydown", (e) => {
  if (e.key === "Enter" || e.key === " ") fileInput.click();
});
["dragenter", "dragover"].forEach((ev) =>
  dropzone.addEventListener(ev, (e) => {
    e.preventDefault();
    dropzone.classList.add("dragover");
  }),
);
["dragleave", "drop"].forEach((ev) =>
  dropzone.addEventListener(ev, (e) => {
    e.preventDefault();
    dropzone.classList.remove("dragover");
  }),
);
dropzone.addEventListener("drop", (e) => {
  const files = Array.from((e.dataTransfer && e.dataTransfer.files) || []);
  if (files.length) handleFiles(files);
});
fileInput.addEventListener("change", () => {
  const files = Array.from(fileInput.files || []);
  // 关键：立即清空 input 值，否则同一文件（尤其同名）再次选择时浏览器认为
  // 值未变化而不触发 change 事件，导致修改内容后重新上传仍是旧数据。
  fileInput.value = "";
  if (files.length) handleFiles(files);
});

async function handleFiles(fileList) {
  const files = fileList.slice(0, 20);
  if (fileList.length > 20) {
    toast("单次最多上传 20 个文件，已截取前 20 个", "error");
  }
  const fd = new FormData();
  files.forEach((f) => fd.append("file", f));
  setUploadStepper(1);
  $("#uploadHint").textContent = `正在上传并解析 ${files.length} 个文件…`;
  $("#uploadResult").innerHTML = "";
  try {
    const data = await fetchJSON("/api/upload", { method: "POST", body: fd });
    uploadState.filename = data.filename;
    uploadState.files = data.files || files.map((f) => f.name);
    uploadState.reports = data.reports;
    uploadState.warnings = data.warnings || [];
    uploadState.format = data.format;

    $("#fileInfo").classList.remove("hidden");
    $("#fileInfo").innerHTML = `
      <span class="fi-tag">${esc(String(data.format).toUpperCase())}</span>
      <span class="fi-name">${esc(uploadState.files.join("、"))}</span>
      <span>解析出 <b>${data.report_count}</b> 份报告</span>`;

    const warnBox = $("#uploadWarn");
    if (uploadState.warnings.length) {
      warnBox.classList.remove("hidden");
      warnBox.innerHTML = `<b>解析提示（${uploadState.warnings.length} 条）</b><ul>${uploadState.warnings
        .map((w) => `<li>${esc(w)}</li>`)
        .join("")}</ul>`;
    } else {
      warnBox.classList.add("hidden");
      warnBox.innerHTML = "";
    }

    renderPreview(data.reports);
    setUploadStepper(2);
    $("#runUploadBtn").disabled = false;
    $("#uploadHint").textContent =
      "解析完成，点击「开始纠错」执行 9 类质控规则检查。";
    toast(`解析完成：${data.report_count} 份报告`, "success");
  } catch (e) {
    setUploadStepper(1, false, true);
    $("#uploadHint").textContent = "文件解析失败，请检查文件格式后重新上传。";
    toast(e.message, "error");
  }
}

function renderPreview(reports) {
  $("#previewCard").classList.remove("hidden");
  $("#previewCount").textContent = `共 ${reports.length} 份`;
  const tbody = $("#previewTable tbody");
  tbody.innerHTML = reports
    .map(
      (r) => `
    <tr>
      <td>${esc(r.report_id)}</td>
      <td><span class="cell-clip" title="${esc(r.source_file || "")}">${esc(r.source_file) || "—"}</span></td>
      <td>${esc(r.exam_item) || "—"}</td>
      <td>${esc(r.gender) || "—"}</td>
      <td>${esc(r.age) || "—"}</td>
      <td><span class="cell-clip" title="${esc(r.finding)}">${esc(r.finding) || "—"}</span></td>
      <td><span class="cell-clip" title="${esc(r.conclusion)}">${esc(r.conclusion) || "—"}</span></td>
    </tr>`,
    )
    .join("");
}

$("#runUploadBtn").addEventListener("click", async () => {
  if (!uploadState.reports.length) return;
  $("#runUploadBtn").disabled = true;
  setUploadStepper(2);
  $("#uploadHint").textContent = "正在执行纠错分析…";
  try {
    const data = await fetchJSON("/api/check/batch", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ reports: uploadState.reports }),
    });
    setUploadStepper(3, true);
    $("#uploadHint").textContent =
      `纠错完成：${data.summary.issue_total} 条问题 / ${data.summary.total} 份报告。`;
    renderResult($("#uploadResult"), data, `文件：${uploadState.filename}`);
    toast("纠错完成", "success");
  } catch (e) {
    setUploadStepper(2, false, true);
    $("#uploadHint").textContent = "纠错执行失败，请重试。";
    toast(e.message, "error");
  } finally {
    $("#runUploadBtn").disabled = false;
  }
});

$("#resetUploadBtn").addEventListener("click", () => {
  uploadState.filename = "";
  uploadState.files = [];
  uploadState.reports = [];
  uploadState.warnings = [];
  fileInput.value = "";
  $("#fileInfo").classList.add("hidden");
  $("#uploadWarn").classList.add("hidden");
  $("#previewCard").classList.add("hidden");
  $("#uploadResult").innerHTML = "";
  $("#runUploadBtn").disabled = true;
  $("#uploadHint").textContent = "请先上传数据文件，解析成功后即可执行纠错。";
  setUploadStepper(-1);
});

/* ================= 测试数据集 ================= */
async function loadDatasets() {
  try {
    const data = await fetchJSON("/api/datasets");
    const grid = $("#datasetGrid");
    if (!data.datasets.length) {
      grid.innerHTML = `<p class="empty-note">暂无预置测试数据集。</p>`;
      return;
    }
    grid.innerHTML = data.datasets
      .map(
        (ds) => `
      <div class="ds-card" id="ds-${esc(ds.id)}">
        <div class="ds-head">
          <h3 class="ds-name">${esc(ds.name)}</h3>
          <span class="ds-cat">${esc(ds.category)}</span>
        </div>
        <p class="ds-desc">${esc(ds.description)}</p>
        <div class="ds-meta">
          <span class="chip">${ds.report_count} 份报告</span>
          ${(ds.tags || []).map((t) => `<span class="chip rule">${esc(t)}</span>`).join("")}
        </div>
        <div class="ds-expect">✓ ${esc(ds.expectation)}</div>
        <div class="ds-actions">
          <button class="btn-small primary" data-run="${esc(ds.id)}">一键纠错</button>
          <button class="btn-small" data-view="${esc(ds.id)}">查看数据</button>
        </div>
      </div>`,
      )
      .join("");

    grid
      .querySelectorAll("[data-run]")
      .forEach((btn) =>
        btn.addEventListener("click", () => runDataset(btn.dataset.run)),
      );
    grid
      .querySelectorAll("[data-view]")
      .forEach((btn) =>
        btn.addEventListener("click", () => viewDataset(btn.dataset.view)),
      );
  } catch (e) {
    $("#datasetGrid").innerHTML =
      `<p class="empty-note">数据集加载失败：${esc(e.message)}</p>`;
  }
}

async function viewDataset(dsId) {
  try {
    const ds = await fetchJSON(`/api/datasets/${encodeURIComponent(dsId)}`);
    openModal(
      ds.name,
      `
      <p class="hint" style="margin-top:0">${esc(ds.description)}</p>
      <div class="table-wrap"><table class="data-table">
        <thead><tr><th>报告ID</th><th>检查项目</th><th>性别</th><th>年龄</th>
          <th>超声所见</th><th>超声提示</th><th>预期规则</th></tr></thead>
        <tbody>${ds.reports
          .map(
            (r) => `
          <tr>
            <td>${esc(r.report_id)}</td>
            <td>${esc(r.exam_item)}</td>
            <td>${esc(r.gender)}</td>
            <td>${esc(r.age)}</td>
            <td><span class="cell-clip" title="${esc(r.finding)}">${esc(r.finding)}</span></td>
            <td><span class="cell-clip" title="${esc(r.conclusion)}">${esc(r.conclusion)}</span></td>
            <td>${
              (r.expect_rules || []).length
                ? r.expect_rules
                    .map((x) => `<span class="chip rule">${esc(x)}</span>`)
                    .join("")
                : '<span class="chip">无</span>'
            }</td>
          </tr>`,
          )
          .join("")}</tbody>
      </table></div>`,
    );
  } catch (e) {
    toast(e.message, "error");
  }
}

async function runDataset(dsId) {
  const card = $(`#ds-${dsId}`);
  if (card) card.classList.add("running");
  renderStepper(
    $("#datasetStepper"),
    makeSteps(["加载数据", "纠错分析", "完成"], 1),
  );
  $("#datasetHint").textContent = "正在执行纠错分析…";
  $("#datasetResult").innerHTML = "";
  try {
    const data = await fetchJSON(
      `/api/datasets/${encodeURIComponent(dsId)}/check`,
      { method: "POST" },
    );
    renderStepper(
      $("#datasetStepper"),
      makeSteps(["加载数据", "纠错分析", "完成"], -1, true),
    );
    const name = data.source && data.source.name ? data.source.name : dsId;
    $("#datasetHint").textContent =
      `「${name}」纠错完成：${data.summary.issue_total} 条问题 / ${data.summary.total} 份报告。`;
    renderResult($("#datasetResult"), data, `测试数据集：${name}`);
    toast("纠错完成", "success");
  } catch (e) {
    renderStepper(
      $("#datasetStepper"),
      makeSteps(["加载数据", "纠错分析", "完成"], 1, false, true),
    );
    $("#datasetHint").textContent = "纠错执行失败，请重试。";
    toast(e.message, "error");
  } finally {
    if (card) card.classList.remove("running");
  }
}

/* ================= 结果展示 ================= */
const FIELD_NAMES = {
  finding: "超声所见",
  conclusion: "超声提示",
  exam_item: "检查项目",
  text: "报告全文",
  gender: "性别",
  report: "报告整体",
};
/* 9 类质控规则（即错误类型，名称与质控 Excel「错误类型」列一致，不可改动） */
const ERROR_TYPES = [
  "描述与结论不一致",
  "分级不规范",
  "语病或表述不规范",
  "数据或单位错误",
  "错别字或字段缺失",
  "模板文字未删除",
  "遗漏检查部位",
  "其他",
  "性别年龄冲突",
];

function renderResult(container, data, sourceLabel) {
  const s = data.summary;

  container.innerHTML = `
    <div class="card">
      <h2>纠错结果 <span class="h-sub">${esc(sourceLabel)} · 引擎 v${esc(data.engine_version)} · 耗时 ${s.elapsed_ms} ms</span></h2>
      ${
        s.issue_total === 0
          ? '<div class="result-banner ok">✓ 全部报告通过质控检查，未发现问题</div>'
          : `<div class="result-banner bad">⚠ 共发现 ${s.issue_total} 条问题，涉及 ${s.reports_with_issues} 份报告</div>`
      }
      <div class="result-summary">
        <div class="stat"><div class="stat-num">${s.total}</div><div class="stat-label">报告总数</div></div>
        <div class="stat ${s.reports_with_issues ? "bad" : "good"}">
          <div class="stat-num">${s.reports_with_issues}</div><div class="stat-label">问题报告数</div></div>
        <div class="stat ${s.issue_total ? "warn" : "good"}">
          <div class="stat-num">${s.issue_total}</div><div class="stat-label">问题总数</div></div>
        <div class="stat"><div class="stat-num">${s.elapsed_ms}<span style="font-size:12px"> ms</span></div>
          <div class="stat-label">引擎处理耗时</div></div>
      </div>
      ${
        Object.keys(s.by_error_type || {}).length
          ? `<div class="ds-meta" style="margin-bottom:14px">${Object.entries(
              s.by_error_type,
            )
              .map(
                ([t, n]) => `<span class="chip rule">${esc(t)} × ${n}</span>`,
              )
              .join("")}</div>`
          : ""
      }
      <div class="btn-row" style="margin-bottom:14px">
        <button class="btn-small primary" data-export>导出纠错结果（Excel）</button>
        <span class="fb-label" style="align-self:center">导出文件包含「质控概览 + 错误明细」两个 Sheet</span>
      </div>
      <div class="filter-bar" id="${container.id}-filters">
        <span class="fb-label">筛选：</span>
        <button class="f-btn active" data-sev="all">全部</button>
        <button class="f-btn" data-sev="high">高风险</button>
        <button class="f-btn" data-sev="medium">中风险</button>
        <select data-etype>
          <option value="all">全部规则</option>
          ${ERROR_TYPES.map((t) => `<option value="${esc(t)}">${esc(t)}</option>`).join("")}
        </select>
        <span class="fb-label" data-count></span>
      </div>
      <div class="issue-list" data-issues></div>
      <h2 style="margin-top:22px">报告明细</h2>
      <div class="table-wrap">
        <table class="data-table report-table">
          <thead><tr><th>报告ID</th><th>检查项目</th><th>性别/年龄</th><th>问题数</th><th>状态</th><th>详情</th></tr></thead>
          <tbody data-reports>
            ${data.results
              .map(
                (r) => `
              <tr>
                <td>${esc(r.report_id)}</td>
                <td>${esc(r.report.exam_item) || "—"}</td>
                <td>${esc(r.report.gender) || "—"}/${esc(r.report.age) || "—"}</td>
                <td>${r.issue_count}</td>
                <td>${
                  r.issue_count
                    ? '<span class="status-bad">发现问题</span>'
                    : '<span class="status-ok">通过</span>'
                }</td>
                <td><span class="rid-link" data-rid="${esc(r.report_id)}">查看</span></td>
              </tr>`,
              )
              .join("")}
          </tbody>
        </table>
      </div>
    </div>`;

  // 问题筛选与列表
  const state = { sev: "all", etype: "all" };
  const issues = [];
  data.results.forEach((r) =>
    r.issues.forEach((it) => issues.push({ ...it, report_id: r.report_id })),
  );
  const listEl = container.querySelector("[data-issues]");
  const countEl = container.querySelector("[data-count]");
  const paint = () => {
    const shown = issues.filter(
      (it) =>
        (state.sev === "all" || it.severity === state.sev) &&
        (state.etype === "all" || it.error_type === state.etype),
    );
    countEl.textContent = `显示 ${shown.length} / ${issues.length} 条问题`;
    if (!issues.length) {
      listEl.innerHTML = '<p class="empty-note">本次纠错未发现问题。</p>';
      return;
    }
    if (!shown.length) {
      listEl.innerHTML = '<p class="empty-note">当前筛选条件下没有问题。</p>';
      return;
    }
    listEl.innerHTML = shown
      .map(
        (it) => `
      <div class="issue-card ${esc(it.severity)}">
        <div class="issue-top">
          <span class="badge rule">${esc(it.error_type || "其他")}</span>
          <span class="badge ${it.severity === "high" ? "sev-high" : "sev-medium"}">
            ${it.severity === "high" ? "高风险" : "中风险"}</span>
          <span class="badge rid">${esc(it.report_id)}</span>
          <span class="issue-field">位置：${esc(FIELD_NAMES[it.field] || it.field)}</span>
        </div>
        <p class="issue-msg">${esc(it.message)}</p>
        <p class="issue-sugg">💡 ${esc(it.suggestion)}</p>
        <p class="issue-frag"><b>原文：</b>${esc(it.fragment)}</p>
      </div>`,
      )
      .join("");
  };
  paint();

  container.querySelectorAll("[data-sev]").forEach((btn) =>
    btn.addEventListener("click", () => {
      container
        .querySelectorAll("[data-sev]")
        .forEach((b) => b.classList.remove("active"));
      btn.classList.add("active");
      state.sev = btn.dataset.sev;
      paint();
    }),
  );
  container.querySelector("[data-etype]").addEventListener("change", (e) => {
    state.etype = e.target.value;
    paint();
  });

  // 报告详情
  container.querySelectorAll("[data-rid]").forEach((el) =>
    el.addEventListener("click", () => {
      const r = data.results.find((x) => x.report_id === el.dataset.rid);
      if (r) showReportDetail(r);
    }),
  );

  // 导出纠错结果
  container
    .querySelector("[data-export]")
    .addEventListener("click", async (e) => {
      const btn = e.currentTarget;
      const old = btn.textContent;
      btn.disabled = true;
      btn.textContent = "正在导出…";
      try {
        await exportResult(data, sourceLabel);
        toast("纠错结果已导出", "success");
      } catch (err) {
        toast(err.message, "error");
      } finally {
        btn.disabled = false;
        btn.textContent = old;
      }
    });
}

/* ================= 结果导出 ================= */
async function exportResult(data, sourceLabel) {
  const reports = data.results.map((r) => ({
    report_id: r.report_id,
    ...r.report,
    source_file: r.report && r.report.source_file,
  }));
  let resp;
  try {
    resp = await fetch(API + "/api/export", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ reports, source: sourceLabel }),
    });
  } catch (e) {
    throw new Error("无法连接后端服务，请确认服务已启动");
  }
  if (!resp.ok) {
    let detail = "";
    try {
      detail = (await resp.json()).detail || "";
    } catch (e) {
      /* 非 JSON 响应 */
    }
    throw new Error(detail || `导出失败（HTTP ${resp.status}）`);
  }
  const blob = await resp.blob();
  const cd = resp.headers.get("Content-Disposition") || "";
  const m = cd.match(/filename\*?=(?:UTF-8'')?"?([^";]+)"?/i);
  const filename = m ? decodeURIComponent(m[1]) : "纠错结果.xlsx";
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function showReportDetail(r) {
  const issuesHtml = r.issues.length
    ? r.issues
        .map(
          (it) => `
        <div class="issue-card ${esc(it.severity)}">
          <div class="issue-top">
            <span class="badge rule">${esc(it.error_type || "其他")}</span>
            <span class="badge ${it.severity === "high" ? "sev-high" : "sev-medium"}">
              ${it.severity === "high" ? "高风险" : "中风险"}</span>
          </div>
          <p class="issue-msg">${esc(it.message)}</p>
          <p class="issue-sugg">💡 ${esc(it.suggestion)}</p>
        </div>`,
        )
        .join("")
    : '<p class="empty-note">该报告通过全部质控规则检查。</p>';

  openModal(
    `报告详情 · ${r.report_id}`,
    `
    <div class="table-wrap"><table class="data-table">
      <tbody>
        <tr><th style="width:110px">检查项目</th><td>${esc(r.report.exam_item) || "—"}</td></tr>
        <tr><th>性别 / 年龄</th><td>${esc(r.report.gender) || "—"} / ${esc(r.report.age) || "—"}</td></tr>
        <tr><th>超声所见</th><td>${esc(r.report.finding) || "—"}</td></tr>
        <tr><th>超声提示</th><td>${esc(r.report.conclusion) || "—"}</td></tr>
      </tbody>
    </table></div>
    <h2 style="margin:18px 0 10px;font-size:15px">质控问题（${r.issue_count}）</h2>
    <div class="issue-list">${issuesHtml}</div>`,
  );
}

/* ================= 弹窗 ================= */
function openModal(title, html) {
  $("#modalTitle").textContent = title;
  $("#modalBody").innerHTML = html;
  $("#modalMask").classList.remove("hidden");
}
$("#modalClose").addEventListener("click", () =>
  $("#modalMask").classList.add("hidden"),
);
$("#modalMask").addEventListener("click", (e) => {
  if (e.target === $("#modalMask")) $("#modalMask").classList.add("hidden");
});
document.addEventListener("keydown", (e) => {
  if (e.key === "Escape") $("#modalMask").classList.add("hidden");
});

/* ================= 帮助页：质控规则（9 类，即错误类型，名称不可改动） ================= */
const RULES_HELP = [
  [
    "描述与结论不一致",
    "所见含结节/囊肿/占位等强病灶词而结论为纯阴性表述、左右侧互相冲突、结论未涵盖所见病灶",
  ],
  [
    "分级不规范",
    "分级前后不一致、同一结论多分级并存（3类/4a类并写）、恶性征象对应分级偏低、分级写法不规范（如“3级”应写“3类”）、分级体系与检查部位不符（如甲状腺报告误用乳腺 BI-RADS、妇科报告误用 C-TIRADS）",
  ],
  [
    "语病或表述不规范",
    "建议语缺时间数值、叠词重复（如“大小大小”）、标点重复、空括号、属性缺描述（如“外形，形态，”后无描述）等书写语病",
  ],
  [
    "数据或单位错误",
    "测量值单位错误/缺失（如径线写作 m 或漏写单位）、未删测量值占位符、径线量纲异常（疑似小数点错）、数值粘连/丢失乘号（如 1.512cm）、测量范围颠倒（如 2.5-1.6cm）、径线数值过小单位 mm 疑应为 cm（如 0.13*0.1mm）",
  ],
  [
    "错别字或字段缺失",
    "命中纠错映射表的错别字（如 回音→回声）、核心字段（所见/结论）缺失",
  ],
  [
    "模板文字未删除",
    "整句模板重复（同句重复出现）、同一部位矛盾描述（见淋巴结又未见淋巴结）、测量值占位符残留、括号残片/无关残留字词（如“(不完整，)”“未了。”）、无关字母串乱码（如 ggun）、院内模板占位符未替换（如“*侧”“*个”“mm×mm”“低回声/高回声/强回声”等斜杠选项未选择）",
  ],
  [
    "遗漏检查部位",
    "检查项目要求多部位，所见仅描述其中一部分，报出缺失部位（如腹部彩超漏双肾、阴囊彩超漏附睾、肝胆胰脾漏脾）",
  ],
  ["其他", "检查项目与实际检查内容不符（如妇科项目写胎儿描述）"],
  [
    "性别年龄冲突",
    "患者性别与检查项目/所见/结论中的特有部位矛盾（男见子宫、女见前列腺等）、年龄与检查内容逻辑冲突（如 79 岁出现妊娠/胎儿描述）、年龄超出合理范围（如 170 岁）",
  ],
];

function renderRuleGrid() {
  $("#ruleGrid").innerHTML = RULES_HELP.map(
    ([name, desc]) => `
    <div class="rule-item">
      <div class="ri-head">
        <span class="badge rule">${esc(name)}</span>
      </div>
      <p class="ri-desc">${esc(desc)}</p>
    </div>`,
  ).join("");
}

/* ================= 初始化 ================= */
setUploadStepper(-1);
renderStepper(
  $("#datasetStepper"),
  makeSteps(["加载数据", "纠错分析", "完成"]),
);
loadHealth();
loadDatasets();
renderRuleGrid();

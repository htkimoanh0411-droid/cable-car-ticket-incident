import React, { useState, useEffect, useCallback, useMemo } from "react";
import * as XLSX from "xlsx";
import {
  AlertTriangle,
  Radio,
  CheckCircle2,
  ClipboardList,
  Send,
  Clock,
  MapPin,
  Users,
  Copy,
  Check,
  X,
  Plus,
  ChevronDown,
  ChevronUp,
  History,
  UploadCloud,
  FileSpreadsheet,
  Download,
  ScanLine,
  AlertCircle,
  Printer,
  LayoutDashboard,
  BarChart3,
  CalendarRange,
  TrendingUp,
  PieChart,
  Trash2,
  Cpu,
} from "lucide-react";

const DEPARTMENTS = [
  { id: "soat_ve", label: "Soát vé / Cổng" },
  { id: "bao_ve", label: "An ninh" },
  { id: "ky_thuat", label: "Kỹ thuật (điện / mạng)" },
  { id: "van_hanh", label: "Vận hành" },
  { id: "quan_ly_ca", label: "Quản lý ca trực" },
  { id: "ke_toan", label: "Kế toán" },
];

const INCIDENT_TYPES = [
  "Mất điện",
  "Mất mạng / lỗi đường truyền",
  "Lỗi phần mềm soát vé",
  "Lỗi cổng soát vé (cơ khí)",
  "Máy in vé / máy quét lỗi",
  "Khác",
];

const STATUS_FLOW = [
  { id: "moi", label: "Mới kích hoạt", color: "#E5484D" },
  { id: "dang_xu_ly", label: "Đang xử lý", color: "#F5A623" },
  { id: "da_lap_bb", label: "Đã lập biên bản", color: "#4F8EF7" },
  { id: "cho_bs", label: "Chờ BCNTT xử lý bổ sung", color: "#F5D90A" },
  { id: "da_bs", label: "Đã xử lý bổ sung", color: "#9B8CFF" },
  { id: "da_gui_kt", label: "Đã gửi kế toán", color: "#3DD68C" },
];

function statusMeta(id) {
  return STATUS_FLOW.find((s) => s.id === id) || STATUS_FLOW[0];
}

function fmtTime(ts) {
  const d = new Date(ts);
  return d.toLocaleString("vi-VN", {
    day: "2-digit",
    month: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
  });
}

function genId(existingIncidents) {
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, "0");
  const d = String(now.getDate()).padStart(2, "0");
  const datePart = `${y}${m}${d}`;
  const todayCount = (existingIncidents || []).filter((inc) => {
    const t = new Date(inc.createdAt);
    return (
      t.getFullYear() === y && t.getMonth() === now.getMonth() && t.getDate() === now.getDate()
    );
  }).length;
  const seq = String(todayCount + 1).padStart(2, "0");
  return `SC-${datePart}-${seq}`;
}

const BOARD_KEY = "incidents-board";

// ---------- Đối soát dữ liệu vé (theo số lượng / loại vé) ----------

function normHeader(h) {
  return String(h)
    .toLowerCase()
    .replace(/đ/g, "d")
    .replace(/Đ/g, "d")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim();
}

// File Access Log thật có vài dòng thông tin tiêu đề (tên báo cáo, ngày giờ chạy,
// khoảng thời gian, địa điểm) trước khi tới dòng cột thật. Hàm này dò đúng dòng cột
// (chứa "UsageDateTime" / "ProductName" / "TicketSerial"...) rồi mới đọc dữ liệu bên dưới.
async function parseAccessLogFile(file) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array", cellDates: true });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", blankrows: true });

  let headerIdx = -1;
  for (let i = 0; i < Math.min(rows.length, 40); i++) {
    const norm = (rows[i] || []).map((c) => normHeader(c));
    if (norm.some((c) => c.includes("usagedatetime")) && norm.some((c) => c.includes("productname"))) {
      headerIdx = i;
      break;
    }
  }
  if (headerIdx === -1) {
    throw new Error("Không tìm thấy dòng tiêu đề cột (UsageDateTime / ProductName) trong file.");
  }

  const headerRow = rows[headerIdx].map((h) => normHeader(h));
  const colIndex = (patterns) => headerRow.findIndex((h) => patterns.some((p) => h.includes(p)));
  const idx = {
    type: colIndex(["usagetypedesc"]),
    qty: colIndex(["groupquantity"]),
    point: colIndex(["accesspointname"]),
    area: colIndex(["accessareaname"]),
    productName: colIndex(["productname"]),
    status: colIndex(["ticketstatus"]),
    media: colIndex(["mediacode"]),
  };
  if (idx.productName === -1) {
    throw new Error("Không tìm thấy cột ProductName trong file.");
  }

  const records = [];
  for (let r = headerIdx + 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row || row.length === 0) continue;
    const productName = String(row[idx.productName] ?? "").trim();
    if (!productName) continue;
    records.push({
      type: idx.type >= 0 ? String(row[idx.type] ?? "").trim() : "",
      qty: idx.qty >= 0 ? Number(row[idx.qty]) || 0 : 1,
      point: idx.point >= 0 ? String(row[idx.point] ?? "").trim() : "",
      area: idx.area >= 0 ? String(row[idx.area] ?? "").trim() : "",
      productName,
      status: idx.status >= 0 ? String(row[idx.status] ?? "").trim() : "",
      media: idx.media >= 0 ? String(row[idx.media] ?? "").trim() : "",
    });
  }
  return records;
}

// Cộng dồn GroupQuantity theo ProductName, chỉ tính các dòng UsageTypeDesc = "Entry"
function aggregateGateRecords(records) {
  const map = new Map();
  let totalRowsUsed = 0;
  records.forEach((r) => {
    if (r.type.toLowerCase() !== "entry") return;
    totalRowsUsed += 1;
    const key = r.productName;
    map.set(key, (map.get(key) || 0) + (r.qty || 0));
  });
  const breakdown = Array.from(map.entries()).map(([type, qty]) => ({ type, qty }));
  return { breakdown, totalRowsUsed, totalRows: records.length };
}

// So sánh số lượng theo loại vé giữa biên bản (nhập tay) và Access Log (đã tổng hợp)
function reconcileByType(manualBreakdown, gateBreakdown) {
  const norm = (s) => String(s || "").trim().toLowerCase();
  const gateMap = new Map(gateBreakdown.map((g) => [norm(g.type), g]));
  const manualMap = new Map(manualBreakdown.map((m) => [norm(m.type), m]));
  const matched = [];
  const mismatch = [];
  const missing = [];
  const extra = [];
  manualBreakdown.forEach((m) => {
    const g = gateMap.get(norm(m.type));
    if (!g) {
      missing.push({ type: m.type, manualQty: m.qty });
      return;
    }
    if (Number(m.qty) === Number(g.qty)) {
      matched.push({ type: m.type, manualQty: m.qty, gateQty: g.qty });
    } else {
      mismatch.push({ type: m.type, manualQty: m.qty, gateQty: g.qty });
    }
  });
  gateBreakdown.forEach((g) => {
    if (!manualMap.has(norm(g.type))) extra.push({ type: g.type, gateQty: g.qty });
  });
  return { matched, mismatch, missing, extra };
}

// File soát vé tay do bộ phận soát vé upload (VD: "Báo cáo quyền lợi") - cấu trúc phẳng,
// 1 dòng tiêu đề duy nhất, cột chính là Media Code (QR) dùng để đối soát.
async function parseManualScanFile(file) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  if (!json.length) return [];
  const headers = Object.keys(json[0]);
  const find = (patterns) => headers.find((h) => patterns.some((p) => normHeader(h).includes(p)));
  const statusKey = find(["trang thai", "status"]);
  const mediaKey = find(["media code", "mediacode", "qr"]);
  const gateKey = find(["vi tri cong", "gate", "vi tri"]);
  const scannerKey = find(["nguoi quet", "scanner"]);
  const reasonKey = find(["ly do", "reason"]);
  const timeKey = find(["thoi gian", "time"]);

  return json
    .map((row) => ({
      status: statusKey ? String(row[statusKey] ?? "").trim() : "",
      mediaCode: mediaKey ? String(row[mediaKey] ?? "").trim() : "",
      gate: gateKey ? String(row[gateKey] ?? "").trim() : "",
      scannedBy: scannerKey ? String(row[scannerKey] ?? "").trim() : "",
      reason: reasonKey ? String(row[reasonKey] ?? "").trim() : "",
      time: timeKey ? String(row[timeKey] ?? "").trim() : "",
    }))
    .filter((r) => r.mediaCode);
}

// Đối soát theo Media Code: so khớp danh sách mã đã soát tay (bộ phận soát vé, lúc bàn giao BCNTT)
// với danh sách mã có trong dữ liệu kế toán cung cấp (Access Log)
function reconcileByMediaCode(manualRecords, gateRecords) {
  const gateSet = new Set(gateRecords.map((r) => r.media).filter(Boolean));
  const matched = [];
  const missing = [];
  manualRecords.forEach((m) => {
    if (gateSet.has(m.mediaCode)) matched.push(m);
    else missing.push(m);
  });
  return { matched, missing, totalManual: manualRecords.length, totalGateCodes: gateSet.size };
}

// Kiểm tra dữ liệu đối soát đã khớp hoàn toàn chưa (không còn thiếu/lệch/thừa).
// Trả về { done, reason } — done=false thì reason giải thích lý do chưa gửi kế toán được.
function reconciliationStatus(inc) {
  if (!inc.gateBreakdown) {
    return { done: false, reason: "Chưa chạy đối soát — cần hoàn tất ở mục \"Đối soát dữ liệu\" trước." };
  }
  const hasManualScan = (inc.manualScanData || []).length > 0;
  if (hasManualScan) {
    const r = inc.mediaCodeResult;
    if (!r) {
      return { done: false, reason: "Chưa có kết quả đối soát theo Media Code." };
    }
    if (r.missing.length > 0) {
      return {
        done: false,
        reason: `Còn ${r.missing.length} mã Media Code chưa khớp (chưa thấy quẹt bù). Yêu cầu kế toán tải lại đúng file Access Log và chạy lại đối soát cho đến khi khớp hết.`,
      };
    }
    return { done: true, reason: "" };
  }
  const manualBreakdown = (inc.report?.breakdown || []).filter((r) => r.type && Number(r.qty) >= 0);
  if (manualBreakdown.length === 0) return { done: true, reason: "" };
  const r = reconcileByType(manualBreakdown, inc.gateBreakdown);
  const totalIssues = r.mismatch.length + r.missing.length + r.extra.length;
  if (totalIssues > 0) {
    return {
      done: false,
      reason: `Còn ${totalIssues} dòng chưa khớp (lệch/thiếu/thừa). Yêu cầu kế toán đối soát lại cho đến khi khớp số liệu.`,
    };
  }
  return { done: true, reason: "" };
}

function downloadReconciliationExcel(inc, result, gateMeta) {
  const wb = XLSX.utils.book_new();
  const totalManual = [...result.matched, ...result.mismatch, ...result.missing].reduce(
    (s, x) => s + (Number(x.manualQty) || 0),
    0
  );
  const totalGate = [...result.matched, ...result.mismatch, ...result.extra].reduce(
    (s, x) => s + (Number(x.gateQty) || 0),
    0
  );
  const summary = [
    { "Chỉ tiêu": "Mã sự cố", "Giá trị": inc.id },
    { "Chỉ tiêu": "Vị trí", "Giá trị": inc.location },
    { "Chỉ tiêu": "File Access Log", "Giá trị": gateMeta?.fileName || "" },
    { "Chỉ tiêu": "Tổng số vé biên bản (thủ công)", "Giá trị": totalManual },
    { "Chỉ tiêu": "Tổng số vé Access Log (Entry)", "Giá trị": totalGate },
    { "Chỉ tiêu": "Loại vé khớp", "Giá trị": result.matched.length },
    { "Chỉ tiêu": "Loại vé lệch số lượng", "Giá trị": result.mismatch.length },
    { "Chỉ tiêu": "Loại vé thiếu (chưa thấy quẹt bù)", "Giá trị": result.missing.length },
    { "Chỉ tiêu": "Loại vé thừa (không có trong biên bản)", "Giá trị": result.extra.length },
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summary), "Tong hop");
  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.json_to_sheet(
      result.matched.map((x) => ({ "Loại vé": x.type, "SL biên bản": x.manualQty, "SL Access Log": x.gateQty }))
    ),
    "Khop"
  );
  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.json_to_sheet(
      result.mismatch.map((x) => ({
        "Loại vé": x.type,
        "SL biên bản": x.manualQty,
        "SL Access Log": x.gateQty,
        "Chênh lệch": x.gateQty - x.manualQty,
      }))
    ),
    "Lech so luong"
  );
  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.json_to_sheet(result.missing.map((x) => ({ "Loại vé": x.type, "SL biên bản": x.manualQty }))),
    "Thieu"
  );
  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.json_to_sheet(result.extra.map((x) => ({ "Loại vé": x.type, "SL Access Log": x.gateQty }))),
    "Thua"
  );
  const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const blob = new Blob([wbout], { type: "application/octet-stream" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `Doi-soat-${inc.id}.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function downloadMediaCodeReconciliationExcel(inc, result) {
  const wb = XLSX.utils.book_new();
  const summary = [
    { "Chỉ tiêu": "Mã sự cố", "Giá trị": inc.id },
    { "Chỉ tiêu": "Vị trí", "Giá trị": inc.location },
    { "Chỉ tiêu": "File Access Log (kế toán)", "Giá trị": inc.gateMeta?.fileName || "" },
    { "Chỉ tiêu": "File soát vé tay (bộ phận soát vé)", "Giá trị": inc.manualScanMeta?.fileName || "" },
    { "Chỉ tiêu": "Tổng số Media Code soát tay", "Giá trị": result.totalManual },
    { "Chỉ tiêu": "Khớp (đã quẹt bù)", "Giá trị": result.matched.length },
    { "Chỉ tiêu": "Thiếu (chưa thấy quẹt bù)", "Giá trị": result.missing.length },
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summary), "Tong hop");
  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.json_to_sheet(
      result.matched.map((x) => ({
        "Media Code": x.mediaCode,
        "Trạng thái": x.status,
        "Cổng/Vị trí": x.gate,
        "Người quét": x.scannedBy,
        "Lý do": x.reason,
        "Thời gian": x.time,
      }))
    ),
    "Khop"
  );
  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.json_to_sheet(
      result.missing.map((x) => ({
        "Media Code": x.mediaCode,
        "Trạng thái": x.status,
        "Cổng/Vị trí": x.gate,
        "Người quét": x.scannedBy,
        "Lý do": x.reason,
        "Thời gian": x.time,
      }))
    ),
    "Thieu"
  );
  const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const blob = new Blob([wbout], { type: "application/octet-stream" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `Doi-soat-MediaCode-${inc.id}.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function nowForDatetimeLocal() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

export default function IncidentApp() {
  const [incidents, setIncidents] = useState(null);
  const [printingId, setPrintingId] = useState(null);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [showForm, setShowForm] = useState(false);
  const [expandedId, setExpandedId] = useState(null);
  const [view, setView] = useState("active"); // active | history
  const [copiedId, setCopiedId] = useState(null);
  const [error, setError] = useState("");
  const [module, setModule] = useState("overview"); // overview | activate | active | recon | reports
  const [filterType, setFilterType] = useState("");
  const [filterLocation, setFilterLocation] = useState("");
  const [filterFrom, setFilterFrom] = useState("");
  const [filterTo, setFilterTo] = useState("");

  // form state
  const [form, setForm] = useState({
    type: INCIDENT_TYPES[0],
    otherDetail: "",
    location: "",
    description: "",
    depts: [],
    reporter: "",
    occurredAt: nowForDatetimeLocal(),
  });

  const load = useCallback(async () => {
    try {
      const res = await window.storage.get(BOARD_KEY, true);
      const list = res ? JSON.parse(res.value) : [];
      setIncidents(list);
    } catch (e) {
      setIncidents([]);
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    load();
    const interval = setInterval(load, 4000);
    return () => clearInterval(interval);
  }, [load]);

  const persist = async (next) => {
    setSaving(true);
    setIncidents(next);
    try {
      await window.storage.set(BOARD_KEY, JSON.stringify(next), true);
    } catch (e) {
      setError("Không lưu được thay đổi. Thử lại.");
    } finally {
      setSaving(false);
    }
  };

  const toggleDept = (id) => {
    setForm((f) => ({
      ...f,
      depts: f.depts.includes(id) ? f.depts.filter((d) => d !== id) : [...f.depts, id],
    }));
  };

  const activateIncident = async () => {
    if (!form.location.trim() || form.depts.length === 0) {
      setError("Cần nhập vị trí và chọn ít nhất một bộ phận phối hợp.");
      return false;
    }
    if (form.type === "Khác" && !form.otherDetail.trim()) {
      setError('Vui lòng nhập mô tả chi tiết cho loại sự cố "Khác".');
      return false;
    }
    setError("");
    const typeLabel = form.type === "Khác" ? `Khác — ${form.otherDetail.trim()}` : form.type;
    const occurredAt = form.occurredAt ? new Date(form.occurredAt).getTime() : Date.now();
    const incident = {
      id: genId(incidents),
      createdAt: Date.now(),
      occurredAt,
      type: typeLabel,
      location: form.location.trim(),
      description: form.description.trim(),
      reporter: form.reporter.trim() || "Chưa ghi tên",
      depts: form.depts,
      status: "moi",
      acks: [],
      report: null,
      history: [{ status: "moi", time: Date.now(), by: form.reporter.trim() || "—" }],
    };
    const next = [incident, ...(incidents || [])];
    await persist(next);
    setForm({
      type: INCIDENT_TYPES[0],
      otherDetail: "",
      location: "",
      description: "",
      depts: [],
      reporter: "",
      occurredAt: nowForDatetimeLocal(),
    });
    setExpandedId(incident.id);
    return incident;
  };

  const deleteIncident = async (id) => {
    const next = (incidents || []).filter((inc) => inc.id !== id);
    await persist(next);
    if (expandedId === id) setExpandedId(null);
  };

  const updateIncident = async (id, updater) => {
    const next = (incidents || []).map((inc) => (inc.id === id ? updater(inc) : inc));
    await persist(next);
  };

  const ackDept = (id, deptId, name, role) => {
    updateIncident(id, (inc) => {
      if (inc.acks.some((a) => a.dept === deptId)) return inc;
      const acks = [...inc.acks, { dept: deptId, time: Date.now(), by: name || "—", role: role || "" }];
      const nextStatus = inc.status === "moi" ? "dang_xu_ly" : inc.status;
      return {
        ...inc,
        acks,
        status: nextStatus,
        history:
          nextStatus !== inc.status
            ? [...inc.history, { status: nextStatus, time: Date.now(), by: name || "—" }]
            : inc.history,
      };
    });
  };

  const setStatus = (id, statusId, by) => {
    updateIncident(id, (inc) => ({
      ...inc,
      status: statusId,
      history: [...inc.history, { status: statusId, time: Date.now(), by: by || "—" }],
    }));
  };

  const saveReport = (id, report) => {
    updateIncident(id, (inc) => ({
      ...inc,
      report,
      status: "da_lap_bb",
      history: [
        ...inc.history,
        { status: "da_lap_bb", time: Date.now(), by: inc.acks[0]?.by || "—" },
      ],
    }));
  };

  const setGateBreakdown = (id, breakdown, meta, mediaCodeResult) => {
    updateIncident(id, (inc) => ({
      ...inc,
      gateBreakdown: breakdown,
      gateMeta: meta,
      mediaCodeResult: mediaCodeResult !== undefined ? mediaCodeResult : inc.mediaCodeResult,
    }));
  };

  const uploadSignedReport = (id, signedReport) => {
    updateIncident(id, (inc) => ({ ...inc, signedReport }));
  };

  const sendHandover = (id, by, manualScanData, manualScanMeta) => {
    updateIncident(id, (inc) => ({
      ...inc,
      handover: { time: Date.now(), by: by || "—" },
      manualScanData: manualScanData || inc.manualScanData || [],
      manualScanMeta: manualScanMeta || inc.manualScanMeta,
      status: "cho_bs",
      history: [...inc.history, { status: "cho_bs", time: Date.now(), by: by || "—" }],
    }));
  };

  const updateHandoverData = (id, manualScanData, manualScanMeta, by) => {
    updateIncident(id, (inc) => ({
      ...inc,
      manualScanData,
      manualScanMeta,
      history: [
        ...inc.history,
        { status: inc.status, time: Date.now(), by: (by || "—") + " (sửa dữ liệu bàn giao)" },
      ],
    }));
  };

  const confirmSupplement = (id, by, role, note) => {
    updateIncident(id, (inc) => ({
      ...inc,
      supplement: { time: Date.now(), by: by || "—", role: role || "", note: note || "" },
      status: "da_bs",
      history: [...inc.history, { status: "da_bs", time: Date.now(), by: by || "—" }],
    }));
  };

  const sendToAccounting = (id) => {
    updateIncident(id, (inc) => ({
      ...inc,
      status: "da_gui_kt",
      history: [...inc.history, { status: "da_gui_kt", time: Date.now(), by: "—" }],
    }));
  };

  const copyBrief = (inc) => {
    const deptNames = inc.depts.map((d) => DEPARTMENTS.find((x) => x.id === d)?.label).join(", ");
    const text =
      `[SỰ CỐ ${inc.id}] ${inc.type}\n` +
      `Vị trí: ${inc.location}\n` +
      `Mô tả: ${inc.description || "—"}\n` +
      `Bộ phận phối hợp: ${deptNames}\n` +
      `Người báo cáo: ${inc.reporter}\n` +
      `Thời gian xảy ra: ${fmtTime(inc.occurredAt || inc.createdAt)}\n` +
      `→ Đề nghị các bộ phận ra hiện trường phối hợp kiểm tra, soát vé qua cổng thủ công.`;
    navigator.clipboard?.writeText(text);
    setCopiedId(inc.id);
    setTimeout(() => setCopiedId(null), 1800);
  };

  const activeList = (incidents || []).filter((i) => i.status !== "da_gui_kt");
  const historyList = (incidents || []).filter((i) => i.status === "da_gui_kt");
  const shown = view === "active" ? activeList : historyList;

  const filteredShown = shown.filter((i) => {
    if (filterType && i.type !== filterType) return false;
    if (filterLocation && !i.location.toLowerCase().includes(filterLocation.toLowerCase())) return false;
    if (filterFrom && i.createdAt < new Date(filterFrom).getTime()) return false;
    if (filterTo && i.createdAt > new Date(filterTo).getTime() + 86400000) return false;
    return true;
  });
  const hasActiveFilter = filterType || filterLocation || filterFrom || filterTo;
  const clearFilters = () => {
    setFilterType("");
    setFilterLocation("");
    setFilterFrom("");
    setFilterTo("");
  };

  // ---- Số liệu tổng quan & cảnh báo SLA ----
  const now = Date.now();
  const SLA_MINUTES = 30;
  const overdueList = activeList.filter(
    (i) => i.status === "moi" || i.status === "dang_xu_ly"
  ).filter((i) => (now - i.createdAt) / 60000 > SLA_MINUTES);

  const all = incidents || [];
  const closedWithDuration = all
    .filter((i) => i.status === "da_gui_kt")
    .map((i) => {
      const closedAt = i.history[i.history.length - 1]?.time || i.createdAt;
      return (closedAt - i.createdAt) / 60000;
    });
  const avgResolveMin = closedWithDuration.length
    ? Math.round(closedWithDuration.reduce((s, m) => s + m, 0) / closedWithDuration.length)
    : null;

  const typeCounts = {};
  all.forEach((i) => {
    typeCounts[i.type] = (typeCounts[i.type] || 0) + 1;
  });
  const topTypes = Object.entries(typeCounts).sort((a, b) => b[1] - a[1]).slice(0, 5);

  const deptCounts = {};
  all.forEach((i) => i.depts.forEach((d) => (deptCounts[d] = (deptCounts[d] || 0) + 1)));

  const reconciledCount = all.filter((i) => i.gateBreakdown).length;
  const mismatchIncidents = all.filter((i) => i.gateBreakdown && !reconciliationStatus(i).done);

  const monthNow = new Date().getMonth();
  const yearNow = new Date().getFullYear();
  const thisMonthCount = all.filter((i) => {
    const d = new Date(i.createdAt);
    return d.getMonth() === monthNow && d.getFullYear() === yearNow;
  }).length;

  const exportMonthlyExcel = (fromDate, toDate) => {
    const from = fromDate ? new Date(fromDate).getTime() : 0;
    const to = toDate ? new Date(toDate).getTime() + 86400000 : Infinity;
    const rows = all
      .filter((i) => i.createdAt >= from && i.createdAt <= to)
      .map((i) => {
        const totalManual = (i.report?.breakdown || []).reduce((s, r) => s + (Number(r.qty) || 0), 0);
        const totalGate = i.gateBreakdown ? i.gateBreakdown.reduce((s, r) => s + (Number(r.qty) || 0), 0) : "";
        const closedAt = i.history[i.history.length - 1]?.time;
        return {
          "Mã sự cố": i.id,
          "Loại sự cố": i.type,
          "Vị trí": i.location,
          "Thời gian xảy ra": fmtTime(i.occurredAt || i.createdAt),
          "Thời gian kích hoạt (nhập hệ thống)": fmtTime(i.createdAt),
          "Trạng thái": statusMeta(i.status).label,
          "Bộ phận phối hợp": i.depts.map((d) => DEPARTMENTS.find((x) => x.id === d)?.label).join(", "),
          "Tổng vé biên bản": totalManual,
          "Tổng vé Access Log": totalGate,
          "Đã đối soát": i.gateBreakdown ? "Có" : "Chưa",
          "Thời gian đóng": closedAt ? fmtTime(closedAt) : "",
        };
      });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), "Bao cao su co");
    const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const blob = new Blob([wbout], { type: "application/octet-stream" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `Bao-cao-su-co-${new Date().toISOString().slice(0, 10)}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const moduleForIncident = (inc) => {
    if (!inc) return "participate";
    if (inc.acks.length < inc.depts.length) return "participate";
    if (!inc.report) return "report";
    if (inc.status === "da_lap_bb" || inc.status === "cho_bs") return "handover";
    if (!inc.gateBreakdown) return "recon";
    return "closing";
  };

  const pendingParticipate = activeList.filter((i) => i.acks.length < i.depts.length).length;
  const pendingReport = activeList.filter((i) => i.acks.length >= i.depts.length && i.depts.length > 0 && !i.report).length;
  const pendingHandover = activeList.filter((i) => i.status === "da_lap_bb" || i.status === "cho_bs").length;
  const pendingRecon = activeList.filter((i) => i.status === "da_bs" && !i.gateBreakdown).length;
  const pendingClosing = activeList.filter((i) => i.status === "da_bs" && i.gateBreakdown).length;

  const NAV_ITEMS = [
    { id: "overview", label: "Tổng quan", icon: LayoutDashboard },
    { id: "activate", label: "Kích hoạt sự cố", icon: Plus },
    { id: "participate", label: "Xác nhận tham gia", icon: Users, badge: pendingParticipate },
    { id: "report", label: "Biên bản sự cố", icon: ClipboardList, badge: pendingReport },
    { id: "handover", label: "Bàn giao BCNTT", icon: Cpu, badge: pendingHandover },
    { id: "recon", label: "Đối soát dữ liệu", icon: ScanLine, badge: pendingRecon },
    { id: "closing", label: "Hoàn tất & Đóng", icon: CheckCircle2, badge: pendingClosing },
    { id: "reports", label: "Báo cáo", icon: BarChart3 },
  ];

  return (
    <>
    <div style={styles.appShell} className="app-root">
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=IBM+Plex+Mono:wght@400;500;600&family=Inter:wght@400;500;600&display=swap');
        * { box-sizing: border-box; }
        body { margin: 0; }
        .btn { cursor: pointer; border: none; transition: filter 0.15s ease, transform 0.1s ease; }
        .btn:hover { filter: brightness(1.12); }
        .btn:active { transform: scale(0.98); }
        .chip { cursor: pointer; user-select: none; transition: all 0.15s ease; }
        .card-enter { animation: slideIn 0.25s ease; }
        @keyframes slideIn { from { opacity: 0; transform: translateY(-6px); } to { opacity: 1; transform: translateY(0); } }
        ::-webkit-scrollbar { width: 8px; }
        ::-webkit-scrollbar-thumb { background: #2A3340; border-radius: 4px; }
        input:focus, textarea:focus { outline: 2px solid #4F8EF7; outline-offset: 1px; }
        .navbtn { cursor: pointer; transition: background 0.15s ease, color 0.15s ease; }
        .navbtn:hover { background: #1B222C; }
        .print-container { display: none; }
        @media print {
          .app-root { display: none !important; }
          .print-container { display: block !important; }
          @page { size: A4; margin: 14mm; }
        }
        @media (max-width: 860px) {
          .sidebar { flex-direction: row !important; width: 100% !important; height: auto !important; overflow-x: auto; border-right: none !important; border-bottom: 1px solid #232B36; }
          .sidebar-title { display: none; }
          .sidebar-nav { flex-direction: row !important; }
          .app-shell { flex-direction: column !important; }
        }
      `}</style>

      <nav style={styles.sidebar} className="sidebar">
        <div style={styles.sidebarHeader}>
          <div style={styles.logoMark}>
            <Radio size={18} color="#E5484D" strokeWidth={2.5} />
          </div>
          <div className="sidebar-title">
            <div style={styles.title}>TRẠM ĐIỀU PHỐI</div>
            <div style={styles.subtitle}>Sự cố soát vé</div>
          </div>
        </div>
        <div style={styles.sidebarNav} className="sidebar-nav">
          {NAV_ITEMS.map((item) => {
            const Icon = item.icon;
            return (
              <div
                key={item.id}
                className="navbtn"
                onClick={() => setModule(item.id)}
                style={{ ...styles.navItem, ...(module === item.id ? styles.navItemActive : {}) }}
              >
                <Icon size={16} />
                <span className="sidebar-title" style={{ flex: 1 }}>{item.label}</span>
                {item.badge > 0 && <span style={styles.navBadge}>{item.badge}</span>}
              </div>
            );
          })}
        </div>
        <div style={styles.sidebarFooter}>
          {saving && <span style={styles.savingTag}>đang lưu…</span>}
          <div style={styles.liveTag}>
            <span style={styles.liveDot} />
            <span className="sidebar-title">TRỰC TUYẾN</span>
          </div>
        </div>
      </nav>

      <div style={styles.contentArea}>
        {error && (
          <div style={styles.errorBar}>
            {error}
            <X size={14} style={{ cursor: "pointer" }} onClick={() => setError("")} />
          </div>
        )}

        {module === "overview" && (
          <OverviewModule
            activeCount={activeList.length}
            thisMonthCount={thisMonthCount}
            avgResolveMin={avgResolveMin}
            overdueList={overdueList}
            topTypes={topTypes}
            reconciledCount={reconciledCount}
            mismatchCount={mismatchIncidents.length}
            totalCount={all.length}
            onGoActivate={() => setModule("activate")}
            onGoActive={() => setModule("participate")}
            onOpenIncident={(id) => {
              const inc = (incidents || []).find((i) => i.id === id);
              setModule(moduleForIncident(inc));
              setExpandedId(id);
            }}
          />
        )}

        {module === "activate" && (
          <ActivateModule
            form={form}
            setForm={setForm}
            toggleDept={toggleDept}
            onActivate={activateIncident}
            onGoToParticipate={() => setModule("participate")}
          />
        )}

        {["participate", "report", "handover", "recon", "closing"].includes(module) && (
          <>
            <div style={styles.moduleHeaderRow}>
              <div>
                <div style={styles.moduleTitle}>
                  {module === "participate" && "Xác nhận tham gia"}
                  {module === "report" && "Biên bản sự cố"}
                  {module === "handover" && "Bàn giao BCNTT xử lý bổ sung"}
                  {module === "recon" && "Đối soát dữ liệu"}
                  {module === "closing" && "Hoàn tất & Đóng sự cố"}
                </div>
                <div style={styles.moduleSub}>
                  {module === "participate" && "Các bộ phận xác nhận đã tiếp nhận và có mặt xử lý sự cố"}
                  {module === "report" && "Ghi nhận số vé soát tay theo từng loại vé"}
                  {module === "handover" && "Chuyển thông tin cho BCNTT soát vé bổ sung khi hệ thống có điện trở lại, chờ xác nhận hoàn tất"}
                  {module === "recon" && "Đối soát số vé biên bản với Access Log quẹt bù"}
                  {module === "closing" && "Xuất biên bản PDF, thông báo và gửi kế toán"}
                </div>
              </div>
              <button className="btn" style={styles.activateBtn} onClick={() => setModule("activate")}>
                <Plus size={16} style={{ marginRight: 6 }} />
                Kích hoạt sự cố
              </button>
            </div>

            <div style={styles.tabs}>
              <button
                className="btn"
                onClick={() => setView("active")}
                style={{ ...styles.tab, ...(view === "active" ? styles.tabActive : {}) }}
              >
                <AlertTriangle size={15} style={{ marginRight: 6 }} />
                Đang xử lý ({activeList.length})
              </button>
              <button
                className="btn"
                onClick={() => setView("history")}
                style={{ ...styles.tab, ...(view === "history" ? styles.tabActive : {}) }}
              >
                <History size={15} style={{ marginRight: 6 }} />
                Đã đóng ({historyList.length})
              </button>
            </div>

            <div style={styles.filterBar}>
              <select
                style={styles.filterSelect}
                value={filterType}
                onChange={(e) => setFilterType(e.target.value)}
              >
                <option value="">Tất cả loại sự cố</option>
                {INCIDENT_TYPES.map((t) => (
                  <option key={t} value={t}>{t}</option>
                ))}
              </select>
              <input
                style={styles.filterInput}
                placeholder="Tìm theo vị trí / cổng..."
                value={filterLocation}
                onChange={(e) => setFilterLocation(e.target.value)}
              />
              <input
                type="date"
                style={styles.filterDate}
                value={filterFrom}
                onChange={(e) => setFilterFrom(e.target.value)}
                title="Từ ngày"
              />
              <input
                type="date"
                style={styles.filterDate}
                value={filterTo}
                onChange={(e) => setFilterTo(e.target.value)}
                title="Đến ngày"
              />
              {hasActiveFilter && (
                <button className="btn" style={styles.clearFilterBtn} onClick={clearFilters}>
                  <X size={13} style={{ marginRight: 4 }} />
                  Xoá lọc
                </button>
              )}
            </div>

            <main style={styles.main}>
              {loading ? (
                <div style={styles.emptyState}>Đang tải bảng sự cố…</div>
              ) : filteredShown.length === 0 ? (
                <div style={styles.emptyState}>
                  {hasActiveFilter
                    ? "Không có sự cố nào khớp với bộ lọc hiện tại."
                    : view === "active"
                    ? "Không có sự cố nào đang xử lý. Mọi cổng hoạt động bình thường."
                    : "Chưa có sự cố nào đã hoàn tất gửi kế toán."}
                </div>
              ) : (
                filteredShown.map((inc) => (
                  <IncidentCard
                    key={inc.id}
                    inc={inc}
                    section={module}
                    expanded={expandedId === inc.id}
                    onToggle={() => setExpandedId(expandedId === inc.id ? null : inc.id)}
                    onAck={ackDept}
                    onSaveReport={saveReport}
                    onSend={sendToAccounting}
                    onCopy={copyBrief}
                    copied={copiedId === inc.id}
                    onSetGateBreakdown={setGateBreakdown}
                    onPrint={() => setPrintingId(inc.id)}
                    onNavigate={setModule}
                    onDelete={deleteIncident}
                    onSendHandover={sendHandover}
                    onUpdateHandoverData={updateHandoverData}
                    onConfirmSupplement={confirmSupplement}
                    onUploadSignedReport={uploadSignedReport}
                  />
                ))
              )}
            </main>
          </>
        )}

        {module === "reports" && (
          <ReportsModule
            all={all}
            topTypes={topTypes}
            deptCounts={deptCounts}
            avgResolveMin={avgResolveMin}
            reconciledCount={reconciledCount}
            mismatchCount={mismatchIncidents.length}
            onExport={exportMonthlyExcel}
          />
        )}
      </div>
    </div>

    {printingId && (
      <PrintableReport
        inc={(incidents || []).find((i) => i.id === printingId)}
        onDone={() => setPrintingId(null)}
      />
    )}
    </>
  );
}

function OverviewModule({
  activeCount,
  thisMonthCount,
  avgResolveMin,
  overdueList,
  topTypes,
  reconciledCount,
  mismatchCount,
  totalCount,
  onGoActivate,
  onGoActive,
  onOpenIncident,
}) {
  return (
    <div style={styles.moduleWrap}>
      <div style={styles.moduleHeaderRow}>
        <div>
          <div style={styles.moduleTitle}>Tổng quan</div>
          <div style={styles.moduleSub}>Tình hình sự cố hệ thống soát vé hôm nay</div>
        </div>
        <button className="btn" style={styles.activateBtn} onClick={onGoActivate}>
          <Plus size={16} style={{ marginRight: 6 }} />
          Kích hoạt sự cố
        </button>
      </div>

      <div style={styles.kpiGrid}>
        <KpiCard label="Đang xử lý" value={activeCount} color="#F5A623" icon={AlertTriangle} onClick={onGoActive} />
        <KpiCard label="Sự cố tháng này" value={thisMonthCount} color="#4F8EF7" icon={CalendarRange} />
        <KpiCard
          label="Thời gian xử lý TB"
          value={avgResolveMin != null ? `${avgResolveMin} phút` : "—"}
          color="#3DD68C"
          icon={Clock}
        />
        <KpiCard
          label="Đã đối soát / Lệch"
          value={`${reconciledCount} / ${mismatchCount}`}
          color={mismatchCount > 0 ? "#E5484D" : "#3DD68C"}
          icon={ScanLine}
        />
      </div>

      {overdueList.length > 0 && (
        <div style={styles.slaBox}>
          <div style={styles.slaHeader}>
            <AlertCircle size={16} color="#E5484D" />
            <span>Cảnh báo SLA — {overdueList.length} sự cố xử lý quá 30 phút chưa cập nhật</span>
          </div>
          <div style={styles.slaList}>
            {overdueList.map((i) => (
              <div key={i.id} className="navbtn" style={styles.slaRow} onClick={() => onOpenIncident(i.id)}>
                <span style={styles.idTag}>{i.id}</span>
                <span style={{ flex: 1 }}>{i.type} — {i.location}</span>
                <span style={{ color: "#E5484D", fontFamily: "'IBM Plex Mono', monospace", fontSize: 11.5 }}>
                  {Math.round((Date.now() - i.createdAt) / 60000)} phút
                </span>
              </div>
            ))}
          </div>
        </div>
      )}

      <div style={styles.overviewGrid}>
        <div style={styles.panelCard}>
          <div style={styles.panelTitle}>
            <PieChart size={14} style={{ marginRight: 6, verticalAlign: -2 }} />
            Loại sự cố hay gặp nhất
          </div>
          {topTypes.length === 0 ? (
            <div style={styles.emptyStateSmall}>Chưa có dữ liệu.</div>
          ) : (
            topTypes.map(([type, count]) => (
              <div key={type} style={styles.barRow}>
                <span style={styles.barLabel}>{type}</span>
                <div style={styles.barTrack}>
                  <div style={{ ...styles.barFill, width: `${(count / topTypes[0][1]) * 100}%` }} />
                </div>
                <span style={styles.barValue}>{count}</span>
              </div>
            ))
          )}
        </div>

        <div style={styles.panelCard}>
          <div style={styles.panelTitle}>
            <TrendingUp size={14} style={{ marginRight: 6, verticalAlign: -2 }} />
            Tổng số sự cố đã ghi nhận
          </div>
          <div style={styles.bigNumber}>{totalCount}</div>
          <div style={styles.moduleSub}>từ trước đến nay, trên toàn bộ hệ thống</div>
        </div>
      </div>
    </div>
  );
}

function KpiCard({ label, value, color, icon: Icon, onClick }) {
  return (
    <div className={onClick ? "navbtn" : ""} style={{ ...styles.kpiCard, borderTopColor: color }} onClick={onClick}>
      <div style={styles.kpiIconRow}>
        <Icon size={16} color={color} />
      </div>
      <div style={styles.kpiValue}>{value}</div>
      <div style={styles.kpiLabel}>{label}</div>
    </div>
  );
}

function ActivateModule({ form, setForm, toggleDept, onActivate, onGoToParticipate }) {
  const [justActivated, setJustActivated] = useState(null);
  const [copied, setCopied] = useState(false);

  const buildCopyText = (inc) => {
    const deptNames = inc.depts.map((d) => DEPARTMENTS.find((x) => x.id === d)?.label).join(", ");
    return (
      `[SỰ CỐ ${inc.id}] ${inc.type}\n` +
      `Vị trí: ${inc.location}\n` +
      `Mô tả: ${inc.description || "—"}\n` +
      `Bộ phận phối hợp: ${deptNames}\n` +
      `Người báo cáo: ${inc.reporter}\n` +
      `Thời gian xảy ra: ${fmtTime(inc.occurredAt || inc.createdAt)}\n` +
      `→ Đề nghị các bộ phận ra hiện trường phối hợp kiểm tra, soát vé qua cổng thủ công.`
    );
  };

  const handleActivate = async () => {
    const result = await onActivate();
    if (result) {
      setJustActivated(result);
      setCopied(false);
    }
  };

  const handleCopy = () => {
    if (!justActivated) return;
    navigator.clipboard?.writeText(buildCopyText(justActivated));
    setCopied(true);
    setTimeout(() => setCopied(false), 1800);
  };

  if (justActivated) {
    const deptNames = justActivated.depts.map((d) => DEPARTMENTS.find((x) => x.id === d)?.label).join(", ");
    return (
      <div style={styles.moduleWrap}>
        <div style={styles.moduleHeaderRow}>
          <div>
            <div style={styles.moduleTitle}>Đã kích hoạt sự cố</div>
            <div style={styles.moduleSub}>Sao chép nội dung dưới đây để gửi ngay vào group Zalo/Viber/Team</div>
          </div>
        </div>

        <div style={styles.formCard}>
          <div style={styles.sectionLabel}>
            <CheckCircle2 size={13} color="#3DD68C" style={{ marginRight: 5, verticalAlign: -2 }} />
            {justActivated.id} — {justActivated.type}
          </div>
          <div style={styles.breakdownReadTable}>
            <div style={styles.breakdownReadRow}><span>Vị trí</span><b>{justActivated.location}</b></div>
            <div style={styles.breakdownReadRow}>
              <span>Thời gian xảy ra</span>
              <b>{fmtTime(justActivated.occurredAt || justActivated.createdAt)}</b>
            </div>
            <div style={styles.breakdownReadRow}><span>Bộ phận phối hợp</span><b>{deptNames}</b></div>
            <div style={styles.breakdownReadRow}><span>Người báo cáo</span><b>{justActivated.reporter}</b></div>
          </div>

          <div style={styles.formActions}>
            <button className="btn" style={styles.confirmBtn} onClick={handleCopy}>
              {copied ? <Check size={15} style={{ marginRight: 6 }} /> : <Copy size={15} style={{ marginRight: 6 }} />}
              {copied ? "Đã sao chép" : "Sao chép nội dung gửi Zalo/Viber/Team"}
            </button>
            <button className="btn" style={styles.cancelBtn} onClick={() => setJustActivated(null)}>
              <Plus size={14} style={{ marginRight: 5 }} />
              Kích hoạt sự cố khác
            </button>
            <button className="btn" style={styles.utilBtn} onClick={onGoToParticipate}>
              <Users size={14} style={{ marginRight: 5 }} />
              Đi tới Xác nhận tham gia
            </button>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div style={styles.moduleWrap}>
      <div style={styles.moduleHeaderRow}>
        <div>
          <div style={styles.moduleTitle}>Kích hoạt sự cố mới</div>
          <div style={styles.moduleSub}>Điền thông tin để thông báo ngay cho các bộ phận phối hợp</div>
        </div>
      </div>

      <div style={styles.formCard}>
        <div style={styles.formGrid}>
          <div>
            <label style={styles.label}>Loại sự cố</label>
            <select
              style={styles.select}
              value={form.type}
              onChange={(e) => setForm((f) => ({ ...f, type: e.target.value }))}
            >
              {INCIDENT_TYPES.map((t) => (
                <option key={t} value={t}>
                  {t}
                </option>
              ))}
            </select>
            {form.type === "Khác" && (
              <input
                style={{ ...styles.input, marginTop: 8 }}
                placeholder="Mô tả chi tiết loại sự cố *"
                value={form.otherDetail}
                onChange={(e) => setForm((f) => ({ ...f, otherDetail: e.target.value }))}
              />
            )}
          </div>
          <div>
            <label style={styles.label}>Vị trí / cổng *</label>
            <input
              style={styles.input}
              placeholder="VD: Cổng soát vé số 2 - khu A"
              value={form.location}
              onChange={(e) => setForm((f) => ({ ...f, location: e.target.value }))}
            />
          </div>
          <div>
            <label style={styles.label}>Người báo cáo</label>
            <input
              style={styles.input}
              placeholder="Tên / chức danh"
              value={form.reporter}
              onChange={(e) => setForm((f) => ({ ...f, reporter: e.target.value }))}
            />
          </div>
          <div>
            <label style={styles.label}>Thời gian xảy ra sự cố *</label>
            <input
              type="datetime-local"
              style={styles.input}
              value={form.occurredAt}
              onChange={(e) => setForm((f) => ({ ...f, occurredAt: e.target.value }))}
            />
          </div>
        </div>
        <label style={styles.label}>Mô tả nhanh</label>
        <textarea
          style={styles.textarea}
          placeholder="VD: Mất điện toàn bộ khu cổng soát vé, khách không quét được vé"
          value={form.description}
          onChange={(e) => setForm((f) => ({ ...f, description: e.target.value }))}
        />
        <label style={styles.label}>Bộ phận cần phối hợp *</label>
        <div style={styles.chipRow}>
          {DEPARTMENTS.map((d) => (
            <div
              key={d.id}
              className="chip"
              onClick={() => toggleDept(d.id)}
              style={{
                ...styles.chip,
                ...(form.depts.includes(d.id) ? styles.chipActive : {}),
              }}
            >
              {d.label}
            </div>
          ))}
        </div>
        <div style={styles.formActions}>
          <button className="btn" style={styles.confirmBtn} onClick={handleActivate}>
            <Send size={15} style={{ marginRight: 6 }} />
            Kích hoạt & thông báo
          </button>
        </div>
      </div>

    </div>
  );
}

function ReportsModule({ all, topTypes, deptCounts, avgResolveMin, reconciledCount, mismatchCount, onExport }) {
  const [fromDate, setFromDate] = useState("");
  const [toDate, setToDate] = useState("");
  const maxDeptCount = Math.max(1, ...Object.values(deptCounts));

  const monthlyTrend = useMemo(() => {
    const months = [];
    const now = new Date();
    for (let i = 5; i >= 0; i--) {
      const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
      months.push({ key: `${d.getFullYear()}-${d.getMonth()}`, label: `Th${d.getMonth() + 1}/${d.getFullYear()}`, count: 0 });
    }
    all.forEach((inc) => {
      const d = new Date(inc.createdAt);
      const key = `${d.getFullYear()}-${d.getMonth()}`;
      const m = months.find((x) => x.key === key);
      if (m) m.count += 1;
    });
    return months;
  }, [all]);
  const maxMonthCount = Math.max(1, ...monthlyTrend.map((m) => m.count));

  return (
    <div style={styles.moduleWrap}>
      <div style={styles.moduleHeaderRow}>
        <div>
          <div style={styles.moduleTitle}>Báo cáo</div>
          <div style={styles.moduleSub}>Thống kê tổng hợp và xuất báo cáo theo khoảng thời gian</div>
        </div>
      </div>

      <div style={styles.panelCard}>
        <div style={styles.panelTitle}>Xuất báo cáo Excel</div>
        <div style={styles.reconHint}>
          Chọn khoảng thời gian (để trống nếu muốn xuất toàn bộ) rồi tải file Excel tổng hợp tất cả sự cố:
          loại, vị trí, bộ phận phối hợp, số vé biên bản/đối soát, trạng thái.
        </div>
        <div style={styles.exportFilterRow}>
          <div>
            <label style={styles.label}>Từ ngày</label>
            <input type="date" style={styles.input} value={fromDate} onChange={(e) => setFromDate(e.target.value)} />
          </div>
          <div>
            <label style={styles.label}>Đến ngày</label>
            <input type="date" style={styles.input} value={toDate} onChange={(e) => setToDate(e.target.value)} />
          </div>
          <button
            className="btn"
            style={{ ...styles.exportBtn, marginTop: 22 }}
            onClick={() => onExport(fromDate, toDate)}
          >
            <Download size={14} style={{ marginRight: 6 }} />
            Xuất Excel
          </button>
        </div>
      </div>

      <div style={styles.panelCard}>
        <div style={styles.panelTitle}>
          <TrendingUp size={14} style={{ marginRight: 6, verticalAlign: -2 }} />
          Xu hướng sự cố theo tháng (6 tháng gần nhất)
        </div>
        <div style={styles.monthTrendRow}>
          {monthlyTrend.map((m) => (
            <div key={m.key} style={styles.monthBarWrap}>
              <div style={styles.monthBarValue}>{m.count}</div>
              <div style={styles.monthBarTrack}>
                <div
                  style={{
                    ...styles.monthBarFill,
                    height: `${(m.count / maxMonthCount) * 100}%`,
                  }}
                />
              </div>
              <div style={styles.monthBarLabel}>{m.label}</div>
            </div>
          ))}
        </div>
      </div>

      <div style={styles.overviewGrid}>
        <div style={styles.panelCard}>
          <div style={styles.panelTitle}>
            <PieChart size={14} style={{ marginRight: 6, verticalAlign: -2 }} />
            Loại sự cố hay gặp nhất
          </div>
          {topTypes.length === 0 ? (
            <div style={styles.emptyStateSmall}>Chưa có dữ liệu.</div>
          ) : (
            topTypes.map(([type, count]) => (
              <div key={type} style={styles.barRow}>
                <span style={styles.barLabel}>{type}</span>
                <div style={styles.barTrack}>
                  <div style={{ ...styles.barFill, width: `${(count / topTypes[0][1]) * 100}%` }} />
                </div>
                <span style={styles.barValue}>{count}</span>
              </div>
            ))
          )}
        </div>

        <div style={styles.panelCard}>
          <div style={styles.panelTitle}>
            <Users size={14} style={{ marginRight: 6, verticalAlign: -2 }} />
            Số lần tham gia theo bộ phận
          </div>
          {Object.keys(deptCounts).length === 0 ? (
            <div style={styles.emptyStateSmall}>Chưa có dữ liệu.</div>
          ) : (
            DEPARTMENTS.filter((d) => deptCounts[d.id]).map((d) => (
              <div key={d.id} style={styles.barRow}>
                <span style={styles.barLabel}>{d.label}</span>
                <div style={styles.barTrack}>
                  <div style={{ ...styles.barFill, width: `${(deptCounts[d.id] / maxDeptCount) * 100}%` }} />
                </div>
                <span style={styles.barValue}>{deptCounts[d.id]}</span>
              </div>
            ))
          )}
        </div>
      </div>

      <div style={styles.overviewGrid}>
        <KpiCard label="Tổng sự cố" value={all.length} color="#4F8EF7" icon={AlertTriangle} />
        <KpiCard label="Thời gian xử lý TB" value={avgResolveMin != null ? `${avgResolveMin} phút` : "—"} color="#3DD68C" icon={Clock} />
        <KpiCard label="Đã đối soát" value={reconciledCount} color="#3DD68C" icon={ScanLine} />
        <KpiCard label="Có chênh lệch" value={mismatchCount} color={mismatchCount > 0 ? "#E5484D" : "#3DD68C"} icon={AlertCircle} />
      </div>
    </div>
  );
}

function IncidentCard({
  inc,
  section,
  expanded,
  onToggle,
  onAck,
  onSaveReport,
  onSend,
  onCopy,
  copied,
  onSetGateBreakdown,
  onPrint,
  onNavigate,
  onDelete,
  onSendHandover,
  onUpdateHandoverData,
  onConfirmSupplement,
  onUploadSignedReport,
}) {
  const meta = statusMeta(inc.status);
  const [ackInputs, setAckInputs] = useState({});
  const setAckField = (deptId, field, value) =>
    setAckInputs((prev) => ({ ...prev, [deptId]: { ...prev[deptId], [field]: value } }));
  const [reportOpen, setReportOpen] = useState(false);
  const [confirmDelete, setConfirmDelete] = useState(false);
  const [handoverBy, setHandoverBy] = useState("");
  const [editingHandover, setEditingHandover] = useState(false);
  const [manualScanBusy, setManualScanBusy] = useState(false);
  const [manualScanErr, setManualScanErr] = useState("");
  const [manualScanParsed, setManualScanParsed] = useState(null); // { records, fileName } chờ chuyển giao
  const [supplementBy, setSupplementBy] = useState("");
  const [supplementRole, setSupplementRole] = useState("");
  const [supplementNote, setSupplementNote] = useState("");
  const [report, setReport] = useState({
    note: "",
    relatedGuest: "",
    breakdown: [{ type: "", qty: "", saleCode: "", mediaCode: "" }],
  });

  const pendingDepts = inc.depts.filter((d) => !inc.acks.some((a) => a.dept === d));

  const STATUS_DOT = {
    participate: inc.acks.length >= inc.depts.length && inc.depts.length > 0,
    report: !!inc.report,
    handover: inc.status === "da_bs" || inc.status === "da_gui_kt",
    recon: !!inc.gateBreakdown,
    closing: inc.status === "da_gui_kt",
  };

  return (
    <div style={{ ...styles.card, borderLeftColor: meta.color }} className="card-enter">
      <div style={styles.cardHeader} onClick={onToggle}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
          <span style={{ ...styles.statusPill, background: meta.color + "22", color: meta.color, borderColor: meta.color + "55" }}>
            {meta.label}
          </span>
          <span style={{ ...styles.stepTabDot, background: STATUS_DOT[section] ? "#3DD68C" : "#5C6572" }} />
          <span style={styles.idTag}>{inc.id}</span>
          <span style={styles.cardType}>{inc.type}</span>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
          <span style={styles.metaText}>
            <MapPin size={13} style={{ marginRight: 4, verticalAlign: -2 }} />
            {inc.location}
          </span>
          <span style={styles.metaText}>
            <Clock size={13} style={{ marginRight: 4, verticalAlign: -2 }} />
            {fmtTime(inc.occurredAt || inc.createdAt)}
          </span>
          <span
            className="navbtn"
            style={styles.deleteIconBtn}
            onClick={(e) => {
              e.stopPropagation();
              setConfirmDelete(true);
            }}
            title="Xóa sự cố (nhập sai thông tin)"
          >
            <Trash2 size={15} color="#8B93A1" />
          </span>
          {expanded ? <ChevronUp size={18} color="#8B93A1" /> : <ChevronDown size={18} color="#8B93A1" />}
        </div>
      </div>

      {confirmDelete && (
        <div style={styles.deleteConfirmBar} onClick={(e) => e.stopPropagation()}>
          <AlertCircle size={14} color="#E5484D" style={{ marginRight: 6, verticalAlign: -2 }} />
          Xóa hẳn sự cố <b>{inc.id}</b>? Toàn bộ biên bản và dữ liệu đối soát đi kèm sẽ mất, không thể khôi phục.
          <div style={{ display: "flex", gap: 8, marginTop: 8 }}>
            <button className="btn" style={styles.deleteConfirmBtn} onClick={() => onDelete(inc.id)}>
              <Trash2 size={13} style={{ marginRight: 5 }} />
              Xóa sự cố
            </button>
            <button className="btn" style={styles.cancelBtn} onClick={() => setConfirmDelete(false)}>
              Hủy
            </button>
          </div>
        </div>
      )}

      {expanded && (
        <div style={styles.cardBody}>
          {inc.description && <div style={styles.descText}>{inc.description}</div>}

          {section === "participate" && (
            <div style={styles.tabPanel}>
              <div style={styles.sectionLabel}>
                <Users size={13} style={{ marginRight: 5, verticalAlign: -2 }} />
                Bộ phận phối hợp
              </div>
              <div style={styles.reconHint}>
                Mỗi bộ phận tự xác nhận độc lập, không cần chờ bộ phận khác — có thể xác nhận cùng lúc theo bất kỳ thứ tự nào.
              </div>
              <div style={styles.deptGrid}>
                {inc.depts.map((d) => {
                  const ack = inc.acks.find((a) => a.dept === d);
                  const dep = DEPARTMENTS.find((x) => x.id === d);
                  const input = ackInputs[d] || { name: "", role: "" };
                  return (
                    <div key={d} style={{ ...styles.deptCard, ...(ack ? styles.deptCardDone : {}) }}>
                      <div style={styles.deptRow}>
                        <div>
                          <div style={styles.deptName}>{dep?.label}</div>
                          {ack && (
                            <div style={styles.deptAckInfo}>
                              đã tiếp nhận · {ack.by}{ack.role ? ` (${ack.role})` : ""} · {fmtTime(ack.time)}
                            </div>
                          )}
                        </div>
                        {ack ? (
                          <CheckCircle2 size={18} color="#3DD68C" />
                        ) : (
                          <span style={styles.waitingTag}>chờ xác nhận</span>
                        )}
                      </div>

                      {!ack && (
                        <div style={styles.ackFormInline}>
                          <input
                            style={styles.inputSmall}
                            placeholder="Tên người tiếp nhận"
                            value={input.name}
                            onChange={(e) => setAckField(d, "name", e.target.value)}
                          />
                          <input
                            style={styles.inputSmall}
                            placeholder="Chức vụ (VD: Tổ trưởng)"
                            value={input.role}
                            onChange={(e) => setAckField(d, "role", e.target.value)}
                          />
                          <button
                            className="btn"
                            style={styles.ackBtn}
                            onClick={() => {
                              onAck(inc.id, d, input.name, input.role);
                              setAckInputs((prev) => ({ ...prev, [d]: { name: "", role: "" } }));
                            }}
                          >
                            Xác nhận tiếp nhận
                          </button>
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>

              {pendingDepts.length === 0 && (
                <div style={styles.tabDoneHint}>
                  <CheckCircle2 size={14} color="#3DD68C" style={{ marginRight: 6, verticalAlign: -2 }} />
                  Tất cả bộ phận đã xác nhận tham gia. Chuyển sang mục "Biên bản sự cố" ở sidebar để tiếp tục.
                </div>
              )}
            </div>
          )}

          {section === "report" && (
            <div style={styles.tabPanel}>
              <div style={styles.sectionLabel}>Biên bản kiểm tra tại cổng</div>
              {inc.report && !reportOpen ? (
                <div style={styles.reportSummary}>
                  <div style={{ marginBottom: 4 }}>Nhân sự tham gia (theo xác nhận tiếp nhận):</div>
                  <div style={styles.breakdownReadTable}>
                    {inc.acks.length === 0 ? (
                      <div style={styles.breakdownReadRow}><span>Chưa có bộ phận nào xác nhận tiếp nhận</span></div>
                    ) : (
                      inc.acks.map((a, i) => (
                        <div key={i} style={styles.breakdownReadRow}>
                          <span>
                            {a.by}{a.role ? ` (${a.role})` : ""} —{" "}
                            {DEPARTMENTS.find((x) => x.id === a.dept)?.label}
                          </span>
                        </div>
                      ))
                    )}
                  </div>
                  <div style={{ marginTop: 10 }}>Số vé soát tay theo loại vé:</div>
                  <div style={styles.breakdownReadTableWide}>
                    <div style={styles.breakdownReadHeaderRow}>
                      <span style={{ flex: 2 }}>Loại vé</span>
                      <span style={{ flex: 1, textAlign: "right" }}>SL</span>
                      <span style={{ flex: 1.4 }}>SaleCode</span>
                      <span style={{ flex: 1.4 }}>MediaCode</span>
                    </div>
                    {(inc.report.breakdown || []).filter((r) => r.type).map((r, i) => (
                      <div key={i} style={styles.breakdownReadRowWide}>
                        <span style={{ flex: 2 }}>{r.type}</span>
                        <span style={{ flex: 1, textAlign: "right", fontWeight: 700 }}>{r.qty || 0}</span>
                        <span style={{ flex: 1.4, color: "#8B93A1" }}>{r.saleCode || "—"}</span>
                        <span style={{ flex: 1.4, color: "#8B93A1" }}>{r.mediaCode || "—"}</span>
                      </div>
                    ))}
                    <div style={styles.breakdownReadRowTotal}>
                      <span>Tổng cộng</span>
                      <b>{(inc.report.breakdown || []).reduce((s, r) => s + (Number(r.qty) || 0), 0)}</b>
                    </div>
                  </div>
                  {inc.report.note && <div style={{ marginTop: 10 }}>Ghi chú: {inc.report.note}</div>}
                  <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
                    <button className="btn" style={styles.editLink} onClick={() => { setReport(inc.report); setReportOpen(true); }}>
                      Sửa biên bản
                    </button>
                  </div>

                  <div style={styles.reconStep}>
                    <div style={styles.reconStepTitle}>Xuất biên bản để in & ký</div>
                    <div style={styles.reconHint}>
                      Xuất biên bản dạng PDF để in, cho các bên liên quan ký tay, sau đó chụp/scan lại và tải bản
                      đã ký lên đây để lưu hồ sơ.
                    </div>
                    <button className="btn" style={styles.utilBtn} onClick={onPrint}>
                      <Printer size={14} style={{ marginRight: 5 }} />
                      Xuất biên bản PDF để in
                    </button>

                    <div style={{ marginTop: 12 }}>
                      <SignedReportUpload inc={inc} onUpload={onUploadSignedReport} />
                    </div>
                  </div>
                </div>
              ) : (
                <>
                  <div style={styles.reconHint}>
                    Nhân sự liên quan đến biên bản lấy tự động từ danh sách các bộ phận đã "Xác nhận tiếp nhận"
                    ở mục "Xác nhận tham gia" — không cần nhập lại tên ở đây.
                  </div>

                  <label style={styles.label}>Khách hàng liên quan (nếu có, tuỳ chọn)</label>
                  <input
                    style={styles.input}
                    placeholder="Họ tên khách + SĐT, nếu sự cố gắn với khiếu nại của khách cụ thể"
                    value={report.relatedGuest}
                    onChange={(e) => setReport((r) => ({ ...r, relatedGuest: e.target.value }))}
                  />

                  <label style={styles.label}>Số vé soát tay theo loại vé</label>
                  <div style={styles.breakdownEditor}>
                    {report.breakdown.map((row, i) => (
                      <div key={i} style={styles.breakdownEditRowWide}>
                        <input
                          style={{ ...styles.input, flex: 2 }}
                          placeholder="Loại vé (VD: NL-VÉ CÁP TREO SWH)"
                          value={row.type}
                          onChange={(e) => {
                            const val = e.target.value;
                            setReport((r) => {
                              const breakdown = [...r.breakdown];
                              breakdown[i] = { ...breakdown[i], type: val };
                              return { ...r, breakdown };
                            });
                          }}
                        />
                        <input
                          style={{ ...styles.input, flex: 0.7 }}
                          placeholder="Số lượng"
                          value={row.qty}
                          onChange={(e) => {
                            const val = e.target.value;
                            setReport((r) => {
                              const breakdown = [...r.breakdown];
                              breakdown[i] = { ...breakdown[i], qty: val };
                              return { ...r, breakdown };
                            });
                          }}
                        />
                        <input
                          style={{ ...styles.input, flex: 1.3 }}
                          placeholder="SaleCode (VD: BIDUOGJF)"
                          value={row.saleCode}
                          onChange={(e) => {
                            const val = e.target.value;
                            setReport((r) => {
                              const breakdown = [...r.breakdown];
                              breakdown[i] = { ...breakdown[i], saleCode: val };
                              return { ...r, breakdown };
                            });
                          }}
                        />
                        <input
                          style={{ ...styles.input, flex: 1.3 }}
                          placeholder="MediaCode (VD: HGNLL4KB39NM28)"
                          value={row.mediaCode}
                          onChange={(e) => {
                            const val = e.target.value;
                            setReport((r) => {
                              const breakdown = [...r.breakdown];
                              breakdown[i] = { ...breakdown[i], mediaCode: val };
                              return { ...r, breakdown };
                            });
                          }}
                        />
                        <button
                          className="btn"
                          style={styles.removeRowBtn}
                          onClick={() =>
                            setReport((r) => ({
                              ...r,
                              breakdown: r.breakdown.filter((_, idx) => idx !== i),
                            }))
                          }
                        >
                          <X size={14} />
                        </button>
                      </div>
                    ))}
                    <button
                      className="btn"
                      style={styles.addRowBtn}
                      onClick={() =>
                        setReport((r) => ({
                          ...r,
                          breakdown: [...r.breakdown, { type: "", qty: "", saleCode: "", mediaCode: "" }],
                        }))
                      }
                    >
                      <Plus size={13} style={{ marginRight: 4 }} />
                      Thêm loại vé
                    </button>
                  </div>

                  <label style={styles.label}>Ghi chú hiện trường</label>
                  <textarea
                    style={styles.textarea}
                    placeholder="Tình trạng xử lý, thời gian khắc phục, bất thường ghi nhận..."
                    value={report.note}
                    onChange={(e) => setReport((r) => ({ ...r, note: e.target.value }))}
                  />

                  <button
                    className="btn"
                    style={styles.confirmBtn}
                    onClick={() => {
                      onSaveReport(inc.id, {
                        ...report,
                        breakdown: report.breakdown.filter((r) => r.type.trim()),
                      });
                      setReportOpen(false);
                      if (onNavigate) onNavigate("handover");
                    }}
                  >
                    Lưu biên bản
                  </button>
                </>
              )}
            </div>
          )}

          {section === "handover" && (
            <div style={styles.tabPanel}>
              <div style={styles.sectionLabel}>
                <Cpu size={13} style={{ marginRight: 5, verticalAlign: -2 }} />
                Bàn giao BCNTT xử lý bổ sung
              </div>

              {!inc.report ? (
                <div style={styles.reconHint}>Cần lập biên bản trước khi bàn giao cho BCNTT.</div>
              ) : inc.status === "da_lap_bb" ? (
                <>
                  <div style={styles.reconHint}>
                    Biên bản đã lập xong. Tải lên dữ liệu soát vé tay (theo Media Code) rồi chuyển thông tin cho
                    BCNTT để soát vé bổ sung ngay khi hệ thống có điện/mạng trở lại.
                  </div>
                  <div style={styles.breakdownReadTable}>
                    <div style={styles.breakdownReadRow}>
                      <span>Tổng vé soát tay theo biên bản</span>
                      <b>{(inc.report.breakdown || []).reduce((s, r) => s + (Number(r.qty) || 0), 0)}</b>
                    </div>
                  </div>

                  <label style={{ ...styles.label, marginTop: 12 }}>
                    Dữ liệu soát vé tay (file "Báo cáo quyền lợi" — cột Media Code)
                  </label>
                  <UploadSlot
                    label="File soát vé tay theo Media Code"
                    fileName={manualScanParsed?.fileName}
                    count={manualScanParsed?.records.length}
                    countLabel="mã đã soát tay"
                    busy={manualScanBusy}
                    onFile={async (file) => {
                      if (!file) return;
                      setManualScanErr("");
                      setManualScanBusy(true);
                      try {
                        const records = await parseManualScanFile(file);
                        if (records.length === 0) {
                          setManualScanErr('Không tìm thấy cột "Media Code" hoặc file rỗng.');
                        } else {
                          setManualScanParsed({ records, fileName: file.name });
                        }
                      } catch (e) {
                        setManualScanErr(e.message || "Không đọc được file.");
                      } finally {
                        setManualScanBusy(false);
                      }
                    }}
                  />
                  {manualScanErr && (
                    <div style={styles.reconError}>
                      <AlertCircle size={13} style={{ marginRight: 5, verticalAlign: -2 }} />
                      {manualScanErr}
                    </div>
                  )}

                  <div style={{ ...styles.ackFormInline, marginTop: 14 }}>
                    <input
                      style={styles.inputSmall}
                      placeholder="Tên người chuyển giao (tuỳ chọn)"
                      value={handoverBy}
                      onChange={(e) => setHandoverBy(e.target.value)}
                    />
                    <button
                      className="btn"
                      style={{ ...styles.confirmBtn, ...(manualScanParsed ? {} : styles.btnDisabled) }}
                      disabled={!manualScanParsed}
                      onClick={() =>
                        onSendHandover(
                          inc.id,
                          handoverBy,
                          manualScanParsed.records,
                          { fileName: manualScanParsed.fileName }
                        )
                      }
                    >
                      <Send size={14} style={{ marginRight: 6 }} />
                      Chuyển thông tin cho BCNTT
                    </button>
                  </div>
                  {!manualScanParsed && (
                    <div style={{ ...styles.reconHint, marginTop: 6 }}>
                      Cần tải lên dữ liệu soát vé tay trước khi chuyển cho BCNTT.
                    </div>
                  )}
                </>
              ) : inc.status === "cho_bs" ? (
                <>
                  <div style={styles.reconHint}>
                    Đã chuyển cho BCNTT lúc <b>{fmtTime(inc.handover?.time)}</b>
                    {inc.handover?.by && inc.handover.by !== "—" ? ` bởi ${inc.handover.by}` : ""} — kèm{" "}
                    <b>{(inc.manualScanData || []).length}</b> mã Media Code đã soát tay
                    {inc.manualScanMeta?.fileName ? ` (file ${inc.manualScanMeta.fileName})` : ""}. Đang chờ BCNTT
                    soát vé bổ sung khi hệ thống có điện/mạng và xác nhận hoàn tất bên dưới.
                  </div>

                  {editingHandover ? (
                    <HandoverFileEditor
                      inc={inc}
                      onCancel={() => setEditingHandover(false)}
                      onSave={(records, fileName) => {
                        onUpdateHandoverData(inc.id, records, { fileName }, handoverBy);
                        setEditingHandover(false);
                      }}
                    />
                  ) : (
                    <button className="btn" style={styles.utilBtn} onClick={() => setEditingHandover(true)}>
                      <FileSpreadsheet size={14} style={{ marginRight: 5 }} />
                      Sửa dữ liệu bàn giao (lỡ tải nhầm file)
                    </button>
                  )}

                  <div style={styles.divider} />

                  <label style={styles.label}>Người xác nhận (BCNTT)</label>
                  <input
                    style={styles.input}
                    placeholder="Tên"
                    value={supplementBy}
                    onChange={(e) => setSupplementBy(e.target.value)}
                  />
                  <label style={styles.label}>Chức vụ (tuỳ chọn)</label>
                  <input
                    style={styles.input}
                    placeholder="VD: Kỹ thuật viên"
                    value={supplementRole}
                    onChange={(e) => setSupplementRole(e.target.value)}
                  />
                  <label style={styles.label}>Ghi chú xử lý bổ sung</label>
                  <textarea
                    style={styles.textarea}
                    placeholder="VD: Đã soát bù xong lúc 14h20 cùng ngày, không phát sinh bất thường"
                    value={supplementNote}
                    onChange={(e) => setSupplementNote(e.target.value)}
                  />
                  <button
                    className="btn"
                    style={styles.confirmBtn}
                    onClick={() => {
                      onConfirmSupplement(inc.id, supplementBy, supplementRole, supplementNote);
                      setSupplementBy("");
                      setSupplementRole("");
                      setSupplementNote("");
                    }}
                  >
                    <CheckCircle2 size={14} style={{ marginRight: 6 }} />
                    Xác nhận hoàn tất xử lý bổ sung
                  </button>
                </>
              ) : (
                <div style={styles.tabDoneHint}>
                  <CheckCircle2 size={14} color="#3DD68C" style={{ marginRight: 6, verticalAlign: -2 }} />
                  BCNTT đã xác nhận hoàn tất xử lý bổ sung lúc {fmtTime(inc.supplement?.time)}
                  {inc.supplement?.by && inc.supplement.by !== "—" ? ` bởi ${inc.supplement.by}` : ""}
                  {inc.supplement?.role ? ` (${inc.supplement.role})` : ""}.
                  {inc.supplement?.note && <div style={{ marginTop: 6 }}>Ghi chú: {inc.supplement.note}</div>}
                  <div style={{ marginTop: 6 }}>
                    Kèm <b>{(inc.manualScanData || []).length}</b> mã Media Code đã soát tay
                    {inc.manualScanMeta?.fileName ? ` (file ${inc.manualScanMeta.fileName})` : ""}.
                  </div>
                  <div style={{ marginTop: 6 }}>Chuyển sang mục "Đối soát dữ liệu" để kế toán đối soát.</div>

                  {editingHandover ? (
                    <div style={{ marginTop: 10 }}>
                      <HandoverFileEditor
                        inc={inc}
                        onCancel={() => setEditingHandover(false)}
                        onSave={(records, fileName) => {
                          onUpdateHandoverData(inc.id, records, { fileName }, handoverBy);
                          setEditingHandover(false);
                        }}
                      />
                    </div>
                  ) : (
                    <button
                      className="btn"
                      style={{ ...styles.utilBtn, marginTop: 10 }}
                      onClick={() => setEditingHandover(true)}
                    >
                      <FileSpreadsheet size={14} style={{ marginRight: 5 }} />
                      Sửa dữ liệu bàn giao (lỡ tải nhầm file)
                    </button>
                  )}
                </div>
              )}
            </div>
          )}

          {section === "recon" && (
            <div style={styles.tabPanel}>
              {inc.status !== "da_bs" && inc.status !== "da_gui_kt" ? (
                <div style={styles.reconBox}>
                  <div style={styles.sectionLabel}>
                    <ScanLine size={13} style={{ marginRight: 5, verticalAlign: -2 }} />
                    Đối soát dữ liệu vé (kế toán)
                  </div>
                  <div style={styles.reconHint}>
                    Chưa thể đối soát — cần BCNTT xác nhận hoàn tất xử lý bổ sung ở mục "Bàn giao BCNTT" trước.
                  </div>
                </div>
              ) : (
                <TicketReconciliation inc={inc} onSetGateBreakdown={onSetGateBreakdown} />
              )}
            </div>
          )}

          {section === "closing" && (
            <div style={styles.tabPanel}>
              <div style={styles.actionRow}>
                <button className="btn" style={styles.utilBtn} onClick={() => onCopy(inc)}>
                  {copied ? <Check size={14} style={{ marginRight: 5 }} /> : <Copy size={14} style={{ marginRight: 5 }} />}
                  {copied ? "Đã sao chép" : "Sao chép nội dung gửi Zalo/Viber/Team"}
                </button>

                {inc.report && (
                  <button className="btn" style={styles.utilBtn} onClick={onPrint}>
                    <Printer size={14} style={{ marginRight: 5 }} />
                    Xuất biên bản PDF
                  </button>
                )}

                {inc.status === "da_bs" && reconciliationStatus(inc).done && (
                  <button className="btn" style={styles.sendBtn} onClick={() => onSend(inc.id)}>
                    <Send size={14} style={{ marginRight: 5 }} />
                    Gửi kế toán
                  </button>
                )}
              </div>

              {inc.status === "da_bs" && !reconciliationStatus(inc).done && (
                <div style={styles.blockedHint}>
                  <AlertCircle size={14} color="#E5484D" style={{ marginRight: 6, verticalAlign: -2 }} />
                  Chưa thể gửi kế toán — {reconciliationStatus(inc).reason}
                  <div style={{ marginTop: 6 }}>
                    Quay lại mục "Đối soát dữ liệu" để tải đúng file và chạy lại cho đến khi khớp hết số liệu.
                  </div>
                </div>
              )}

              {inc.status === "da_gui_kt" && (
                <div style={styles.tabDoneHint}>
                  <CheckCircle2 size={14} color="#3DD68C" style={{ marginRight: 6, verticalAlign: -2 }} />
                  Sự cố đã hoàn tất — đã gửi kế toán.
                </div>
              )}

              <IncidentDetailReport inc={inc} />

              <div style={{ ...styles.sectionLabel, marginTop: 16 }}>Lịch sử xử lý</div>
              <div style={styles.timeline}>
                {inc.history.map((h, i) => (
                  <div key={i} style={styles.timelineRow}>
                    <span style={{ ...styles.timelineDot, background: statusMeta(h.status).color }} />
                    <span style={styles.timelineLabel}>{statusMeta(h.status).label}</span>
                    <span style={styles.timelineTime}>{fmtTime(h.time)}</span>
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

function IncidentDetailReport({ inc }) {
  const [open, setOpen] = useState({ report: false, manual: false, recon: false });
  const toggle = (key) => setOpen((o) => ({ ...o, [key]: !o[key] }));

  const breakdown = (inc.report?.breakdown || []).filter((r) => r.type);
  const totalManual = breakdown.reduce((s, r) => s + (Number(r.qty) || 0), 0);
  const manualScan = inc.manualScanData || [];
  const matchedList = inc.mediaCodeResult?.matched || [];
  const missingList = inc.mediaCodeResult?.missing || [];

  return (
    <div style={{ ...styles.reconBox, marginTop: 16 }}>
      <div style={styles.sectionLabel}>
        <FileSpreadsheet size={13} style={{ marginRight: 5, verticalAlign: -2 }} />
        Báo cáo chi tiết sự cố
      </div>
      <div style={styles.reconHint}>
        Xem lại toàn bộ dữ liệu đã ghi nhận cho sự cố này: biên bản, danh mục soát vé do bộ phận soát vé tải lên,
        và danh mục đối soát khớp do kế toán xác nhận.
      </div>

      {/* Biên bản */}
      <div style={styles.detailBlock}>
        <div className="navbtn" style={styles.detailBlockHeader} onClick={() => toggle("report")}>
          <span>
            <ClipboardList size={14} style={{ marginRight: 6, verticalAlign: -2 }} />
            Biên bản sự cố {inc.report ? `— ${totalManual} vé` : "— chưa lập"}
          </span>
          {open.report ? <ChevronUp size={16} color="#8B93A1" /> : <ChevronDown size={16} color="#8B93A1" />}
        </div>
        {open.report && (
          <div style={styles.detailBlockBody}>
            {!inc.report ? (
              <div style={styles.emptyStateSmall}>Sự cố chưa lập biên bản.</div>
            ) : (
              <>
                {breakdown.length === 0 ? (
                  <div style={styles.emptyStateSmall}>Chưa có dữ liệu loại vé.</div>
                ) : (
                  <div style={styles.compareTable}>
                    <div style={{ ...styles.compareRow, ...styles.compareHeaderRow }}>
                      <span style={{ flex: 2 }}>Loại vé</span>
                      <span style={{ flex: 1, textAlign: "right" }}>SL</span>
                      <span style={{ flex: 1.4 }}>SaleCode</span>
                      <span style={{ flex: 1.4 }}>MediaCode</span>
                    </div>
                    {breakdown.map((r, i) => (
                      <div key={i} style={styles.compareRow}>
                        <span style={{ flex: 2 }}>{r.type}</span>
                        <span style={{ flex: 1, textAlign: "right" }}>{r.qty}</span>
                        <span style={{ flex: 1.4, color: "#8B93A1" }}>{r.saleCode || "—"}</span>
                        <span style={{ flex: 1.4, color: "#8B93A1" }}>{r.mediaCode || "—"}</span>
                      </div>
                    ))}
                  </div>
                )}
                {inc.report.note && <div style={{ marginTop: 8, fontSize: 12.5 }}>Ghi chú: {inc.report.note}</div>}
              </>
            )}
          </div>
        )}
      </div>

      {/* Danh mục soát vé (soát vé up) */}
      <div style={styles.detailBlock}>
        <div className="navbtn" style={styles.detailBlockHeader} onClick={() => toggle("manual")}>
          <span>
            <UploadCloud size={14} style={{ marginRight: 6, verticalAlign: -2 }} />
            Danh mục soát vé (bộ phận soát vé tải lên) — {manualScan.length} mã
          </span>
          {open.manual ? <ChevronUp size={16} color="#8B93A1" /> : <ChevronDown size={16} color="#8B93A1" />}
        </div>
        {open.manual && (
          <div style={styles.detailBlockBody}>
            {manualScan.length === 0 ? (
              <div style={styles.emptyStateSmall}>Chưa có dữ liệu soát vé tay.</div>
            ) : (
              <>
                {inc.manualScanMeta?.fileName && (
                  <div style={{ fontSize: 11.5, color: "#5C6572", marginBottom: 6 }}>
                    Nguồn file: {inc.manualScanMeta.fileName}
                  </div>
                )}
                <div style={styles.compareTable}>
                  <div style={{ ...styles.compareRow, ...styles.compareHeaderRow }}>
                    <span style={{ flex: 2 }}>Media Code</span>
                    <span style={{ flex: 1.5 }}>Trạng thái</span>
                    <span style={{ flex: 2 }}>Cổng/Vị trí</span>
                    <span style={{ flex: 1.5 }}>Người quét</span>
                  </div>
                  {manualScan.slice(0, 30).map((m, i) => (
                    <div key={i} style={styles.compareRow}>
                      <span style={{ flex: 2, fontFamily: "'IBM Plex Mono', monospace" }}>{m.mediaCode}</span>
                      <span style={{ flex: 1.5 }}>{m.status || "—"}</span>
                      <span style={{ flex: 2 }}>{m.gate || "—"}</span>
                      <span style={{ flex: 1.5 }}>{m.scannedBy || "—"}</span>
                    </div>
                  ))}
                  {manualScan.length > 30 && (
                    <div style={{ fontSize: 11.5, color: "#5C6572", marginTop: 6 }}>
                      + {manualScan.length - 30} mã khác — xem đầy đủ trong file xuất Excel ở mục Đối soát dữ liệu
                    </div>
                  )}
                </div>
              </>
            )}
          </div>
        )}
      </div>

      {/* Danh mục đối soát khớp (kế toán up) */}
      <div style={styles.detailBlock}>
        <div className="navbtn" style={styles.detailBlockHeader} onClick={() => toggle("recon")}>
          <span>
            <ScanLine size={14} style={{ marginRight: 6, verticalAlign: -2 }} />
            Danh mục đối soát (kế toán tải lên) — {matchedList.length} khớp
            {missingList.length > 0 ? `, ${missingList.length} thiếu` : ""}
          </span>
          {open.recon ? <ChevronUp size={16} color="#8B93A1" /> : <ChevronDown size={16} color="#8B93A1" />}
        </div>
        {open.recon && (
          <div style={styles.detailBlockBody}>
            {!inc.gateBreakdown ? (
              <div style={styles.emptyStateSmall}>Chưa có dữ liệu đối soát từ kế toán.</div>
            ) : (
              <>
                {inc.gateMeta?.fileName && (
                  <div style={{ fontSize: 11.5, color: "#5C6572", marginBottom: 6 }}>
                    Nguồn file: {inc.gateMeta.fileName}
                  </div>
                )}
                {matchedList.length > 0 ? (
                  <div style={styles.compareTable}>
                    <div style={{ ...styles.compareRow, ...styles.compareHeaderRow }}>
                      <span style={{ flex: 2 }}>Media Code (đã khớp)</span>
                      <span style={{ flex: 1.5 }}>Trạng thái</span>
                      <span style={{ flex: 2 }}>Cổng/Vị trí</span>
                    </div>
                    {matchedList.slice(0, 30).map((m, i) => (
                      <div key={i} style={styles.compareRow}>
                        <span style={{ flex: 2, fontFamily: "'IBM Plex Mono', monospace" }}>{m.mediaCode}</span>
                        <span style={{ flex: 1.5 }}>{m.status || "—"}</span>
                        <span style={{ flex: 2 }}>{m.gate || "—"}</span>
                      </div>
                    ))}
                    {matchedList.length > 30 && (
                      <div style={{ fontSize: 11.5, color: "#5C6572", marginTop: 6 }}>
                        + {matchedList.length - 30} mã khác — xem đầy đủ trong file xuất Excel
                      </div>
                    )}
                  </div>
                ) : (
                  <div style={styles.compareTable}>
                    <div style={{ ...styles.compareRow, ...styles.compareHeaderRow }}>
                      <span style={{ flex: 2 }}>Loại vé</span>
                      <span style={{ flex: 1, textAlign: "right" }}>SL (Access Log)</span>
                    </div>
                    {inc.gateBreakdown.map((r, i) => (
                      <div key={i} style={styles.compareRow}>
                        <span style={{ flex: 2 }}>{r.type}</span>
                        <span style={{ flex: 1, textAlign: "right" }}>{r.qty}</span>
                      </div>
                    ))}
                  </div>
                )}
              </>
            )}
          </div>
        )}
      </div>
    </div>
  );
}

function TicketReconciliation({ inc, onSetGateBreakdown }) {
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState("");
  const [pendingUpload, setPendingUpload] = useState(null); // { records, breakdown, meta } vừa tải, chưa chạy đối soát
  const [running, setRunning] = useState(false);

  const hasManualScan = (inc.manualScanData || []).length > 0;
  const savedBreakdown = inc.gateBreakdown || null;
  const manualBreakdown = (inc.report?.breakdown || []).filter((r) => r.type && Number(r.qty) >= 0);
  const mediaResult = inc.mediaCodeResult || null;

  // Chỉ dùng để so loại vé/số lượng khi sự cố cũ không có dữ liệu Media Code
  const typeResult = useMemo(() => {
    if (hasManualScan) return null;
    if (!savedBreakdown || manualBreakdown.length === 0) return null;
    return reconcileByType(manualBreakdown, savedBreakdown);
  }, [hasManualScan, savedBreakdown, manualBreakdown]);

  const handleFile = async (file) => {
    if (!file) return;
    setErr("");
    setBusy(true);
    try {
      const records = await parseAccessLogFile(file);
      const { breakdown, totalRowsUsed, totalRows } = aggregateGateRecords(records);
      if (breakdown.length === 0) {
        setErr('Không tìm thấy dòng nào có UsageTypeDesc = "Entry" trong file.');
        setBusy(false);
        return;
      }
      setPendingUpload({
        records,
        breakdown,
        meta: { fileName: file.name, totalRowsUsed, totalRows },
      });
    } catch (e) {
      setErr(e.message || "Không đọc được file Access Log.");
    } finally {
      setBusy(false);
    }
  };

  const runReconciliation = () => {
    if (!pendingUpload) return;
    setRunning(true);
    const media = hasManualScan ? reconcileByMediaCode(inc.manualScanData, pendingUpload.records) : null;
    onSetGateBreakdown(inc.id, pendingUpload.breakdown, pendingUpload.meta, media);
    setPendingUpload(null);
    setTimeout(() => setRunning(false), 300);
  };

  if (!inc.report) {
    return (
      <div style={styles.reconBox}>
        <div style={styles.sectionLabel}>
          <ScanLine size={13} style={{ marginRight: 5, verticalAlign: -2 }} />
          Đối soát dữ liệu vé (kế toán)
        </div>
        <div style={styles.reconHint}>Cần lập biên bản (số lượng vé theo loại) trước khi đối soát.</div>
      </div>
    );
  }

  const canRun = !!pendingUpload || !!savedBreakdown;
  const result = hasManualScan ? mediaResult : typeResult;

  return (
    <div style={styles.reconBox}>
      <div style={styles.sectionLabel}>
        <ScanLine size={13} style={{ marginRight: 5, verticalAlign: -2 }} />
        Đối soát dữ liệu vé (kế toán)
      </div>

      {hasManualScan && (
        <div style={styles.reconHint}>
          Đối soát theo <b>Media Code</b>: so khớp {inc.manualScanData.length} mã đã soát tay
          {inc.manualScanMeta?.fileName ? ` (từ file ${inc.manualScanMeta.fileName})` : ""} với dữ liệu Access Log
          kế toán tải lên bên dưới.
        </div>
      )}
      {!hasManualScan && (
        <div style={styles.reconHint}>
          Sự cố này chưa có dữ liệu soát vé tay theo Media Code (thiếu ở bước "Bàn giao BCNTT") — tạm đối soát
          theo số lượng/loại vé.
        </div>
      )}

      {/* Bước 1 — Tải dữ liệu */}
      <div style={styles.reconStep}>
        <div style={styles.reconStepTitle}>Bước 1 — Tải dữ liệu Access Log</div>
        <div style={styles.reconHint}>
          Tải file Access Log vé soát bù (đã xuất riêng cho sự cố này). App sẽ đọc cột MediaCode để đối soát,
          đồng thời cộng dồn GroupQuantity theo loại vé (chỉ tính dòng "Entry") để tham khảo.
        </div>
        <UploadSlot
          label="File Access Log vé soát bù"
          fileName={pendingUpload?.meta.fileName || (savedBreakdown ? inc.gateMeta?.fileName : null)}
          count={
            pendingUpload
              ? pendingUpload.breakdown.reduce((s, r) => s + r.qty, 0)
              : savedBreakdown
              ? savedBreakdown.reduce((s, r) => s + r.qty, 0)
              : null
          }
          countLabel="vé (Entry)"
          busy={busy}
          onFile={handleFile}
        />
        {err && (
          <div style={styles.reconError}>
            <AlertCircle size={13} style={{ marginRight: 5, verticalAlign: -2 }} />
            {err}
          </div>
        )}
        {pendingUpload && (
          <div style={styles.reconPreviewConfirm}>
            <div style={{ fontSize: 12.5, color: "#C4CAD3" }}>
              Đã đọc <b>{pendingUpload.meta.totalRowsUsed}</b> dòng "Entry" / {pendingUpload.meta.totalRows} dòng
              dữ liệu, gồm <b>{pendingUpload.breakdown.length}</b> loại vé. Sẵn sàng để chạy đối soát ở bước dưới.
            </div>
          </div>
        )}
      </div>

      {/* Bước 2 — Chạy đối soát */}
      <div style={styles.reconStep}>
        <div style={styles.reconStepTitle}>Bước 2 — Chạy đối soát</div>
        <div style={styles.reconHint}>
          {hasManualScan
            ? "So khớp từng Media Code đã soát tay với dữ liệu Access Log vừa tải."
            : "So khớp số lượng theo từng loại vé giữa biên bản thủ công và dữ liệu Access Log vừa tải."}
        </div>
        <button
          className="btn"
          style={{ ...styles.confirmBtn, ...(canRun ? {} : styles.btnDisabled) }}
          disabled={!canRun}
          onClick={runReconciliation}
        >
          <ScanLine size={14} style={{ marginRight: 6 }} />
          {running ? "Đang chạy đối soát…" : pendingUpload ? "Chạy đối soát với dữ liệu vừa tải" : "Chạy lại đối soát"}
        </button>
        {!pendingUpload && savedBreakdown && (
          <div style={{ ...styles.reconHint, marginTop: 8 }}>
            Đã chạy đối soát lần gần nhất với file <b>{inc.gateMeta?.fileName || "—"}</b>. Tải file mới ở Bước 1
            nếu cần chạy lại.
          </div>
        )}
      </div>

      {/* Bước 3 — Kết quả đối soát */}
      <div style={styles.reconStep}>
        <div style={styles.reconStepTitle}>Bước 3 — Kết quả đối soát</div>
        {!result ? (
          <div style={styles.reconHint}>Chưa có kết quả — hoàn tất Bước 1 và Bước 2 trước.</div>
        ) : hasManualScan ? (
          <>
            <div style={styles.reconSummaryRow}>
              <ReconChip label="Khớp (đã quẹt bù)" count={result.matched.length} color="#3DD68C" />
              <ReconChip label="Thiếu (chưa quẹt bù)" count={result.missing.length} color="#E5484D" />
            </div>

            {result.missing.length > 0 ? (
              <div style={styles.blockedHint}>
                <AlertCircle size={14} color="#E5484D" style={{ marginRight: 6, verticalAlign: -2 }} />
                Dữ liệu chưa khớp — còn <b>{result.missing.length}</b> mã Media Code chưa thấy quẹt bù. Yêu cầu kế
                toán kiểm tra lại file Access Log (đúng khoảng thời gian/đã quẹt bù đầy đủ chưa) rồi tải lại và
                chạy lại đối soát ở Bước 1–2 cho đến khi khớp hết mới được gửi kế toán.
              </div>
            ) : (
              <div style={styles.tabDoneHint}>
                <CheckCircle2 size={14} color="#3DD68C" style={{ marginRight: 6, verticalAlign: -2 }} />
                Dữ liệu đã khớp hoàn toàn — có thể chuyển sang mục "Hoàn tất & Đóng" để gửi kế toán.
              </div>
            )}

            {result.missing.length > 0 && (
              <MediaCodeMissingTable missing={result.missing} />
            )}

            <button
              className="btn"
              style={styles.exportBtn}
              onClick={() => downloadMediaCodeReconciliationExcel(inc, result)}
            >
              <Download size={14} style={{ marginRight: 6 }} />
              Xuất Excel kết quả đối soát
            </button>
          </>
        ) : (
          <>
            <div style={styles.reconSummaryRow}>
              <ReconChip label="Khớp" count={result.matched.length} color="#3DD68C" />
              <ReconChip label="Lệch số lượng" count={result.mismatch.length} color="#F5A623" />
              <ReconChip label="Thiếu (chưa quẹt bù)" count={result.missing.length} color="#E5484D" />
              <ReconChip label="Thừa (ngoài biên bản)" count={result.extra.length} color="#4F8EF7" />
            </div>

            {result.mismatch.length + result.missing.length + result.extra.length > 0 ? (
              <div style={styles.blockedHint}>
                <AlertCircle size={14} color="#E5484D" style={{ marginRight: 6, verticalAlign: -2 }} />
                Dữ liệu chưa khớp — yêu cầu kế toán kiểm tra lại và chạy lại đối soát cho đến khi khớp hết mới
                được gửi kế toán.
              </div>
            ) : (
              <div style={styles.tabDoneHint}>
                <CheckCircle2 size={14} color="#3DD68C" style={{ marginRight: 6, verticalAlign: -2 }} />
                Dữ liệu đã khớp hoàn toàn — có thể chuyển sang mục "Hoàn tất & Đóng" để gửi kế toán.
              </div>
            )}

            <TypeCompareTable manualBreakdown={manualBreakdown} gateBreakdown={savedBreakdown} result={result} />

            <button
              className="btn"
              style={styles.exportBtn}
              onClick={() => downloadReconciliationExcel(inc, result, inc.gateMeta)}
            >
              <Download size={14} style={{ marginRight: 6 }} />
              Xuất Excel kết quả đối soát
            </button>
          </>
        )}
      </div>
    </div>
  );
}

function MediaCodeMissingTable({ missing }) {
  const shown = missing.slice(0, 8);
  return (
    <div style={styles.compareTable}>
      <div style={{ ...styles.compareRow, ...styles.compareHeaderRow }}>
        <span style={{ flex: 2 }}>Media Code</span>
        <span style={{ flex: 2 }}>Cổng/Vị trí</span>
        <span style={{ flex: 2 }}>Người quét</span>
      </div>
      {shown.map((m, i) => (
        <div key={i} style={styles.compareRow}>
          <span style={{ flex: 2, fontFamily: "'IBM Plex Mono', monospace" }}>{m.mediaCode}</span>
          <span style={{ flex: 2 }}>{m.gate || "—"}</span>
          <span style={{ flex: 2 }}>{m.scannedBy || "—"}</span>
        </div>
      ))}
      {missing.length > shown.length && (
        <div style={{ fontSize: 11.5, color: "#5C6572", marginTop: 6 }}>
          + {missing.length - shown.length} mã khác — xem đầy đủ trong file xuất
        </div>
      )}
    </div>
  );
}

function TypeCompareTable({ result }) {
  const rows = [
    ...result.matched.map((x) => ({ ...x, tag: "Khớp", color: "#3DD68C" })),
    ...result.mismatch.map((x) => ({ ...x, tag: "Lệch", color: "#F5A623" })),
    ...result.missing.map((x) => ({ ...x, tag: "Thiếu", color: "#E5484D", gateQty: "—" })),
    ...result.extra.map((x) => ({ ...x, tag: "Thừa", color: "#4F8EF7", manualQty: "—" })),
  ];
  return (
    <div style={styles.compareTable}>
      <div style={{ ...styles.compareRow, ...styles.compareHeaderRow }}>
        <span style={{ flex: 2 }}>Loại vé</span>
        <span style={{ flex: 1, textAlign: "right" }}>SL biên bản</span>
        <span style={{ flex: 1, textAlign: "right" }}>SL Access Log</span>
        <span style={{ flex: 1, textAlign: "right" }}>Trạng thái</span>
      </div>
      {rows.map((r, i) => (
        <div key={i} style={styles.compareRow}>
          <span style={{ flex: 2 }}>{r.type}</span>
          <span style={{ flex: 1, textAlign: "right", fontFamily: "'IBM Plex Mono', monospace" }}>
            {r.manualQty ?? "—"}
          </span>
          <span style={{ flex: 1, textAlign: "right", fontFamily: "'IBM Plex Mono', monospace" }}>
            {r.gateQty ?? "—"}
          </span>
          <span style={{ flex: 1, textAlign: "right", color: r.color, fontWeight: 700, fontSize: 11.5 }}>
            {r.tag}
          </span>
        </div>
      ))}
    </div>
  );
}

const MAX_SIGNED_FILE_BYTES = 900_000; // ~900KB gốc — dữ liệu này nằm chung 1 kho lưu trữ với tất cả sự cố khác

function readFileAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(new Error("Không đọc được file."));
    reader.readAsDataURL(file);
  });
}

function HandoverFileEditor({ inc, onCancel, onSave }) {
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState("");
  const [parsed, setParsed] = useState(null);

  const handleFile = async (file) => {
    if (!file) return;
    setErr("");
    setBusy(true);
    try {
      const records = await parseManualScanFile(file);
      if (records.length === 0) {
        setErr('Không tìm thấy cột "Media Code" hoặc file rỗng.');
      } else {
        setParsed({ records, fileName: file.name });
      }
    } catch (e) {
      setErr(e.message || "Không đọc được file.");
    } finally {
      setBusy(false);
    }
  };

  return (
    <div style={{ ...styles.reconPreviewConfirm, marginTop: 10 }}>
      <div style={{ fontSize: 12.5, color: "#C4CAD3", marginBottom: 8 }}>
        Dữ liệu hiện tại: <b>{(inc.manualScanData || []).length}</b> mã
        {inc.manualScanMeta?.fileName ? ` (file ${inc.manualScanMeta.fileName})` : ""}. Tải lên file đúng để thay
        thế toàn bộ.
      </div>
      <UploadSlot
        label="File soát vé tay theo Media Code (thay thế)"
        fileName={parsed?.fileName}
        count={parsed?.records.length}
        countLabel="mã đã soát tay"
        busy={busy}
        onFile={handleFile}
      />
      {err && (
        <div style={styles.reconError}>
          <AlertCircle size={13} style={{ marginRight: 5, verticalAlign: -2 }} />
          {err}
        </div>
      )}
      <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
        <button
          className="btn"
          style={{ ...styles.confirmBtn, ...(parsed ? {} : styles.btnDisabled) }}
          disabled={!parsed}
          onClick={() => onSave(parsed.records, parsed.fileName)}
        >
          Lưu thay thế dữ liệu
        </button>
        <button className="btn" style={styles.cancelBtn} onClick={onCancel}>
          Hủy
        </button>
      </div>
    </div>
  );
}

function SignedReportUpload({ inc, onUpload }) {
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState("");

  const handleFile = async (file) => {
    if (!file) return;
    setErr("");
    if (file.size > MAX_SIGNED_FILE_BYTES) {
      setErr(
        `File quá lớn (${(file.size / 1_000_000).toFixed(1)}MB). Vui lòng chụp/scan nén dưới ` +
          `${(MAX_SIGNED_FILE_BYTES / 1_000_000).toFixed(1)}MB rồi tải lại.`
      );
      return;
    }
    setBusy(true);
    try {
      const dataUrl = await readFileAsDataUrl(file);
      onUpload(inc.id, {
        fileName: file.name,
        mimeType: file.type,
        dataUrl,
        uploadedAt: Date.now(),
      });
    } catch (e) {
      setErr(e.message || "Không đọc được file.");
    } finally {
      setBusy(false);
    }
  };

  const signed = inc.signedReport;

  return (
    <div>
      <label htmlFor={`signed-upload-${inc.id}`} style={styles.uploadSlot}>
        <input
          id={`signed-upload-${inc.id}`}
          type="file"
          accept="image/*,.pdf"
          style={{ display: "none" }}
          onChange={(e) => handleFile(e.target.files?.[0])}
        />
        <div style={styles.uploadIcon}>
          {signed ? <FileSpreadsheet size={16} color="#3DD68C" /> : <UploadCloud size={16} color="#8B93A1" />}
        </div>
        <div style={{ flex: 1, minWidth: 0 }}>
          <div style={styles.uploadLabel}>Biên bản đã ký (ảnh chụp / scan / PDF)</div>
          <div style={styles.uploadStatus}>
            {busy
              ? "Đang tải lên…"
              : signed
              ? `${signed.fileName} · đã tải lúc ${fmtTime(signed.uploadedAt)}`
              : "Chưa có file — bấm để tải lên"}
          </div>
        </div>
      </label>
      {err && (
        <div style={styles.reconError}>
          <AlertCircle size={13} style={{ marginRight: 5, verticalAlign: -2 }} />
          {err}
        </div>
      )}
      {signed && (
        <a
          href={signed.dataUrl}
          target="_blank"
          rel="noreferrer"
          style={{ ...styles.editLink, display: "inline-block", marginTop: 8, textDecoration: "none" }}
        >
          Xem / tải lại bản đã ký
        </a>
      )}
    </div>
  );
}

function UploadSlot({ label, fileName, count, countLabel, busy, onFile }) {
  const inputId = "upload-" + label.replace(/\s+/g, "-").slice(0, 12) + Math.random().toString(36).slice(2, 6);
  return (
    <label htmlFor={inputId} style={styles.uploadSlot}>
      <input
        id={inputId}
        type="file"
        accept=".xlsx,.xls,.csv"
        style={{ display: "none" }}
        onChange={(e) => onFile(e.target.files?.[0])}
      />
      <div style={styles.uploadIcon}>
        {fileName ? <FileSpreadsheet size={16} color="#3DD68C" /> : <UploadCloud size={16} color="#8B93A1" />}
      </div>
      <div style={{ flex: 1, minWidth: 0 }}>
        <div style={styles.uploadLabel}>{label}</div>
        <div style={styles.uploadStatus}>
          {busy
            ? "Đang đọc file…"
            : fileName
            ? `${fileName} · ${count} ${countLabel || ""}`
            : "Chưa có file — bấm để tải lên"}
        </div>
      </div>
    </label>
  );
}

function ReconChip({ label, count, color }) {
  return (
    <div style={{ ...styles.reconChip, borderColor: color + "55", background: color + "15" }}>
      <span style={{ color, fontWeight: 700, fontFamily: "'IBM Plex Mono', monospace" }}>{count}</span>
      <span style={{ color: "#C4CAD3", marginLeft: 6, fontSize: 12 }}>{label}</span>
    </div>
  );
}

function PrintableReport({ inc, onDone }) {
  useEffect(() => {
    if (!inc) {
      onDone();
      return;
    }
    const t = setTimeout(() => window.print(), 150);
    const handleAfterPrint = () => onDone();
    window.addEventListener("afterprint", handleAfterPrint);
    return () => {
      clearTimeout(t);
      window.removeEventListener("afterprint", handleAfterPrint);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [inc]);

  if (!inc) return null;

  const manualBreakdown = (inc.report?.breakdown || []).filter((r) => r.type);
  const hasRecon = !!inc.gateBreakdown;
  // Ưu tiên số liệu đã được kế toán đối soát (Access Log); nếu chưa đối soát thì tạm dùng số liệu biên bản thủ công
  const ticketBreakdown = hasRecon ? inc.gateBreakdown : manualBreakdown;
  const ticketBreakdownSource = hasRecon
    ? "theo báo cáo đối soát của kế toán"
    : "theo biên bản thủ công — chưa đối soát";
  const totalManual = manualBreakdown.reduce((s, r) => s + (Number(r.qty) || 0), 0);
  const totalGate = hasRecon ? inc.gateBreakdown.reduce((s, r) => s + (Number(r.qty) || 0), 0) : 0;
  const totalTicketBreakdown = ticketBreakdown.reduce((s, r) => s + (Number(r.qty) || 0), 0);

  const closedEntry = inc.history.find((h) => h.status === "da_lap_bb");
  const closedAt = closedEntry?.time || null;

  const d = new Date(inc.occurredAt || inc.createdAt);
  const gio = d.getHours();
  const phut = d.getMinutes();
  const ngay = d.getDate();
  const thang = d.getMonth() + 1;
  const nam = d.getFullYear();

  const attendees = inc.acks.map((a) => ({
    name: a.by,
    dept: DEPARTMENTS.find((x) => x.id === a.dept)?.label || a.dept,
    role: a.role || "",
  }));
  // đảm bảo tối thiểu 4 dòng như mẫu giấy, phần còn trống để điền tay nếu cần
  while (attendees.length < 4) attendees.push({ name: "", dept: "", role: "" });

  const contentLines = [];
  if (inc.description) contentLines.push(inc.description.trim());
  if (inc.report?.note) contentLines.push(inc.report.note.trim());
  const ticketCodes = manualBreakdown.filter((r) => r.saleCode || r.mediaCode);
  if (ticketCodes.length > 0) {
    contentLines.push(
      "Chứng cứ kèm theo — mã vé ghi nhận: " +
        ticketCodes
          .map((r) =>
            [
              r.type,
              r.saleCode && `SaleCode ${r.saleCode}`,
              r.mediaCode && `MediaCode ${r.mediaCode}`,
            ]
              .filter(Boolean)
              .join(" / ")
          )
          .join("; ") +
        "."
    );
  }
  if (hasRecon) {
    contentLines.push(
      `Đối soát với Access Log quẹt bù: tổng ${totalGate} vé` +
        (inc.gateMeta?.fileName ? ` (nguồn: ${inc.gateMeta.fileName})` : "") +
        `, chênh lệch ${totalGate - totalManual} vé so với biên bản.`
    );
  }
  contentLines.push("Vậy chúng tôi tiến hành lập biên bản sự việc để lưu chứng từ.");

  return (
    <div className="print-container" style={pstyles.page}>
      <div style={pstyles.topRow}>
        <div>Ban hành lần: 01</div>
        <div>ST.SVE.01/B.03</div>
      </div>

      <div style={pstyles.headerRow}>
        <div style={pstyles.headerLeft}>
          <div style={pstyles.logoText}>SUNWORLD</div>
          <div style={pstyles.logoSub}>BÀ NÀ HILLS</div>
        </div>
        <div style={pstyles.headerRight}>
          <div style={pstyles.companyLine}>CÔNG TY CPDV CÁP TREO BÀ NÀ</div>
          <div style={pstyles.companyLine}>BỘ PHẬN SOÁT VÉ CÁP TREO</div>
        </div>
      </div>

      <div style={pstyles.formTitle}>BIÊN BẢN SỰ VIỆC</div>

      <div style={pstyles.row}>
        <span style={pstyles.fieldLabelItalic}>V/v:</span>
        <span style={pstyles.fieldValue}>{inc.type}{inc.description ? ` — ${inc.location}` : ""}</span>
      </div>

      <div style={pstyles.row}>
        <span style={pstyles.fieldLabelItalic}>Hôm nay, vào lúc</span>
        <span style={pstyles.fieldValueShort}>{gio}</span>
        <span style={pstyles.fieldLabelItalic}>giờ</span>
        <span style={pstyles.fieldValueShort}>{phut}</span>
        <span style={pstyles.fieldLabelItalic}>phút, ngày</span>
        <span style={pstyles.fieldValueShort}>{ngay}</span>
        <span style={pstyles.fieldLabelItalic}>tháng</span>
        <span style={pstyles.fieldValueShort}>{thang}</span>
        <span style={pstyles.fieldLabelItalic}>năm</span>
        <span style={pstyles.fieldValueShort}>{nam}</span>
        <span style={pstyles.fieldLabelItalic}>, tại</span>
        <span style={pstyles.fieldValue}>{inc.location}</span>
      </div>

      {closedAt && (
        <div style={pstyles.row}>
          <span style={pstyles.fieldLabelItalic}>Đóng sự cố lúc:</span>
          <span style={pstyles.fieldValue}>{fmtTime(closedAt)}</span>
          <span style={{ ...pstyles.fieldLabelItalic, fontSize: 10.5, fontStyle: "normal" }}>
            (ngay khi hoàn thành lập biên bản)
          </span>
        </div>
      )}

      <div style={{ ...pstyles.fieldLabelItalic, marginTop: 8 }}>Chúng tôi gồm có:</div>
      {attendees.map((a, i) => (
        <div key={i} style={pstyles.row}>
          <span style={pstyles.fieldLabelItalic}>Ông/Bà:</span>
          <span style={{ ...pstyles.fieldValue, flex: 2 }}>{a.name}</span>
          <span style={pstyles.fieldLabelItalic}>Bộ phận công tác:</span>
          <span style={{ ...pstyles.fieldValue, flex: 2 }}>{a.dept}</span>
          <span style={pstyles.fieldLabelItalic}>Chức vụ:</span>
          <span style={{ ...pstyles.fieldValue, flex: 1 }}>{a.role}</span>
        </div>
      ))}

      <div style={{ ...pstyles.fieldLabelItalic, marginTop: 8 }}>
        Số lượng vé theo tên vé ({ticketBreakdownSource}):
      </div>
      <table style={pstyles.table}>
        <thead>
          <tr>
            <th style={pstyles.th}>Tên vé</th>
            <th style={{ ...pstyles.th, textAlign: "right", width: 90 }}>Số lượng</th>
          </tr>
        </thead>
        <tbody>
          {ticketBreakdown.length === 0 ? (
            <tr>
              <td style={pstyles.td} colSpan={2}>—</td>
            </tr>
          ) : (
            ticketBreakdown.map((r, i) => (
              <tr key={i}>
                <td style={pstyles.td}>{r.type}</td>
                <td style={{ ...pstyles.td, textAlign: "right" }}>{r.qty}</td>
              </tr>
            ))
          )}
          <tr>
            <td style={{ ...pstyles.td, fontWeight: 700 }}>Tổng cộng</td>
            <td style={{ ...pstyles.td, textAlign: "right", fontWeight: 700 }}>{totalTicketBreakdown}</td>
          </tr>
        </tbody>
      </table>

      <div style={{ ...pstyles.fieldLabelItalic, marginTop: 10 }}>
        Tiến hành lập biên bản nội dung như sau:
      </div>
      <div style={pstyles.contentBox}>
        {contentLines.map((line, i) => (
          <div key={i} style={pstyles.contentLine}>
            {line}
          </div>
        ))}
      </div>

      <div style={pstyles.closingLine}>
        Biên bản kết thúc vào lúc …… giờ …… phút, cùng ngày và đọc lại cho các bên liên quan cùng nghe
        và thống nhất kí vào biên bản.
      </div>

      <div style={pstyles.sigGrid}>
        {inc.depts.map((d) => {
          const dep = DEPARTMENTS.find((x) => x.id === d);
          const ack = inc.acks.find((a) => a.dept === d);
          return (
            <div key={d} style={pstyles.sigBlock}>
              <div style={pstyles.sigLabel}>{dep?.label}</div>
              <div style={pstyles.sigSub}>(Ký, ghi rõ họ tên)</div>
              <div style={pstyles.sigNameSpace}>{ack?.by || ""}</div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

const pstyles = {
  page: {
    background: "#fff",
    color: "#111",
    fontFamily: "'Times New Roman', Times, serif",
    fontSize: 12.5,
    lineHeight: 1.6,
    padding: "10mm 8mm",
    maxWidth: "210mm",
    margin: "0 auto",
  },
  topRow: {
    display: "flex",
    justifyContent: "space-between",
    fontSize: 11.5,
    marginBottom: 10,
  },
  headerRow: { display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 4 },
  headerLeft: { width: "35%" },
  logoText: { fontSize: 19, fontWeight: 700, letterSpacing: "0.03em" },
  logoSub: { fontSize: 10, letterSpacing: "0.12em", color: "#333" },
  headerRight: { width: "65%", textAlign: "center" },
  companyLine: { fontWeight: 700, fontSize: 12.5 },
  formTitle: {
    fontSize: 16,
    fontWeight: 700,
    textAlign: "center",
    letterSpacing: "0.04em",
    margin: "14px 0 14px 0",
  },
  row: { display: "flex", flexWrap: "wrap", alignItems: "baseline", gap: 4, marginBottom: 6 },
  fieldLabelItalic: { fontStyle: "italic" },
  fieldValue: {
    borderBottom: "1px dotted #999",
    paddingBottom: 1,
    flex: "1 1 auto",
    minWidth: 40,
  },
  fieldValueShort: {
    borderBottom: "1px dotted #999",
    paddingBottom: 1,
    minWidth: 22,
    textAlign: "center",
  },
  table: { width: "100%", borderCollapse: "collapse", marginTop: 6, marginBottom: 4 },
  th: {
    border: "1px solid #999",
    padding: "4px 8px",
    textAlign: "left",
    fontSize: 11.5,
    background: "#f0f0f0",
  },
  td: { border: "1px solid #999", padding: "4px 8px", fontSize: 12 },
  contentBox: { marginTop: 4, minHeight: "40mm" },
  contentLine: { marginBottom: 6 },
  closingLine: { marginTop: 10, fontSize: 12 },
  sigGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(3, 1fr)",
    gap: "22px 16px",
    marginTop: 26,
  },
  sigBlock: { textAlign: "center" },
  sigLabel: { fontWeight: 700, fontSize: 13 },
  sigSub: { fontStyle: "italic", fontSize: 11, color: "#333", marginTop: 2 },
  sigNameSpace: { marginTop: 50, fontWeight: 600, borderTop: "1px solid #111", paddingTop: 4 },
};

const styles = {
  appShell: {
    minHeight: "100vh",
    display: "flex",
    background: "#0F1319",
    color: "#E8EAED",
    fontFamily: "'Inter', system-ui, sans-serif",
  },
  sidebar: {
    width: 232,
    minWidth: 232,
    height: "100vh",
    position: "sticky",
    top: 0,
    display: "flex",
    flexDirection: "column",
    background: "#12161D",
    borderRight: "1px solid #232B36",
  },
  sidebarHeader: {
    display: "flex",
    alignItems: "center",
    gap: 10,
    padding: "18px 16px",
    borderBottom: "1px solid #1C222B",
  },
  sidebarNav: {
    flex: 1,
    display: "flex",
    flexDirection: "column",
    gap: 2,
    padding: "10px",
    overflowY: "auto",
  },
  navItem: {
    display: "flex",
    alignItems: "center",
    gap: 10,
    padding: "10px 12px",
    borderRadius: 8,
    color: "#8B93A1",
    fontSize: 13.5,
    fontWeight: 600,
    whiteSpace: "nowrap",
  },
  navItemActive: { background: "#1B222C", color: "#E8EAED" },
  navBadge: {
    background: "#E5484D",
    color: "#fff",
    fontSize: 10.5,
    fontWeight: 700,
    borderRadius: 10,
    padding: "1px 7px",
    fontFamily: "'IBM Plex Mono', monospace",
  },
  sidebarFooter: {
    padding: "14px 16px",
    borderTop: "1px solid #1C222B",
    display: "flex",
    flexDirection: "column",
    gap: 4,
  },
  contentArea: { flex: 1, minWidth: 0, padding: "22px 26px 60px 26px" },
  moduleWrap: { display: "flex", flexDirection: "column", gap: 16 },
  moduleHeaderRow: { display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 12 },
  moduleTitle: { fontFamily: "'Space Grotesk', sans-serif", fontSize: 21, fontWeight: 700 },
  moduleSub: { fontSize: 13, color: "#8B93A1", marginTop: 4 },
  kpiGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 14 },
  kpiCard: {
    background: "#151B23",
    border: "1px solid #232B36",
    borderTop: "3px solid",
    borderRadius: 10,
    padding: "14px 16px",
  },
  kpiIconRow: { marginBottom: 8 },
  kpiValue: { fontFamily: "'Space Grotesk', sans-serif", fontSize: 24, fontWeight: 700 },
  kpiLabel: { fontSize: 12.5, color: "#8B93A1", marginTop: 2 },
  slaBox: {
    background: "#2A1416",
    border: "1px solid #E5484D55",
    borderRadius: 10,
    padding: "14px 16px",
  },
  slaHeader: { display: "flex", alignItems: "center", gap: 8, fontWeight: 700, fontSize: 13.5, color: "#FF9B9E" },
  slaList: { marginTop: 10, display: "flex", flexDirection: "column", gap: 6 },
  slaRow: {
    display: "flex",
    alignItems: "center",
    gap: 10,
    padding: "8px 10px",
    background: "#151B23",
    borderRadius: 7,
    fontSize: 12.5,
  },
  overviewGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(260px, 1fr))", gap: 14 },
  panelCard: {
    background: "#151B23",
    border: "1px solid #232B36",
    borderRadius: 10,
    padding: "16px 18px",
  },
  panelTitle: { fontSize: 13, fontWeight: 700, marginBottom: 12, color: "#C4CAD3" },
  emptyStateSmall: { fontSize: 12.5, color: "#5C6572" },
  barRow: { display: "flex", alignItems: "center", gap: 10, marginBottom: 10, fontSize: 12.5 },
  barLabel: { width: "38%", color: "#C4CAD3", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" },
  barTrack: { flex: 1, height: 8, background: "#0F1319", borderRadius: 4, overflow: "hidden" },
  barFill: { height: "100%", background: "#4F8EF7", borderRadius: 4 },
  barValue: { width: 28, textAlign: "right", fontFamily: "'IBM Plex Mono', monospace", color: "#8B93A1" },
  bigNumber: { fontFamily: "'Space Grotesk', sans-serif", fontSize: 40, fontWeight: 700, marginTop: 6 },
  exportFilterRow: { display: "flex", gap: 14, alignItems: "flex-end", flexWrap: "wrap", marginTop: 10 },
  monthTrendRow: {
    display: "flex",
    alignItems: "flex-end",
    gap: 14,
    height: 120,
    marginTop: 6,
    padding: "0 6px",
  },
  monthBarWrap: {
    flex: 1,
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    height: "100%",
    justifyContent: "flex-end",
    gap: 4,
  },
  monthBarValue: { fontSize: 11.5, color: "#8B93A1", fontFamily: "'IBM Plex Mono', monospace" },
  monthBarTrack: {
    width: "60%",
    minWidth: 18,
    height: 70,
    display: "flex",
    alignItems: "flex-end",
    background: "#0F1319",
    borderRadius: 4,
    overflow: "hidden",
  },
  monthBarFill: { width: "100%", background: "#4F8EF7", borderRadius: "4px 4px 0 0", minHeight: 3 },
  monthBarLabel: { fontSize: 10.5, color: "#5C6572" },
  filterBar: {
    display: "flex",
    gap: 8,
    flexWrap: "wrap",
    padding: "0 0 14px 0",
    alignItems: "center",
  },
  filterSelect: {
    padding: "8px 10px",
    borderRadius: 7,
    border: "1px solid #2A3340",
    background: "#151B23",
    color: "#E8EAED",
    fontSize: 12.5,
  },
  filterInput: {
    padding: "8px 10px",
    borderRadius: 7,
    border: "1px solid #2A3340",
    background: "#151B23",
    color: "#E8EAED",
    fontSize: 12.5,
    flex: "1 1 180px",
  },
  filterDate: {
    padding: "8px 10px",
    borderRadius: 7,
    border: "1px solid #2A3340",
    background: "#151B23",
    color: "#E8EAED",
    fontSize: 12.5,
  },
  clearFilterBtn: {
    display: "flex",
    alignItems: "center",
    padding: "8px 12px",
    borderRadius: 7,
    background: "#1B222C",
    border: "1px solid #2A3340",
    color: "#8B93A1",
    fontSize: 12.5,
  },
  page: {
    minHeight: "100vh",
    background: "#0F1319",
    color: "#E8EAED",
    fontFamily: "'Inter', system-ui, sans-serif",
    paddingBottom: 40,
  },
  header: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    padding: "18px 22px",
    borderBottom: "1px solid #232B36",
    background: "#12161D",
  },
  headerLeft: { display: "flex", alignItems: "center", gap: 12 },
  logoMark: {
    width: 38,
    height: 38,
    borderRadius: 10,
    background: "#1B222C",
    border: "1px solid #2A3340",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
  },
  title: {
    fontFamily: "'Space Grotesk', sans-serif",
    fontWeight: 700,
    fontSize: 16,
    letterSpacing: "0.02em",
  },
  subtitle: { fontSize: 12, color: "#8B93A1", marginTop: 2 },
  headerRight: { display: "flex", alignItems: "center", gap: 12 },
  savingTag: { fontSize: 11, color: "#8B93A1", fontFamily: "'IBM Plex Mono', monospace" },
  liveTag: {
    display: "flex",
    alignItems: "center",
    gap: 6,
    fontSize: 11,
    fontFamily: "'IBM Plex Mono', monospace",
    color: "#3DD68C",
    letterSpacing: "0.05em",
  },
  liveDot: {
    width: 7,
    height: 7,
    borderRadius: "50%",
    background: "#3DD68C",
    boxShadow: "0 0 6px #3DD68C",
  },
  tabs: {
    display: "flex",
    alignItems: "center",
    gap: 8,
    padding: "0 0 14px 0",
    borderBottom: "1px solid #1C222B",
    marginBottom: 4,
  },
  tab: {
    display: "flex",
    alignItems: "center",
    padding: "8px 14px",
    borderRadius: 8,
    background: "transparent",
    color: "#8B93A1",
    fontSize: 13,
    fontWeight: 600,
  },
  tabActive: { background: "#1B222C", color: "#E8EAED", border: "1px solid #2A3340" },
  activateBtn: {
    display: "flex",
    alignItems: "center",
    padding: "9px 16px",
    borderRadius: 8,
    background: "#E5484D",
    color: "#fff",
    fontSize: 13,
    fontWeight: 700,
  },
  errorBar: {
    margin: "0 0 12px 0",
    padding: "10px 14px",
    background: "#2A1416",
    border: "1px solid #E5484D55",
    color: "#FF9B9E",
    borderRadius: 8,
    fontSize: 13,
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
  },
  formCard: {
    margin: "0",
    padding: 20,
    background: "#151B23",
    border: "1px solid #232B36",
    borderRadius: 12,
  },
  formTitle: {
    fontFamily: "'Space Grotesk', sans-serif",
    fontSize: 13,
    fontWeight: 700,
    letterSpacing: "0.08em",
    color: "#F5A623",
    marginBottom: 14,
  },
  formGrid: { display: "grid", gridTemplateColumns: "1fr 1.4fr 1fr", gap: 14, marginBottom: 4 },
  formGrid3: { display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 14, marginBottom: 4 },
  label: { display: "block", fontSize: 11.5, color: "#8B93A1", marginBottom: 5, marginTop: 10 },
  input: {
    width: "100%",
    padding: "9px 11px",
    borderRadius: 7,
    border: "1px solid #2A3340",
    background: "#0F1319",
    color: "#E8EAED",
    fontSize: 13.5,
  },
  select: {
    width: "100%",
    padding: "9px 11px",
    borderRadius: 7,
    border: "1px solid #2A3340",
    background: "#0F1319",
    color: "#E8EAED",
    fontSize: 13.5,
  },
  textarea: {
    width: "100%",
    minHeight: 60,
    padding: "9px 11px",
    borderRadius: 7,
    border: "1px solid #2A3340",
    background: "#0F1319",
    color: "#E8EAED",
    fontSize: 13.5,
    fontFamily: "inherit",
    resize: "vertical",
  },
  chipRow: { display: "flex", flexWrap: "wrap", gap: 8, marginTop: 4 },
  chip: {
    padding: "7px 13px",
    borderRadius: 20,
    border: "1px solid #2A3340",
    fontSize: 12.5,
    color: "#8B93A1",
    background: "#0F1319",
  },
  chipActive: { background: "#4F8EF722", borderColor: "#4F8EF7", color: "#8FB6FF" },
  formActions: { display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 18 },
  cancelBtn: {
    padding: "9px 16px",
    borderRadius: 8,
    background: "transparent",
    border: "1px solid #2A3340",
    color: "#8B93A1",
    fontSize: 13,
    fontWeight: 600,
  },
  confirmBtn: {
    display: "flex",
    alignItems: "center",
    padding: "9px 18px",
    borderRadius: 8,
    background: "#4F8EF7",
    color: "#fff",
    fontSize: 13,
    fontWeight: 700,
    marginTop: 14,
  },
  main: { padding: "0", display: "flex", flexDirection: "column", gap: 12 },
  emptyState: {
    padding: "60px 20px",
    textAlign: "center",
    color: "#5C6572",
    fontSize: 14,
    border: "1px dashed #232B36",
    borderRadius: 12,
    marginTop: 12,
  },
  card: {
    background: "#151B23",
    border: "1px solid #232B36",
    borderLeft: "4px solid",
    borderRadius: 10,
    overflow: "hidden",
  },
  deleteIconBtn: {
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    width: 26,
    height: 26,
    borderRadius: 6,
    cursor: "pointer",
  },
  deleteConfirmBar: {
    margin: "0 16px 14px 16px",
    padding: "10px 14px",
    background: "#2A1416",
    border: "1px solid #E5484D55",
    borderRadius: 8,
    fontSize: 12.5,
    color: "#FF9B9E",
  },
  deleteConfirmBtn: {
    display: "flex",
    alignItems: "center",
    padding: "7px 14px",
    borderRadius: 7,
    background: "#E5484D",
    color: "#fff",
    fontSize: 12.5,
    fontWeight: 700,
  },
  cardHeader: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    padding: "14px 16px",
    cursor: "pointer",
    flexWrap: "wrap",
    gap: 8,
  },
  statusPill: {
    fontSize: 11,
    fontWeight: 700,
    padding: "4px 10px",
    borderRadius: 20,
    border: "1px solid",
    fontFamily: "'IBM Plex Mono', monospace",
    letterSpacing: "0.02em",
  },
  idTag: {
    fontFamily: "'IBM Plex Mono', monospace",
    fontSize: 12,
    color: "#5C6572",
  },
  cardType: { fontSize: 13.5, fontWeight: 600 },
  metaText: { fontSize: 12.5, color: "#8B93A1" },
  cardBody: { padding: "0 16px 18px 16px", borderTop: "1px solid #1C222B" },
  descText: { fontSize: 13.5, color: "#C4CAD3", margin: "14px 0", lineHeight: 1.5 },
  sectionLabel: {
    fontSize: 11.5,
    fontWeight: 700,
    color: "#8B93A1",
    letterSpacing: "0.05em",
    marginTop: 16,
    marginBottom: 8,
    textTransform: "uppercase",
  },
  deptGrid: { display: "flex", flexDirection: "column", gap: 8 },
  deptCard: {
    background: "#0F1319",
    border: "1px solid #232B36",
    borderRadius: 8,
    overflow: "hidden",
  },
  deptCardDone: { borderColor: "#3DD68C33" },
  deptRow: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    padding: "9px 12px",
  },
  ackFormInline: {
    display: "flex",
    gap: 8,
    flexWrap: "wrap",
    padding: "0 12px 12px 12px",
    borderTop: "1px solid #1C222B",
    paddingTop: 10,
    marginTop: 2,
  },
  deptRowDone: { borderColor: "#3DD68C33" },
  deptName: { fontSize: 13, fontWeight: 600 },
  deptAckInfo: { fontSize: 11.5, color: "#5C6572", marginTop: 2 },
  waitingTag: {
    fontSize: 11,
    color: "#F5A623",
    fontFamily: "'IBM Plex Mono', monospace",
  },
  ackForm: { display: "flex", gap: 8, marginTop: 10, flexWrap: "wrap" },
  selectSmall: {
    flex: "1 1 160px",
    padding: "8px 10px",
    borderRadius: 7,
    border: "1px solid #2A3340",
    background: "#0F1319",
    color: "#E8EAED",
    fontSize: 13,
  },
  inputSmall: {
    flex: "1 1 160px",
    padding: "8px 10px",
    borderRadius: 7,
    border: "1px solid #2A3340",
    background: "#0F1319",
    color: "#E8EAED",
    fontSize: 13,
  },
  ackBtn: {
    padding: "8px 14px",
    borderRadius: 7,
    background: "#3DD68C",
    color: "#0F1319",
    fontSize: 12.5,
    fontWeight: 700,
  },
  divider: { height: 1, background: "#1C222B", margin: "16px 0" },
  stepTabs: {
    display: "flex",
    gap: 4,
    marginTop: 14,
    marginBottom: 4,
    borderBottom: "1px solid #1C222B",
    overflowX: "auto",
  },
  stepTab: {
    display: "flex",
    alignItems: "center",
    gap: 6,
    padding: "9px 12px",
    fontSize: 12.5,
    fontWeight: 600,
    color: "#8B93A1",
    borderBottom: "2px solid transparent",
    whiteSpace: "nowrap",
  },
  stepTabActive: { color: "#E8EAED", borderBottom: "2px solid #4F8EF7" },
  stepTabDot: { width: 6, height: 6, borderRadius: "50%", flexShrink: 0 },
  stepTabBadge: {
    fontSize: 10.5,
    color: "#5C6572",
    fontFamily: "'IBM Plex Mono', monospace",
    marginLeft: 2,
  },
  tabPanel: { paddingTop: 14 },
  tabDoneHint: {
    marginTop: 10,
    padding: "9px 12px",
    background: "#0F1319",
    border: "1px solid #232B36",
    borderRadius: 7,
    fontSize: 12.5,
    color: "#8B93A1",
  },
  actionRow: { display: "flex", gap: 10, flexWrap: "wrap" },
  utilBtn: {
    display: "flex",
    alignItems: "center",
    padding: "8px 14px",
    borderRadius: 7,
    background: "#1B222C",
    border: "1px solid #2A3340",
    color: "#C4CAD3",
    fontSize: 12.5,
    fontWeight: 600,
  },
  sendBtn: {
    display: "flex",
    alignItems: "center",
    padding: "8px 14px",
    borderRadius: 7,
    background: "#3DD68C",
    color: "#0F1319",
    fontSize: 12.5,
    fontWeight: 700,
  },
  reportBox: {
    marginTop: 14,
    padding: 14,
    background: "#0F1319",
    border: "1px solid #232B36",
    borderRadius: 10,
  },
  reportSummary: { fontSize: 13, color: "#C4CAD3", display: "flex", flexDirection: "column", gap: 4 },
  editLink: {
    marginTop: 8,
    padding: "6px 12px",
    borderRadius: 6,
    background: "transparent",
    border: "1px solid #2A3340",
    color: "#8FB6FF",
    fontSize: 12,
    alignSelf: "flex-start",
  },
  reconBox: {
    marginTop: 16,
    padding: 14,
    background: "#0F1319",
    border: "1px solid #232B36",
    borderRadius: 10,
  },
  reconHint: { fontSize: 12, color: "#8B93A1", marginBottom: 12, lineHeight: 1.5 },
  reconStep: {
    marginTop: 16,
    paddingTop: 16,
    borderTop: "1px solid #1C222B",
  },
  reconStepTitle: {
    fontSize: 12.5,
    fontWeight: 700,
    color: "#4F8EF7",
    marginBottom: 8,
    textTransform: "uppercase",
    letterSpacing: "0.03em",
  },
  btnDisabled: { opacity: 0.4, cursor: "not-allowed" },
  blockedHint: {
    marginTop: 10,
    padding: "10px 14px",
    background: "#2A1416",
    border: "1px solid #E5484D55",
    borderRadius: 8,
    fontSize: 12.5,
    color: "#FF9B9E",
    lineHeight: 1.5,
  },
  detailBlock: {
    marginTop: 10,
    background: "#151B23",
    border: "1px solid #232B36",
    borderRadius: 8,
    overflow: "hidden",
  },
  detailBlockHeader: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    padding: "10px 14px",
    fontSize: 12.5,
    fontWeight: 600,
    cursor: "pointer",
  },
  detailBlockBody: {
    padding: "0 14px 14px 14px",
    borderTop: "1px solid #1C222B",
    paddingTop: 10,
  },
  reconUploadRow: { display: "flex", gap: 10, flexWrap: "wrap" },
  uploadSlot: {
    flex: "1 1 240px",
    display: "flex",
    alignItems: "center",
    gap: 10,
    padding: "10px 12px",
    background: "#151B23",
    border: "1px dashed #2A3340",
    borderRadius: 8,
    cursor: "pointer",
  },
  uploadIcon: {
    width: 30,
    height: 30,
    borderRadius: 7,
    background: "#1B222C",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    flexShrink: 0,
  },
  uploadLabel: { fontSize: 12.5, fontWeight: 600, color: "#C4CAD3" },
  uploadStatus: {
    fontSize: 11.5,
    color: "#5C6572",
    marginTop: 2,
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
  },
  reconError: {
    marginTop: 10,
    padding: "8px 12px",
    background: "#2A1416",
    border: "1px solid #E5484D55",
    color: "#FF9B9E",
    borderRadius: 7,
    fontSize: 12.5,
  },
  reconSummaryRow: { display: "flex", gap: 8, flexWrap: "wrap", marginTop: 14 },
  reconPreviewConfirm: {
    marginTop: 12,
    padding: 12,
    background: "#151B23",
    border: "1px solid #4F8EF755",
    borderRadius: 8,
  },
  compareTable: {
    marginTop: 12,
    display: "flex",
    flexDirection: "column",
    gap: 2,
    padding: "6px 12px",
    background: "#151B23",
    borderRadius: 8,
    border: "1px solid #232B36",
  },
  compareRow: {
    display: "flex",
    alignItems: "center",
    padding: "7px 0",
    fontSize: 12.5,
    borderBottom: "1px solid #1C222B",
    gap: 8,
  },
  compareHeaderRow: {
    fontSize: 11,
    color: "#5C6572",
    fontWeight: 700,
    textTransform: "uppercase",
    letterSpacing: "0.04em",
    borderBottom: "1px solid #232B36",
  },
  breakdownReadTable: {
    marginTop: 4,
    display: "flex",
    flexDirection: "column",
    gap: 4,
    padding: "8px 12px",
    background: "#0F1319",
    borderRadius: 7,
    border: "1px solid #232B36",
  },
  breakdownReadTableWide: {
    marginTop: 4,
    display: "flex",
    flexDirection: "column",
    gap: 5,
    padding: "8px 12px",
    background: "#0F1319",
    borderRadius: 7,
    border: "1px solid #232B36",
  },
  breakdownReadHeaderRow: {
    display: "flex",
    gap: 8,
    fontSize: 11,
    color: "#5C6572",
    fontWeight: 700,
    textTransform: "uppercase",
    letterSpacing: "0.03em",
    paddingBottom: 4,
    borderBottom: "1px solid #232B36",
  },
  breakdownReadRowWide: { display: "flex", gap: 8, fontSize: 12.5 },
  breakdownReadRow: { display: "flex", justifyContent: "space-between", fontSize: 12.5 },
  breakdownReadRowTotal: {
    display: "flex",
    justifyContent: "space-between",
    fontSize: 12.5,
    marginTop: 4,
    paddingTop: 6,
    borderTop: "1px solid #232B36",
  },
  breakdownEditor: { display: "flex", flexDirection: "column", gap: 8, marginTop: 4, marginBottom: 4 },
  breakdownEditRow: { display: "flex", gap: 8, alignItems: "center" },
  breakdownEditRowWide: { display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" },
  removeRowBtn: {
    padding: "8px 10px",
    borderRadius: 7,
    background: "#1B222C",
    border: "1px solid #2A3340",
    color: "#8B93A1",
  },
  addRowBtn: {
    display: "flex",
    alignItems: "center",
    alignSelf: "flex-start",
    padding: "7px 12px",
    borderRadius: 7,
    background: "#1B222C",
    border: "1px solid #2A3340",
    color: "#8FB6FF",
    fontSize: 12.5,
    fontWeight: 600,
  },
  reconChip: {
    display: "flex",
    alignItems: "center",
    padding: "6px 12px",
    borderRadius: 20,
    border: "1px solid",
  },
  reconPreview: {
    marginTop: 12,
    display: "flex",
    flexDirection: "column",
    gap: 5,
    padding: "10px 12px",
    background: "#151B23",
    borderRadius: 8,
    border: "1px solid #232B36",
  },
  reconPreviewRow: { display: "flex", alignItems: "center", gap: 10, fontSize: 12 },
  reconPreviewTag: {
    fontSize: 10.5,
    fontWeight: 700,
    padding: "2px 8px",
    borderRadius: 10,
    border: "1px solid",
    fontFamily: "'IBM Plex Mono', monospace",
    flexShrink: 0,
  },
  reconPreviewCode: {
    fontFamily: "'IBM Plex Mono', monospace",
    color: "#E8EAED",
    flexShrink: 0,
  },
  reconPreviewNote: { color: "#8B93A1", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" },
  reconPreviewMore: { fontSize: 11.5, color: "#5C6572", marginTop: 4 },
  exportBtn: {
    display: "flex",
    alignItems: "center",
    padding: "9px 16px",
    borderRadius: 8,
    background: "#4F8EF7",
    color: "#fff",
    fontSize: 12.5,
    fontWeight: 700,
    marginTop: 14,
  },
  timeline: { marginTop: 16, display: "flex", flexDirection: "column", gap: 6 },
  timelineRow: { display: "flex", alignItems: "center", gap: 8 },
  timelineDot: { width: 6, height: 6, borderRadius: "50%" },
  timelineLabel: { fontSize: 12, color: "#8B93A1", flex: 1 },
  timelineTime: { fontSize: 11.5, color: "#5C6572", fontFamily: "'IBM Plex Mono', monospace" },
};

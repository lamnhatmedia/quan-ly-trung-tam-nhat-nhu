import React, { useState, useEffect, useMemo, useRef } from "react";
import {
  BarChart, Bar, LineChart, Line, PieChart, Pie, Cell, XAxis, YAxis,
  CartesianGrid, Tooltip, Legend, ResponsiveContainer,
} from "recharts";
import * as XLSX from "xlsx";
import { createClient } from "@supabase/supabase-js";
import {
  LayoutDashboard, Users, GraduationCap, CalendarDays,
  Star, Wallet, Receipt, UserCog, Settings2, Search, Plus, Pencil,
  Trash2, X, ChevronLeft, ChevronRight, Download, AlertTriangle,
  RefreshCw, Menu, ArrowUpDown, FileSpreadsheet, School, DoorOpen,
  TrendingUp, TrendingDown, DollarSign, CheckCircle2, XCircle, Clock3,
  BadgeCheck, PiggyBank, ClipboardList, ChevronDown, Printer, CloudCog,
  BookOpen, Eye, Upload, FileUp, AlertCircle, Info,
} from "lucide-react";

/* ============================================================================
   UTILITIES
============================================================================ */
const cx = (...a) => a.filter(Boolean).join(" ");
const uid = (p = "id") => `${p}_${Date.now().toString(36)}${Math.random().toString(36).slice(2, 7)}`;
const vnd = (n) => (Number(n) || 0).toLocaleString("vi-VN") + "đ";
const fmtDate = (d) => { if (!d) return ""; const dt = new Date(d); return dt.toLocaleDateString("vi-VN"); };
const todayISO = () => new Date().toISOString().slice(0, 10);
const daysBetween = (a, b) => Math.round((new Date(b) - new Date(a)) / 86400000);
const clamp = (n, min, max) => Math.min(max, Math.max(min, n));
const WEEKDAYS = ["CN", "Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5", "Thứ 6", "Thứ 7"];

// ---- Lưu trữ dữ liệu bền vững trên trình duyệt (localStorage) ----
// Đổi APP_STORAGE_PREFIX khi cấu trúc dữ liệu thay đổi lớn để tránh xung đột dữ liệu cũ.
const APP_STORAGE_PREFIX = "ntn_data_v1__";

// ---- Đồng bộ dữ liệu Cloud bằng Supabase ----
const RAW_SUPABASE_URL = import.meta.env.VITE_SUPABASE_URL || "";
const SUPABASE_URL = RAW_SUPABASE_URL.trim().replace(/\/+$/, "").replace(/\/rest\/v1$/i, "");
const SUPABASE_ANON_KEY = import.meta.env.VITE_SUPABASE_ANON_KEY || import.meta.env.VITE_SUPABASE_PUBLISHABLE_KEY || "";
const supabase = SUPABASE_URL && SUPABASE_ANON_KEY ? createClient(SUPABASE_URL, SUPABASE_ANON_KEY) : null;
const CLOUD_ROW_ID = "main";
const CLOUD_DATA_VERSION = 1;
function usePersistentState(key, initialValue) {
  const [state, setState] = useState(() => {
    try {
      const raw = localStorage.getItem(APP_STORAGE_PREFIX + key);
      if (raw != null) return JSON.parse(raw);
    } catch (e) { /* dữ liệu lưu bị hỏng, dùng dữ liệu mặc định */ }
    return typeof initialValue === "function" ? initialValue() : initialValue;
  });
  useEffect(() => {
    try { localStorage.setItem(APP_STORAGE_PREFIX + key, JSON.stringify(state)); } catch (e) { /* bộ nhớ đầy hoặc không khả dụng */ }
  }, [key, state]);
  return [state, setState];
}
function resetAppData() {
  if (!window.confirm("Xoá toàn bộ dữ liệu đã lưu trên trình duyệt này và đưa ứng dụng về trạng thái trống ban đầu? Hành động này không thể hoàn tác.")) return;
  Object.keys(localStorage).filter((k) => k.startsWith(APP_STORAGE_PREFIX)).forEach((k) => localStorage.removeItem(k));
  window.location.reload();
}
function monthsBetweenInclusive(startYM, endYM) {
  if (!startYM || !endYM) return [];
  const res = [];
  let [sy, sm] = startYM.split("-").map(Number);
  const [ey, em] = endYM.split("-").map(Number);
  let guard = 0;
  while ((sy < ey || (sy === ey && sm <= em)) && guard < 240) {
    res.push(`${sy}-${String(sm).padStart(2, "0")}`);
    sm++; if (sm > 12) { sm = 1; sy++; }
    guard++;
  }
  return res;
}

function download(filename, wb) { XLSX.writeFile(wb, filename); }
function exportRows(rows, sheetName, filename) {
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName.slice(0, 31));
  download(filename, wb);
}

/* ============================================================================
   DANH MỤC HỆ THỐNG (không phải dữ liệu mẫu — dùng để phân loại)
============================================================================ */
const KHOI = [
  { id: "K1", ten: "Tiểu học" },
  { id: "K2", ten: "THCS" },
  { id: "K3", ten: "THPT" },
];

/* ============================================================================
   ĐĂNG KÝ MÔN HỌC & THU TIỀN — mô hình dữ liệu: Classes → Subjects → StudentEnrollments
   Học phí = số môn đăng ký × cấu hình học phí/môn/tháng (KHÔNG hard-code, xem tuitionConfig)
============================================================================ */
// Danh sách "kỳ phải thu" (theo từng môn, từng tháng) của 1 học sinh, tính từ tháng nhập học tới tháng `uptoMonth`.
// Tự động cập nhật theo môn đăng ký hiện tại — không cần lưu hoá đơn tĩnh.
// Tháng bắt đầu tính học phí cho 1 lượt đăng ký = tháng muộn hơn giữa "ngày nhập học" của học sinh
// và "ngày bắt đầu lớp học" (nếu lớp có cấu hình ngày bắt đầu). Nhờ vậy học phí chỉ được tính từ
// khi lớp thực sự khai giảng, thay vì lùi về những tháng trước đó.
function enrollStartMonth(student, lop) {
  const studentStart = (student.ngayNhapHoc || todayISO()).slice(0, 7);
  const classStart = lop?.ngayBatDau ? lop.ngayBatDau.slice(0, 7) : null;
  return classStart && classStart > studentStart ? classStart : studentStart;
}
function studentPeriods(student, enrollments, tuitionConfig, uptoMonth, classes = []) {
  if (!student) return [];
  const myEnroll = enrollments.filter((e) => e.hocSinhId === student.id);
  if (!myEnroll.length) return [];
  const periods = [];
  myEnroll.forEach((en) => {
    const lop = classes.find((c) => c.id === en.lopId);
    const startMonth = enrollStartMonth(student, lop);
    if (startMonth > uptoMonth) return;
    const months = monthsBetweenInclusive(startMonth, uptoMonth);
    months.forEach((thang) => periods.push({ hocSinhId: student.id, lopId: en.lopId, monHocId: en.monHocId, thang, phaiThu: tuitionConfig.hocPhiMon }));
  });
  return periods;
}
function buildAllPeriods(students, enrollments, tuitionConfig, uptoMonth, classes = []) {
  const list = [];
  students.filter((s) => s.trangThai === "Đang học").forEach((st) => list.push(...studentPeriods(st, enrollments, tuitionConfig, uptoMonth, classes)));
  return list;
}
function allocatedOf(hocSinhId, lopId, monHocId, thang, paymentAllocations) {
  return paymentAllocations.filter((a) => a.hocSinhId === hocSinhId && a.lopId === lopId && a.monHocId === monHocId && a.thang === thang).reduce((s, a) => s + a.soTien, 0);
}
function periodStatus(phaiThu, daThu) { if (daThu <= 0) return "Chưa đóng"; if (daThu < phaiThu) return "Đóng một phần"; return "Đã đóng"; }
function periodsForMonth(student, enrollments, tuitionConfig, thang, classes = []) {
  const myEnroll = enrollments.filter((e) => e.hocSinhId === student.id);
  if (!myEnroll.length) return [];
  return myEnroll
    .filter((en) => thang >= enrollStartMonth(student, classes.find((c) => c.id === en.lopId)))
    .map((en) => ({ hocSinhId: student.id, lopId: en.lopId, monHocId: en.monHocId, thang, phaiThu: tuitionConfig.hocPhiMon }));
}
function lastNMonths(n, uptoMonth) {
  const [y, m] = uptoMonth.split("-").map(Number);
  const arr = [];
  for (let i = n - 1; i >= 0; i--) {
    let mm = m - i, yy = y;
    while (mm <= 0) { mm += 12; yy -= 1; }
    arr.push(`${yy}-${String(mm).padStart(2, "0")}`);
  }
  return arr;
}

/* ============================================================================
   ROLES & PERMISSIONS
============================================================================ */
const ROLES = ["Admin", "Kế toán", "Giáo viên", "Trợ giảng"];
// access: 'full' | 'view' | 'none'
const PERMISSIONS = {
  dashboard: { Admin: "full", "Kế toán": "full", "Giáo viên": "view", "Trợ giảng": "view" },
  students: { Admin: "full", "Kế toán": "view", "Giáo viên": "view", "Trợ giảng": "view" },
  classes: { Admin: "full", "Kế toán": "view", "Giáo viên": "view", "Trợ giảng": "view" },
  resources: { Admin: "full", "Kế toán": "view", "Giáo viên": "none", "Trợ giảng": "none" },
  schedule: { Admin: "full", "Kế toán": "view", "Giáo viên": "view", "Trợ giảng": "view" },
  thuTien: { Admin: "full", "Kế toán": "full", "Giáo viên": "view", "Trợ giảng": "none" },
  transactions: { Admin: "full", "Kế toán": "full", "Giáo viên": "none", "Trợ giảng": "none" },
  reports: { Admin: "full", "Kế toán": "full", "Giáo viên": "view", "Trợ giảng": "none" },
  users: { Admin: "full", "Kế toán": "none", "Giáo viên": "none", "Trợ giảng": "none" },
  sync: { Admin: "full", "Kế toán": "none", "Giáo viên": "none", "Trợ giảng": "none" },
};
const can = (role, moduleKey, level = "view") => {
  const acc = PERMISSIONS[moduleKey]?.[role] || "none";
  if (level === "view") return acc !== "none";
  return acc === "full";
};

/* ============================================================================
   GENERIC UI PRIMITIVES
============================================================================ */
function Badge({ children, color = "slate" }) {
  const map = {
    slate: "bg-slate-100 text-slate-700 border-slate-200",
    green: "bg-emerald-50 text-emerald-700 border-emerald-200",
    amber: "bg-amber-50 text-amber-700 border-amber-200",
    red: "bg-rose-50 text-rose-700 border-rose-200",
    blue: "bg-sky-50 text-sky-700 border-sky-200",
    teal: "bg-teal-50 text-teal-700 border-teal-200",
    purple: "bg-violet-50 text-violet-700 border-violet-200",
  };
  return <span className={cx("inline-flex items-center px-2 py-0.5 rounded-full text-xs font-medium border", map[color])}>{children}</span>;
}

function statusBadge(status) {
  const M = {
    "Đang học": "green", "Nghỉ học": "slate",
    "Có mặt": "green", "Vắng": "red", "Vắng phép": "amber", "Đi trễ": "blue",
    "Tốt": "green", "Khá": "teal", "Đạt": "amber", "Chưa đạt": "red",
    "Đã thu đủ": "green", "Thu một phần": "amber", "Chưa thu": "red",
    "Hoạt động": "green", "Khóa": "slate", "Thu": "green", "Chi": "red",
  };
  return <Badge color={M[status] || "slate"}>{status}</Badge>;
}

function Card({ children, className }) { return <div className={cx("bg-white rounded-xl border border-slate-200 shadow-sm", className)}>{children}</div>; }

function KPICard({ icon: Icon, label, value, sub, tone = "teal" }) {
  const tones = {
    teal: "bg-teal-50 text-teal-700", amber: "bg-amber-50 text-amber-700",
    rose: "bg-rose-50 text-rose-700", sky: "bg-sky-50 text-sky-700",
    violet: "bg-violet-50 text-violet-700", emerald: "bg-emerald-50 text-emerald-700",
  };
  return (
    <Card className="p-4 flex items-start gap-3 min-w-0">
      <div className={cx("w-10 h-10 rounded-lg flex items-center justify-center shrink-0", tones[tone])}><Icon size={20} /></div>
      <div className="min-w-0">
        <p className="text-xs text-slate-500 truncate">{label}</p>
        <p className="text-lg font-semibold text-slate-800 truncate">{value}</p>
        {sub && <p className="text-xs text-slate-400 mt-0.5 truncate">{sub}</p>}
      </div>
    </Card>
  );
}

function EmptyState({ title = "Chưa có dữ liệu", desc = "Hãy thêm bản ghi mới để bắt đầu.", icon: Icon = ClipboardList }) {
  return (
    <div className="flex flex-col items-center justify-center py-14 text-center text-slate-400">
      <Icon size={36} className="mb-2 opacity-60" />
      <p className="font-medium text-slate-500">{title}</p>
      <p className="text-sm">{desc}</p>
    </div>
  );
}

function Modal({ open, onClose, title, children, wide }) {
  if (!open) return null;
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/40" onClick={onClose}>
      <div className={cx("bg-white rounded-xl shadow-xl w-full max-h-[90vh] overflow-y-auto", wide ? "max-w-3xl" : "max-w-lg")} onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between px-5 py-4 border-b border-slate-100 sticky top-0 bg-white z-10">
          <h3 className="font-semibold text-slate-800">{title}</h3>
          <button onClick={onClose} className="text-slate-400 hover:text-slate-600"><X size={18} /></button>
        </div>
        <div className="p-5">{children}</div>
      </div>
    </div>
  );
}

function ConfirmDialog({ open, onCancel, onConfirm, text }) {
  if (!open) return null;
  return (
    <div className="fixed inset-0 z-[60] flex items-center justify-center p-4 bg-slate-900/40" onClick={onCancel}>
      <div className="bg-white rounded-xl shadow-xl w-full max-w-sm p-5" onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center gap-2 text-rose-600 mb-2"><AlertTriangle size={20} /><h3 className="font-semibold">Xác nhận xoá</h3></div>
        <p className="text-sm text-slate-600 mb-4">{text}</p>
        <div className="flex justify-end gap-2">
          <button onClick={onCancel} className="px-3 py-1.5 rounded-lg text-sm border border-slate-200 hover:bg-slate-50">Huỷ</button>
          <button onClick={onConfirm} className="px-3 py-1.5 rounded-lg text-sm bg-rose-600 text-white hover:bg-rose-700">Xoá</button>
        </div>
      </div>
    </div>
  );
}

function Field({ label, required, error, children }) {
  return (
    <div className="mb-3">
      <label className="block text-sm font-medium text-slate-600 mb-1">{label}{required && <span className="text-rose-500"> *</span>}</label>
      {children}
      {error && <p className="text-xs text-rose-500 mt-1">{error}</p>}
    </div>
  );
}
const inputCls = "w-full px-3 py-2 rounded-lg border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-teal-500/40 focus:border-teal-500";

function TextInput(props) { return <input {...props} className={cx(inputCls, props.className)} />; }
function Select({ options, ...props }) {
  return (
    <select {...props} className={cx(inputCls, "bg-white", props.className)}>
      {options.map((o) => <option key={o.value} value={o.value}>{o.label}</option>)}
    </select>
  );
}
function TextArea(props) { return <textarea {...props} rows={3} className={cx(inputCls, props.className)} />; }

/* Generic searchable / sortable / paginated table */
function DataTable({ columns, rows, searchKeys = [], pageSize = 8, exportName, filterBar, emptyTitle }) {
  const [q, setQ] = useState("");
  const [sortKey, setSortKey] = useState(null);
  const [sortDir, setSortDir] = useState("asc");
  const [page, setPage] = useState(1);

  const filtered = useMemo(() => {
    let d = rows;
    if (q.trim()) {
      const ql = q.toLowerCase();
      d = d.filter((r) => searchKeys.some((k) => String(r[k] ?? "").toLowerCase().includes(ql)));
    }
    if (sortKey) {
      d = [...d].sort((a, b) => {
        const av = a[sortKey], bv = b[sortKey];
        if (typeof av === "number" && typeof bv === "number") return sortDir === "asc" ? av - bv : bv - av;
        return sortDir === "asc" ? String(av ?? "").localeCompare(String(bv ?? ""), "vi") : String(bv ?? "").localeCompare(String(av ?? ""), "vi");
      });
    }
    return d;
  }, [rows, q, sortKey, sortDir, searchKeys]);

  useEffect(() => { setPage(1); }, [q, rows.length]);
  const totalPages = Math.max(1, Math.ceil(filtered.length / pageSize));
  const pageRows = filtered.slice((page - 1) * pageSize, page * pageSize);

  function toggleSort(key) {
    if (sortKey === key) setSortDir((d) => (d === "asc" ? "desc" : "asc"));
    else { setSortKey(key); setSortDir("asc"); }
  }
  function doExport() {
    exportRows(filtered.map((r) => { const o = {}; columns.forEach((c) => { o[c.label] = c.exportValue ? c.exportValue(r) : (typeof r[c.key] !== "object" ? r[c.key] : ""); }); return o; }), exportName || "data", `${exportName || "data"}.xlsx`);
  }

  return (
    <div>
      <div className="flex flex-col sm:flex-row sm:items-center gap-2 mb-3">
        <div className="relative flex-1">
          <Search size={16} className="absolute left-3 top-2.5 text-slate-400" />
          <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Tìm kiếm..." className="w-full pl-9 pr-3 py-2 rounded-lg border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-teal-500/40" />
        </div>
        {filterBar}
        <button onClick={doExport} className="flex items-center gap-1.5 px-3 py-2 rounded-lg text-sm border border-slate-200 hover:bg-slate-50 text-slate-600 shrink-0"><FileSpreadsheet size={15} /> Xuất Excel</button>
      </div>
      <div className="overflow-x-auto border border-slate-200 rounded-xl">
        <table className="w-full text-sm min-w-[640px]">
          <thead>
            <tr className="bg-slate-50 text-slate-500 text-xs uppercase">
              {columns.map((c) => (
                <th key={c.key} className={cx("px-3 py-2.5 text-left font-medium whitespace-nowrap", c.sortable && "cursor-pointer select-none")} onClick={() => c.sortable && toggleSort(c.key)}>
                  <span className="inline-flex items-center gap-1">{c.label}{c.sortable && <ArrowUpDown size={11} className="opacity-50" />}</span>
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {pageRows.map((r, i) => (
              <tr key={r.id || i} className="border-t border-slate-100 hover:bg-slate-50/70">
                {columns.map((c) => <td key={c.key} className="px-3 py-2.5 align-middle whitespace-nowrap">{c.render ? c.render(r) : r[c.key]}</td>)}
              </tr>
            ))}
          </tbody>
        </table>
        {filtered.length === 0 && <EmptyState title={emptyTitle || "Không tìm thấy dữ liệu"} desc="Thử thay đổi từ khoá tìm kiếm hoặc thêm mới bản ghi." />}
      </div>
      {filtered.length > 0 && (
        <div className="flex items-center justify-between mt-3 text-sm text-slate-500">
          <span>{filtered.length} bản ghi · Trang {page}/{totalPages}</span>
          <div className="flex gap-1">
            <button disabled={page <= 1} onClick={() => setPage((p) => p - 1)} className="p-1.5 rounded-lg border border-slate-200 disabled:opacity-40 hover:bg-slate-50"><ChevronLeft size={15} /></button>
            <button disabled={page >= totalPages} onClick={() => setPage((p) => p + 1)} className="p-1.5 rounded-lg border border-slate-200 disabled:opacity-40 hover:bg-slate-50"><ChevronRight size={15} /></button>
          </div>
        </div>
      )}
    </div>
  );
}

function SectionHeader({ title, desc, actions }) {
  return (
    <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 mb-4">
      <div><h2 className="text-lg font-semibold text-slate-800">{title}</h2>{desc && <p className="text-sm text-slate-500">{desc}</p>}</div>
      <div className="flex gap-2">{actions}</div>
    </div>
  );
}

function PrimaryButton({ children, onClick, icon: Icon = Plus, className }) {
  return <button onClick={onClick} className={cx("flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-sm font-medium bg-teal-700 text-white hover:bg-teal-800 shrink-0", className)}><Icon size={15} />{children}</button>;
}
function IconBtn({ onClick, icon: Icon, tone = "slate", title }) {
  const tones = { slate: "hover:bg-slate-100 text-slate-500", rose: "hover:bg-rose-50 text-rose-500", teal: "hover:bg-teal-50 text-teal-600" };
  return <button title={title} onClick={onClick} className={cx("p-1.5 rounded-md", tones[tone])}><Icon size={15} /></button>;
}

/* ============================================================================
   ĐĂNG NHẬP NỘI BỘ — không dùng Supabase Auth
============================================================================ */
const LOGIN_USERNAME = "admin";
const LOGIN_PASSWORD = "123456";
const LOGIN_STORAGE_KEY = "ntn_authenticated_v1";

function LoginScreen({ onLogin }) {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [showPassword, setShowPassword] = useState(false);
  const [error, setError] = useState("");
  function submit(e) {
    e.preventDefault();
    if (username.trim() === LOGIN_USERNAME && password === LOGIN_PASSWORD) {
      localStorage.setItem(LOGIN_STORAGE_KEY, "true");
      onLogin();
    } else setError("Tài khoản hoặc mật khẩu không đúng.");
  }
  return (
    <div className="min-h-screen bg-slate-50 flex items-center justify-center p-4" style={{ fontFamily: "Inter, ui-sans-serif, system-ui" }}>
      <div className="w-full max-w-md bg-white rounded-2xl shadow-xl border border-slate-200 overflow-hidden">
        <div className="bg-teal-950 px-6 py-7 text-center">
          <div className="mx-auto w-14 h-14 rounded-2xl bg-amber-400 text-teal-950 flex items-center justify-center font-bold text-xl mb-3">NN</div>
          <h1 className="text-xl font-bold text-white">Quản lý Trung tâm Nhật Như</h1>
          <p className="text-sm text-teal-200 mt-1">Đăng nhập để tiếp tục</p>
        </div>
        <form onSubmit={submit} className="p-6 space-y-4">
          <div><label className="block text-sm font-medium text-slate-600 mb-1.5">Tài khoản</label>
            <input value={username} onChange={e=>setUsername(e.target.value)} autoFocus autoComplete="username" placeholder="Nhập tài khoản" className={inputCls}/></div>
          <div><label className="block text-sm font-medium text-slate-600 mb-1.5">Mật khẩu</label>
            <div className="relative"><input type={showPassword?"text":"password"} value={password} onChange={e=>setPassword(e.target.value)} autoComplete="current-password" placeholder="Nhập mật khẩu" className={cx(inputCls,"pr-16")}/>
              <button type="button" onClick={()=>setShowPassword(v=>!v)} className="absolute right-2 top-1/2 -translate-y-1/2 px-2 py-1 text-xs text-slate-500">{showPassword?"Ẩn":"Hiện"}</button>
            </div>
          </div>
          {error && <div className="rounded-lg bg-rose-50 border border-rose-200 px-3 py-2.5 text-sm text-rose-700">{error}</div>}
          <button className="w-full py-2.5 rounded-lg bg-teal-700 text-white font-medium hover:bg-teal-800">Đăng nhập</button>
        </form>
      </div>
    </div>
  );
}

/* ============================================================================
   MAIN APP
============================================================================ */
export default function App() {
  useEffect(() => { document.title = "Quản lý Trung tâm Nhật Như"; }, []);

  /* ---------------- core data state (tự động lưu vào localStorage của trình duyệt) ----------------
     Ứng dụng khởi động với dữ liệu TRỐNG (không có dữ liệu mẫu/demo). Mọi thay đổi được lưu tự động
     vào localStorage của trình duyệt này qua usePersistentState. */
  const [students, setStudents] = usePersistentState("students", []);
  const [classes, setClasses] = usePersistentState("classes", []);
  const [teachers, setTeachers] = usePersistentState("teachers", []);
  const [assistants, setAssistants] = usePersistentState("assistants", []);
  const [rooms, setRooms] = usePersistentState("rooms", []);
  const [subjects, setSubjects] = usePersistentState("subjects", []);
  const [tuitionConfig, setTuitionConfig] = usePersistentState("tuitionConfig", { hocPhiMon: 400000 });
  const [enrollments, setEnrollments] = usePersistentState("enrollments", []);
  const [payments, setPayments] = usePersistentState("payments", []);
  const [paymentAllocations, setPaymentAllocations] = usePersistentState("paymentAllocations", []);
  const [transactions, setTransactions] = usePersistentState("transactions", []);
  const [users, setUsers] = usePersistentState("users", []);
  const [syncLog, setSyncLog] = useState([{ time: new Date().toLocaleString("vi-VN"), msg: "Ứng dụng khởi tạo — dữ liệu được lưu tự động trên trình duyệt này." }]);
  const [sheetsCfg, setSheetsCfg] = usePersistentState("sheetsCfg", { url: "", lastSync: null, status: "offline" });

  // Cloud sync: localStorage vẫn là cache cục bộ; Supabase là bản sao chung giữa các thiết bị.
  const [cloudStatus, setCloudStatus] = useState(supabase ? "loading" : "not_configured");
  const [cloudError, setCloudError] = useState("");
  const [lastCloudSync, setLastCloudSync] = useState(null);
  const cloudReadyRef = useRef(false);
  const syncTimerRef = useRef(null);

  const cloudSnapshot = useMemo(() => ({
    version: CLOUD_DATA_VERSION,
    updatedAt: new Date().toISOString(),
    students, classes, teachers, assistants, rooms, subjects, tuitionConfig,
    enrollments, payments, paymentAllocations, transactions, users,
  }), [students, classes, teachers, assistants, rooms, subjects, tuitionConfig, enrollments, payments, paymentAllocations, transactions, users]);

  async function loadCloudData() {
    if (!supabase) {
      setCloudStatus("not_configured");
      setCloudError("Chưa cấu hình VITE_SUPABASE_URL và VITE_SUPABASE_ANON_KEY.");
      cloudReadyRef.current = true;
      return;
    }
    setCloudStatus("loading"); setCloudError("");
    const { data, error } = await supabase.from("app_data").select("data, updated_at").eq("id", CLOUD_ROW_ID).maybeSingle();
    if (error) {
      setCloudStatus("error"); setCloudError(error.message); cloudReadyRef.current = true; return;
    }
    if (data?.data) {
      const d = data.data;
      if (Array.isArray(d.students)) setStudents(d.students);
      if (Array.isArray(d.classes)) setClasses(d.classes);
      if (Array.isArray(d.teachers)) setTeachers(d.teachers);
      if (Array.isArray(d.assistants)) setAssistants(d.assistants);
      if (Array.isArray(d.rooms)) setRooms(d.rooms);
      if (Array.isArray(d.subjects)) setSubjects(d.subjects);
      if (d.tuitionConfig) setTuitionConfig(d.tuitionConfig);
      if (Array.isArray(d.enrollments)) setEnrollments(d.enrollments);
      if (Array.isArray(d.payments)) setPayments(d.payments);
      if (Array.isArray(d.paymentAllocations)) setPaymentAllocations(d.paymentAllocations);
      if (Array.isArray(d.transactions)) setTransactions(d.transactions);
      if (Array.isArray(d.users)) setUsers(d.users);
      setLastCloudSync(data.updated_at || new Date().toISOString());
      setCloudStatus("connected");
    } else {
      // Chưa có dữ liệu cloud: giữ dữ liệu local hiện tại và tạo bản sao đầu tiên khi có thay đổi/đồng bộ.
      setCloudStatus("connected");
    }
    cloudReadyRef.current = true;
  }

  async function saveCloudData(snapshot = cloudSnapshot, silent = false) {
    if (!supabase || !cloudReadyRef.current) return false;
    if (!silent) setCloudStatus("saving");
    const now = new Date().toISOString();
    const payload = { ...snapshot, updatedAt: now };
    const { error } = await supabase.from("app_data").upsert(
      { id: CLOUD_ROW_ID, data: payload, updated_at: now },
      { onConflict: "id" }
    );
    if (error) {
      setCloudStatus("error"); setCloudError(error.message); return false;
    }
    setLastCloudSync(now); setCloudStatus("connected"); setCloudError("");
    return true;
  }

  useEffect(() => {
    loadCloudData();
    return () => { if (syncTimerRef.current) clearTimeout(syncTimerRef.current); };
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  useEffect(() => {
    if (!cloudReadyRef.current || !supabase) return;
    if (syncTimerRef.current) clearTimeout(syncTimerRef.current);
    syncTimerRef.current = setTimeout(() => saveCloudData(cloudSnapshot, true), 700);
  }, [cloudSnapshot]); // eslint-disable-line react-hooks/exhaustive-deps

  async function syncNow() {
    return saveCloudData(cloudSnapshot);
  }

  /* ---------------- ui / auth state ---------------- */
  const [role, setRole] = usePersistentState("role", "Admin");
  const currentUser = users.find((u) => u.vaiTro === role) || users[0];
  const [view, setView] = useState("dashboard");
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [isLoggedIn, setIsLoggedIn] = useState(() => localStorage.getItem(LOGIN_STORAGE_KEY) === "true");

  const lookups = { classes, teachers, assistants, rooms, students, subjects };

  const NAV = [
    { key: "dashboard", label: "Dashboard", icon: LayoutDashboard },
    { key: "students", label: "Học sinh", icon: Users },
    { key: "classes", label: "Lớp học", icon: School },
    { key: "resources", label: "GV/Trợ giảng/Phòng/Môn", icon: UserCog },
    { key: "schedule", label: "Thời khoá biểu", icon: CalendarDays },
    { key: "thuTien", label: "Thu tiền", icon: Receipt },
    { key: "transactions", label: "Thu chi", icon: Wallet },
    { key: "reports", label: "Báo cáo", icon: FileSpreadsheet },
    { key: "users", label: "Người dùng", icon: UserCog },
    { key: "sync", label: "Đồng bộ Google Sheets", icon: CloudCog },
  ].filter((n) => can(role, n.key, "view"));

  useEffect(() => { if (!NAV.find((n) => n.key === view)) setView(NAV[0]?.key || "dashboard"); }, [role]); // eslint-disable-line

  function logSync(msg) { setSyncLog((l) => [{ time: new Date().toLocaleString("vi-VN"), msg }, ...l].slice(0, 30)); }

  const logout = () => { localStorage.removeItem(LOGIN_STORAGE_KEY); setIsLoggedIn(false); };

  const ctx = { students, setStudents, classes, setClasses, teachers, setTeachers, assistants, setAssistants, rooms, setRooms, subjects, setSubjects, tuitionConfig, setTuitionConfig, enrollments, setEnrollments, payments, setPayments, paymentAllocations, setPaymentAllocations, transactions, setTransactions, users, setUsers, lookups, role, logSync, sheetsCfg, setSheetsCfg, syncLog, resetAppData, cloudStatus, cloudError, lastCloudSync, syncNow, logout };

  if (!isLoggedIn) return <LoginScreen onLogin={() => setIsLoggedIn(true)} />;

  return (
    <div className="min-h-screen bg-slate-50 text-slate-800 flex" style={{ fontFamily: "Inter, ui-sans-serif, system-ui" }}>
      {/* Sidebar */}
      <aside className={cx("fixed lg:static z-40 inset-y-0 left-0 w-64 bg-teal-950 text-teal-50 flex flex-col transition-transform duration-200", sidebarOpen ? "translate-x-0" : "-translate-x-full lg:translate-x-0")}>
        <div className="flex items-center gap-2 px-5 h-16 border-b border-teal-900 shrink-0">
          <div className="w-8 h-8 rounded-lg bg-amber-400 text-teal-950 flex items-center justify-center font-bold">NN</div>
          <div>
            <p className="font-semibold leading-tight text-sm">Trung tâm Nhật Như</p>
          </div>
          <button className="ml-auto lg:hidden text-teal-300" onClick={() => setSidebarOpen(false)}><X size={18} /></button>
        </div>
        <nav className="flex-1 overflow-y-auto py-3 px-2 space-y-0.5">
          {NAV.map((n) => (
            <button key={n.key} onClick={() => { setView(n.key); setSidebarOpen(false); }}
              className={cx("w-full flex items-center gap-2.5 px-3 py-2.5 rounded-lg text-sm transition relative", view === n.key ? "bg-teal-800 text-white font-medium" : "text-teal-200 hover:bg-teal-900")}>
              {view === n.key && <span className="absolute left-0 top-1.5 bottom-1.5 w-1 bg-amber-400 rounded-full" />}
              <n.icon size={16} /> {n.label}
            </button>
          ))}
        </nav>
        <div className="p-3 border-t border-teal-900 text-xs text-teal-300">
          <p className="truncate">{currentUser?.hoTen}</p>
          <p className="text-teal-400">{role}</p>
        </div>
      </aside>
      {sidebarOpen && <div className="fixed inset-0 bg-slate-900/40 z-30 lg:hidden" onClick={() => setSidebarOpen(false)} />}

      <div className="flex-1 min-w-0 flex flex-col">
        {/* Topbar */}
        <header className="h-16 bg-white border-b border-slate-200 flex items-center gap-3 px-4 lg:px-6 sticky top-0 z-20">
          <button className="lg:hidden text-slate-500" onClick={() => setSidebarOpen(true)}><Menu size={20} /></button>
          <h1 className="font-semibold text-slate-800">{NAV.find((n) => n.key === view)?.label}</h1>
          <div className="ml-auto flex items-center gap-3">
            <div className={cx("hidden sm:flex items-center gap-1.5 text-xs px-2.5 py-1.5 rounded-full", cloudStatus === "connected" ? "bg-emerald-50 text-emerald-700" : cloudStatus === "saving" || cloudStatus === "loading" ? "bg-amber-50 text-amber-700" : "bg-rose-50 text-rose-700")} title={cloudError || "Dữ liệu được tự động đồng bộ lên Supabase."}>
              <CloudCog size={13} />
              {cloudStatus === "connected" ? "Đã đồng bộ Cloud" : cloudStatus === "saving" ? "Đang lưu Cloud…" : cloudStatus === "loading" ? "Đang tải Cloud…" : cloudStatus === "not_configured" ? "Chưa cấu hình Cloud" : "Lỗi đồng bộ"}
            </div>
            <button onClick={syncNow} className="hidden sm:inline-flex items-center gap-1.5 text-xs px-2.5 py-1.5 rounded-lg border border-slate-200 hover:bg-slate-50" title="Đồng bộ ngay">
              <RefreshCw size={13} /> Đồng bộ
            </button>
            <label className="text-xs text-slate-400 hidden md:block">Vai trò</label>
            <select value={role} onChange={(e) => setRole(e.target.value)} className="text-sm border border-slate-200 rounded-lg px-2.5 py-1.5 bg-white">
              {ROLES.map((r) => <option key={r}>{r}</option>)}
            </select>
            <button onClick={logout} className="px-2.5 py-1.5 rounded-lg text-sm border border-slate-200 text-slate-600 hover:bg-slate-50">Đăng xuất</button>
          </div>
        </header>
        <main className="flex-1 p-4 lg:p-6 max-w-[1400px] w-full mx-auto">
          {view === "dashboard" && <Dashboard {...ctx} />}
          {view === "students" && <StudentsPage {...ctx} />}
          {view === "classes" && <ClassesPage {...ctx} />}
          {view === "resources" && <ResourcesPage {...ctx} />}
          {view === "schedule" && <SchedulePage {...ctx} />}
          {view === "thuTien" && <ThuTienPage {...ctx} />}
          {view === "transactions" && <TransactionsPage {...ctx} />}
          {view === "reports" && <ReportsPage {...ctx} />}
          {view === "users" && <UsersPage {...ctx} />}
          {view === "sync" && <SyncPage {...ctx} />}
        </main>
      </div>
    </div>
  );
}

/* ============================================================================
   DASHBOARD
============================================================================ */
function Dashboard({ students, classes, subjects, enrollments, tuitionConfig, teachers, payments, paymentAllocations, transactions }) {
  const activeStudents = students.filter((s) => s.trangThai === "Đang học");
  const thisMonth = todayISO().slice(0, 7);

  const allPeriods = useMemo(() => buildAllPeriods(students, enrollments, tuitionConfig, thisMonth, classes).map((p) => ({ ...p, daThu: allocatedOf(p.hocSinhId, p.lopId, p.monHocId, p.thang, paymentAllocations) })), [students, enrollments, tuitionConfig, paymentAllocations, thisMonth, classes]);
  const periodsThisMonth = allPeriods.filter((p) => p.thang === thisMonth);
  const phaiThu = periodsThisMonth.reduce((s, p) => s + p.phaiThu, 0);
  const daThu = periodsThisMonth.reduce((s, p) => s + Math.min(p.daThu, p.phaiThu), 0);
  const conNo = allPeriods.reduce((s, p) => s + Math.max(0, p.phaiThu - p.daThu), 0);

  const revenueOther = transactions.filter((t) => t.loai === "Thu" && t.ngay.startsWith(thisMonth)).reduce((s, t) => s + t.soTien, 0);
  const chiPhi = transactions.filter((t) => t.loai === "Chi" && t.ngay.startsWith(thisMonth)).reduce((s, t) => s + t.soTien, 0);
  const doanhThu = daThu + revenueOther;
  const loiNhuan = doanhThu - chiPhi;

  // last 6 months revenue/expense chart (theo ngày thực thu tiền)
  const monthsArr = Array.from({ length: 6 }).map((_, i) => { const d = new Date(); d.setMonth(d.getMonth() - (5 - i)); return d.toISOString().slice(0, 7); });
  const chartData = monthsArr.map((m) => {
    const payM = payments.filter((p) => p.ngayThu.startsWith(m));
    const rev = payM.reduce((s, p) => s + p.tongThucThu, 0) + transactions.filter((t) => t.loai === "Thu" && t.ngay.startsWith(m)).reduce((s, t) => s + t.soTien, 0);
    const exp = transactions.filter((t) => t.loai === "Chi" && t.ngay.startsWith(m)).reduce((s, t) => s + t.soTien, 0);
    return { thang: m.slice(5) + "/" + m.slice(0, 4), "Doanh thu": rev, "Chi phí": exp };
  });

  const classSize = classes.map((c) => ({ ten: c.tenLop, siSo: students.filter((s) => s.lopId === c.id && s.trangThai === "Đang học").length }));
  const revByClass = classes.map((c) => ({ ten: c.tenLop, doanhThu: paymentAllocations.filter((a) => a.lopId === c.id).reduce((s, a) => s + a.soTien, 0) }));
  const revBySubject = subjects.map((m) => ({ ten: m.ten, doanhThu: paymentAllocations.filter((a) => a.monHocId === m.id).reduce((s, a) => s + a.soTien, 0) }));

  return (
    <div className="space-y-6">
      <div className="grid grid-cols-2 md:grid-cols-3 xl:grid-cols-5 gap-3">
        <KPICard icon={Users} tone="teal" label="Học sinh đang học" value={activeStudents.length} sub={`${students.length} tổng số`} />
        <KPICard icon={School} tone="sky" label="Lớp học" value={classes.length} sub={`${teachers.length} giáo viên`} />
        <KPICard icon={BookOpen} tone="violet" label="Môn học" value={subjects.length} />
        <KPICard icon={DollarSign} tone="amber" label="Học phí phải thu (tháng)" value={vnd(phaiThu)} />
        <KPICard icon={CheckCircle2} tone="teal" label="Học phí đã thu (tháng)" value={vnd(daThu)} />
        <KPICard icon={PiggyBank} tone="rose" label="Tổng còn thiếu (luỹ kế)" value={vnd(conNo)} />
        <KPICard icon={TrendingUp} tone="emerald" label="Doanh thu tháng" value={vnd(doanhThu)} />
        <KPICard icon={TrendingDown} tone="rose" label="Chi phí tháng" value={vnd(chiPhi)} />
        <KPICard icon={Wallet} tone="violet" label="Lợi nhuận tháng" value={vnd(loiNhuan)} />
      </div>

      <Card className="p-4">
        <p className="font-medium text-slate-700 mb-3 text-sm">Doanh thu &amp; Chi phí 6 tháng gần nhất</p>
        <ResponsiveContainer width="100%" height={260}>
          <BarChart data={chartData}>
            <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" />
            <XAxis dataKey="thang" tick={{ fontSize: 12 }} />
            <YAxis tick={{ fontSize: 11 }} tickFormatter={(v) => (v / 1e6) + "tr"} />
            <Tooltip formatter={(v) => vnd(v)} />
            <Legend />
            <Bar dataKey="Doanh thu" fill="#0f766e" radius={[4, 4, 0, 0]} />
            <Bar dataKey="Chi phí" fill="#e11d48" radius={[4, 4, 0, 0]} />
          </BarChart>
        </ResponsiveContainer>
      </Card>

      <div className="grid lg:grid-cols-2 gap-4">
        <Card className="p-4">
          <p className="font-medium text-slate-700 mb-3 text-sm">Doanh thu học phí theo lớp (luỹ kế)</p>
          <ResponsiveContainer width="100%" height={260}>
            <BarChart data={revByClass}>
              <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" />
              <XAxis dataKey="ten" tick={{ fontSize: 11 }} interval={0} angle={-15} textAnchor="end" height={60} />
              <YAxis tick={{ fontSize: 11 }} tickFormatter={(v) => (v / 1e6) + "tr"} />
              <Tooltip formatter={(v) => vnd(v)} />
              <Bar dataKey="doanhThu" name="Doanh thu" fill="#0f766e" radius={[4, 4, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </Card>
        <Card className="p-4">
          <p className="font-medium text-slate-700 mb-3 text-sm">Doanh thu học phí theo môn (luỹ kế)</p>
          <ResponsiveContainer width="100%" height={260}>
            <BarChart data={revBySubject}>
              <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" />
              <XAxis dataKey="ten" tick={{ fontSize: 11 }} interval={0} angle={-15} textAnchor="end" height={60} />
              <YAxis tick={{ fontSize: 11 }} tickFormatter={(v) => (v / 1e6) + "tr"} />
              <Tooltip formatter={(v) => vnd(v)} />
              <Bar dataKey="doanhThu" name="Doanh thu" fill="#d97706" radius={[4, 4, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </Card>
      </div>

      <Card className="p-4">
        <p className="font-medium text-slate-700 mb-3 text-sm">Sĩ số theo lớp</p>
        <ResponsiveContainer width="100%" height={260}>
          <BarChart data={classSize}>
            <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" />
            <XAxis dataKey="ten" tick={{ fontSize: 11 }} interval={0} angle={-15} textAnchor="end" height={60} />
            <YAxis allowDecimals={false} tick={{ fontSize: 11 }} />
            <Tooltip />
            <Bar dataKey="siSo" fill="#d97706" radius={[4, 4, 0, 0]} />
          </BarChart>
        </ResponsiveContainer>
      </Card>
    </div>
  );
}

/* ============================================================================
   STUDENTS
============================================================================ */
function StudentsPage({ students, setStudents, classes, setClasses, subjects, setSubjects, enrollments, setEnrollments, tuitionConfig, paymentAllocations, role }) {
  const editable = can(role, "students", "full");
  const [modal, setModal] = useState(null); // {mode:'add'|'edit', data}
  const [del, setDel] = useState(null);
  const [detail, setDetail] = useState(null);
  const [classFilter, setClassFilter] = useState("");
  const [importOpen, setImportOpen] = useState(false);
  const [monthFilter, setMonthFilter] = useState(todayISO().slice(0, 7));

  const rows = useMemo(() => (classFilter ? students.filter((s) => s.lopId === classFilter) : students), [students, classFilter]);
  const classOf = (id) => classes.find((c) => c.id === id)?.tenLop || "—";
  const monthPeriods = (r) => periodsForMonth(r, enrollments, tuitionConfig, monthFilter, classes).map((p) => ({ ...p, daThu: Math.min(p.phaiThu, allocatedOf(p.hocSinhId, p.lopId, p.monHocId, p.thang, paymentAllocations)) }));
  const subjNamesForMonth = (r) => [...new Set(monthPeriods(r).map((p) => subjects.find((s) => s.id === p.monHocId)?.ten).filter(Boolean))];

  const columns = [
    { key: "maHS", label: "Mã HS", sortable: true },
    { key: "hoTen", label: "Họ tên", sortable: true, render: (r) => <button onClick={() => setDetail(r)} className="text-teal-700 font-medium hover:underline">{r.hoTen}</button> },
    { key: "gioiTinh", label: "Giới tính" },
    { key: "lopId", label: "Lớp", render: (r) => classOf(r.lopId), exportValue: (r) => classOf(r.lopId) },
    { key: "monDangKy", label: "Môn đăng ký (tháng)", render: (r) => {
      const names = subjNamesForMonth(r);
      return <div className="flex flex-wrap gap-1">{names.map((n, i) => <Badge key={i} color="teal">{n}</Badge>)}{!names.length && <span className="text-slate-300 text-xs">—</span>}</div>;
    }, exportValue: (r) => subjNamesForMonth(r).join(", ") },
    { key: "phaiThuThang", label: "Phải thu (tháng)", render: (r) => { const p = monthPeriods(r); return p.length ? vnd(p.reduce((s, x) => s + x.phaiThu, 0)) : <span className="text-slate-300 text-xs">—</span>; }, exportValue: (r) => monthPeriods(r).reduce((s, x) => s + x.phaiThu, 0) },
    { key: "daThuThang", label: "Đã thu (tháng)", render: (r) => { const p = monthPeriods(r); return p.length ? vnd(p.reduce((s, x) => s + x.daThu, 0)) : <span className="text-slate-300 text-xs">—</span>; }, exportValue: (r) => monthPeriods(r).reduce((s, x) => s + x.daThu, 0) },
    { key: "conThieuThang", label: "Còn thiếu (tháng)", render: (r) => {
      const p = monthPeriods(r);
      if (!p.length) return <span className="text-slate-300 text-xs">—</span>;
      const con = p.reduce((s, x) => s + Math.max(0, x.phaiThu - x.daThu), 0);
      return <span className={con > 0 ? "text-rose-600 font-medium" : "text-teal-700"}>{vnd(con)}</span>;
    }, exportValue: (r) => monthPeriods(r).reduce((s, x) => s + Math.max(0, x.phaiThu - x.daThu), 0) },
    { key: "trangThaiThang", label: "Trạng thái (tháng)", render: (r) => {
      const p = monthPeriods(r);
      if (!p.length) return <span className="text-slate-300 text-xs">—</span>;
      const phaiThu = p.reduce((s, x) => s + x.phaiThu, 0), daThu = p.reduce((s, x) => s + x.daThu, 0);
      return statusBadge(periodStatus(phaiThu, daThu));
    }, exportValue: (r) => {
      const p = monthPeriods(r);
      if (!p.length) return "";
      const phaiThu = p.reduce((s, x) => s + x.phaiThu, 0), daThu = p.reduce((s, x) => s + x.daThu, 0);
      return periodStatus(phaiThu, daThu);
    } },
    { key: "tenPH", label: "Phụ huynh" },
    { key: "sdtPH", label: "SĐT PH" },
    { key: "trangThai", label: "Trạng thái", render: (r) => statusBadge(r.trangThai), exportValue: (r) => r.trangThai },
    ...(editable ? [{ key: "actions", label: "", render: (r) => (
      <div className="flex gap-1">
        <IconBtn icon={Pencil} tone="teal" title="Sửa" onClick={() => setModal({ mode: "edit", data: r })} />
        <IconBtn icon={Trash2} tone="rose" title="Xoá" onClick={() => setDel(r)} />
      </div>
    ) }] : []),
  ];

  return (
    <div>
      <SectionHeader title="Danh sách học sinh" desc={`${rows.length} học sinh — chọn tháng để xem môn đăng ký và tình hình học phí đúng theo tháng đó`}
        actions={editable && (
          <>
            <button onClick={() => setImportOpen(true)} className="flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-sm font-medium border border-slate-200 text-slate-600 hover:bg-slate-50 shrink-0"><FileUp size={15} /> Nhập từ Excel</button>
            <PrimaryButton onClick={() => setModal({ mode: "add", data: null })}>Thêm học sinh</PrimaryButton>
          </>
        )} />
      <DataTable columns={columns} rows={rows} searchKeys={["maHS", "hoTen", "sdtPH", "tenPH"]} exportName="HocSinh" emptyTitle="Chưa có học sinh nào"
        filterBar={
          <div className="flex gap-2">
            <Select value={classFilter} onChange={(e) => setClassFilter(e.target.value)} options={[{ value: "", label: "Tất cả lớp" }, ...classes.map((c) => ({ value: c.id, label: c.tenLop }))]} className="sm:w-48" />
            <TextInput type="month" value={monthFilter} onChange={(e) => setMonthFilter(e.target.value)} className="w-40" />
          </div>
        } />

      <Modal open={!!modal} onClose={() => setModal(null)} title={modal?.mode === "add" ? "Thêm học sinh" : "Sửa thông tin học sinh"} wide>
        {modal && <StudentForm initial={modal.data} classes={classes} students={students} subjects={subjects} enrollments={enrollments}
          onCancel={() => setModal(null)}
          onSubmit={(data) => {
            const { enrolledSubjectIds, ...stuData } = data;
            let studentId = stuData.id;
            if (modal.mode === "add") { studentId = uid("hs"); setStudents((p) => [...p, { ...stuData, id: studentId }]); }
            else setStudents((p) => p.map((s) => s.id === stuData.id ? stuData : s));
            setEnrollments((prev) => [
              ...prev.filter((e) => e.hocSinhId !== studentId),
              ...(enrolledSubjectIds || []).map((monHocId) => ({ id: uid("dk"), hocSinhId: studentId, lopId: stuData.lopId, monHocId, ngayDangKy: todayISO() })),
            ]);
            setModal(null);
          }} />}
      </Modal>
      <ConfirmDialog open={!!del} onCancel={() => setDel(null)} text={`Xoá học sinh "${del?.hoTen}"? Dữ liệu đăng ký môn học và học phí liên quan sẽ vẫn được lưu ở lịch sử thu tiền.`}
        onConfirm={() => { setStudents((p) => p.filter((s) => s.id !== del.id)); setEnrollments((p) => p.filter((e) => e.hocSinhId !== del.id)); setDel(null); }} />

      <Modal open={!!detail} onClose={() => setDetail(null)} title={`Hồ sơ học sinh — ${detail?.hoTen || ""}`} wide>
        {detail && <StudentProfile student={detail} classes={classes} subjects={subjects} enrollments={enrollments} tuitionConfig={tuitionConfig} paymentAllocations={paymentAllocations} />}
      </Modal>

      <Modal open={importOpen} onClose={() => setImportOpen(false)} title="Nhập danh sách học sinh từ Excel" wide>
        {importOpen && <ImportStudentsModal students={students} classes={classes} subjects={subjects}
          setStudents={setStudents} setClasses={setClasses} setSubjects={setSubjects} setEnrollments={setEnrollments}
          onClose={() => setImportOpen(false)} />}
      </Modal>
    </div>
  );
}

function StudentForm({ initial, classes, students, subjects, enrollments, onCancel, onSubmit }) {
  const [f, setF] = useState(initial || { maHS: "", hoTen: "", gioiTinh: "Nam", ngaySinh: "", lopId: classes[0]?.id || "", tenPH: "", sdtPH: "", diaChi: "", ngayNhapHoc: todayISO(), trangThai: "Đang học" });
  const [subIds, setSubIds] = useState(() => initial ? enrollments.filter((e) => e.hocSinhId === initial.id).map((e) => e.monHocId) : []);
  const [errs, setErrs] = useState({});
  const lop = classes.find((c) => c.id === f.lopId);
  const available = subjects.filter((s) => (lop?.subjectIds || []).includes(s.id));
  function set(k, v) { setF((p) => ({ ...p, [k]: v })); if (k === "lopId") setSubIds([]); }
  function toggleSub(id) { setSubIds((prev) => prev.includes(id) ? prev.filter((x) => x !== id) : (prev.length >= 3 ? prev : [...prev, id])); }
  function validate() {
    const e = {};
    if (!f.maHS.trim()) e.maHS = "Bắt buộc nhập mã học sinh";
    else if (students.some((s) => s.maHS === f.maHS && s.id !== f.id)) e.maHS = "Mã học sinh đã tồn tại";
    if (!f.hoTen.trim()) e.hoTen = "Bắt buộc nhập họ tên";
    if (!f.lopId) e.lopId = "Chọn lớp học";
    if (f.sdtPH.trim() && !/^0\d{9,10}$/.test(f.sdtPH)) e.sdtPH = "SĐT không hợp lệ";
    if (subIds.length < 1 || subIds.length > 3) e.subIds = "Chọn từ 1 đến 3 môn học (trong số môn lớp đang mở)";
    setErrs(e);
    return Object.keys(e).length === 0;
  }
  return (
    <div>
      <div className="grid sm:grid-cols-2 gap-x-4">
        <Field label="Mã học sinh" required error={errs.maHS}><TextInput value={f.maHS} onChange={(e) => set("maHS", e.target.value)} placeholder="HS0001" /></Field>
        <Field label="Họ và tên" required error={errs.hoTen}><TextInput value={f.hoTen} onChange={(e) => set("hoTen", e.target.value)} /></Field>
        <Field label="Giới tính"><Select value={f.gioiTinh} onChange={(e) => set("gioiTinh", e.target.value)} options={[{ value: "Nam", label: "Nam" }, { value: "Nữ", label: "Nữ" }]} /></Field>
        <Field label="Ngày sinh"><TextInput type="date" value={f.ngaySinh} onChange={(e) => set("ngaySinh", e.target.value)} /></Field>
        <Field label="Lớp học" required error={errs.lopId}><Select value={f.lopId} onChange={(e) => set("lopId", e.target.value)} options={classes.map((c) => ({ value: c.id, label: c.tenLop }))} /></Field>
        <Field label="Trạng thái"><Select value={f.trangThai} onChange={(e) => set("trangThai", e.target.value)} options={[{ value: "Đang học", label: "Đang học" }, { value: "Nghỉ học", label: "Nghỉ học" }]} /></Field>
        <Field label="Tên phụ huynh"><TextInput value={f.tenPH} onChange={(e) => set("tenPH", e.target.value)} /></Field>
        <Field label="SĐT phụ huynh" error={errs.sdtPH}><TextInput value={f.sdtPH} onChange={(e) => set("sdtPH", e.target.value)} placeholder="09xxxxxxxx" /></Field>
        <Field label="Ngày nhập học"><TextInput type="date" value={f.ngayNhapHoc} onChange={(e) => set("ngayNhapHoc", e.target.value)} /></Field>
        <Field label="Địa chỉ"><TextInput value={f.diaChi} onChange={(e) => set("diaChi", e.target.value)} /></Field>
      </div>
      <Field label="Môn học đăng ký (1–3 môn, trong số môn lớp đang mở)" error={errs.subIds}>
        {available.length === 0 ? (
          <p className="text-xs text-amber-600">Lớp này chưa mở môn học nào. Vào mục Lớp học để cấu hình môn học trước.</p>
        ) : (
          <div className="flex flex-wrap gap-2">
            {available.map((s) => (
              <label key={s.id} className={cx("flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg border text-sm cursor-pointer", subIds.includes(s.id) ? "bg-teal-50 border-teal-300 text-teal-700" : "border-slate-200 text-slate-600")}>
                <input type="checkbox" className="hidden" checked={subIds.includes(s.id)} onChange={() => toggleSub(s.id)} />{s.ten}
              </label>
            ))}
          </div>
        )}
      </Field>
      <div className="flex justify-end gap-2 pt-2 border-t border-slate-100 mt-2">
        <button onClick={onCancel} className="px-3.5 py-2 rounded-lg text-sm border border-slate-200 hover:bg-slate-50">Huỷ</button>
        <button onClick={() => validate() && onSubmit({ ...f, enrolledSubjectIds: subIds })} className="px-3.5 py-2 rounded-lg text-sm bg-teal-700 text-white hover:bg-teal-800">Lưu</button>
      </div>
    </div>
  );
}

function StudentProfile({ student, classes, subjects, enrollments, tuitionConfig, paymentAllocations }) {
  const lop = classes.find((c) => c.id === student.lopId);
  const myEnroll = enrollments.filter((e) => e.hocSinhId === student.id);
  const periods = studentPeriods(student, enrollments, tuitionConfig, todayISO().slice(0, 7), classes).map((p) => ({ ...p, daThu: allocatedOf(p.hocSinhId, p.lopId, p.monHocId, p.thang, paymentAllocations) }));
  const debt = periods.reduce((s, p) => s + Math.max(0, p.phaiThu - p.daThu), 0);
  return (
    <div className="space-y-4 text-sm">
      <div className="grid sm:grid-cols-3 gap-3">
        <div><p className="text-slate-400 text-xs">Lớp</p><p className="font-medium">{lop?.tenLop}</p></div>
        <div><p className="text-slate-400 text-xs">Phụ huynh</p><p className="font-medium">{student.tenPH} — {student.sdtPH}</p></div>
        <div><p className="text-slate-400 text-xs">Trạng thái</p>{statusBadge(student.trangThai)}</div>
      </div>
      <div>
        <p className="text-slate-400 text-xs mb-1">Môn đăng ký</p>
        <div className="flex flex-wrap gap-1.5">
          {myEnroll.map((e) => <Badge key={e.id} color="teal">{subjects.find((s) => s.id === e.monHocId)?.ten}</Badge>)}
          {!myEnroll.length && <span className="text-xs text-slate-400">Chưa đăng ký môn nào</span>}
        </div>
      </div>
      <div className="grid grid-cols-1 gap-3">
        <Card className="p-3 text-center"><p className="text-xs text-slate-400">Công nợ luỹ kế</p><p className="text-lg font-semibold text-rose-600">{vnd(debt)}</p></Card>
      </div>
    </div>
  );
}

/* ============================================================================
   NHẬP DANH SÁCH HỌC SINH TỪ EXCEL
============================================================================ */
const IMPORT_KNOWN_HEADERS = {
  mahs: "maHS", mahocsinh: "maHS", mahs001: "maHS",
  hoten: "hoTen", hovaten: "hoTen", tenhocsinh: "hoTen",
  gioitinh: "gioiTinh",
  ngaysinh: "ngaySinh",
  sdtphuhuynh: "sdtPH", sodienthoaiphuhuynh: "sdtPH", sdt: "sdtPH", sodienthoai: "sdtPH",
  tenphuhuynh: "tenPH", phuhuynh: "tenPH", hotenphuhuynh: "tenPH",
  lop: "lopRaw", malop: "lopRaw", tenlop: "lopRaw", lophoc: "lopRaw",
  ngaynhaphoc: "ngayNhapHoc", ngayvaohoc: "ngayNhapHoc",
  diachi: "diaChi",
};
function normalizeHeader(s) {
  return String(s ?? "")
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/đ/gi, "d")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}
const IMPORT_CHECK_VALUES = ["✓", "✔", "x", "có", "co", "1", "true", "yes", "y", "dangky", "đăngký"];
function isMarkChecked(v) {
  if (v === null || v === undefined || v === "") return false;
  if (typeof v === "number") return v === 1;
  if (typeof v === "boolean") return v === true;
  const s = String(v).trim().toLowerCase();
  return IMPORT_CHECK_VALUES.includes(s) || IMPORT_CHECK_VALUES.includes(normalizeHeader(s));
}
function parseImportDate(v) {
  if (v === null || v === undefined || v === "") return "";
  if (v instanceof Date && !isNaN(v)) return v.toISOString().slice(0, 10);
  if (typeof v === "number") {
    const d = new Date(Math.round((v - 25569) * 86400 * 1000));
    if (!isNaN(d)) return d.toISOString().slice(0, 10);
  }
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) return `${m[1]}-${String(m[2]).padStart(2, "0")}-${String(m[3]).padStart(2, "0")}`;
  m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m) return `${m[3]}-${String(m[1]).padStart(2, "0")}-${String(m[2]).padStart(2, "0")}`;
  return "";
}
function downloadStudentImportTemplate() {
  const rows = [
    { "Mã HS": "HS0001", "Họ tên": "Nguyễn Văn An", "SĐT phụ huynh": "0901234567", "Tên phụ huynh": "Nguyễn Văn Bình", "Lớp": "TH-T4A", "Ngày nhập học": "2025-09-01", "Toán": "x", "Tiếng Anh": "", "Ngữ văn": "x" },
    { "Mã HS": "HS0002", "Họ tên": "Trần Thị Bích", "SĐT phụ huynh": "0912345678", "Tên phụ huynh": "Trần Văn Cường", "Lớp": "TH-T4A", "Ngày nhập học": "2025-09-01", "Toán": "x", "Tiếng Anh": "x", "Ngữ văn": "" },
  ];
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "MauNhapHocSinh");
  download("Mau_Nhap_Hoc_Sinh.xlsx", wb);
}

function ImportStudentsModal({ students, classes, subjects, setStudents, setClasses, setSubjects, setEnrollments, onClose }) {
  const [fileName, setFileName] = useState("");
  const [parseError, setParseError] = useState("");
  const [previewRows, setPreviewRows] = useState(null); // null = chưa chọn file
  const [subjectHeaders, setSubjectHeaders] = useState([]);
  const [importing, setImporting] = useState(false);
  const [result, setResult] = useState(null);
  const fileInputRef = useRef(null);

  function handleFile(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    setFileName(file.name);
    setParseError("");
    setResult(null);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const wb = XLSX.read(data, { type: "array", cellDates: true });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const rawRows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        if (!rawRows.length) { setParseError("File Excel không có dữ liệu (hoặc thiếu dòng tiêu đề)."); setPreviewRows(null); return; }
        buildPreview(rawRows);
      } catch (err) {
        setParseError("Không thể đọc file. Vui lòng kiểm tra định dạng file Excel (.xlsx/.xls).");
        setPreviewRows(null);
      }
    };
    reader.onerror = () => setParseError("Không thể đọc file đã chọn.");
    reader.readAsArrayBuffer(file);
  }

  function buildPreview(rawRows) {
    const headers = Object.keys(rawRows[0] || {});
    const fieldOfHeader = {}; // header -> known field name
    const subjHeaders = [];
    headers.forEach((h) => {
      const norm = normalizeHeader(h);
      if (IMPORT_KNOWN_HEADERS[norm]) fieldOfHeader[h] = IMPORT_KNOWN_HEADERS[norm];
      else if (h.trim()) subjHeaders.push(h);
    });
    setSubjectHeaders(subjHeaders);

    const existingCodes = new Set(students.map((s) => s.maHS.trim().toLowerCase()));
    const seenInFile = new Set();

    const rows = rawRows.map((raw, idx) => {
      let maHS = "", hoTen = "", gioiTinh = "", ngaySinh = "", sdtPH = "", tenPH = "", lopRaw = "", ngayNhapHoc = "", diaChi = "";
      Object.entries(fieldOfHeader).forEach(([header, field]) => {
        const val = raw[header];
        if (field === "maHS") maHS = String(val ?? "").trim();
        else if (field === "hoTen") hoTen = String(val ?? "").trim();
        else if (field === "gioiTinh") gioiTinh = String(val ?? "").trim();
        else if (field === "ngaySinh") ngaySinh = parseImportDate(val);
        else if (field === "sdtPH") sdtPH = String(val ?? "").trim();
        else if (field === "tenPH") tenPH = String(val ?? "").trim();
        else if (field === "lopRaw") lopRaw = String(val ?? "").trim();
        else if (field === "ngayNhapHoc") ngayNhapHoc = parseImportDate(val);
        else if (field === "diaChi") diaChi = String(val ?? "").trim();
      });
      const monDangKy = subjHeaders.filter((h) => isMarkChecked(raw[h]));

      const errs = [];
      if (!maHS) errs.push("Thiếu Mã HS");
      if (!hoTen) errs.push("Thiếu Họ tên");
      if (!lopRaw) errs.push("Thiếu Lớp");
      const codeKey = maHS.trim().toLowerCase();
      let dup = false;
      if (maHS) {
        if (existingCodes.has(codeKey)) { errs.push("Mã HS đã tồn tại trong hệ thống"); dup = true; }
        else if (seenInFile.has(codeKey)) { errs.push("Mã HS trùng trong file"); dup = true; }
        else seenInFile.add(codeKey);
      }

      return {
        _row: idx + 2, // dòng excel (tính cả header)
        maHS, hoTen, gioiTinh: gioiTinh || "Nam", ngaySinh, sdtPH, tenPH, lopRaw,
        ngayNhapHoc: ngayNhapHoc || todayISO(), diaChi, monDangKy,
        status: errs.length ? "error" : "ok",
        errors: errs,
      };
    });
    setPreviewRows(rows);
  }

  const validRows = previewRows ? previewRows.filter((r) => r.status === "ok") : [];
  const errorCount = previewRows ? previewRows.length - validRows.length : 0;

  function commitImport() {
    if (!validRows.length) return;
    setImporting(true);

    let classList = classes.map((c) => ({ ...c, subjectIds: [...(c.subjectIds || [])] }));
    let subjectList = subjects.map((s) => ({ ...s }));
    const newStudents = [];
    const newEnrollments = [];

    function findOrCreateClass(raw) {
      const key = raw.trim().toLowerCase();
      let lop = classList.find((c) => (c.maLop || "").trim().toLowerCase() === key || (c.tenLop || "").trim().toLowerCase() === key);
      if (!lop) {
        lop = { id: uid("lop"), maLop: raw.trim(), tenLop: raw.trim(), khoiId: KHOI[0].id, gvId: "", troId: "", phongId: "", hocPhiBuoi: 0, mon: "", subjectIds: [], lich: [] };
        classList = [...classList, lop];
      }
      return lop;
    }
    function findOrCreateSubject(name) {
      const key = name.trim().toLowerCase();
      let mon = subjectList.find((s) => s.ten.trim().toLowerCase() === key);
      if (!mon) {
        mon = { id: uid("mon"), ten: name.trim() };
        subjectList = [...subjectList, mon];
      }
      return mon;
    }

    validRows.forEach((r) => {
      const lop = findOrCreateClass(r.lopRaw);
      const studentId = uid("hs");
      newStudents.push({
        id: studentId, maHS: r.maHS, hoTen: r.hoTen, gioiTinh: r.gioiTinh || "Nam",
        ngaySinh: r.ngaySinh || "", lopId: lop.id, tenPH: r.tenPH, sdtPH: r.sdtPH,
        diaChi: r.diaChi || "", ngayNhapHoc: r.ngayNhapHoc, trangThai: "Đang học",
      });
      r.monDangKy.forEach((subjName) => {
        const mon = findOrCreateSubject(subjName);
        if (!lop.subjectIds.includes(mon.id)) lop.subjectIds = [...lop.subjectIds, mon.id];
        newEnrollments.push({ id: uid("dk"), hocSinhId: studentId, lopId: lop.id, monHocId: mon.id, ngayDangKy: r.ngayNhapHoc });
      });
    });

    setClasses(classList);
    setSubjects(subjectList);
    setStudents((prev) => [...prev, ...newStudents]);
    setEnrollments((prev) => [...prev, ...newEnrollments]);

    setImporting(false);
    setResult({ imported: newStudents.length, skipped: errorCount });
    setPreviewRows(null);
    setFileName("");
    if (fileInputRef.current) fileInputRef.current.value = "";
  }

  return (
    <div className="space-y-4">
      <Card className="p-4 bg-teal-50/60 border-teal-100">
        <p className="text-sm text-teal-800 flex items-start gap-2"><Info size={16} className="mt-0.5 shrink-0" />
          Hỗ trợ các cột: <b>Mã HS, Họ tên, SĐT phụ huynh, Tên phụ huynh, Lớp, Ngày nhập học</b> và các cột <b>môn học</b> (mỗi cột 1 môn — đánh dấu ✓ / Có / x / 1 để tự động đăng ký môn đó cho học sinh). Nếu Lớp hoặc Môn học chưa có trong hệ thống, ứng dụng sẽ tự tạo mới.
        </p>
      </Card>

      <div className="flex flex-wrap items-center gap-2">
        <button onClick={downloadStudentImportTemplate} className="flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-sm border border-slate-200 hover:bg-slate-50"><Download size={15} /> Tải file Excel mẫu</button>
        <label className="flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-sm bg-teal-700 text-white hover:bg-teal-800 cursor-pointer">
          <Upload size={15} /> Chọn file Excel (.xlsx/.xls)
          <input ref={fileInputRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={handleFile} />
        </label>
        {fileName && <span className="text-xs text-slate-500">Đã chọn: {fileName}</span>}
      </div>

      {parseError && (
        <Card className="p-3 border-rose-200 bg-rose-50 text-sm text-rose-700 flex items-center gap-2"><AlertCircle size={16} /> {parseError}</Card>
      )}

      {result && (
        <Card className="p-3 border-emerald-200 bg-emerald-50 text-sm text-emerald-700 flex items-center gap-2">
          <CheckCircle2 size={16} /> Đã nhập thành công {result.imported} học sinh{result.skipped > 0 ? ` — bỏ qua ${result.skipped} dòng lỗi/trùng mã` : ""}.
        </Card>
      )}

      {previewRows && (
        <div>
          <div className="flex items-center justify-between mb-2">
            <p className="text-sm font-medium text-slate-700">Xem trước — {previewRows.length} dòng ({validRows.length} hợp lệ{errorCount > 0 ? `, ${errorCount} lỗi/trùng mã` : ""})</p>
          </div>
          <div className="border border-slate-200 rounded-lg overflow-auto max-h-80">
            <table className="w-full text-sm min-w-[760px]">
              <thead className="sticky top-0 bg-slate-50 text-slate-500 text-xs uppercase">
                <tr>
                  <th className="px-2 py-2 text-left">Dòng</th>
                  <th className="px-2 py-2 text-left">Mã HS</th>
                  <th className="px-2 py-2 text-left">Họ tên</th>
                  <th className="px-2 py-2 text-left">Lớp</th>
                  <th className="px-2 py-2 text-left">Môn đăng ký</th>
                  <th className="px-2 py-2 text-left">Trạng thái</th>
                </tr>
              </thead>
              <tbody>
                {previewRows.map((r, i) => (
                  <tr key={i} className={cx("border-t border-slate-100", r.status === "error" && "bg-rose-50/60")}>
                    <td className="px-2 py-2 text-slate-400">{r._row}</td>
                    <td className="px-2 py-2">{r.maHS || <span className="text-slate-300">—</span>}</td>
                    <td className="px-2 py-2">{r.hoTen || <span className="text-slate-300">—</span>}</td>
                    <td className="px-2 py-2">{r.lopRaw || <span className="text-slate-300">—</span>}</td>
                    <td className="px-2 py-2">
                      <div className="flex flex-wrap gap-1">{r.monDangKy.map((n, j) => <Badge key={j} color="teal">{n}</Badge>)}{!r.monDangKy.length && <span className="text-slate-300 text-xs">—</span>}</div>
                    </td>
                    <td className="px-2 py-2">
                      {r.status === "ok" ? <Badge color="green">Hợp lệ</Badge> : <Badge color="red">{r.errors.join("; ")}</Badge>}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      <div className="flex justify-end gap-2 pt-2 border-t border-slate-100">
        <button onClick={onClose} className="px-3.5 py-2 rounded-lg text-sm border border-slate-200 hover:bg-slate-50">Đóng</button>
        {previewRows && (
          <button disabled={!validRows.length || importing} onClick={commitImport}
            className="px-3.5 py-2 rounded-lg text-sm bg-teal-700 text-white hover:bg-teal-800 disabled:opacity-50 disabled:cursor-not-allowed">
            {importing ? "Đang nhập..." : `Nhập ${validRows.length} học sinh hợp lệ`}
          </button>
        )}
      </div>
    </div>
  );
}

/* ============================================================================
   CLASSES
============================================================================ */
function ClassesPage({ classes, setClasses, students, subjects, setSubjects, enrollments, paymentAllocations, tuitionConfig, teachers, assistants, rooms, role }) {
  const editable = can(role, "classes", "full");
  const [modal, setModal] = useState(null);
  const [del, setDel] = useState(null);
  const [detailClass, setDetailClass] = useState(null);
  const nameOf = (arr, id, key = "hoTen") => arr.find((x) => x.id === id)?.[key] || "—";
  const subjNames = (ids) => (ids || []).map((id) => subjects.find((s) => s.id === id)?.ten).filter(Boolean);
  const columns = [
    { key: "maLop", label: "Mã lớp", sortable: true },
    { key: "tenLop", label: "Tên lớp", sortable: true, render: (r) => <button onClick={() => setDetailClass(r)} className="text-teal-700 font-medium hover:underline">{r.tenLop}</button> },
    { key: "khoiId", label: "Khối", render: (r) => KHOI.find((k) => k.id === r.khoiId)?.ten, exportValue: (r) => KHOI.find((k) => k.id === r.khoiId)?.ten },
    { key: "ngayBatDau", label: "Ngày bắt đầu", sortable: true, render: (r) => r.ngayBatDau ? fmtDate(r.ngayBatDau) : "—", exportValue: (r) => r.ngayBatDau || "" },
    { key: "siso", label: "Sĩ số", render: (r) => students.filter((s) => s.lopId === r.id && s.trangThai === "Đang học").length, exportValue: (r) => students.filter((s) => s.lopId === r.id && s.trangThai === "Đang học").length },
    { key: "subjectIds", label: "Môn đang mở", render: (r) => (
      <div className="flex flex-wrap gap-1">{subjNames(r.subjectIds).map((n, i) => <Badge key={i} color="teal">{n}</Badge>)}{!subjNames(r.subjectIds).length && <span className="text-slate-300 text-xs">—</span>}</div>
    ), exportValue: (r) => subjNames(r.subjectIds).join(", ") },
    { key: "gvId", label: "Giáo viên", render: (r) => nameOf(teachers, r.gvId), exportValue: (r) => nameOf(teachers, r.gvId) },
    { key: "phongId", label: "Phòng", render: (r) => nameOf(rooms, r.phongId, "tenPhong"), exportValue: (r) => nameOf(rooms, r.phongId, "tenPhong") },
    { key: "lich", label: "Lịch học", render: (r) => r.lich.map((l) => WEEKDAYS[l.thu === 0 ? 0 : l.thu - 1]).join(", "), exportValue: (r) => r.lich.map((l) => `${WEEKDAYS[l.thu === 0 ? 0 : l.thu - 1]} ${l.gioBD}-${l.gioKT}`).join(" | ") },
    ...(editable ? [{ key: "actions", label: "", render: (r) => (
      <div className="flex gap-1"><IconBtn icon={Eye} tone="slate" title="Xem học sinh" onClick={() => setDetailClass(r)} /><IconBtn icon={Pencil} tone="teal" onClick={() => setModal({ mode: "edit", data: r })} /><IconBtn icon={Trash2} tone="rose" onClick={() => setDel(r)} /></div>
    ) }] : []),
  ];
  return (
    <div>
      <SectionHeader title="Danh sách lớp học" desc={`${classes.length} lớp — bấm vào tên lớp để xem danh sách học sinh, môn đăng ký và tình hình học phí`} actions={editable && <PrimaryButton onClick={() => setModal({ mode: "add", data: null })}>Thêm lớp</PrimaryButton>} />
      <DataTable columns={columns} rows={classes} searchKeys={["maLop", "tenLop"]} exportName="LopHoc" emptyTitle="Chưa có lớp học nào" />
      <Modal open={!!modal} onClose={() => setModal(null)} title={modal?.mode === "add" ? "Thêm lớp học" : "Sửa lớp học"} wide>
        {modal && <ClassForm initial={modal.data} classes={classes} teachers={teachers} assistants={assistants} rooms={rooms} subjects={subjects} setSubjects={setSubjects}
          onCancel={() => setModal(null)}
          onSubmit={(data) => { if (modal.mode === "add") setClasses((p) => [...p, { ...data, id: uid("lop") }]); else setClasses((p) => p.map((c) => c.id === data.id ? data : c)); setModal(null); }} />}
      </Modal>
      <ConfirmDialog open={!!del} onCancel={() => setDel(null)} text={`Xoá lớp "${del?.tenLop}"? Học sinh trong lớp cần được chuyển lớp khác trước.`}
        onConfirm={() => { if (students.some((s) => s.lopId === del.id && s.trangThai === "Đang học")) { alert("Không thể xoá: lớp vẫn còn học sinh đang học."); setDel(null); return; } setClasses((p) => p.filter((c) => c.id !== del.id)); setDel(null); }} />

      <Modal open={!!detailClass} onClose={() => setDetailClass(null)} title={`Lớp ${detailClass?.tenLop || ""} — Danh sách học sinh, môn đăng ký & học phí`} wide>
        {detailClass && <ClassRoster lop={detailClass} classes={classes} students={students} subjects={subjects} enrollments={enrollments} paymentAllocations={paymentAllocations} tuitionConfig={tuitionConfig} />}
      </Modal>
    </div>
  );
}
function ClassRoster({ lop, classes, students, subjects, enrollments, paymentAllocations, tuitionConfig }) {
  const roster = students.filter((s) => s.lopId === lop.id);
  const lopSubjects = subjects.filter((s) => (lop.subjectIds || []).includes(s.id));
  const thisMonth = todayISO().slice(0, 7);
  // Nếu lớp có ngày bắt đầu ở tương lai, mặc định mở đúng tháng khai giảng thay vì tháng hiện tại.
  const defaultMonth = lop.ngayBatDau && lop.ngayBatDau.slice(0, 7) > thisMonth ? lop.ngayBatDau.slice(0, 7) : thisMonth;
  const [thang, setThang] = useState(defaultMonth);
  const [detailStu, setDetailStu] = useState(null);

  const rows = roster.map((s) => {
    const myPeriods = periodsForMonth(s, enrollments, tuitionConfig, thang, classes).filter((p) => p.lopId === lop.id);
    const myIds = myPeriods.map((p) => p.monHocId);
    const soMon = myIds.length;
    const phaiDong = myPeriods.reduce((sum, p) => sum + p.phaiThu, 0);
    const daDong = myIds.reduce((sum, monHocId) => sum + allocatedOf(s.id, lop.id, monHocId, thang, paymentAllocations), 0);
    const trangThaiDong = periodStatus(phaiDong, daDong);
    return { student: s, myIds, soMon, phaiDong, daDong: Math.min(daDong, phaiDong), trangThaiDong };
  });

  return (
    <div>
      <div className="flex flex-wrap items-center gap-1.5 mb-3">
        <span className="text-sm text-slate-500 mr-1">Môn đang mở:</span>
        {lopSubjects.map((s) => <Badge key={s.id} color="teal">{s.ten}</Badge>)}
        {!lopSubjects.length && <span className="text-xs text-slate-400">Chưa mở môn nào</span>}
        {lop.ngayBatDau && <span className="text-xs text-slate-400">· Khai giảng {fmtDate(lop.ngayBatDau)}</span>}
        <div className="flex items-center gap-1.5 ml-auto">
          <span className="text-xs text-slate-400">Học phí tháng</span>
          <TextInput type="month" value={thang} onChange={(e) => setThang(e.target.value)} className="w-36" />
        </div>
      </div>
      <div className="border border-slate-200 rounded-lg overflow-hidden overflow-x-auto">
        <table className="w-full text-sm min-w-[780px]">
          <thead>
            <tr className="bg-slate-50 text-slate-500 text-xs uppercase">
              <th className="px-3 py-2 text-left">STT</th>
              <th className="px-3 py-2 text-left">Họ và tên học sinh</th>
              <th className="px-3 py-2 text-left">SĐT phụ huynh</th>
              {lopSubjects.map((s) => <th key={s.id} className="px-2 py-2 text-center">{s.ten}</th>)}
              <th className="px-2 py-2 text-center">Số môn đăng ký</th>
              <th className="px-3 py-2 text-right">Học phí phải đóng</th>
              <th className="px-3 py-2 text-right">Số tiền đã đóng</th>
              <th className="px-3 py-2 text-left">Trạng thái</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((r, i) => (
              <tr key={r.student.id} className="border-t border-slate-100 hover:bg-slate-50/70">
                <td className="px-3 py-2">{i + 1}</td>
                <td className="px-3 py-2"><button onClick={() => setDetailStu(r)} className="text-teal-700 font-medium hover:underline">{r.student.hoTen}</button></td>
                <td className="px-3 py-2">{r.student.sdtPH}</td>
                {lopSubjects.map((s) => <td key={s.id} className="px-2 py-2 text-center font-medium text-teal-600">{r.myIds.includes(s.id) ? "X" : ""}</td>)}
                <td className="px-2 py-2 text-center">{r.soMon}</td>
                <td className="px-3 py-2 text-right">{vnd(r.phaiDong)}</td>
                <td className="px-3 py-2 text-right text-teal-700 font-medium">{vnd(r.daDong)}</td>
                <td className="px-3 py-2">{r.soMon > 0 ? statusBadge(r.trangThaiDong) : <span className="text-xs text-slate-300">Chưa đến kỳ</span>}</td>
              </tr>
            ))}
          </tbody>
        </table>
        {roster.length === 0 && <EmptyState title="Lớp chưa có học sinh" />}
      </div>

      <Modal open={!!detailStu} onClose={() => setDetailStu(null)} title={`Lịch sử đóng tiền — ${detailStu?.student?.hoTen || ""}`}>
        {detailStu && <StudentPaymentHistory student={detailStu.student} classes={classes} subjects={subjects} enrollments={enrollments} paymentAllocations={paymentAllocations} tuitionConfig={tuitionConfig} />}
      </Modal>
    </div>
  );
}
function StudentPaymentHistory({ student, classes, subjects, enrollments, paymentAllocations, tuitionConfig }) {
  const periods = studentPeriods(student, enrollments, tuitionConfig, todayISO().slice(0, 7), classes)
    .map((p) => ({ ...p, daThu: allocatedOf(p.hocSinhId, p.lopId, p.monHocId, p.thang, paymentAllocations) }))
    .sort((a, b) => b.thang.localeCompare(a.thang) || a.monHocId.localeCompare(b.monHocId));
  return (
    <div className="space-y-3 text-sm">
      {periods.length === 0 ? <EmptyState title="Chưa có kỳ học phí nào" desc="Học sinh chưa đăng ký môn học." /> : (
        <div className="border border-slate-200 rounded-lg overflow-hidden">
          <table className="w-full text-sm">
            <thead><tr className="bg-slate-50 text-slate-500 text-xs uppercase">
              <th className="px-2 py-2 text-left">Tháng</th><th className="px-2 py-2 text-left">Môn</th><th className="px-2 py-2 text-right">Phải đóng</th><th className="px-2 py-2 text-right">Đã đóng</th><th className="px-2 py-2 text-left">Trạng thái</th>
            </tr></thead>
            <tbody>
              {periods.map((p, i) => (
                <tr key={i} className="border-t border-slate-100">
                  <td className="px-2 py-2">{p.thang}</td>
                  <td className="px-2 py-2">{subjects.find((s) => s.id === p.monHocId)?.ten}</td>
                  <td className="px-2 py-2 text-right">{vnd(p.phaiThu)}</td>
                  <td className="px-2 py-2 text-right text-teal-700">{vnd(Math.min(p.daThu, p.phaiThu))}</td>
                  <td className="px-2 py-2">{statusBadge(periodStatus(p.phaiThu, p.daThu))}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

function ClassForm({ initial, classes, teachers, assistants, rooms, subjects, setSubjects, onCancel, onSubmit }) {
  const [f, setF] = useState(initial || { maLop: "", tenLop: "", khoiId: KHOI[0].id, gvId: teachers[0]?.id, troId: assistants[0]?.id, phongId: rooms[0]?.id, hocPhiBuoi: 100000, ngayBatDau: todayISO(), mon: "", subjectIds: [], lich: [{ thu: 2, gioBD: "18:00", gioKT: "19:30" }] });
  const [errs, setErrs] = useState({});
  const [newSubj, setNewSubj] = useState("");
  function set(k, v) { setF((p) => ({ ...p, [k]: v })); }
  function setSlot(i, k, v) { setF((p) => ({ ...p, lich: p.lich.map((l, idx) => idx === i ? { ...l, [k]: v } : l) })); }
  function addSlot() { setF((p) => ({ ...p, lich: [...p.lich, { thu: 2, gioBD: "18:00", gioKT: "19:30" }] })); }
  function rmSlot(i) { setF((p) => ({ ...p, lich: p.lich.filter((_, idx) => idx !== i) })); }
  function toggleSubj(id) { setF((p) => ({ ...p, subjectIds: p.subjectIds.includes(id) ? p.subjectIds.filter((x) => x !== id) : (p.subjectIds.length >= 3 ? p.subjectIds : [...p.subjectIds, id]) })); }
  function addNewSubject() {
    const name = newSubj.trim();
    if (!name) return;
    const exist = subjects.find((s) => s.ten.toLowerCase() === name.toLowerCase());
    if (exist) { if (!f.subjectIds.includes(exist.id)) toggleSubj(exist.id); setNewSubj(""); return; }
    if (f.subjectIds.length >= 3) { alert("Mỗi lớp chỉ mở tối đa 3 môn."); return; }
    const id = uid("mon");
    setSubjects((p) => [...p, { id, ten: name }]);
    setF((p) => ({ ...p, subjectIds: [...p.subjectIds, id] }));
    setNewSubj("");
  }
  function validate() {
    const e = {};
    if (!f.maLop.trim()) e.maLop = "Bắt buộc";
    else if (classes.some((c) => c.maLop === f.maLop && c.id !== f.id)) e.maLop = "Mã lớp đã tồn tại";
    if (!f.tenLop.trim()) e.tenLop = "Bắt buộc";
    if (!f.mon.trim()) e.mon = "Bắt buộc nhập môn học";
    if (!f.hocPhiBuoi || f.hocPhiBuoi <= 0) e.hocPhiBuoi = "Học phí phải > 0";
    if (!f.subjectIds || f.subjectIds.length < 1 || f.subjectIds.length > 3) e.subjectIds = "Chọn 1 đến 3 môn học đang mở cho lớp";
    // schedule conflict: same phong or same gv, overlapping day+time, excluding self
    const conflict = classes.find((c) => c.id !== f.id && (c.phongId === f.phongId || c.gvId === f.gvId) && c.lich.some((s1) => f.lich.some((s2) => s1.thu === s2.thu && timeOverlap(s1, s2))));
    if (conflict) e.lich = `Trùng lịch với lớp "${conflict.tenLop}" (cùng phòng hoặc cùng giáo viên)`;
    setErrs(e);
    return Object.keys(e).length === 0;
  }
  return (
    <div>
      <div className="grid sm:grid-cols-2 gap-x-4">
        <Field label="Mã lớp" required error={errs.maLop}><TextInput value={f.maLop} onChange={(e) => set("maLop", e.target.value)} /></Field>
        <Field label="Tên lớp" required error={errs.tenLop}><TextInput value={f.tenLop} onChange={(e) => set("tenLop", e.target.value)} /></Field>
        <Field label="Khối"><Select value={f.khoiId} onChange={(e) => set("khoiId", e.target.value)} options={KHOI.map((k) => ({ value: k.id, label: k.ten }))} /></Field>
        <Field label="Môn học chính (dùng cho Điểm số/Đánh giá)" required error={errs.mon}><TextInput value={f.mon} onChange={(e) => set("mon", e.target.value)} /></Field>
        <Field label="Giáo viên"><Select value={f.gvId} onChange={(e) => set("gvId", e.target.value)} options={teachers.map((t) => ({ value: t.id, label: t.hoTen }))} /></Field>
        <Field label="Trợ giảng"><Select value={f.troId} onChange={(e) => set("troId", e.target.value)} options={assistants.map((t) => ({ value: t.id, label: t.hoTen }))} /></Field>
        <Field label="Phòng học"><Select value={f.phongId} onChange={(e) => set("phongId", e.target.value)} options={rooms.map((t) => ({ value: t.id, label: t.tenPhong }))} /></Field>
        <Field label="Học phí / buổi (đ)" required error={errs.hocPhiBuoi}><TextInput type="number" value={f.hocPhiBuoi} onChange={(e) => set("hocPhiBuoi", Number(e.target.value))} /></Field>
        <Field label="Ngày bắt đầu lớp học">
          <TextInput type="date" value={f.ngayBatDau || ""} onChange={(e) => set("ngayBatDau", e.target.value)} />
          <p className="text-xs text-slate-400 mt-1">Dùng để tính học phí — chỉ thu học phí từ tháng khai giảng trở đi</p>
        </Field>
      </div>
      <Field label="Môn học đang mở cho lớp (1–3 môn — dùng để học sinh đăng ký & tính học phí)" error={errs.subjectIds}>
        <div className="flex flex-wrap gap-2 mb-2">
          {subjects.map((s) => (
            <label key={s.id} className={cx("flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg border text-sm cursor-pointer", f.subjectIds.includes(s.id) ? "bg-teal-50 border-teal-300 text-teal-700" : "border-slate-200 text-slate-600")}>
              <input type="checkbox" className="hidden" checked={f.subjectIds.includes(s.id)} onChange={() => toggleSubj(s.id)} />{s.ten}
            </label>
          ))}
        </div>
        <div className="flex gap-2">
          <TextInput value={newSubj} onChange={(e) => setNewSubj(e.target.value)} placeholder="Thêm môn học mới..." className="flex-1" />
          <button type="button" onClick={addNewSubject} className="px-3 py-2 rounded-lg text-sm border border-slate-200 hover:bg-slate-50 shrink-0">+ Thêm môn</button>
        </div>
      </Field>
      <Field label="Lịch học trong tuần" error={errs.lich}>
        <div className="space-y-2">
          {f.lich.map((s, i) => (
            <div key={i} className="flex flex-wrap items-center gap-2">
              <Select value={s.thu} onChange={(e) => setSlot(i, "thu", Number(e.target.value))} options={[{ value: 0, label: "CN" }, { value: 2, label: "Thứ 2" }, { value: 3, label: "Thứ 3" }, { value: 4, label: "Thứ 4" }, { value: 5, label: "Thứ 5" }, { value: 6, label: "Thứ 6" }, { value: 7, label: "Thứ 7" }]} className="w-28" />
              <TextInput type="time" value={s.gioBD} onChange={(e) => setSlot(i, "gioBD", e.target.value)} className="w-28" />
              <span className="text-slate-400 text-sm">đến</span>
              <TextInput type="time" value={s.gioKT} onChange={(e) => setSlot(i, "gioKT", e.target.value)} className="w-28" />
              {f.lich.length > 1 && <button onClick={() => rmSlot(i)} className="text-rose-500 text-xs">Xoá</button>}
            </div>
          ))}
          <button onClick={addSlot} className="text-teal-700 text-xs font-medium flex items-center gap-1"><Plus size={12} /> Thêm buổi trong tuần</button>
        </div>
      </Field>
      <div className="flex justify-end gap-2 pt-2 border-t border-slate-100 mt-2">
        <button onClick={onCancel} className="px-3.5 py-2 rounded-lg text-sm border border-slate-200 hover:bg-slate-50">Huỷ</button>
        <button onClick={() => validate() && onSubmit(f)} className="px-3.5 py-2 rounded-lg text-sm bg-teal-700 text-white hover:bg-teal-800">Lưu</button>
      </div>
    </div>
  );
}
function timeOverlap(a, b) { return a.gioBD < b.gioKT && b.gioBD < a.gioKT; }

/* ============================================================================
   RESOURCES (Teachers / Assistants / Rooms) - simple tabbed generic CRUD
============================================================================ */
function ResourcesPage({ teachers, setTeachers, assistants, setAssistants, rooms, setRooms, subjects, setSubjects, classes, role }) {
  const editable = can(role, "resources", "full");
  const [tab, setTab] = useState("gv");
  return (
    <div>
      <div className="flex gap-2 mb-4 flex-wrap">
        {[["gv", "Giáo viên"], ["tro", "Trợ giảng"], ["phong", "Phòng học"], ["mon", "Môn học"]].map(([k, l]) => (
          <button key={k} onClick={() => setTab(k)} className={cx("px-3.5 py-1.5 rounded-lg text-sm font-medium", tab === k ? "bg-teal-700 text-white" : "bg-white border border-slate-200 text-slate-600")}>{l}</button>
        ))}
      </div>
      {tab === "gv" && <SimpleEntity title="giáo viên" editable={editable} items={teachers} setItems={setTeachers} usedCheck={(id) => classes.some((c) => c.gvId === id)}
        columns={[{ key: "maGV", label: "Mã GV" }, { key: "hoTen", label: "Họ tên" }, { key: "sdt", label: "SĐT" }, { key: "email", label: "Email" }, { key: "chuyenMon", label: "Chuyên môn" }]}
        fields={[{ name: "maGV", label: "Mã GV", required: true }, { name: "hoTen", label: "Họ tên", required: true }, { name: "sdt", label: "Số điện thoại" }, { name: "email", label: "Email" }, { name: "chuyenMon", label: "Chuyên môn" }]} />}
      {tab === "tro" && <SimpleEntity title="trợ giảng" editable={editable} items={assistants} setItems={setAssistants} usedCheck={(id) => classes.some((c) => c.troId === id)}
        columns={[{ key: "maTro", label: "Mã TG" }, { key: "hoTen", label: "Họ tên" }, { key: "sdt", label: "SĐT" }]}
        fields={[{ name: "maTro", label: "Mã trợ giảng", required: true }, { name: "hoTen", label: "Họ tên", required: true }, { name: "sdt", label: "Số điện thoại" }]} />}
      {tab === "phong" && <SimpleEntity title="phòng học" editable={editable} items={rooms} setItems={setRooms} usedCheck={(id) => classes.some((c) => c.phongId === id)}
        columns={[{ key: "tenPhong", label: "Tên phòng" }, { key: "succhua", label: "Sức chứa" }]}
        fields={[{ name: "tenPhong", label: "Tên phòng", required: true }, { name: "succhua", label: "Sức chứa", type: "number" }]} />}
      {tab === "mon" && <SimpleEntity title="môn học" editable={editable} items={subjects} setItems={setSubjects} usedCheck={(id) => classes.some((c) => (c.subjectIds || []).includes(id))}
        columns={[{ key: "ten", label: "Tên môn học" }]}
        fields={[{ name: "ten", label: "Tên môn học", required: true }]} />}
    </div>
  );
}

function SimpleEntity({ title, items, setItems, columns, fields, editable, usedCheck }) {
  const [modal, setModal] = useState(null);
  const [del, setDel] = useState(null);
  const cols = [...columns, ...(editable ? [{ key: "actions", label: "", render: (r) => <div className="flex gap-1"><IconBtn icon={Pencil} tone="teal" onClick={() => setModal({ mode: "edit", data: r })} /><IconBtn icon={Trash2} tone="rose" onClick={() => setDel(r)} /></div> }] : [])];
  return (
    <div>
      <SectionHeader title={`Danh sách ${title}`} actions={editable && <PrimaryButton onClick={() => setModal({ mode: "add", data: null })}>Thêm {title}</PrimaryButton>} />
      <DataTable columns={cols} rows={items} searchKeys={fields.map((f) => f.name)} exportName={title} emptyTitle={`Chưa có ${title}`} />
      <Modal open={!!modal} onClose={() => setModal(null)} title={modal?.mode === "add" ? `Thêm ${title}` : `Sửa ${title}`}>
        {modal && <GenericForm initial={modal.data} fields={fields}
          onCancel={() => setModal(null)}
          onSubmit={(data) => { if (modal.mode === "add") setItems((p) => [...p, { ...data, id: uid("r") }]); else setItems((p) => p.map((x) => x.id === data.id ? data : x)); setModal(null); }} />}
      </Modal>
      <ConfirmDialog open={!!del} onCancel={() => setDel(null)} text={`Xoá "${del?.hoTen || del?.tenPhong}"?`}
        onConfirm={() => { if (usedCheck && usedCheck(del.id)) { alert("Không thể xoá: đang được gán cho một lớp học."); setDel(null); return; } setItems((p) => p.filter((x) => x.id !== del.id)); setDel(null); }} />
    </div>
  );
}

function GenericForm({ initial, fields, onCancel, onSubmit }) {
  const init = {}; fields.forEach((f) => (init[f.name] = initial?.[f.name] ?? (f.type === "number" ? 0 : "")));
  const [f, setF] = useState(initial || init);
  const [errs, setErrs] = useState({});
  function validate() { const e = {}; fields.forEach((fl) => { if (fl.required && !String(f[fl.name] ?? "").trim()) e[fl.name] = "Bắt buộc nhập"; }); setErrs(e); return Object.keys(e).length === 0; }
  return (
    <div>
      {fields.map((fl) => (
        <Field key={fl.name} label={fl.label} required={fl.required} error={errs[fl.name]}>
          <TextInput type={fl.type || "text"} value={f[fl.name]} onChange={(e) => setF((p) => ({ ...p, [fl.name]: fl.type === "number" ? Number(e.target.value) : e.target.value }))} />
        </Field>
      ))}
      <div className="flex justify-end gap-2 pt-2 border-t border-slate-100 mt-2">
        <button onClick={onCancel} className="px-3.5 py-2 rounded-lg text-sm border border-slate-200 hover:bg-slate-50">Huỷ</button>
        <button onClick={() => validate() && onSubmit(f)} className="px-3.5 py-2 rounded-lg text-sm bg-teal-700 text-white hover:bg-teal-800">Lưu</button>
      </div>
    </div>
  );
}

/* ============================================================================
   SCHEDULE
============================================================================ */
function SchedulePage({ classes, teachers, rooms }) {
  const conflicts = useMemo(() => {
    const list = [];
    for (let i = 0; i < classes.length; i++) for (let j = i + 1; j < classes.length; j++) {
      const a = classes[i], b = classes[j];
      if (a.phongId !== b.phongId && a.gvId !== b.gvId) continue;
      a.lich.forEach((s1) => b.lich.forEach((s2) => { if (s1.thu === s2.thu && timeOverlap(s1, s2)) list.push({ a, b, s1, reason: a.phongId === b.phongId ? "Trùng phòng học" : "Trùng giáo viên" }); }));
    }
    return list;
  }, [classes]);

  const gvName = (id) => teachers.find((t) => t.id === id)?.hoTen || "—";
  const roomName = (id) => rooms.find((r) => r.id === id)?.tenPhong || "—";
  const dayList = [2, 3, 4, 5, 6, 7, 0];

  return (
    <div className="space-y-4">
      {conflicts.length > 0 && (
        <div className="bg-rose-50 border border-rose-200 text-rose-700 rounded-xl p-3 text-sm flex gap-2 items-start">
          <AlertTriangle size={16} className="mt-0.5 shrink-0" />
          <div>
            <p className="font-medium">Phát hiện {conflicts.length} lịch bị trùng:</p>
            <ul className="list-disc pl-4 mt-1 space-y-0.5">
              {conflicts.map((c, i) => <li key={i}>{c.reason}: <b>{c.a.tenLop}</b> và <b>{c.b.tenLop}</b> — {WEEKDAYS[c.s1.thu === 0 ? 0 : c.s1.thu - 1]} {c.s1.gioBD}</li>)}
            </ul>
          </div>
        </div>
      )}
      <Card className="p-4 overflow-x-auto">
        <p className="font-medium text-slate-700 mb-3 text-sm">Thời khoá biểu tuần</p>
        <div className="grid grid-cols-7 gap-2 min-w-[900px]">
          {dayList.map((d) => (
            <div key={d}>
              <p className="text-xs font-semibold text-center text-slate-500 mb-2 pb-1 border-b border-slate-200">{WEEKDAYS[d === 0 ? 0 : d - 1]}</p>
              <div className="space-y-1.5">
                {classes.filter((c) => c.lich.some((s) => s.thu === d)).flatMap((c) => c.lich.filter((s) => s.thu === d).map((s, i) => (
                  <div key={c.id + i} className="rounded-lg bg-teal-50 border border-teal-100 p-2 text-[11px]">
                    <p className="font-semibold text-teal-800">{c.tenLop}</p>
                    <p className="text-teal-600">{s.gioBD}-{s.gioKT}</p>
                    <p className="text-teal-500">{gvName(c.gvId)} · {roomName(c.phongId)}</p>
                  </div>
                )))}
                {!classes.some((c) => c.lich.some((s) => s.thu === d)) && <p className="text-[11px] text-slate-300 text-center pt-4">—</p>}
              </div>
            </div>
          ))}
        </div>
      </Card>
    </div>
  );
}

/* ============================================================================
   THU TIỀN — gộp "Học phí" + "Hoá đơn" thành 1 mục duy nhất.
   Dữ liệu: Classes → Subjects → StudentEnrollments, Payments (phiếu thu), PaymentAllocations (phân bổ theo môn/tháng)
============================================================================ */
function ThuTienPage({ students, classes, subjects, enrollments, payments, setPayments, paymentAllocations, setPaymentAllocations, tuitionConfig, setTuitionConfig, role }) {
  const editable = can(role, "thuTien", "full");
  const [tab, setTab] = useState("phaiThu");
  const [payModal, setPayModal] = useState(null); // student object
  const [detailPayment, setDetailPayment] = useState(null);
  const [cfgFee, setCfgFee] = useState(tuitionConfig.hocPhiMon);
  const thisMonth = todayISO().slice(0, 7);
  const stuName = (id) => students.find((s) => s.id === id)?.hoTen || "—";
  const lopName = (id) => classes.find((c) => c.id === id)?.tenLop || "—";

  const periodRows = useMemo(() => {
    const rows = [];
    students.filter((s) => s.trangThai === "Đang học").forEach((st) => {
      const periods = studentPeriods(st, enrollments, tuitionConfig, thisMonth, classes);
      const byMonth = {};
      periods.forEach((p) => {
        if (!byMonth[p.thang]) byMonth[p.thang] = { hocSinhId: st.id, lopId: st.lopId, thang: p.thang, soMon: 0, phaiThu: 0, daThu: 0 };
        byMonth[p.thang].soMon += 1;
        byMonth[p.thang].phaiThu += p.phaiThu;
        byMonth[p.thang].daThu += Math.min(p.phaiThu, allocatedOf(p.hocSinhId, p.lopId, p.monHocId, p.thang, paymentAllocations));
      });
      Object.values(byMonth).forEach((r) => rows.push(r));
    });
    return rows;
  }, [students, enrollments, tuitionConfig, paymentAllocations, thisMonth, classes]);

  const [monthFilter, setMonthFilter] = useState(thisMonth);
  const [statusFilter, setStatusFilter] = useState("");
  const filteredRows = periodRows
    .filter((r) => !monthFilter || r.thang === monthFilter)
    .map((r) => ({ ...r, conThieu: Math.max(0, r.phaiThu - r.daThu), trangThai: periodStatus(r.phaiThu, r.daThu) }))
    .filter((r) => !statusFilter || r.trangThai === statusFilter)
    .sort((a, b) => b.thang.localeCompare(a.thang));

  const receivableColumns = [
    { key: "hocSinhId", label: "Học sinh", render: (r) => stuName(r.hocSinhId), exportValue: (r) => stuName(r.hocSinhId) },
    { key: "lopId", label: "Lớp", render: (r) => lopName(r.lopId), exportValue: (r) => lopName(r.lopId) },
    { key: "thang", label: "Tháng" },
    { key: "soMon", label: "Số môn" },
    { key: "phaiThu", label: "Phải thu", render: (r) => vnd(r.phaiThu), exportValue: (r) => r.phaiThu },
    { key: "daThu", label: "Đã thu", render: (r) => vnd(r.daThu), exportValue: (r) => r.daThu },
    { key: "conThieu", label: "Còn thiếu", render: (r) => vnd(r.conThieu), exportValue: (r) => r.conThieu },
    { key: "trangThai", label: "Trạng thái", render: (r) => statusBadge(r.trangThai), exportValue: (r) => r.trangThai },
    ...(editable ? [{ key: "actions", label: "", render: (r) => (
      <button onClick={() => setPayModal(students.find((s) => s.id === r.hocSinhId))} className="text-xs px-2 py-1 rounded-md bg-teal-50 text-teal-700 hover:bg-teal-100 font-medium">Thu tiền</button>
    ) }] : []),
  ];

  const paymentColumns = [
    { key: "maPhieu", label: "Mã phiếu", sortable: true },
    { key: "hocSinhId", label: "Học sinh", render: (r) => stuName(r.hocSinhId), exportValue: (r) => stuName(r.hocSinhId) },
    { key: "ngayThu", label: "Ngày thu", render: (r) => fmtDate(r.ngayThu), sortable: true },
    { key: "tongThucThu", label: "Thực thu", render: (r) => vnd(r.tongThucThu), exportValue: (r) => r.tongThucThu },
    { key: "phuongThuc", label: "Phương thức" },
    { key: "ghiChu", label: "Ghi chú" },
    { key: "actions", label: "", render: (r) => (
      <div className="flex gap-1">
        <IconBtn icon={ClipboardList} tone="teal" title="Chi tiết phiếu thu" onClick={() => setDetailPayment(r)} />
        <IconBtn icon={Printer} tone="slate" title="In phiếu thu" onClick={() => printReceipt(r, paymentAllocations, students, classes, subjects)} />
      </div>
    ) },
  ];

  return (
    <div>
      <SectionHeader title="Thu tiền" desc="Khoản phải thu, tạo phiếu thu và lịch sử thu tiền học phí — tất cả trong một nơi"
        actions={editable && students.length > 0 && <PrimaryButton onClick={() => setPayModal(students.find((s) => s.trangThai === "Đang học") || students[0])}>Tạo phiếu thu</PrimaryButton>} />

      {editable && (
        <Card className="p-3 mb-4 flex flex-wrap items-center gap-2 bg-amber-50/50 border-amber-100">
          <span className="text-sm text-slate-600">Cấu hình học phí / môn / tháng:</span>
          <TextInput type="number" value={cfgFee} onChange={(e) => setCfgFee(Number(e.target.value))} className="w-36" />
          <button onClick={() => setTuitionConfig({ hocPhiMon: Number(cfgFee) || 0 })} className="px-3 py-1.5 rounded-lg text-sm bg-teal-700 text-white hover:bg-teal-800">Cập nhật</button>
          <span className="text-xs text-slate-400">Đang áp dụng: {vnd(tuitionConfig.hocPhiMon)}/môn/tháng — học phí mỗi học sinh = số môn đăng ký × mức này</span>
        </Card>
      )}

      <div className="flex gap-2 mb-4">
        <button onClick={() => setTab("phaiThu")} className={cx("px-3.5 py-1.5 rounded-lg text-sm font-medium", tab === "phaiThu" ? "bg-teal-700 text-white" : "bg-white border border-slate-200 text-slate-600")}>Khoản phải thu</button>
        <button onClick={() => setTab("lichSu")} className={cx("px-3.5 py-1.5 rounded-lg text-sm font-medium", tab === "lichSu" ? "bg-teal-700 text-white" : "bg-white border border-slate-200 text-slate-600")}>Lịch sử thu tiền</button>
      </div>

      {tab === "phaiThu" && (
        <DataTable columns={receivableColumns} rows={filteredRows} searchKeys={[]} exportName="KhoanPhaiThu" emptyTitle="Không có khoản phải thu"
          filterBar={<div className="flex gap-2">
            <TextInput type="month" value={monthFilter} onChange={(e) => setMonthFilter(e.target.value)} className="w-40" />
            <Select value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)} options={[{ value: "", label: "Tất cả trạng thái" }, { value: "Đã đóng", label: "Đã đóng" }, { value: "Đóng một phần", label: "Đóng một phần" }, { value: "Chưa đóng", label: "Chưa đóng" }]} className="w-44" />
          </div>} />
      )}

      {tab === "lichSu" && (
        <DataTable columns={paymentColumns} rows={payments} searchKeys={["maPhieu"]} exportName="LichSuThuTien" emptyTitle="Chưa có phiếu thu nào" />
      )}

      <Modal open={!!payModal} onClose={() => setPayModal(null)} title={`Tạo phiếu thu — ${payModal?.hoTen || ""}`} wide>
        {payModal && <ReceiptForm student={payModal} students={students} classes={classes} subjects={subjects} enrollments={enrollments} tuitionConfig={tuitionConfig} paymentAllocations={paymentAllocations}
          onCancel={() => setPayModal(null)}
          onSubmit={({ payment, allocations }) => { setPayments((p) => [...p, payment]); setPaymentAllocations((p) => [...p, ...allocations]); setPayModal(null); }} />}
      </Modal>

      <Modal open={!!detailPayment} onClose={() => setDetailPayment(null)} title={`Chi tiết phiếu thu — ${detailPayment?.maPhieu || ""}`}>
        {detailPayment && <ReceiptDetail payment={detailPayment} allocations={paymentAllocations.filter((a) => a.paymentId === detailPayment.id)} classes={classes} subjects={subjects} />}
      </Modal>
    </div>
  );
}

function ReceiptForm({ student, students, classes, subjects, enrollments, tuitionConfig, paymentAllocations, onCancel, onSubmit }) {
  const [lopId, setLopId] = useState(student.lopId || classes[0]?.id || "");
  const studentsInClass = useMemo(() => students.filter((s) => s.trangThai === "Đang học" && s.lopId === lopId), [students, lopId]);
  const [hocSinhId, setHocSinhId] = useState(student.id);
  const stu = students.find((s) => s.id === hocSinhId);
  const lop = classes.find((c) => c.id === lopId);
  const thisMonth = todayISO().slice(0, 7);

  function changeLop(newLopId) {
    setLopId(newLopId);
    const firstInClass = students.find((s) => s.trangThai === "Đang học" && s.lopId === newLopId);
    setHocSinhId(firstInClass?.id || "");
  }

  const periods = useMemo(() => studentPeriods(stu, enrollments, tuitionConfig, thisMonth, classes)
    .map((p) => ({ ...p, daThu: allocatedOf(p.hocSinhId, p.lopId, p.monHocId, p.thang, paymentAllocations) }))
    .map((p) => ({ ...p, conThieu: Math.max(0, p.phaiThu - p.daThu) }))
    .filter((p) => p.conThieu > 0)
    .sort((a, b) => a.thang.localeCompare(b.thang) || a.monHocId.localeCompare(b.monHocId)), [stu, enrollments, tuitionConfig, paymentAllocations, thisMonth, classes]);

  const [checked, setChecked] = useState(() => new Set(periods.map((_, i) => i)));
  useEffect(() => { setChecked(new Set(periods.map((_, i) => i))); }, [hocSinhId]); // eslint-disable-line

  const selected = periods.filter((_, i) => checked.has(i));
  const suggestTotal = selected.reduce((s, p) => s + p.conThieu, 0);
  const [ngayThu, setNgayThu] = useState(todayISO());
  const [phuongThuc, setPhuongThuc] = useState("Tiền mặt");
  const [ghiChu, setGhiChu] = useState("");
  const [thucThu, setThucThu] = useState(suggestTotal);
  useEffect(() => { setThucThu(suggestTotal); }, [suggestTotal]); // eslint-disable-line
  const [err, setErr] = useState("");

  function toggle(i) { setChecked((prev) => { const n = new Set(prev); n.has(i) ? n.delete(i) : n.add(i); return n; }); }

  function submit() {
    if (!hocSinhId) { setErr("Chọn học sinh cần thu tiền"); return; }
    const amt = Number(thucThu);
    if (!amt || amt <= 0) { setErr("Số tiền thực thu phải lớn hơn 0"); return; }
    if (selected.length === 0) { setErr("Chọn ít nhất một kỳ học phí để thu tiền"); return; }
    let remain = amt;
    const allocations = [];
    const paymentId = uid("tt");
    selected.slice().sort((a, b) => a.thang.localeCompare(b.thang)).forEach((p) => {
      if (remain <= 0) return;
      const pay = Math.min(remain, p.conThieu);
      if (pay > 0) allocations.push({ id: uid("pb"), paymentId, hocSinhId: p.hocSinhId, lopId: p.lopId, monHocId: p.monHocId, thang: p.thang, soTien: pay });
      remain -= pay;
    });
    const maPhieu = `PT${todayISO().replace(/-/g, "").slice(0, 8)}-${stu?.maHS || ""}${Math.floor(Math.random() * 90 + 10)}`;
    onSubmit({ payment: { id: paymentId, maPhieu, hocSinhId, ngayThu, phuongThuc, ghiChu, tongThucThu: amt - Math.max(0, remain) }, allocations });
  }

  return (
    <div>
      <div className="grid sm:grid-cols-2 gap-x-4">
        <Field label="Lớp"><Select value={lopId} onChange={(e) => changeLop(e.target.value)} options={classes.map((c) => ({ value: c.id, label: c.tenLop }))} /></Field>
        <Field label="Học sinh" error={!studentsInClass.length ? "Lớp này chưa có học sinh đang học" : undefined}>
          <Select value={hocSinhId} onChange={(e) => setHocSinhId(e.target.value)} options={studentsInClass.length ? studentsInClass.map((s) => ({ value: s.id, label: `${s.hoTen} (${s.maHS})` })) : [{ value: "", label: "— Không có học sinh —" }]} />
        </Field>
      </div>
      <p className="text-sm font-medium text-slate-600 mb-2">Chọn kỳ học phí còn thiếu (theo môn/tháng)</p>
      <div className="border border-slate-200 rounded-lg overflow-hidden mb-3 overflow-x-auto">
        {periods.length === 0 ? <p className="text-sm text-slate-400 p-4 text-center">Học sinh không còn khoản nợ nào.</p> : (
          <table className="w-full text-sm min-w-[480px]">
            <thead><tr className="bg-slate-50 text-slate-500 text-xs uppercase">
              <th className="px-2 py-2"></th><th className="px-2 py-2 text-left">Môn</th><th className="px-2 py-2 text-left">Tháng</th>
              <th className="px-2 py-2 text-left">Phải thu</th><th className="px-2 py-2 text-left">Đã thu</th><th className="px-2 py-2 text-left">Còn thiếu</th>
            </tr></thead>
            <tbody>
              {periods.map((p, i) => (
                <tr key={i} className="border-t border-slate-100">
                  <td className="px-2 py-2"><input type="checkbox" checked={checked.has(i)} onChange={() => toggle(i)} /></td>
                  <td className="px-2 py-2">{subjects.find((m) => m.id === p.monHocId)?.ten || "—"}</td>
                  <td className="px-2 py-2">{p.thang}</td>
                  <td className="px-2 py-2">{vnd(p.phaiThu)}</td>
                  <td className="px-2 py-2">{vnd(p.daThu)}</td>
                  <td className="px-2 py-2 text-rose-600 font-medium">{vnd(p.conThieu)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>
      <div className="grid sm:grid-cols-2 gap-x-4">
        <Field label="Số tiền thực thu" required error={err}><TextInput type="number" value={thucThu} onChange={(e) => setThucThu(e.target.value)} /></Field>
        <Field label="Ngày thu"><TextInput type="date" value={ngayThu} onChange={(e) => setNgayThu(e.target.value)} /></Field>
        <Field label="Phương thức"><Select value={phuongThuc} onChange={(e) => setPhuongThuc(e.target.value)} options={["Tiền mặt", "Chuyển khoản", "Ví điện tử"].map((v) => ({ value: v, label: v }))} /></Field>
        <Field label="Ghi chú"><TextInput value={ghiChu} onChange={(e) => setGhiChu(e.target.value)} /></Field>
      </div>
      <Card className="p-3 bg-slate-50 border-slate-200 text-sm mb-2"><p>Tổng còn thiếu (các kỳ đã chọn): <b className="text-rose-600">{vnd(suggestTotal)}</b></p><p className="text-xs text-slate-400 mt-1">Có thể thu nhiều kỳ/tháng cùng lúc; nếu số tiền thực thu ít hơn tổng còn thiếu, hệ thống sẽ tự phân bổ lần lượt theo kỳ cũ nhất trước.</p></Card>
      <div className="flex justify-end gap-2 pt-2 border-t border-slate-100 mt-2">
        <button onClick={onCancel} className="px-3.5 py-2 rounded-lg text-sm border border-slate-200 hover:bg-slate-50">Huỷ</button>
        <button onClick={submit} className="px-3.5 py-2 rounded-lg text-sm bg-teal-700 text-white hover:bg-teal-800">Xác nhận thu tiền</button>
      </div>
    </div>
  );
}

function ReceiptDetail({ payment, allocations, classes, subjects }) {
  const total = allocations.reduce((s, a) => s + a.soTien, 0);
  return (
    <div className="space-y-3 text-sm">
      <div className="grid sm:grid-cols-2 gap-2">
        <div><p className="text-slate-400 text-xs">Ngày thu</p><p className="font-medium">{fmtDate(payment.ngayThu)}</p></div>
        <div><p className="text-slate-400 text-xs">Phương thức</p><p className="font-medium">{payment.phuongThuc}</p></div>
      </div>
      {payment.ghiChu && <div><p className="text-slate-400 text-xs">Ghi chú</p><p>{payment.ghiChu}</p></div>}
      <div className="border border-slate-200 rounded-lg overflow-hidden">
        <table className="w-full text-sm">
          <thead><tr className="bg-slate-50 text-slate-500 text-xs uppercase"><th className="px-2 py-2 text-left">Lớp</th><th className="px-2 py-2 text-left">Môn</th><th className="px-2 py-2 text-left">Tháng</th><th className="px-2 py-2 text-right">Số tiền</th></tr></thead>
          <tbody>{allocations.map((a) => (<tr key={a.id} className="border-t border-slate-100"><td className="px-2 py-2">{classes.find((c) => c.id === a.lopId)?.tenLop}</td><td className="px-2 py-2">{subjects.find((m) => m.id === a.monHocId)?.ten}</td><td className="px-2 py-2">{a.thang}</td><td className="px-2 py-2 text-right">{vnd(a.soTien)}</td></tr>))}</tbody>
        </table>
      </div>
      <p className="text-right font-semibold text-teal-700">Tổng thực thu: {vnd(total)}</p>
    </div>
  );
}

function printReceipt(payment, allAllocations, students, classes, subjects) {
  const stu = students.find((s) => s.id === payment.hocSinhId);
  const rows = allAllocations.filter((a) => a.paymentId === payment.id);
  const w = window.open("", "_blank");
  w.document.write(`<html><head><title>${payment.maPhieu}</title><style>body{font-family:Arial;padding:32px;color:#1e293b} h2{margin-bottom:0} table{width:100%;border-collapse:collapse;margin-top:16px} td,th{border:1px solid #cbd5e1;padding:8px;font-size:14px;text-align:left}</style></head><body>
  <h2>PHIẾU THU HỌC PHÍ</h2><p>Mã phiếu: <b>${payment.maPhieu}</b> — Ngày thu: ${fmtDate(payment.ngayThu)}</p>
  <p>Học sinh: <b>${stu?.hoTen || ""}</b> (${stu?.maHS || ""})</p>
  <table><tr><th>Lớp</th><th>Môn</th><th>Tháng</th><th>Số tiền</th></tr>
  ${rows.map((a) => `<tr><td>${classes.find((c) => c.id === a.lopId)?.tenLop || ""}</td><td>${subjects.find((m) => m.id === a.monHocId)?.ten || ""}</td><td>${a.thang}</td><td>${vnd(a.soTien)}</td></tr>`).join("")}
  <tr><td colspan="3"><b>Tổng thực thu</b></td><td><b>${vnd(payment.tongThucThu)}</b></td></tr>
  </table><p>Phương thức: ${payment.phuongThuc}</p>${payment.ghiChu ? `<p>Ghi chú: ${payment.ghiChu}</p>` : ""}
  <script>window.print()</script></body></html>`);
  w.document.close();
}

/* ============================================================================
   TRANSACTIONS (Thu chi)
============================================================================ */
const TC_CAT = { Thu: ["Học phí khác", "Thu khác"], Chi: ["Lương giáo viên", "Lương trợ giảng", "Văn phòng phẩm", "Điện nước", "Thuê phòng", "Chi khác"] };
function TransactionsPage({ transactions, setTransactions, role }) {
  const editable = can(role, "transactions", "full");
  const [modal, setModal] = useState(null);
  const [del, setDel] = useState(null);
  const [loaiFilter, setLoaiFilter] = useState("");
  const rows = loaiFilter ? transactions.filter((t) => t.loai === loaiFilter) : transactions;
  const columns = [
    { key: "ngay", label: "Ngày", render: (r) => fmtDate(r.ngay), sortable: true },
    { key: "loai", label: "Loại", render: (r) => statusBadge(r.loai), exportValue: (r) => r.loai },
    { key: "danhMuc", label: "Danh mục" },
    { key: "soTien", label: "Số tiền", render: (r) => <span className={r.loai === "Thu" ? "text-emerald-600 font-medium" : "text-rose-600 font-medium"}>{r.loai === "Thu" ? "+" : "-"}{vnd(r.soTien)}</span>, exportValue: (r) => r.soTien },
    { key: "ghiChu", label: "Ghi chú" },
    ...(editable ? [{ key: "actions", label: "", render: (r) => <div className="flex gap-1"><IconBtn icon={Pencil} tone="teal" onClick={() => setModal({ mode: "edit", data: r })} /><IconBtn icon={Trash2} tone="rose" onClick={() => setDel(r)} /></div> }] : []),
  ];
  return (
    <div>
      <SectionHeader title="Sổ thu chi" desc={`${rows.length} giao dịch`} actions={editable && <PrimaryButton onClick={() => setModal({ mode: "add", data: null })}>Thêm giao dịch</PrimaryButton>} />
      <DataTable columns={columns} rows={rows} searchKeys={["danhMuc", "ghiChu"]} exportName="ThuChi" emptyTitle="Chưa có giao dịch nào"
        filterBar={<Select value={loaiFilter} onChange={(e) => setLoaiFilter(e.target.value)} options={[{ value: "", label: "Tất cả" }, { value: "Thu", label: "Khoản thu" }, { value: "Chi", label: "Khoản chi" }]} className="sm:w-40" />} />
      <Modal open={!!modal} onClose={() => setModal(null)} title={modal?.mode === "add" ? "Thêm giao dịch" : "Sửa giao dịch"}>
        {modal && <TransactionForm initial={modal.data} onCancel={() => setModal(null)}
          onSubmit={(data) => { if (modal.mode === "add") setTransactions((p) => [...p, { ...data, id: uid("tc") }]); else setTransactions((p) => p.map((t) => t.id === data.id ? data : t)); setModal(null); }} />}
      </Modal>
      <ConfirmDialog open={!!del} onCancel={() => setDel(null)} text="Xoá giao dịch này?" onConfirm={() => { setTransactions((p) => p.filter((t) => t.id !== del.id)); setDel(null); }} />
    </div>
  );
}
function TransactionForm({ initial, onCancel, onSubmit }) {
  const [f, setF] = useState(initial || { loai: "Chi", danhMuc: TC_CAT.Chi[0], soTien: 0, ngay: todayISO(), ghiChu: "" });
  const [err, setErr] = useState("");
  function submit() { if (!f.soTien || f.soTien <= 0) { setErr("Số tiền phải lớn hơn 0"); return; } onSubmit(f); }
  return (
    <div>
      <Field label="Loại giao dịch"><Select value={f.loai} onChange={(e) => setF((p) => ({ ...p, loai: e.target.value, danhMuc: TC_CAT[e.target.value][0] }))} options={[{ value: "Thu", label: "Khoản thu" }, { value: "Chi", label: "Khoản chi" }]} /></Field>
      <Field label="Danh mục"><Select value={f.danhMuc} onChange={(e) => setF((p) => ({ ...p, danhMuc: e.target.value }))} options={TC_CAT[f.loai].map((c) => ({ value: c, label: c }))} /></Field>
      <Field label="Số tiền (đ)" required error={err}><TextInput type="number" value={f.soTien} onChange={(e) => setF((p) => ({ ...p, soTien: Number(e.target.value) }))} /></Field>
      <Field label="Ngày"><TextInput type="date" value={f.ngay} onChange={(e) => setF((p) => ({ ...p, ngay: e.target.value }))} /></Field>
      <Field label="Ghi chú"><TextArea value={f.ghiChu} onChange={(e) => setF((p) => ({ ...p, ghiChu: e.target.value }))} /></Field>
      <div className="flex justify-end gap-2 pt-2 border-t border-slate-100 mt-2">
        <button onClick={onCancel} className="px-3.5 py-2 rounded-lg text-sm border border-slate-200 hover:bg-slate-50">Huỷ</button>
        <button onClick={submit} className="px-3.5 py-2 rounded-lg text-sm bg-teal-700 text-white hover:bg-teal-800">Lưu</button>
      </div>
    </div>
  );
}

/* ============================================================================
   REPORTS
============================================================================ */
function ReportsPage({ transactions, payments, paymentAllocations, students, classes, subjects, enrollments, tuitionConfig }) {
  const [thang, setThang] = useState(todayISO().slice(0, 7));
  const revenueHP = payments.filter((p) => p.ngayThu.startsWith(thang)).reduce((s, p) => s + p.tongThucThu, 0);
  const revenueOther = transactions.filter((t) => t.loai === "Thu" && t.ngay.startsWith(thang)).reduce((s, t) => s + t.soTien, 0);
  const chi = transactions.filter((t) => t.loai === "Chi" && t.ngay.startsWith(thang)).reduce((s, t) => s + t.soTien, 0);
  const doanhThu = revenueHP + revenueOther;
  const loiNhuan = doanhThu - chi;
  const allPeriods = buildAllPeriods(students, enrollments, tuitionConfig, todayISO().slice(0, 7), classes).map((p) => ({ ...p, daThu: allocatedOf(p.hocSinhId, p.lopId, p.monHocId, p.thang, paymentAllocations) }));
  const congNo = allPeriods.reduce((s, p) => s + Math.max(0, p.phaiThu - p.daThu), 0);

  const byClass = classes.map((c) => {
    const roster = students.filter((s) => s.lopId === c.id);
    const debt = allPeriods.filter((p) => p.lopId === c.id).reduce((s, p) => s + Math.max(0, p.phaiThu - p.daThu), 0);
    return { ten: c.tenLop, siso: roster.filter((s) => s.trangThai === "Đang học").length, debt };
  });

  function printReport() {
    const w = window.open("", "_blank");
    w.document.write(`<html><head><title>Báo cáo ${thang}</title><style>body{font-family:Arial;padding:32px;color:#1e293b}table{width:100%;border-collapse:collapse;margin-top:12px}td,th{border:1px solid #cbd5e1;padding:6px;font-size:13px}</style></head><body>
    <h2>BÁO CÁO TÀI CHÍNH THÁNG ${thang}</h2>
    <p>Doanh thu học phí: <b>${vnd(revenueHP)}</b> · Doanh thu khác: <b>${vnd(revenueOther)}</b></p>
    <p>Chi phí: <b>${vnd(chi)}</b> · Lợi nhuận: <b>${vnd(loiNhuan)}</b> · Công nợ hiện tại: <b>${vnd(congNo)}</b></p>
    <table><tr><th>Lớp</th><th>Sĩ số</th><th>Công nợ</th></tr>
    ${byClass.map((c) => `<tr><td>${c.ten}</td><td>${c.siso}</td><td>${vnd(c.debt)}</td></tr>`).join("")}
    </table><script>window.print()</script></body></html>`);
    w.document.close();
  }

  return (
    <div className="space-y-4">
      <div className="flex flex-wrap items-center gap-2">
        <TextInput type="month" value={thang} onChange={(e) => setThang(e.target.value)} className="w-44" />
        <button onClick={printReport} className="flex items-center gap-1.5 px-3 py-2 rounded-lg text-sm border border-slate-200 hover:bg-slate-50"><Printer size={15} /> In báo cáo A4</button>
        <button onClick={() => exportRows(byClass.map((c) => ({ "Lớp": c.ten, "Sĩ số": c.siso, "Công nợ": c.debt })), "BaoCao", `BaoCao_${thang}.xlsx`)} className="flex items-center gap-1.5 px-3 py-2 rounded-lg text-sm border border-slate-200 hover:bg-slate-50"><FileSpreadsheet size={15} /> Xuất Excel</button>
      </div>
      <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
        <KPICard icon={TrendingUp} tone="emerald" label="Doanh thu" value={vnd(doanhThu)} />
        <KPICard icon={TrendingDown} tone="rose" label="Chi phí" value={vnd(chi)} />
        <KPICard icon={Wallet} tone="violet" label="Lợi nhuận" value={vnd(loiNhuan)} />
        <KPICard icon={PiggyBank} tone="amber" label="Tổng công nợ" value={vnd(congNo)} />
      </div>
      <Card className="p-0 overflow-x-auto">
        <table className="w-full text-sm min-w-[560px]">
          <thead><tr className="bg-slate-50 text-slate-500 text-xs uppercase"><th className="px-3 py-2.5 text-left">Lớp</th><th className="px-3 py-2.5 text-left">Sĩ số</th><th className="px-3 py-2.5 text-left">Công nợ</th></tr></thead>
          <tbody>{byClass.map((c, i) => (<tr key={i} className="border-t border-slate-100"><td className="px-3 py-2.5">{c.ten}</td><td className="px-3 py-2.5">{c.siso}</td><td className="px-3 py-2.5">{vnd(c.debt)}</td></tr>))}</tbody>
        </table>
      </Card>
      <p className="text-xs text-slate-400">Báo cáo học phí gửi phụ huynh có thể in trực tiếp từ hồ sơ học sinh (mục Học sinh → xem hồ sơ) hoặc từ phiếu thu (mục Thu tiền → Lịch sử thu tiền).</p>
    </div>
  );
}

/* ============================================================================
   USERS
============================================================================ */
function UsersPage({ users, setUsers, teachers, assistants, role }) {
  const editable = can(role, "users", "full");
  const [modal, setModal] = useState(null);
  const [del, setDel] = useState(null);
  const columns = [
    { key: "hoTen", label: "Họ tên", sortable: true }, { key: "email", label: "Email" },
    { key: "vaiTro", label: "Vai trò", render: (r) => <Badge color="purple">{r.vaiTro}</Badge> },
    { key: "trangThai", label: "Trạng thái", render: (r) => statusBadge(r.trangThai), exportValue: (r) => r.trangThai },
    ...(editable ? [{ key: "actions", label: "", render: (r) => <div className="flex gap-1"><IconBtn icon={Pencil} tone="teal" onClick={() => setModal({ mode: "edit", data: r })} /><IconBtn icon={Trash2} tone="rose" onClick={() => setDel(r)} /></div> }] : []),
  ];
  return (
    <div>
      <SectionHeader title="Người dùng hệ thống" desc={`${users.length} tài khoản · Phân quyền theo 4 vai trò: Admin, Kế toán, Giáo viên, Trợ giảng`} actions={editable && <PrimaryButton onClick={() => setModal({ mode: "add", data: null })}>Thêm người dùng</PrimaryButton>} />
      <DataTable columns={columns} rows={users} searchKeys={["hoTen", "email"]} exportName="NguoiDung" emptyTitle="Chưa có người dùng" />
      <Modal open={!!modal} onClose={() => setModal(null)} title={modal?.mode === "add" ? "Thêm người dùng" : "Sửa người dùng"}>
        {modal && <UserForm initial={modal.data} teachers={teachers} assistants={assistants} onCancel={() => setModal(null)}
          onSubmit={(data) => { if (modal.mode === "add") setUsers((p) => [...p, { ...data, id: uid("u") }]); else setUsers((p) => p.map((u) => u.id === data.id ? data : u)); setModal(null); }} />}
      </Modal>
      <ConfirmDialog open={!!del} onCancel={() => setDel(null)} text={`Xoá người dùng "${del?.hoTen}"?`} onConfirm={() => { setUsers((p) => p.filter((u) => u.id !== del.id)); setDel(null); }} />
    </div>
  );
}
function UserForm({ initial, teachers, assistants, onCancel, onSubmit }) {
  const [f, setF] = useState(initial || { hoTen: "", email: "", vaiTro: "Giáo viên", lienKetId: "", trangThai: "Hoạt động" });
  const [err, setErr] = useState({});
  function validate() { const e = {}; if (!f.hoTen.trim()) e.hoTen = "Bắt buộc"; if (!/^\S+@\S+\.\S+$/.test(f.email)) e.email = "Email không hợp lệ"; setErr(e); return !Object.keys(e).length; }
  return (
    <div>
      <Field label="Họ tên" required error={err.hoTen}><TextInput value={f.hoTen} onChange={(e) => setF((p) => ({ ...p, hoTen: e.target.value }))} /></Field>
      <Field label="Email" required error={err.email}><TextInput value={f.email} onChange={(e) => setF((p) => ({ ...p, email: e.target.value }))} /></Field>
      <Field label="Vai trò"><Select value={f.vaiTro} onChange={(e) => setF((p) => ({ ...p, vaiTro: e.target.value }))} options={ROLES.map((r) => ({ value: r, label: r }))} /></Field>
      {f.vaiTro === "Giáo viên" && <Field label="Liên kết hồ sơ giáo viên"><Select value={f.lienKetId} onChange={(e) => setF((p) => ({ ...p, lienKetId: e.target.value }))} options={[{ value: "", label: "—" }, ...teachers.map((t) => ({ value: t.id, label: t.hoTen }))]} /></Field>}
      {f.vaiTro === "Trợ giảng" && <Field label="Liên kết hồ sơ trợ giảng"><Select value={f.lienKetId} onChange={(e) => setF((p) => ({ ...p, lienKetId: e.target.value }))} options={[{ value: "", label: "—" }, ...assistants.map((t) => ({ value: t.id, label: t.hoTen }))]} /></Field>}
      <Field label="Trạng thái"><Select value={f.trangThai} onChange={(e) => setF((p) => ({ ...p, trangThai: e.target.value }))} options={[{ value: "Hoạt động", label: "Hoạt động" }, { value: "Khóa", label: "Khoá" }]} /></Field>
      <div className="flex justify-end gap-2 pt-2 border-t border-slate-100 mt-2">
        <button onClick={onCancel} className="px-3.5 py-2 rounded-lg text-sm border border-slate-200 hover:bg-slate-50">Huỷ</button>
        <button onClick={() => validate() && onSubmit(f)} className="px-3.5 py-2 rounded-lg text-sm bg-teal-700 text-white hover:bg-teal-800">Lưu</button>
      </div>
    </div>
  );
}

/* ============================================================================
   GOOGLE SHEETS SYNC
============================================================================ */
const SHEET_NAMES = ["Users", "Students", "Classes", "Subjects", "Enrollments", "Teachers", "Assistants", "Rooms", "Schedules", "Payments", "PaymentAllocations", "Transactions", "Settings"];
function SyncPage({ sheetsCfg, setSheetsCfg, logSync, syncLog, students, classes, subjects, enrollments, teachers, assistants, rooms, payments, paymentAllocations, transactions, users, tuitionConfig, resetAppData }) {
  const [url, setUrl] = useState(sheetsCfg.url);
  const [busy, setBusy] = useState(false);

  const dataByTable = { Users: users, Students: students, Classes: classes, Subjects: subjects, Enrollments: enrollments, Teachers: teachers, Assistants: assistants, Rooms: rooms, Payments: payments, PaymentAllocations: paymentAllocations, Transactions: transactions, Settings: [{ key: "lastSync", value: sheetsCfg.lastSync || "" }, { key: "hocPhiMon", value: tuitionConfig?.hocPhiMon ?? "" }] };

  async function testConnection() {
    if (!url.trim()) { logSync("Chưa nhập URL Web App của Google Apps Script."); return; }
    setBusy(true);
    try {
      const res = await fetch(url, { method: "GET" });
      if (res.ok) { setSheetsCfg((p) => ({ ...p, url, status: "connected" })); logSync("Kết nối Google Sheets thành công."); }
      else throw new Error("HTTP " + res.status);
    } catch (e) { setSheetsCfg((p) => ({ ...p, url, status: "error" })); logSync("Kết nối thất bại: " + e.message + ". Ứng dụng tiếp tục hoạt động ở chế độ offline."); }
    setBusy(false);
  }

  async function pushSync() {
    setBusy(true);
    try {
      if (sheetsCfg.status === "connected") {
        await fetch(sheetsCfg.url, { method: "POST", body: JSON.stringify({ action: "sync", data: dataByTable }) });
        logSync("Đã đẩy toàn bộ dữ liệu lên Google Sheets.");
      } else {
        logSync("Chưa kết nối Google Sheets — đã lưu bản sao lưu (backup) offline trong trình duyệt.");
      }
      setSheetsCfg((p) => ({ ...p, lastSync: new Date().toLocaleString("vi-VN") }));
    } catch (e) { logSync("Đồng bộ lỗi: " + e.message); }
    setBusy(false);
  }

  function exportAllSheets() {
    const wb = XLSX.utils.book_new();
    SHEET_NAMES.forEach((name) => { const rows = dataByTable[name] || []; const ws = XLSX.utils.json_to_sheet(rows.length ? rows.map(flattenForSheet) : [{}]); XLSX.utils.book_append_sheet(wb, ws, name.slice(0, 31)); });
    download("GoogleSheets_Backup.xlsx", wb);
    logSync(`Đã xuất bản sao lưu đầy đủ (${SHEET_NAMES.length} sheet) ra file Excel.`);
  }

  return (
    <div className="space-y-4">
      <Card className="p-4 bg-teal-50/60 border-teal-100">
        <p className="font-medium text-teal-800 mb-1 text-sm flex items-center gap-1.5"><CheckCircle2 size={15} /> Dữ liệu được lưu tự động</p>
        <p className="text-xs text-teal-700/80">Mọi thay đổi (thêm/sửa/xoá học sinh, lớp, thu tiền...) được lưu ngay vào bộ nhớ trình duyệt (localStorage) của máy này — không mất khi tải lại trang hoặc tắt trình duyệt. Đây không phải dữ liệu demo tạm thời. Lưu ý: dữ liệu gắn với trình duyệt/máy hiện tại; dùng mục "Sao lưu toàn bộ" bên dưới hoặc kết nối Google Sheets nếu cần chia sẻ / truy cập từ máy khác.</p>
      </Card>

      <Card className="p-4">
        <p className="font-medium text-slate-700 mb-1 text-sm">Kết nối Google Sheets (qua Google Apps Script Web App)</p>
        <p className="text-xs text-slate-500 mb-3">Triển khai Apps Script làm Web App (doGet/doPost) trỏ tới Google Sheet của bạn, dán URL bên dưới. Nếu chưa kết nối, ứng dụng vẫn hoạt động đầy đủ, dữ liệu được lưu trong bộ nhớ trình duyệt.</p>
        <div className="flex flex-col sm:flex-row gap-2">
          <TextInput value={url} onChange={(e) => setUrl(e.target.value)} placeholder="https://script.google.com/macros/s/xxx/exec" className="flex-1" />
          <button disabled={busy} onClick={testConnection} className="px-3.5 py-2 rounded-lg text-sm border border-slate-200 hover:bg-slate-50 shrink-0">Kiểm tra kết nối</button>
          <PrimaryButton icon={RefreshCw} onClick={pushSync} className={busy ? "opacity-60" : ""}>{busy ? "Đang đồng bộ..." : "Đồng bộ ngay"}</PrimaryButton>
        </div>
        <div className="flex items-center gap-2 mt-3 text-xs">
          <span className={cx("w-2 h-2 rounded-full", sheetsCfg.status === "connected" ? "bg-emerald-500" : sheetsCfg.status === "error" ? "bg-rose-500" : "bg-slate-300")} />
          Trạng thái: {sheetsCfg.status === "connected" ? "Đã kết nối" : sheetsCfg.status === "error" ? "Lỗi kết nối" : "Chưa kết nối (offline)"}
          {sheetsCfg.lastSync && <span className="text-slate-400">· Lần đồng bộ gần nhất: {sheetsCfg.lastSync}</span>}
        </div>
      </Card>

      <Card className="p-4">
        <div className="flex items-center justify-between mb-3">
          <p className="font-medium text-slate-700 text-sm">{SHEET_NAMES.length} bảng dữ liệu (Sheets)</p>
          <button onClick={exportAllSheets} className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-sm border border-slate-200 hover:bg-slate-50"><Download size={14} /> Sao lưu toàn bộ (.xlsx)</button>
        </div>
        <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-5 gap-2">
          {SHEET_NAMES.map((n) => (
            <div key={n} className="border border-slate-200 rounded-lg px-3 py-2 text-sm flex items-center justify-between">
              <span>{n}</span><span className="text-xs text-slate-400">{(dataByTable[n] || []).length}</span>
            </div>
          ))}
        </div>
      </Card>

      <Card className="p-4">
        <p className="font-medium text-slate-700 mb-2 text-sm">Nhật ký đồng bộ</p>
        <div className="space-y-1.5 max-h-52 overflow-y-auto text-xs text-slate-500">
          {syncLog.map((l, i) => <div key={i} className="flex gap-2"><span className="text-slate-300 shrink-0">{l.time}</span><span>{l.msg}</span></div>)}
        </div>
      </Card>

      <Card className="p-4 border-rose-100">
        <p className="font-medium text-rose-700 mb-1 text-sm">Vùng nguy hiểm</p>
        <p className="text-xs text-slate-500 mb-3">Xoá toàn bộ dữ liệu đã lưu trên trình duyệt này (học sinh, lớp, thu tiền...) và đưa ứng dụng về trạng thái trống ban đầu. Không thể hoàn tác.</p>
        <button onClick={resetAppData} className="px-3.5 py-2 rounded-lg text-sm border border-rose-200 text-rose-600 hover:bg-rose-50">Xoá toàn bộ dữ liệu</button>
      </Card>
    </div>
  );
}
function flattenForSheet(row) { const o = {}; Object.entries(row).forEach(([k, v]) => { o[k] = typeof v === "object" ? JSON.stringify(v) : v; }); return o; }

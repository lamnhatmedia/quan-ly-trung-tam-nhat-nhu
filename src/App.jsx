import React, { useEffect, useMemo, useRef, useState } from "react";
import { createClient } from "@supabase/supabase-js";
import * as XLSX from "xlsx";
import {
  Users, School, BookOpen, Wallet, Plus, Pencil, Trash2, Upload,
  RefreshCw, LogOut, Menu, X, CheckCircle2, AlertTriangle,
  FileSpreadsheet, Cloud, CloudOff
} from "lucide-react";

const SUPABASE_URL = import.meta.env.VITE_SUPABASE_URL || "";
const SUPABASE_ANON_KEY =
  import.meta.env.VITE_SUPABASE_ANON_KEY ||
  import.meta.env.VITE_SUPABASE_PUBLISHABLE_KEY ||
  "";

const supabase =
  SUPABASE_URL && SUPABASE_ANON_KEY
    ? createClient(SUPABASE_URL, SUPABASE_ANON_KEY)
    : null;

const CLOUD_ROW_ID = "main";
const STORAGE_KEY = "nhatnhu_data_v2";
const LOGIN_USER = "admin";
const LOGIN_PASS = "123456";

const emptyData = {
  students: [],
  classes: [],
  subjects: [],
  enrollments: [],
  payments: [],
  transactions: [],
  settings: { tuitionPerSubject: 400000 },
};

const uid = (prefix = "id") =>
  `${prefix}_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 8)}`;

const today = () => new Date().toISOString().slice(0, 10);
const currentMonth = () => today().slice(0, 7);
const money = (n) => `${(Number(n) || 0).toLocaleString("vi-VN")}đ`;

const normalize = (v) =>
  String(v ?? "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");

const isChecked = (v) =>
  ["x", "✓", "yes", "true", "1", "co", "có", "dang ky", "đăng ký"].includes(
    normalize(v)
  );

const standardHeaders = new Set([
  "ma", "ma hs", "mahs", "ma so", "maso", "ten", "ho ten", "hoten",
  "sdt ph", "sdt phu huynh", "so dien thoai phu huynh",
  "phu huynh", "ten phu huynh", "lop", "ma lop", "trang thai",
]);

function sanitizeData(raw) {
  const d = raw && typeof raw === "object" ? raw : {};
  return {
    students: Array.isArray(d.students) ? d.students : [],
    classes: Array.isArray(d.classes) ? d.classes : [],
    subjects: Array.isArray(d.subjects) ? d.subjects : [],
    enrollments: Array.isArray(d.enrollments) ? d.enrollments : [],
    payments: Array.isArray(d.payments) ? d.payments : [],
    transactions: Array.isArray(d.transactions) ? d.transactions : [],
    settings: {
      tuitionPerSubject:
        Number(d.settings?.tuitionPerSubject) ||
        emptyData.settings.tuitionPerSubject,
    },
  };
}

function Badge({ children, tone = "slate" }) {
  const cls = {
    slate: "bg-slate-100 text-slate-600 border-slate-200",
    teal: "bg-teal-50 text-teal-700 border-teal-200",
    green: "bg-emerald-50 text-emerald-700 border-emerald-200",
    red: "bg-rose-50 text-rose-700 border-rose-200",
    amber: "bg-amber-50 text-amber-700 border-amber-200",
  }[tone];

  return (
    <span className={`inline-flex px-2 py-0.5 rounded-full border text-xs ${cls}`}>
      {children}
    </span>
  );
}

function Modal({ title, onClose, children }) {
  return (
    <div
      className="fixed inset-0 z-50 bg-slate-900/40 flex items-center justify-center p-4"
      onMouseDown={onClose}
    >
      <div
        className="bg-white rounded-2xl shadow-xl w-full max-w-2xl max-h-[90vh] overflow-y-auto"
        onMouseDown={(e) => e.stopPropagation()}
      >
        <div className="sticky top-0 bg-white border-b px-5 py-4 flex items-center justify-between">
          <h3 className="font-semibold text-slate-800">{title}</h3>
          <button onClick={onClose} className="p-1 rounded-lg hover:bg-slate-100">
            <X size={18} />
          </button>
        </div>
        <div className="p-5">{children}</div>
      </div>
    </div>
  );
}

function Input({ className = "", ...props }) {
  return (
    <input
      {...props}
      className={`w-full px-3 py-2 border border-slate-200 rounded-lg text-sm outline-none focus:ring-2 focus:ring-teal-500/30 ${className}`}
    />
  );
}

function Select({ className = "", children, ...props }) {
  return (
    <select
      {...props}
      className={`w-full px-3 py-2 border border-slate-200 rounded-lg text-sm bg-white outline-none focus:ring-2 focus:ring-teal-500/30 ${className}`}
    >
      {children}
    </select>
  );
}

function Button({
  children,
  icon: Icon,
  secondary = false,
  danger = false,
  ...props
}) {
  const cls = danger
    ? "bg-rose-600 hover:bg-rose-700 text-white"
    : secondary
      ? "bg-white border border-slate-200 hover:bg-slate-50 text-slate-700"
      : "bg-teal-700 hover:bg-teal-800 text-white";

  return (
    <button
      {...props}
      className={`inline-flex items-center justify-center gap-1.5 px-3.5 py-2 rounded-lg text-sm font-medium disabled:opacity-50 ${cls} ${props.className || ""}`}
    >
      {Icon && <Icon size={15} />}
      {children}
    </button>
  );
}

function Kpi({ icon: Icon, label, value }) {
  return (
    <div className="bg-white border rounded-xl p-4 flex gap-3">
      <div className="w-10 h-10 rounded-lg bg-teal-50 text-teal-700 flex items-center justify-center">
        <Icon size={19} />
      </div>
      <div>
        <p className="text-xs text-slate-400">{label}</p>
        <p className="text-lg font-semibold">{value}</p>
      </div>
    </div>
  );
}

function Row({ label, value }) {
  return (
    <div className="flex justify-between border-b last:border-0 pb-2 last:pb-0">
      <span className="text-slate-500">{label}</span>
      <b>{value}</b>
    </div>
  );
}

function Login({ onSuccess }) {
  const [user, setUser] = useState("");
  const [pass, setPass] = useState("");
  const [error, setError] = useState("");

  function login() {
    if (user === LOGIN_USER && pass === LOGIN_PASS) {
      sessionStorage.setItem("nhatnhu_login", "1");
      onSuccess();
    } else {
      setError("Tài khoản hoặc mật khẩu không đúng.");
    }
  }

  return (
    <div className="min-h-screen bg-slate-50 flex items-center justify-center p-4">
      <div className="w-full max-w-sm bg-white rounded-2xl shadow-sm border border-slate-200 p-7">
        <div className="w-12 h-12 rounded-xl bg-amber-400 text-teal-950 flex items-center justify-center font-bold text-lg mx-auto mb-4">
          NN
        </div>
        <h1 className="text-xl font-bold text-center text-slate-800">
          Trung tâm Nhật Như
        </h1>
        <p className="text-sm text-slate-400 text-center mt-1 mb-6">
          Quản lý trung tâm
        </p>

        {error && (
          <div className="mb-3 p-3 rounded-lg bg-rose-50 text-rose-700 text-sm">
            {error}
          </div>
        )}

        <div className="space-y-3">
          <Input
            value={user}
            onChange={(e) => setUser(e.target.value)}
            placeholder="Tài khoản"
            autoComplete="username"
          />
          <Input
            type="password"
            value={pass}
            onChange={(e) => setPass(e.target.value)}
            placeholder="Mật khẩu"
            autoComplete="current-password"
            onKeyDown={(e) => e.key === "Enter" && login()}
          />
          <Button className="w-full" onClick={login}>
            Đăng nhập
          </Button>
        </div>
      </div>
    </div>
  );
}

function Dashboard({ data, activeStudents, monthPaid }) {
  return (
    <div className="space-y-5">
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
        <Kpi icon={Users} label="Học sinh đang học" value={activeStudents} />
        <Kpi icon={School} label="Lớp học" value={data.classes.length} />
        <Kpi icon={BookOpen} label="Môn học" value={data.subjects.length} />
        <Kpi icon={Wallet} label="Đã thu tháng này" value={money(monthPaid)} />
      </div>

      <div className="grid lg:grid-cols-2 gap-4">
        <div className="bg-white border rounded-xl p-5">
          <h2 className="font-semibold mb-4">Tổng quan</h2>
          <div className="space-y-3 text-sm">
            <Row label="Học sinh" value={data.students.length} />
            <Row label="Lớp" value={data.classes.length} />
            <Row label="Lượt đăng ký môn" value={data.enrollments.length} />
            <Row label="Phiếu thu" value={data.payments.length} />
          </div>
        </div>

        <div className="bg-white border rounded-xl p-5">
          <h2 className="font-semibold mb-4">Đồng bộ dữ liệu</h2>
          <div className="flex items-center gap-2 text-sm text-emerald-700 bg-emerald-50 p-3 rounded-lg">
            <CheckCircle2 size={17} />
            Dữ liệu được tự động lưu lên Supabase Cloud.
          </div>
          <p className="text-xs text-slate-400 mt-3">
            Không có dữ liệu demo. Dữ liệu chỉ xuất hiện sau khi bạn nhập Excel
            hoặc thêm trực tiếp.
          </p>
        </div>
      </div>
    </div>
  );
}

function Students({ data, set }) {
  const [q, setQ] = useState("");
  const [month, setMonth] = useState(currentMonth());
  const [classId, setClassId] = useState("");
  const [modal, setModal] = useState(null);

  const rows = useMemo(
    () =>
      data.students.filter((s) => {
        const text = `${s.maHS} ${s.hoTen} ${s.sdtPH}`.toLowerCase();
        return (
          (!q || text.includes(q.toLowerCase())) &&
          (!classId || s.lopId === classId)
        );
      }),
    [data.students, q, classId]
  );

  const subjectsOf = (studentId) =>
    data.enrollments
      .filter((e) => e.hocSinhId === studentId)
      .map((e) => data.subjects.find((m) => m.id === e.monHocId)?.ten)
      .filter(Boolean);

  const paidOf = (studentId) =>
    data.payments
      .filter((p) => p.hocSinhId === studentId && p.thang === month)
      .reduce((s, p) => s + Number(p.soTien || 0), 0);

  const dueOf = (studentId) =>
    subjectsOf(studentId).length * Number(data.settings.tuitionPerSubject || 0);

  function remove(student) {
    if (!confirm(`Xóa học sinh "${student.hoTen}"?`)) return;
    set("students", (p) => p.filter((s) => s.id !== student.id));
    set("enrollments", (p) => p.filter((e) => e.hocSinhId !== student.id));
    set("payments", (p) => p.filter((p) => p.hocSinhId !== student.id));
  }

  return (
    <div className="space-y-4">
      <div className="flex flex-col lg:flex-row gap-2 lg:items-center">
        <div>
          <h2 className="text-lg font-semibold">
            Danh sách học sinh, môn đăng ký & học phí
          </h2>
          <p className="text-sm text-slate-400">{rows.length} học sinh</p>
        </div>

        <div className="lg:ml-auto flex flex-wrap gap-2">
          <Input
            value={q}
            onChange={(e) => setQ(e.target.value)}
            placeholder="Tìm mã, tên, SĐT..."
            className="w-56"
          />
          <Select
            value={classId}
            onChange={(e) => setClassId(e.target.value)}
            className="w-44"
          >
            <option value="">Tất cả lớp</option>
            {data.classes.map((c) => (
              <option key={c.id} value={c.id}>
                {c.tenLop}
              </option>
            ))}
          </Select>
          <Input
            type="month"
            value={month}
            onChange={(e) => setMonth(e.target.value)}
            className="w-40"
          />
          <Button icon={Plus} onClick={() => setModal({ mode: "add" })}>
            Thêm học sinh
          </Button>
        </div>
      </div>

      <div className="bg-white border rounded-xl overflow-x-auto">
        <table className="w-full min-w-[1000px] text-sm">
          <thead>
            <tr className="bg-slate-50 text-slate-500 text-xs">
              <th className="p-3 text-left">Mã số</th>
              <th className="p-3 text-left">Họ tên</th>
              <th className="p-3 text-left">Lớp</th>
              <th className="p-3 text-left">SĐT phụ huynh</th>
              <th className="p-3 text-left">Môn đăng ký</th>
              <th className="p-3 text-right">
                Phải thu T{Number(month.slice(5))}
              </th>
              <th className="p-3 text-right">Đã thu</th>
              <th className="p-3 text-right">Còn thiếu</th>
              <th className="p-3">Trạng thái</th>
              <th className="p-3"></th>
            </tr>
          </thead>
          <tbody>
            {rows.map((s) => {
              const subs = subjectsOf(s.id);
              const due = dueOf(s.id);
              const paid = Math.min(due, paidOf(s.id));
              const debt = Math.max(0, due - paid);

              return (
                <tr key={s.id} className="border-t hover:bg-slate-50">
                  <td className="p-3">{s.maHS}</td>
                  <td className="p-3 font-medium">{s.hoTen}</td>
                  <td className="p-3">
                    {data.classes.find((c) => c.id === s.lopId)?.tenLop || "—"}
                  </td>
                  <td className="p-3">{s.sdtPH || "—"}</td>
                  <td className="p-3">
                    <div className="flex gap-1 flex-wrap">
                      {subs.map((x) => (
                        <Badge key={x} tone="teal">
                          {x}
                        </Badge>
                      ))}
                    </div>
                  </td>
                  <td className="p-3 text-right">{money(due)}</td>
                  <td className="p-3 text-right text-emerald-700">
                    {money(paid)}
                  </td>
                  <td className="p-3 text-right text-rose-600">
                    {money(debt)}
                  </td>
                  <td className="p-3">
                    {due === 0 ? (
                      <Badge>Chưa đăng ký</Badge>
                    ) : debt === 0 ? (
                      <Badge tone="green">Đã đóng</Badge>
                    ) : paid > 0 ? (
                      <Badge tone="amber">Một phần</Badge>
                    ) : (
                      <Badge tone="red">Chưa đóng</Badge>
                    )}
                  </td>
                  <td className="p-3">
                    <div className="flex gap-1">
                      <button
                        onClick={() => setModal({ mode: "edit", student: s })}
                        className="p-1.5 text-teal-700 hover:bg-teal-50 rounded"
                      >
                        <Pencil size={15} />
                      </button>
                      <button
                        onClick={() => remove(s)}
                        className="p-1.5 text-rose-600 hover:bg-rose-50 rounded"
                      >
                        <Trash2 size={15} />
                      </button>
                    </div>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>

        {!rows.length && (
          <div className="p-10 text-center text-slate-400">
            Chưa có học sinh. Hãy nhập Excel hoặc thêm học sinh.
          </div>
        )}
      </div>

      {modal && (
        <StudentModal
          data={data}
          set={set}
          modal={modal}
          onClose={() => setModal(null)}
        />
      )}
    </div>
  );
}

function StudentModal({ data, set, modal, onClose }) {
  const initial =
    modal.student || {
      maHS: "",
      hoTen: "",
      sdtPH: "",
      tenPH: "",
      lopId: data.classes[0]?.id || "",
      trangThai: "Đang học",
    };

  const [f, setF] = useState(initial);

  const [selected, setSelected] = useState(() =>
    modal.student
      ? data.enrollments
          .filter((e) => e.hocSinhId === modal.student.id)
          .map((e) => e.monHocId)
      : []
  );

  const classObj = data.classes.find((c) => c.id === f.lopId);
  const available = data.subjects.filter(
    (s) => !classObj?.subjectIds?.length || classObj.subjectIds.includes(s.id)
  );

  function save() {
    if (!f.maHS.trim() || !f.hoTen.trim()) {
      alert("Vui lòng nhập mã số và họ tên.");
      return;
    }

    const id = f.id || uid("hs");
    const student = {
      ...f,
      id,
      maHS: f.maHS.trim(),
      hoTen: f.hoTen.trim(),
    };

    set("students", (p) =>
      f.id ? p.map((x) => (x.id === id ? student : x)) : [...p, student]
    );

    set("enrollments", (p) => [
      ...p.filter((e) => e.hocSinhId !== id),
      ...selected.map((monHocId) => ({
        id: uid("dk"),
        hocSinhId: id,
        lopId: student.lopId,
        monHocId,
      })),
    ]);

    onClose();
  }

  return (
    <Modal
      title={modal.mode === "add" ? "Thêm học sinh" : "Sửa học sinh"}
      onClose={onClose}
    >
      <div className="grid sm:grid-cols-2 gap-3">
        <label className="text-sm">
          Mã số
          <Input
            value={f.maHS}
            onChange={(e) => setF({ ...f, maHS: e.target.value })}
          />
        </label>

        <label className="text-sm">
          Họ tên
          <Input
            value={f.hoTen}
            onChange={(e) => setF({ ...f, hoTen: e.target.value })}
          />
        </label>

        <label className="text-sm">
          SĐT phụ huynh
          <Input
            value={f.sdtPH}
            onChange={(e) => setF({ ...f, sdtPH: e.target.value })}
          />
        </label>

        <label className="text-sm">
          Tên phụ huynh
          <Input
            value={f.tenPH}
            onChange={(e) => setF({ ...f, tenPH: e.target.value })}
          />
        </label>

        <label className="text-sm">
          Lớp
          <Select
            value={f.lopId}
            onChange={(e) => setF({ ...f, lopId: e.target.value })}
          >
            {data.classes.map((c) => (
              <option key={c.id} value={c.id}>
                {c.tenLop}
              </option>
            ))}
          </Select>
        </label>

        <label className="text-sm">
          Trạng thái
          <Select
            value={f.trangThai}
            onChange={(e) => setF({ ...f, trangThai: e.target.value })}
          >
            <option>Đang học</option>
            <option>Nghỉ học</option>
          </Select>
        </label>
      </div>

      <div className="mt-4">
        <p className="text-sm font-medium mb-2">Môn đăng ký</p>
        <div className="flex flex-wrap gap-2">
          {available.map((s) => (
            <label
              key={s.id}
              className={`px-3 py-2 border rounded-lg text-sm cursor-pointer ${
                selected.includes(s.id)
                  ? "bg-teal-50 border-teal-300 text-teal-700"
                  : "border-slate-200"
              }`}
            >
              <input
                type="checkbox"
                className="mr-2"
                checked={selected.includes(s.id)}
                onChange={() =>
                  setSelected((p) =>
                    p.includes(s.id)
                      ? p.filter((x) => x !== s.id)
                      : [...p, s.id]
                  )
                }
              />
              {s.ten}
            </label>
          ))}
        </div>
      </div>

      <div className="flex justify-end gap-2 mt-5">
        <Button secondary onClick={onClose}>
          Hủy
        </Button>
        <Button onClick={save}>Lưu</Button>
      </div>
    </Modal>
  );
}

function Classes({ data, set }) {
  const [modal, setModal] = useState(null);
  const [q, setQ] = useState("");

  const rows = data.classes.filter((c) =>
    `${c.maLop} ${c.tenLop}`.toLowerCase().includes(q.toLowerCase())
  );

  function remove(c) {
    if (data.students.some((s) => s.lopId === c.id)) {
      alert("Không thể xóa lớp đang có học sinh.");
      return;
    }

    if (confirm(`Xóa lớp "${c.tenLop}"?`)) {
      set("classes", (p) => p.filter((x) => x.id !== c.id));
    }
  }

  return (
    <div className="space-y-4">
      <div className="flex gap-2 items-center">
        <div>
          <h2 className="text-lg font-semibold">Lớp học</h2>
          <p className="text-sm text-slate-400">{rows.length} lớp</p>
        </div>

        <div className="ml-auto flex gap-2">
          <Input
            value={q}
            onChange={(e) => setQ(e.target.value)}
            placeholder="Tìm lớp..."
            className="w-52"
          />
          <Button icon={Plus} onClick={() => setModal({ mode: "add" })}>
            Thêm lớp
          </Button>
        </div>
      </div>

      <div className="bg-white border rounded-xl overflow-x-auto">
        <table className="w-full text-sm">
          <thead>
            <tr className="bg-slate-50 text-xs text-slate-500">
              <th className="p-3 text-left">Mã lớp</th>
              <th className="p-3 text-left">Tên lớp</th>
              <th className="p-3">Khối</th>
              <th className="p-3">Sĩ số</th>
              <th className="p-3 text-left">Môn mở</th>
              <th className="p-3"></th>
            </tr>
          </thead>

          <tbody>
            {rows.map((c) => (
              <tr key={c.id} className="border-t">
                <td className="p-3">{c.maLop}</td>
                <td className="p-3 font-medium">{c.tenLop}</td>
                <td className="p-3 text-center">{c.khoi || "—"}</td>
                <td className="p-3 text-center">
                  {
                    data.students.filter(
                      (s) => s.lopId === c.id && s.trangThai !== "Nghỉ học"
                    ).length
                  }
                </td>
                <td className="p-3">
                  <div className="flex gap-1 flex-wrap">
                    {(c.subjectIds || []).map((id) => (
                      <Badge key={id} tone="teal">
                        {data.subjects.find((s) => s.id === id)?.ten}
                      </Badge>
                    ))}
                  </div>
                </td>
                <td className="p-3">
                  <div className="flex gap-1">
                    <button
                      className="p-1.5 text-teal-700"
                      onClick={() => setModal({ mode: "edit", class: c })}
                    >
                      <Pencil size={15} />
                    </button>
                    <button
                      className="p-1.5 text-rose-600"
                      onClick={() => remove(c)}
                    >
                      <Trash2 size={15} />
                    </button>
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
        </table>

        {!rows.length && (
          <div className="p-10 text-center text-slate-400">
            Chưa có lớp.
          </div>
        )}
      </div>

      {modal && (
        <ClassModal
          data={data}
          set={set}
          modal={modal}
          onClose={() => setModal(null)}
        />
      )}
    </div>
  );
}

function ClassModal({ data, set, modal, onClose }) {
  const f0 = modal.class || {
    maLop: "",
    tenLop: "",
    khoi: "",
    subjectIds: [],
  };

  const [f, setF] = useState(f0);

  function save() {
    if (!f.maLop.trim() || !f.tenLop.trim()) {
      alert("Nhập mã lớp và tên lớp.");
      return;
    }

    const id = f.id || uid("lop");
    const item = { ...f, id };

    set("classes", (p) =>
      f.id ? p.map((x) => (x.id === id ? item : x)) : [...p, item]
    );

    onClose();
  }

  return (
    <Modal
      title={modal.mode === "add" ? "Thêm lớp" : "Sửa lớp"}
      onClose={onClose}
    >
      <div className="grid sm:grid-cols-2 gap-3">
        <label className="text-sm">
          Mã lớp
          <Input
            value={f.maLop}
            onChange={(e) => setF({ ...f, maLop: e.target.value })}
          />
        </label>

        <label className="text-sm">
          Tên lớp
          <Input
            value={f.tenLop}
            onChange={(e) => setF({ ...f, tenLop: e.target.value })}
          />
        </label>

        <label className="text-sm">
          Khối
          <Input
            value={f.khoi}
            onChange={(e) => setF({ ...f, khoi: e.target.value })}
          />
        </label>
      </div>

      <p className="text-sm font-medium mt-4 mb-2">Môn đang mở</p>

      <div className="flex flex-wrap gap-2">
        {data.subjects.map((s) => (
          <label
            key={s.id}
            className="border rounded-lg px-3 py-2 text-sm"
          >
            <input
              type="checkbox"
              className="mr-2"
              checked={(f.subjectIds || []).includes(s.id)}
              onChange={() =>
                setF({
                  ...f,
                  subjectIds: (f.subjectIds || []).includes(s.id)
                    ? f.subjectIds.filter((x) => x !== s.id)
                    : [...(f.subjectIds || []), s.id],
                })
              }
            />
            {s.ten}
          </label>
        ))}
      </div>

      <div className="flex justify-end gap-2 mt-5">
        <Button secondary onClick={onClose}>
          Hủy
        </Button>
        <Button onClick={save}>Lưu</Button>
      </div>
    </Modal>
  );
}

function Subjects({ data, set }) {
  const [name, setName] = useState("");

  function add() {
    const ten = name.trim();
    if (!ten) return;

    if (data.subjects.some((s) => normalize(s.ten) === normalize(ten))) {
      alert("Môn học đã tồn tại.");
      return;
    }

    set("subjects", (p) => [...p, { id: uid("mon"), ten }]);
    setName("");
  }

  function remove(s) {
    if (data.enrollments.some((e) => e.monHocId === s.id)) {
      alert("Không thể xóa môn đang có học sinh đăng ký.");
      return;
    }

    if (confirm(`Xóa môn "${s.ten}"?`)) {
      set("subjects", (p) => p.filter((x) => x.id !== s.id));
    }
  }

  return (
    <div className="space-y-4">
      <div>
        <h2 className="text-lg font-semibold">Môn học</h2>
        <p className="text-sm text-slate-400">
          Quản lý danh mục môn dùng khi nhập Excel và đăng ký học.
        </p>
      </div>

      <div className="flex gap-2 max-w-xl">
        <Input
          value={name}
          onChange={(e) => setName(e.target.value)}
          placeholder="Tên môn học mới"
        />
        <Button onClick={add} icon={Plus}>
          Thêm
        </Button>
      </div>

      <div className="grid sm:grid-cols-2 lg:grid-cols-4 gap-3">
        {data.subjects.map((s) => (
          <div
            key={s.id}
            className="bg-white border rounded-xl p-4 flex justify-between items-center"
          >
            <span>{s.ten}</span>
            <button onClick={() => remove(s)} className="text-rose-600">
              <Trash2 size={16} />
            </button>
          </div>
        ))}
      </div>
    </div>
  );
}

function Fees({ data, set }) {
  const [month, setMonth] = useState(currentMonth());
  const [fee, setFee] = useState(data.settings.tuitionPerSubject);

  const rows = data.students
    .filter((s) => s.trangThai !== "Nghỉ học")
    .map((s) => {
      const count = data.enrollments.filter(
        (e) => e.hocSinhId === s.id
      ).length;

      const due =
        count * Number(data.settings.tuitionPerSubject || 0);

      const paid = data.payments
        .filter(
          (p) => p.hocSinhId === s.id && p.thang === month
        )
        .reduce((a, p) => a + Number(p.soTien || 0), 0);

      return {
        ...s,
        count,
        due,
        paid: Math.min(due, paid),
        debt: Math.max(0, due - paid),
      };
    });

  function pay(s) {
    const amount = prompt(
      `Thu tiền cho ${s.hoTen}. Còn thiếu ${money(s.debt)}:`,
      String(s.debt)
    );

    if (amount === null) return;

    const n = Number(amount);
    if (!n || n <= 0) return;

    set("payments", (p) => [
      ...p,
      {
        id: uid("pt"),
        hocSinhId: s.id,
        thang: month,
        soTien: n,
        ngayThu: today(),
        phuongThuc: "Tiền mặt",
        ghiChu: "",
      },
    ]);
  }

  return (
    <div className="space-y-4">
      <div className="flex flex-wrap items-center gap-2">
        <div>
          <h2 className="text-lg font-semibold">Học phí</h2>
          <p className="text-sm text-slate-400">
            Theo dõi và thu học phí theo từng tháng.
          </p>
        </div>

        <div className="ml-auto flex gap-2">
          <Input
            type="month"
            value={month}
            onChange={(e) => setMonth(e.target.value)}
            className="w-40"
          />

          <Input
            type="number"
            value={fee}
            onChange={(e) => setFee(Number(e.target.value))}
            className="w-40"
          />

          <Button
            onClick={() =>
              set("settings", (p) => ({
                ...p,
                tuitionPerSubject: Number(fee) || 0,
              }))
            }
          >
            Lưu mức phí
          </Button>
        </div>
      </div>

      <div className="bg-white border rounded-xl overflow-x-auto">
        <table className="w-full min-w-[800px] text-sm">
          <thead>
            <tr className="bg-slate-50 text-xs text-slate-500">
              <th className="p-3 text-left">Học sinh</th>
              <th className="p-3">Môn</th>
              <th className="p-3 text-right">Phải thu</th>
              <th className="p-3 text-right">Đã thu</th>
              <th className="p-3 text-right">Còn thiếu</th>
              <th className="p-3">Trạng thái</th>
              <th className="p-3"></th>
            </tr>
          </thead>

          <tbody>
            {rows.map((s) => (
              <tr key={s.id} className="border-t">
                <td className="p-3">
                  <b>{s.hoTen}</b>
                  <div className="text-xs text-slate-400">{s.maHS}</div>
                </td>
                <td className="p-3 text-center">{s.count}</td>
                <td className="p-3 text-right">{money(s.due)}</td>
                <td className="p-3 text-right text-emerald-700">
                  {money(s.paid)}
                </td>
                <td className="p-3 text-right text-rose-600">
                  {money(s.debt)}
                </td>
                <td className="p-3 text-center">
                  {s.due === 0 ? (
                    <Badge>Chưa đăng ký</Badge>
                  ) : s.debt === 0 ? (
                    <Badge tone="green">Đã đóng</Badge>
                  ) : s.paid > 0 ? (
                    <Badge tone="amber">Một phần</Badge>
                  ) : (
                    <Badge tone="red">Chưa đóng</Badge>
                  )}
                </td>
                <td className="p-3 text-right">
                  {s.debt > 0 && (
                    <Button onClick={() => pay(s)}>Thu tiền</Button>
                  )}
                </td>
              </tr>
            ))}
          </tbody>
        </table>

        {!rows.length && (
          <div className="p-10 text-center text-slate-400">
            Chưa có học sinh.
          </div>
        )}
      </div>
    </div>
  );
}

function ImportExcel({ data, setData }) {
  const [result, setResult] = useState("");
  const [busy, setBusy] = useState(false);

  async function handleFile(file) {
    if (!file) return;

    setBusy(true);
    setResult("");

    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });

      if (!rows.length) {
        throw new Error("File Excel không có dữ liệu.");
      }

      const next = sanitizeData(data);
      const classMap = new Map(
        next.classes.map((c) => [normalize(c.maLop), c])
      );
      const subjectMap = new Map(
        next.subjects.map((s) => [normalize(s.ten), s])
      );

      let added = 0;
      let updated = 0;
      let enrollCount = 0;

      for (const row of rows) {
        const keys = Object.keys(row);

        const get = (names) => {
          const key = keys.find((k) => names.includes(normalize(k)));
          return key ? row[key] : "";
        };

        const maHS = String(
          get(["ma", "ma hs", "mahs", "ma so", "maso"])
        ).trim();

        const hoTen = String(
          get(["ten", "ho ten", "hoten"])
        ).trim();

        if (!maHS && !hoTen) continue;

        const sdtPH = String(
          get([
            "sdt ph",
            "sdt phu huynh",
            "so dien thoai phu huynh",
          ])
        ).trim();

        const tenPH = String(
          get(["phu huynh", "ten phu huynh"])
        ).trim();

        const lopText = String(
          get(["lop", "ma lop"])
        ).trim();

        let lop = classMap.get(normalize(lopText));

        if (!lop && lopText) {
          lop = {
            id: uid("lop"),
            maLop: lopText,
            tenLop: lopText,
            khoi: "",
            subjectIds: [],
          };

          next.classes.push(lop);
          classMap.set(normalize(lopText), lop);
        }

        const existing = next.students.find(
          (s) => maHS && s.maHS === maHS
        );

        const student = {
          ...(existing || {}),
          id: existing?.id || uid("hs"),
          maHS: maHS || existing?.maHS || uid("ma"),
          hoTen: hoTen || existing?.hoTen || "",
          sdtPH,
          tenPH,
          lopId: lop?.id || existing?.lopId || "",
          trangThai: existing?.trangThai || "Đang học",
        };

        if (existing) {
          next.students = next.students.map((s) =>
            s.id === existing.id ? student : s
          );
          updated++;
        } else {
          next.students.push(student);
          added++;
        }

        for (const key of keys) {
          if (standardHeaders.has(normalize(key))) continue;

          const subjectName = String(key).trim();
          if (!subjectName || !isChecked(row[key])) continue;

          let subject = subjectMap.get(normalize(subjectName));

          if (!subject) {
            subject = {
              id: uid("mon"),
              ten: subjectName,
            };

            next.subjects.push(subject);
            subjectMap.set(normalize(subjectName), subject);
          }

          if (
            lop &&
            !(lop.subjectIds || []).includes(subject.id)
          ) {
            lop.subjectIds = [
              ...(lop.subjectIds || []),
              subject.id,
            ];
          }

          if (
            !next.enrollments.some(
              (e) =>
                e.hocSinhId === student.id &&
                e.monHocId === subject.id
            )
          ) {
            next.enrollments.push({
              id: uid("dk"),
              hocSinhId: student.id,
              lopId: student.lopId,
              monHocId: subject.id,
            });

            enrollCount++;
          }
        }
      }

      setData(next);

      setResult(
        `Đã nhập ${added} học sinh mới, cập nhật ${updated}, thêm ${enrollCount} lượt đăng ký môn.`
      );
    } catch (e) {
      setResult(`Lỗi: ${e.message}`);
    } finally {
      setBusy(false);
    }
  }

  function exportTemplate() {
    const headers = [
      "Mã số",
      "Tên",
      "SĐT phụ huynh",
      "Tên phụ huynh",
      "Lớp",
      ...data.subjects.map((s) => s.ten),
    ];

    const ws = XLSX.utils.aoa_to_sheet([
      headers,
      [
        "HS001",
        "Nguyễn Văn A",
        "0900000000",
        "Nguyễn Văn B",
        "9A",
        ...data.subjects.map(() => "X"),
      ],
    ]);

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "HocSinh");
    XLSX.writeFile(wb, "Mau_Nhap_Hoc_Sinh.xlsx");
  }

  return (
    <div className="space-y-5">
      <div>
        <h2 className="text-lg font-semibold">
          Nhập danh sách lớp từ Excel
        </h2>
        <p className="text-sm text-slate-400">
          Hỗ trợ mã số, tên, SĐT phụ huynh, lớp và các cột môn học
          đánh dấu X/✓/1.
        </p>
      </div>

      <div className="grid lg:grid-cols-2 gap-4">
        <div className="bg-white border rounded-xl p-5">
          <div className="flex items-center gap-3 mb-4">
            <div className="w-10 h-10 rounded-lg bg-emerald-50 text-emerald-700 flex items-center justify-center">
              <Upload size={18} />
            </div>
            <div>
              <b>Chọn file Excel</b>
              <p className="text-xs text-slate-400">
                .xlsx hoặc .xls
              </p>
            </div>
          </div>

          <input
            type="file"
            accept=".xlsx,.xls"
            disabled={busy}
            onChange={(e) => handleFile(e.target.files?.[0])}
          />

          {busy && (
            <p className="text-sm text-teal-700 mt-3">
              Đang đọc dữ liệu...
            </p>
          )}

          {result && (
            <div className="mt-4 p-3 rounded-lg bg-slate-50 text-sm">
              {result}
            </div>
          )}
        </div>

        <div className="bg-white border rounded-xl p-5">
          <b>Cấu trúc đề nghị</b>

          <div className="mt-3 overflow-x-auto">
            <table className="text-sm">
              <thead>
                <tr>
                  <th className="text-left p-2">Mã số</th>
                  <th className="text-left p-2">Tên</th>
                  <th className="text-left p-2">SĐT PH</th>
                  <th className="text-left p-2">Lớp</th>
                  {data.subjects.slice(0, 3).map((s) => (
                    <th key={s.id} className="p-2">
                      {s.ten}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                <tr>
                  <td className="p-2">HS001</td>
                  <td className="p-2">Nguyễn Văn A</td>
                  <td className="p-2">090...</td>
                  <td className="p-2">9A</td>
                  {data.subjects.slice(0, 3).map((s) => (
                    <td key={s.id} className="p-2">
                      X
                    </td>
                  ))}
                </tr>
              </tbody>
            </table>
          </div>

          <Button
            secondary
            icon={FileSpreadsheet}
            className="mt-4"
            onClick={exportTemplate}
          >
            Tải file mẫu
          </Button>
        </div>
      </div>
    </div>
  );
}

function Manager({ onLogout }) {
  const [data, setData] = useState(() => {
    try {
      return sanitizeData(
        JSON.parse(localStorage.getItem(STORAGE_KEY) || "null") ||
          emptyData
      );
    } catch {
      return emptyData;
    }
  });

  const [hydrated, setHydrated] = useState(false);
  const [cloudStatus, setCloudStatus] = useState("loading");
  const [cloudError, setCloudError] = useState("");
  const [view, setView] = useState("dashboard");
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [syncing, setSyncing] = useState(false);
  const saveTimer = useRef(null);

  useEffect(() => {
    let cancelled = false;

    async function loadCloud() {
      if (!supabase) {
        setCloudStatus("error");
        setCloudError(
          "Thiếu VITE_SUPABASE_URL hoặc VITE_SUPABASE_ANON_KEY."
        );
        setHydrated(true);
        return;
      }

      try {
        const { data: row, error } = await supabase
          .from("app_data")
          .select("id,data,updated_at")
          .eq("id", CLOUD_ROW_ID)
          .maybeSingle();

        if (error) throw error;

        if (!cancelled && row?.data) {
          setData(sanitizeData(row.data));
          localStorage.setItem(
            STORAGE_KEY,
            JSON.stringify(row.data)
          );
        }

        if (!cancelled) setCloudStatus("connected");
      } catch (e) {
        if (!cancelled) {
          setCloudStatus("error");
          setCloudError(
            e?.message || "Không thể đọc dữ liệu Cloud."
          );
          console.error("Supabase load error:", e);
        }
      } finally {
        if (!cancelled) setHydrated(true);
      }
    }

    loadCloud();

    return () => {
      cancelled = true;
    };
  }, []);

  useEffect(() => {
    if (!hydrated) return;

    localStorage.setItem(STORAGE_KEY, JSON.stringify(data));

    clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(() => saveCloud(data), 700);

    return () => clearTimeout(saveTimer.current);
  }, [data, hydrated]);

  async function saveCloud(nextData) {
    if (!supabase || !hydrated) return;

    setSyncing(true);

    try {
      const { error } = await supabase
        .from("app_data")
        .upsert(
          {
            id: CLOUD_ROW_ID,
            data: nextData,
            updated_at: new Date().toISOString(),
          },
          { onConflict: "id" }
        );

      if (error) throw error;

      setCloudStatus("connected");
      setCloudError("");
    } catch (e) {
      setCloudStatus("error");
      setCloudError(
        e?.message || "Không thể lưu dữ liệu Cloud."
      );
      console.error("Supabase sync error:", e);
    } finally {
      setSyncing(false);
    }
  }

  async function manualSync() {
    await saveCloud(data);
  }

  function set(key, value) {
    setData((d) => ({
      ...d,
      [key]:
        typeof value === "function"
          ? value(d[key])
          : value,
    }));
  }

  const activeStudents = data.students.filter(
    (s) => s.trangThai !== "Nghỉ học"
  ).length;

  const month = currentMonth();

  const monthPaid = data.payments
    .filter((p) => p.thang === month)
    .reduce((s, p) => s + Number(p.soTien || 0), 0);

  const nav = [
    ["dashboard", "Dashboard", LayoutDashboard],
    ["students", "Học sinh", Users],
    ["classes", "Lớp học", School],
    ["subjects", "Môn học", BookOpen],
    ["fees", "Học phí", Wallet],
    ["import", "Nhập Excel", FileSpreadsheet],
  ];

  return (
    <div className="min-h-screen bg-slate-50 text-slate-800 flex">
      <aside
        className={`fixed lg:static z-40 inset-y-0 left-0 w-64 bg-teal-950 text-teal-50 flex flex-col transition-transform ${
          sidebarOpen
            ? "translate-x-0"
            : "-translate-x-full lg:translate-x-0"
        }`}
      >
        <div className="h-16 px-5 border-b border-teal-900 flex items-center gap-2">
          <div className="w-9 h-9 rounded-lg bg-amber-400 text-teal-950 flex items-center justify-center font-bold">
            NN
          </div>
          <div>
            <p className="font-semibold text-sm">
              Trung tâm Nhật Như
            </p>
            <p className="text-[11px] text-teal-300">
              Quản lý trung tâm
            </p>
          </div>
        </div>

        <nav className="p-2 space-y-1 flex-1">
          {nav.map(([key, label, Icon]) => (
            <button
              key={key}
              onClick={() => {
                setView(key);
                setSidebarOpen(false);
              }}
              className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm ${
                view === key
                  ? "bg-teal-800 text-white"
                  : "text-teal-200 hover:bg-teal-900"
              }`}
            >
              <Icon size={17} />
              {label}
            </button>
          ))}
        </nav>

        <div className="p-3 border-t border-teal-900">
          <div
            className={`text-xs flex items-center gap-1.5 ${
              cloudStatus === "connected"
                ? "text-emerald-300"
                : cloudStatus === "error"
                  ? "text-rose-300"
                  : "text-amber-300"
            }`}
          >
            {cloudStatus === "connected" ? (
              <Cloud size={14} />
            ) : (
              <CloudOff size={14} />
            )}
            {cloudStatus === "connected"
              ? "Đã đồng bộ Cloud"
              : cloudStatus === "error"
                ? "Lỗi đồng bộ"
                : "Đang kết nối..."}
          </div>
        </div>
      </aside>

      {sidebarOpen && (
        <div
          className="fixed inset-0 bg-slate-900/40 z-30 lg:hidden"
          onClick={() => setSidebarOpen(false)}
        />
      )}

      <div className="flex-1 min-w-0">
        <header className="h-16 bg-white border-b border-slate-200 flex items-center px-4 lg:px-6 gap-3 sticky top-0 z-20">
          <button
            className="lg:hidden"
            onClick={() => setSidebarOpen(true)}
          >
            <Menu size={20} />
          </button>

          <h1 className="font-semibold">
            {nav.find((n) => n[0] === view)?.[1]}
          </h1>

          <div className="ml-auto flex items-center gap-2">
            <div
              className={`hidden sm:flex items-center gap-1.5 text-xs px-2.5 py-1.5 rounded-full ${
                cloudStatus === "connected"
                  ? "bg-emerald-50 text-emerald-700"
                  : "bg-rose-50 text-rose-700"
              }`}
            >
              {cloudStatus === "connected" ? (
                <CheckCircle2 size={13} />
              ) : (
                <AlertTriangle size={13} />
              )}
              {syncing
                ? "Đang lưu..."
                : cloudStatus === "connected"
                  ? "Đã đồng bộ"
                  : "Lỗi đồng bộ"}
            </div>

            <Button
              secondary
              icon={RefreshCw}
              onClick={manualSync}
              disabled={syncing}
            >
              Đồng bộ
            </Button>

            <Button
              secondary
              icon={LogOut}
              onClick={onLogout}
            >
              Đăng xuất
            </Button>
          </div>
        </header>

        {cloudError && (
          <div className="mx-4 lg:mx-6 mt-4 p-3 rounded-xl border border-rose-200 bg-rose-50 text-rose-700 text-sm">
            <b>Lỗi Supabase:</b> {cloudError}
          </div>
        )}

        <main className="p-4 lg:p-6 max-w-[1450px] mx-auto">
          {view === "dashboard" && (
            <Dashboard
              data={data}
              activeStudents={activeStudents}
              monthPaid={monthPaid}
            />
          )}

          {view === "students" && (
            <Students data={data} set={set} />
          )}

          {view === "classes" && (
            <Classes data={data} set={set} />
          )}

          {view === "subjects" && (
            <Subjects data={data} set={set} />
          )}

          {view === "fees" && (
            <Fees data={data} set={set} />
          )}

          {view === "import" && (
            <ImportExcel data={data} setData={setData} />
          )}
        </main>
      </div>
    </div>
  );
}

function LayoutDashboard(props) {
  return (
    <svg
      {...props}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="2"
    >
      <rect x="3" y="3" width="7" height="7" rx="1" />
      <rect x="14" y="3" width="7" height="7" rx="1" />
      <rect x="3" y="14" width="7" height="7" rx="1" />
      <rect x="14" y="14" width="7" height="7" rx="1" />
    </svg>
  );
}

export default function App() {
  const [loggedIn, setLoggedIn] = useState(
    () => sessionStorage.getItem("nhatnhu_login") === "1"
  );

  if (!loggedIn) {
    return <Login onSuccess={() => setLoggedIn(true)} />;
  }

  return (
    <Manager
      onLogout={() => {
        sessionStorage.removeItem("nhatnhu_login");
        setLoggedIn(false);
      }}
    />
  );
}

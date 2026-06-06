import { useState, useRef, useEffect } from "react";
import * as XLSX from "xlsx";

const BLUE = "#1a56db"; const BLUE_LIGHT = "#eff4ff"; const BLUE_DARK = "#1e429f"; const RED = "#b91c1c";
const REMARKS = [
  "1. Lead time: 30 days after PO confirmation (to be confirmed)",
  "2. Validity of quotation: within 15 days",
  "3. Prices subject to change without notice",
  "4. Price included Packaging, Freight",
  "5. Price not included: China Mainland Import Duty + VAT"
];

function parseBOM(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
        let productName = "", partNumber = "", bomCost = 0;
        const row1 = data[0] || [];
        for (let i = 0; i < row1.length; i++) {
          const c = String(row1[i]);
          if (c.includes("产品名称")) { const a = c.replace(/.*产品名称[：:]\s*/, "").trim(); if (a) { productName = a; } else { for (let j = i+1; j < row1.length; j++) { if (row1[j]) { productName = String(row1[j]).trim(); break; } } } break; }
        }
        for (let i = 0; i < row1.length; i++) {
          const c = String(row1[i]);
          if (c.includes("产品型号")) { const a = c.replace(/.*产品型号[：:]\s*/, "").trim(); if (a) { partNumber = a; } else { for (let j = i+1; j < row1.length; j++) { if (row1[j]) { partNumber = String(row1[j]).trim(); break; } } } break; }
        }
        for (let r = 0; r < data.length; r++) {
          const row = data[r];
          if (row.some(c => String(c).trim() === "合计")) { const v = parseFloat(row[32]); if (!isNaN(v) && v > 0) { bomCost = v; } break; }
        }
        resolve({ productName, partNumber, bomCost });
      } catch(err) { reject(err); }
    };
    reader.readAsArrayBuffer(file);
  });
}

const fmtVND = n => Number(n).toLocaleString("vi-VN") + " ₫";
const fmtUSD = n => "$" + Number(n).toFixed(4);
const sellVND = (cost, margin) => Math.round(parseFloat(cost||0) * parseFloat(margin||1));
const sellUSD = (cost, margin, rate) => rate > 0 ? sellVND(cost, margin) / parseFloat(rate||1) : 0;

async function apiLogin(username, password) {
  try {
    const res = await fetch("/api/login", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ username, password }) });
    if (!res.ok && res.headers.get("content-type")?.includes("text/html")) {
      return { error: "Lỗi server — API không hoạt động (HTTP " + res.status + ")" };
    }
    return res.json();
  } catch {
    return { error: "Không kết nối được server. Kiểm tra lại mạng." };
  }
}
async function apiGetRecords(userId) {
  const res = await fetch(`/api/records?user_id=${userId}`);
  return res.json();
}
async function apiSaveRecord(data) {
  const res = await fetch("/api/records", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(data) });
  return res.json();
}
async function apiUpdateRecord(data) {
  const res = await fetch("/api/records", { method: "PUT", headers: { "Content-Type": "application/json" }, body: JSON.stringify(data) });
  return res.json();
}
async function apiDeleteRecord(id) {
  const res = await fetch("/api/records", { method: "DELETE", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ id }) });
  return res.json();
}

function genPDF(rec) {
  const usd = sellUSD(rec.bom_cost||rec.bomCost, rec.margin, rec.exchange_rate||rec.rate).toFixed(4);
  const dateStr = rec.date ? rec.date.split("-").reverse().join("/") : "";
  const html = `<!DOCTYPE html><html><head><meta charset="utf-8"/><style>
    *{margin:0;padding:0;box-sizing:border-box;}
    body{font-family:Arial,sans-serif;font-size:12px;color:#111;padding:32px;}
    .header{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:8px;}
    .logo{font-size:28px;font-weight:900;color:${RED};letter-spacing:1px;}
    .co-name{font-size:20px;font-weight:700;text-align:center;margin-bottom:2px;}
    .co-sub{font-size:13px;text-align:center;color:#333;}
    .addr{font-size:10px;color:#444;text-align:right;line-height:1.6;}
    .divider{border:none;border-top:2px solid ${RED};margin:8px 0 16px;}
    .title{text-align:center;font-size:16px;font-weight:700;letter-spacing:2px;margin-bottom:16px;}
    .meta{display:grid;grid-template-columns:1fr 1fr;gap:4px 24px;margin-bottom:16px;font-size:11px;}
    .meta-row{display:flex;gap:8px;} .meta-label{font-weight:700;white-space:nowrap;}
    table{width:100%;border-collapse:collapse;margin-bottom:16px;}
    th{background:#222;color:#fff;padding:7px 10px;text-align:center;font-size:11px;}
    td{padding:10px;border:1px solid #ccc;vertical-align:middle;font-size:11px;}
    td:first-child,th:first-child{text-align:center;width:40px;}
    td:nth-child(2),th:nth-child(2){width:110px;text-align:center;}
    td:nth-child(4),th:nth-child(4){text-align:right;width:90px;}
    td:nth-child(5),th:nth-child(5){text-align:center;width:70px;}
    .remarks{font-size:10px;line-height:1.8;} .remarks b{font-size:11px;}
  </style></head><body>
  <div class="header">
    <div><div class="logo">COMART</div></div>
    <div><div class="co-name">越 南 有 限 公 司</div><div class="co-sub">COMART Việt Nam</div></div>
    <div class="addr">越南北寧省 Nénh坊雲中工業園區 CN-06-06 號工廠<br/>Factory CN-06-06, Van Trung Industrial Park,<br/>Nenh Ward, Nénh, Bắc Ninh<br/>TEL: +886-2-89111133 FAX: +886-2-86653056 http://www.comart.com.tw</div>
  </div>
  <hr class="divider"/>
  <div class="title">QUOTATION</div>
  <div class="meta">
    <div class="meta-row"><span class="meta-label">To</span><span>${rec.to_customer||rec.to||""}</span></div>
    <div class="meta-row"><span class="meta-label">TERM:</span><span>${rec.term||""}</span></div>
    <div class="meta-row"><span class="meta-label">ATTENTION:</span><span>${rec.attention||""}</span></div>
    <div class="meta-row"><span class="meta-label">DATE:</span><span>${dateStr}</span></div>
    <div class="meta-row"><span class="meta-label">Tel</span><span>${rec.tel||""}</span></div>
    <div class="meta-row"><span class="meta-label">CURRENCY:</span><span>USD</span></div>
  </div>
  <table>
    <thead><tr><th>No</th><th>Item</th><th>Description</th><th>Price/Pcs</th><th>QTY</th><th>Remark</th></tr></thead>
    <tbody><tr>
      <td>1</td><td>${rec.part_number||rec.partNumber||""}</td><td>${rec.product_name||rec.productName||""}</td>
      <td>$${parseFloat(usd).toFixed(4)}</td><td>${rec.qty||""}</td><td>${rec.remark||""}</td>
    </tr></tbody>
  </table>
  <div class="remarks"><b>Remarks</b><br/>${REMARKS.join("<br/>")}</div>
  </body></html>`;
  const w = window.open("", "_blank", "width=900,height=700");
  w.document.write(html); w.document.close();
  setTimeout(() => w.print(), 600);
}

function Btn({ children, onClick, style, disabled, title }) {
  return <button onClick={onClick} disabled={disabled} title={title} style={style}>{children}</button>;
}

function LoginPage({ onLogin }) {
  const [u, setU] = useState(""); const [p, setP] = useState(""); const [err, setErr] = useState(""); const [loading, setLoading] = useState(false);
  const submit = async () => {
    setLoading(true); setErr("");
    const res = await apiLogin(u, p);
    if (res.user) onLogin(res.user);
    else setErr(res.error || "Lỗi đăng nhập");
    setLoading(false);
  };
  return (
    <div style={{ minHeight: 500, display: "flex", alignItems: "center", justifyContent: "center", background: "var(--color-background-tertiary)" }}>
      <div style={{ width: 380, background: "var(--color-background-primary)", border: "0.5px solid var(--color-border-tertiary)", borderRadius: 16, overflow: "hidden" }}>
        <div style={{ background: BLUE, padding: "2rem", textAlign: "center" }}>
          <div style={{ width: 52, height: 52, borderRadius: 14, background: "rgba(255,255,255,0.15)", display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 12px" }}>
            <i className="ti ti-table-import" style={{ fontSize: 26, color: "#fff" }} aria-hidden="true" />
          </div>
          <h2 style={{ margin: "0 0 4px", fontSize: 20, fontWeight: 500, color: "#fff" }}>BOM Price Manager</h2>
          <p style={{ margin: 0, fontSize: 13, color: "rgba(255,255,255,0.6)" }}>COMART Việt Nam</p>
        </div>
        <div style={{ padding: "1.5rem" }}>
          {[{ l: "Tài khoản", v: u, s: setU, t: "text", p: "username" }, { l: "Mật khẩu", v: p, s: setP, t: "password", p: "••••••••" }].map(f => (
            <div key={f.l} style={{ marginBottom: 12 }}>
              <label style={{ fontSize: 12, color: "var(--color-text-secondary)", display: "block", marginBottom: 4 }}>{f.l}</label>
              <input type={f.t} value={f.v} onChange={e => f.s(e.target.value)} placeholder={f.p} style={{ width: "100%", boxSizing: "border-box" }} onKeyDown={e => e.key === "Enter" && submit()} />
            </div>
          ))}
          {err && <p style={{ fontSize: 13, color: "var(--color-text-danger)", marginBottom: 12 }}>{err}</p>}
          <button onClick={submit} disabled={loading} style={{ width: "100%", background: BLUE, color: "#fff", border: "none", padding: "10px", borderRadius: 8, fontSize: 14, cursor: "pointer", fontWeight: 500 }}>
            {loading ? "Đang đăng nhập..." : "Đăng nhập"}
          </button>
        </div>
      </div>
    </div>
  );
}

function UploadModal({ onClose, onSave, initial, userId }) {
  const empty = { productName: "", partNumber: "", bomCost: "", margin: "0.9", rate: "26000", date: new Date().toISOString().slice(0, 10), notes: "", fileName: "", to: "", attention: "", term: "DAP — DONGGUAN, CHINA", tel: "", qty: "1000", remark: "Price included: Package + Freight" };
  const [form, setForm] = useState(initial ? { ...empty, productName: initial.product_name, partNumber: initial.part_number, bomCost: String(initial.bom_cost), margin: String(initial.margin), rate: String(initial.exchange_rate), date: initial.date, notes: initial.notes, to: initial.to_customer, attention: initial.attention, term: initial.term, tel: initial.tel, qty: initial.qty, remark: initial.remark, fileName: initial.file_name } : empty);
  const [uploading, setUploading] = useState(false); const [uploadErr, setUploadErr] = useState(""); const [step, setStep] = useState(initial ? 2 : 1); const [saving, setSaving] = useState(false);
  const fileRef = useRef();
  const setF = (k, v) => setForm(f => ({ ...f, [k]: v }));

  const handleFile = async e => {
    const file = e.target.files[0]; if (!file) return;
    setUploading(true); setUploadErr("");
    try {
      const res = await parseBOM(file);
      setForm(f => ({ ...f, productName: res.productName || f.productName, partNumber: res.partNumber || f.partNumber, bomCost: res.bomCost ? String(res.bomCost) : f.bomCost, fileName: file.name }));
      setStep(2);
    } catch { setUploadErr("Không đọc được file. Kiểm tra lại định dạng."); }
    setUploading(false); e.target.value = "";
  };

  const valid = form.productName && form.partNumber && form.bomCost;
  const sv = sellVND(form.bomCost, form.margin);
  const su = sellUSD(form.bomCost, form.margin, form.rate);

  const handleSave = async () => {
    setSaving(true);
    const payload = { user_id: userId, part_number: form.partNumber, product_name: form.productName, bom_cost: parseFloat(form.bomCost), margin: parseFloat(form.margin), exchange_rate: parseFloat(form.rate), qty: form.qty, date: form.date, to_customer: form.to, attention: form.attention, term: form.term, tel: form.tel, remark: form.remark, notes: form.notes, file_name: form.fileName };
    if (initial) { await apiUpdateRecord({ id: initial.id, ...payload }); onSave({ id: initial.id, ...payload }); }
    else { const res = await apiSaveRecord(payload); onSave({ id: res.id, ...payload }); }
    setSaving(false);
  };

  const Inp = ({ label, fkey, type, placeholder, hint, span }) => (
    <div style={{ gridColumn: `span ${span || 1}` }}>
      <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 4 }}>
        <label style={{ fontSize: 12, color: "var(--color-text-secondary)" }}>{label}</label>
        {hint && <span style={{ fontSize: 11, color: BLUE }}>{hint}</span>}
      </div>
      <input type={type || "text"} value={form[fkey]} onChange={e => setF(fkey, e.target.value)} placeholder={placeholder} style={{ width: "100%", boxSizing: "border-box", borderColor: hint ? BLUE : undefined }} />
    </div>
  );

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.55)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 100, padding: 16 }}>
      <div style={{ width: "100%", maxWidth: 580, background: "var(--color-background-primary)", borderRadius: 16, border: "0.5px solid var(--color-border-tertiary)", overflow: "hidden", maxHeight: "92vh", display: "flex", flexDirection: "column" }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "1rem 1.25rem", borderBottom: "0.5px solid var(--color-border-tertiary)", flexShrink: 0 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <i className="ti ti-file-upload" style={{ fontSize: 18, color: BLUE }} aria-hidden="true" />
            <span style={{ fontWeight: 500 }}>{initial ? "Chỉnh sửa báo giá" : "Thêm báo giá mới"}</span>
          </div>
          <button onClick={onClose}><i className="ti ti-x" aria-hidden="true" /></button>
        </div>
        <div style={{ overflowY: "auto", padding: "1.25rem", flex: 1 }}>
          <div style={{ display: "flex", gap: 6, marginBottom: "1.25rem" }}>
            {[{ n: 1, l: "Upload BOM" }, { n: 2, l: "Sản phẩm" }, { n: 3, l: "Báo giá" }, { n: 4, l: "Xác nhận" }].map(s => (
              <div key={s.n} style={{ flex: 1 }}>
                <div style={{ height: 3, borderRadius: 4, background: step >= s.n ? BLUE : "var(--color-border-tertiary)", marginBottom: 4 }} />
                <span style={{ fontSize: 10, color: step >= s.n ? BLUE : "var(--color-text-tertiary)", fontWeight: step === s.n ? 500 : 400 }}>{s.l}</span>
              </div>
            ))}
          </div>

          {step === 1 && (
            <div>
              <div onClick={() => fileRef.current.click()} style={{ border: `1.5px dashed ${BLUE}`, borderRadius: 12, padding: "2.5rem 1rem", textAlign: "center", cursor: "pointer", background: BLUE_LIGHT }}>
                <div style={{ width: 56, height: 56, borderRadius: 12, background: "rgba(26,86,219,0.12)", display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 12px" }}>
                  <i className="ti ti-file-spreadsheet" style={{ fontSize: 28, color: BLUE }} aria-hidden="true" />
                </div>
                <p style={{ margin: "0 0 4px", fontWeight: 500, color: "var(--color-text-primary)" }}>{uploading ? "Đang đọc file..." : "Click để chọn file Excel BOM"}</p>
                <p style={{ margin: 0, fontSize: 13, color: "var(--color-text-secondary)" }}>Tự đọc Tên SP, P/N và BOM Cost</p>
                <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={handleFile} />
              </div>
              {uploadErr && <p style={{ color: "var(--color-text-danger)", fontSize: 13, marginTop: 8 }}>{uploadErr}</p>}
              <div style={{ display: "flex", alignItems: "center", gap: 8, margin: "1rem 0" }}>
                <div style={{ flex: 1, height: "0.5px", background: "var(--color-border-tertiary)" }} />
                <span style={{ fontSize: 12, color: "var(--color-text-tertiary)" }}>hoặc nhập tay</span>
                <div style={{ flex: 1, height: "0.5px", background: "var(--color-border-tertiary)" }} />
              </div>
              <button onClick={() => setStep(2)} style={{ width: "100%" }}>Nhập thủ công</button>
            </div>
          )}

          {step === 2 && (
            <div>
              {form.fileName && (
                <div style={{ display: "flex", alignItems: "center", gap: 8, padding: "8px 12px", background: BLUE_LIGHT, borderRadius: 8, marginBottom: "1rem", border: `0.5px solid ${BLUE}` }}>
                  <i className="ti ti-file-check" style={{ color: BLUE, fontSize: 16 }} aria-hidden="true" />
                  <span style={{ fontSize: 13, color: BLUE, flex: 1, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{form.fileName}</span>
                  <button onClick={() => fileRef.current.click()} style={{ fontSize: 12, padding: "2px 8px" }}>Đổi file</button>
                  <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={handleFile} />
                </div>
              )}
              {form.fileName && (
                <div style={{ display: "flex", alignItems: "center", gap: 6, padding: "6px 10px", background: "#f0fdf4", borderRadius: 6, marginBottom: "1rem", border: "0.5px solid #86efac" }}>
                  <i className="ti ti-sparkles" style={{ fontSize: 13, color: "#16a34a" }} aria-hidden="true" />
                  <span style={{ fontSize: 12, color: "#16a34a" }}>Tự động điền từ file BOM — có thể chỉnh sửa</span>
                </div>
              )}
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                <Inp label="Mã P/N" fkey="partNumber" placeholder="8SHB010068" hint={form.fileName ? "Tự điền ✓" : ""} />
                <Inp label="Ngày báo giá" fkey="date" type="date" />
                <Inp label="Tên sản phẩm" fkey="productName" placeholder="Tên sản phẩm" hint={form.fileName ? "Tự điền ✓" : ""} span={2} />
                <Inp label="BOM Cost (VND)" fkey="bomCost" placeholder="92217" type="number" hint={form.fileName ? "Tự điền ✓" : ""} />
                <Inp label="Margin" fkey="margin" placeholder="0.9" type="number" />
                <Inp label="Exchange Rate (VND/USD)" fkey="rate" placeholder="26000" type="number" />
                <Inp label="QTY" fkey="qty" placeholder="1000" />
              </div>
              {form.bomCost && (
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginTop: "1rem" }}>
                  <div style={{ background: "var(--color-background-secondary)", borderRadius: 10, padding: "10px 14px" }}>
                    <div style={{ fontSize: 11, color: "var(--color-text-tertiary)", marginBottom: 4 }}>Selling Price VND</div>
                    <div style={{ fontSize: 16, fontWeight: 500 }}>{fmtVND(sv)}</div>
                  </div>
                  <div style={{ background: BLUE, borderRadius: 10, padding: "10px 14px" }}>
                    <div style={{ fontSize: 11, color: "rgba(255,255,255,0.6)", marginBottom: 4 }}>Selling Price USD</div>
                    <div style={{ fontSize: 16, fontWeight: 500, color: "#fff" }}>{fmtUSD(su)}</div>
                  </div>
                </div>
              )}
              <div style={{ display: "flex", gap: 8, marginTop: "1rem" }}>
                {!initial && <button onClick={() => setStep(1)}>← Quay lại</button>}
                {initial && <button onClick={onClose}>Hủy</button>}
                <button onClick={() => setStep(3)} disabled={!valid} style={{ flex: 1, background: BLUE, color: "#fff", border: "none", borderRadius: 8, padding: "9px", cursor: "pointer", fontSize: 14 }}>Tiếp theo →</button>
              </div>
            </div>
          )}

          {step === 3 && (
            <div>
              <p style={{ fontSize: 13, color: "var(--color-text-secondary)", marginBottom: "1rem" }}>Thông tin này sẽ xuất hiện trên PDF Quotation gửi khách.</p>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                <Inp label="To (Tên khách hàng)" fkey="to" placeholder="HENGQUN" span={2} />
                <Inp label="Attention" fkey="attention" placeholder="Tinkie Zhang" />
                <Inp label="Tel" fkey="tel" placeholder="+86..." />
                <Inp label="Term" fkey="term" placeholder="DAP — DONGGUAN, CHINA" span={2} />
                <Inp label="Remark dòng sản phẩm" fkey="remark" placeholder="Price included: Package + Freight" span={2} />
                <Inp label="Ghi chú nội bộ" fkey="notes" placeholder="Ghi chú..." span={2} />
              </div>
              <div style={{ display: "flex", gap: 8, marginTop: "1rem" }}>
                <button onClick={() => setStep(2)}>← Quay lại</button>
                <button onClick={() => setStep(4)} style={{ flex: 1, background: BLUE, color: "#fff", border: "none", borderRadius: 8, padding: "9px", cursor: "pointer", fontSize: 14 }}>Xem trước →</button>
              </div>
            </div>
          )}

          {step === 4 && (
            <div>
              <div style={{ background: "var(--color-background-secondary)", borderRadius: 12, padding: "1rem", marginBottom: "1rem" }}>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0.75rem", marginBottom: "0.75rem" }}>
                  {[{ l: "P/N", v: form.partNumber }, { l: "Ngày", v: form.date }, { l: "Tên sản phẩm", v: form.productName, s: 2 }, { l: "BOM Cost", v: fmtVND(form.bomCost) }, { l: "Margin", v: (parseFloat(form.margin) * 100).toFixed(0) + "%" }, { l: "QTY", v: form.qty }, { l: "To", v: form.to }, { l: "Term", v: form.term, s: 2 }].map(r => (
                    <div key={r.l} style={{ gridColumn: `span ${r.s || 1}` }}>
                      <div style={{ fontSize: 11, color: "var(--color-text-tertiary)", marginBottom: 2 }}>{r.l}</div>
                      <div style={{ fontWeight: 500, fontSize: 13 }}>{r.v || "—"}</div>
                    </div>
                  ))}
                </div>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, paddingTop: "0.75rem", borderTop: "0.5px solid var(--color-border-tertiary)" }}>
                  <div style={{ background: "var(--color-background-primary)", borderRadius: 8, padding: "10px 14px" }}>
                    <div style={{ fontSize: 11, color: "var(--color-text-tertiary)", marginBottom: 4 }}>Selling Price VND</div>
                    <div style={{ fontSize: 18, fontWeight: 500 }}>{fmtVND(sv)}</div>
                  </div>
                  <div style={{ background: BLUE, borderRadius: 8, padding: "10px 14px" }}>
                    <div style={{ fontSize: 11, color: "rgba(255,255,255,0.6)", marginBottom: 4 }}>Selling Price USD</div>
                    <div style={{ fontSize: 18, fontWeight: 500, color: "#fff" }}>{fmtUSD(su)}</div>
                  </div>
                </div>
              </div>
              <div style={{ display: "flex", gap: 8 }}>
                <button onClick={() => setStep(3)}>← Sửa lại</button>
                <button onClick={handleSave} disabled={saving} style={{ flex: 1, background: BLUE, color: "#fff", border: "none", borderRadius: 8, padding: "10px", cursor: "pointer", fontSize: 14, fontWeight: 500 }}>
                  {saving ? "Đang lưu..." : initial ? "Cập nhật" : "Lưu báo giá"}
                </button>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

export default function App() {
  const [user, setUser] = useState(null);
  const [records, setRecords] = useState([]);
  const [showModal, setShowModal] = useState(false);
  const [editRec, setEditRec] = useState(null);
  const [search, setSearch] = useState("");
  const [copied, setCopied] = useState(null);
  const [loading, setLoading] = useState(false);

  useEffect(() => {
    if (user) { setLoading(true); apiGetRecords(user.id).then(res => { setRecords(res.records || []); setLoading(false); }); }
  }, [user]);

  const saveRecord = rec => {
    setRecords(r => editRec ? r.map(x => x.id === rec.id ? rec : x) : [rec, ...r]);
    setShowModal(false); setEditRec(null);
  };
  const deleteRec = async id => { await apiDeleteRecord(id); setRecords(r => r.filter(x => x.id !== id)); };
  const copyQuote = rec => {
    const usd = sellUSD(rec.bom_cost, rec.margin, rec.exchange_rate).toFixed(4);
    const text = `BÁO GIÁ - ${rec.date}\nSản phẩm: ${rec.product_name}\nMã P/N: ${rec.part_number}\nSelling Price: $${usd} USD\nQTY: ${rec.qty || ""}\nTerm: ${rec.term || ""}`;
    navigator.clipboard.writeText(text).then(() => { setCopied(rec.id); setTimeout(() => setCopied(null), 2000); });
  };

  const filtered = records.filter(r => !search || (r.product_name||"").toLowerCase().includes(search.toLowerCase()) || (r.part_number||"").toLowerCase().includes(search.toLowerCase()));

  if (!user) return <LoginPage onLogin={setUser} />;

  return (
    <div style={{ minHeight: 500, background: "var(--color-background-tertiary)", paddingBottom: "2rem" }}>
      {showModal && <UploadModal onClose={() => { setShowModal(false); setEditRec(null); }} onSave={saveRecord} initial={editRec} userId={user.id} />}
      <div style={{ background: BLUE, padding: "0.875rem 1.5rem", display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: "1.5rem" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <i className="ti ti-table-import" style={{ fontSize: 20, color: "#fff" }} aria-hidden="true" />
          <span style={{ fontSize: 16, fontWeight: 500, color: "#fff" }}>BOM Price Manager</span>
          <span style={{ fontSize: 12, color: "rgba(255,255,255,0.5)" }}>COMART Việt Nam</span>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <span style={{ fontSize: 13, color: "rgba(255,255,255,0.8)" }}>{user.name}</span>
          <button onClick={() => setUser(null)} style={{ fontSize: 12, padding: "4px 10px", background: "rgba(255,255,255,0.1)", color: "#fff", border: "0.5px solid rgba(255,255,255,0.25)", borderRadius: 6, cursor: "pointer" }}>Đăng xuất</button>
        </div>
      </div>

      <div style={{ padding: "0 1.5rem" }}>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 12, marginBottom: "1.5rem" }}>
          {[
            { label: "Tổng báo giá", value: records.length, icon: "ti-file-description", accent: false },
            { label: "Tổng Selling USD", value: "$" + records.reduce((s, r) => s + sellUSD(r.bom_cost, r.margin, r.exchange_rate), 0).toFixed(2), icon: "ti-currency-dollar", accent: true },
            { label: "Tổng Selling VND", value: records.reduce((s, r) => s + sellVND(r.bom_cost, r.margin), 0).toLocaleString("vi-VN") + " ₫", icon: "ti-coin", accent: false },
          ].map(m => (
            <div key={m.label} style={{ background: m.accent ? BLUE : "var(--color-background-primary)", border: `0.5px solid ${m.accent ? BLUE : "var(--color-border-tertiary)"}`, borderRadius: 12, padding: "1rem 1.25rem" }}>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 8 }}>
                <span style={{ fontSize: 12, color: m.accent ? "rgba(255,255,255,0.7)" : "var(--color-text-secondary)" }}>{m.label}</span>
                <div style={{ width: 28, height: 28, borderRadius: 8, background: m.accent ? "rgba(255,255,255,0.15)" : BLUE_LIGHT, display: "flex", alignItems: "center", justifyContent: "center" }}>
                  <i className={`ti ${m.icon}`} style={{ fontSize: 14, color: m.accent ? "#fff" : BLUE }} aria-hidden="true" />
                </div>
              </div>
              <div style={{ fontSize: 20, fontWeight: 500, color: m.accent ? "#fff" : "var(--color-text-primary)" }}>{m.value}</div>
            </div>
          ))}
        </div>

        <div style={{ display: "flex", gap: 10, marginBottom: "1.25rem", alignItems: "center" }}>
          <button onClick={() => { setEditRec(null); setShowModal(true); }} style={{ display: "flex", alignItems: "center", gap: 8, background: BLUE, color: "#fff", border: "none", borderRadius: 8, padding: "9px 16px", fontSize: 14, cursor: "pointer", fontWeight: 500 }}>
            <i className="ti ti-upload" aria-hidden="true" /> Upload BOM
          </button>
          <input value={search} onChange={e => setSearch(e.target.value)} placeholder="Tìm kiếm P/N hoặc tên sản phẩm..." style={{ flex: 1 }} />
        </div>

        {loading ? (
          <div style={{ textAlign: "center", padding: "3rem", color: "var(--color-text-secondary)" }}>Đang tải...</div>
        ) : filtered.length === 0 ? (
          <div style={{ background: "var(--color-background-primary)", border: "0.5px solid var(--color-border-tertiary)", borderRadius: 12, padding: "3rem", textAlign: "center" }}>
            <div style={{ width: 64, height: 64, borderRadius: 16, background: BLUE_LIGHT, display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 12px" }}>
              <i className="ti ti-file-upload" style={{ fontSize: 30, color: BLUE }} aria-hidden="true" />
            </div>
            <p style={{ fontWeight: 500, margin: "0 0 6px" }}>{records.length === 0 ? "Chưa có báo giá nào" : "Không tìm thấy kết quả"}</p>
            <p style={{ fontSize: 13, color: "var(--color-text-secondary)", margin: "0 0 16px" }}>Upload file Excel BOM để tự động điền giá và xuất Quotation PDF</p>
            {records.length === 0 && <button onClick={() => setShowModal(true)} style={{ background: BLUE, color: "#fff", border: "none", borderRadius: 8, padding: "9px 20px", cursor: "pointer", fontSize: 14, fontWeight: 500 }}><i className="ti ti-upload" style={{ marginRight: 6 }} aria-hidden="true" />Upload BOM đầu tiên</button>}
          </div>
        ) : (
          <div style={{ background: "var(--color-background-primary)", border: "0.5px solid var(--color-border-tertiary)", borderRadius: 12, overflow: "hidden" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13, tableLayout: "fixed" }}>
              <colgroup><col style={{ width: "13%" }} /><col style={{ width: "22%" }} /><col style={{ width: "10%" }} /><col style={{ width: "8%" }} /><col style={{ width: "12%" }} /><col style={{ width: "11%" }} /><col style={{ width: "9%" }} /><col style={{ width: "15%" }} /></colgroup>
              <thead>
                <tr style={{ background: BLUE_LIGHT }}>
                  {["P/N", "Tên sản phẩm", "Cost (VND)", "Margin", "Selling VND", "Selling USD", "Ngày", "Thao tác"].map(h => (
                    <th key={h} style={{ padding: "10px 12px", textAlign: "left", fontWeight: 500, fontSize: 12, color: BLUE_DARK, borderBottom: `1px solid ${BLUE}22` }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {filtered.map((r, i) => (
                  <tr key={r.id} style={{ borderBottom: "0.5px solid var(--color-border-tertiary)", background: i % 2 === 0 ? "var(--color-background-primary)" : "var(--color-background-secondary)" }}>
                    <td style={{ padding: "10px 12px", fontWeight: 500, color: BLUE, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{r.part_number}</td>
                    <td style={{ padding: "10px 12px", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }} title={r.product_name}>{r.product_name}</td>
                    <td style={{ padding: "10px 12px", color: "var(--color-text-secondary)" }}>{Number(r.bom_cost).toLocaleString("vi-VN")}</td>
                    <td style={{ padding: "10px 12px" }}><span style={{ background: BLUE_LIGHT, color: BLUE, borderRadius: 6, padding: "2px 7px", fontSize: 11, fontWeight: 500 }}>{(r.margin * 100).toFixed(0)}%</span></td>
                    <td style={{ padding: "10px 12px", fontSize: 12 }}>{fmtVND(sellVND(r.bom_cost, r.margin))}</td>
                    <td style={{ padding: "10px 12px" }}><span style={{ fontWeight: 500, color: "#fff", background: BLUE, borderRadius: 6, padding: "3px 8px", fontSize: 12 }}>{fmtUSD(sellUSD(r.bom_cost, r.margin, r.exchange_rate))}</span></td>
                    <td style={{ padding: "10px 12px", color: "var(--color-text-secondary)", fontSize: 12 }}>{r.date}</td>
                    <td style={{ padding: "8px 10px" }}>
                      <div style={{ display: "flex", gap: 3 }}>
                        <button onClick={() => genPDF(r)} title="Export PDF" style={{ padding: "4px 7px", fontSize: 11, color: RED, borderColor: RED }}><i className="ti ti-file-type-pdf" aria-hidden="true" /></button>
                        <button onClick={() => copyQuote(r)} title="Copy báo giá" style={{ padding: "4px 7px", fontSize: 11, color: copied === r.id ? "#16a34a" : BLUE, borderColor: copied === r.id ? "#16a34a" : BLUE }}><i className={`ti ${copied === r.id ? "ti-check" : "ti-copy"}`} aria-hidden="true" /></button>
                        <button onClick={() => { setEditRec(r); setShowModal(true); }} title="Sửa" style={{ padding: "4px 7px", fontSize: 11 }}><i className="ti ti-edit" aria-hidden="true" /></button>
                        <button onClick={() => deleteRec(r.id)} title="Xóa" style={{ padding: "4px 7px", fontSize: 11, color: "var(--color-text-danger)" }}><i className="ti ti-trash" aria-hidden="true" /></button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}
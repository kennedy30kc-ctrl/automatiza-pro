import { useState, useEffect, useRef, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  LogIn, LogOut, CreditCard, QrCode, CheckCircle2, AlertCircle,
  Home, Settings, Users, Building2, Phone, Mail, MapPin,
  Eye, EyeOff, Edit2, Trash2, Shield, BarChart3, ChevronLeft,
  ChevronRight, RefreshCw, Check, X, AlertTriangle, Key, Bell, Download, FileSpreadsheet
} from "lucide-react";

// ══════════════════════════════════════════
// LOGOS
// ══════════════════════════════════════════
const LogoAP = ({ size = 100 }) => (
  <svg width={size} height={size} viewBox="0 0 200 200" xmlns="http://www.w3.org/2000/svg">
    <defs>
      <linearGradient id="apG" x1="0%" y1="100%" x2="100%" y2="0%">
        <stop offset="0%" stopColor="#86efac"/><stop offset="50%" stopColor="#22d3ee"/><stop offset="100%" stopColor="#60a5fa"/>
      </linearGradient>
    </defs>
    <circle cx="100" cy="100" r="99" fill="#0d1a2a"/>
    <circle cx="100" cy="100" r="94" fill="none" stroke="url(#apG)" strokeWidth="8"/>
    <circle cx="100" cy="100" r="63" fill="white"/>
    {[0,36,72,108,144,180,216,252,288,324].map((a,i)=>(
      <rect key={i} x="88" y="15" width="24" height="22" rx="3" fill="white" transform={`rotate(${a} 100 100)`}/>
    ))}
    <circle cx="100" cy="100" r="46" fill="#0d1a2a"/>
    <path d="M97 58 A42 42 0 1 0 97 142" fill="none" stroke="#65a30d" strokeWidth="15" strokeLinecap="round"/>
    <path d="M106 86 L128 86 L148 68" fill="none" stroke="white" strokeWidth="4.5" strokeLinecap="round" strokeLinejoin="round"/>
    <circle cx="152" cy="65" r="6" fill="white"/>
    <path d="M110 100 L152 100" fill="none" stroke="#65a30d" strokeWidth="4.5" strokeLinecap="round"/>
    <circle cx="156" cy="100" r="6" fill="#65a30d"/>
    <path d="M106 114 L128 114 L148 132" fill="none" stroke="white" strokeWidth="4.5" strokeLinecap="round" strokeLinejoin="round"/>
    <circle cx="152" cy="135" r="6" fill="white"/>
  </svg>
);

const LogoNA = ({ h = 60 }) => (
  <svg width={h*2.8} height={h} viewBox="0 0 280 100" xmlns="http://www.w3.org/2000/svg">
    <rect width="280" height="100" fill="#111" rx="8"/>
    <path d="M28 18 L16 18 L16 82 L28 82" fill="none" stroke="#1e6e35" strokeWidth="6" strokeLinecap="round" strokeLinejoin="round"/>
    <rect x="22" y="60" width="9" height="18" fill="#2ea84f" rx="1.5"/>
    <rect x="34" y="50" width="9" height="28" fill="#2ea84f" rx="1.5"/>
    <rect x="46" y="38" width="9" height="40" fill="#2ea84f" rx="1.5"/>
    <line x1="55" y1="78" x2="55" y2="14" stroke="#2ea84f" strokeWidth="5.5" strokeLinecap="round"/>
    <polygon points="55,6 49,20 61,20" fill="#2ea84f"/>
    <text x="74" y="52" fontFamily="Arial Black,Arial" fontWeight="900" fontSize="31" fill="white">NÓMINA</text>
    <text x="74" y="84" fontFamily="Arial Black,Arial" fontWeight="900" fontSize="31" fill="#2ea84f">ACTIVA</text>
  </svg>
);

// ══════════════════════════════════════════
// STORAGE
// ══════════════════════════════════════════
const CLAVE_MAESTRA = "NOMACTIVA2025";
const st = {
  async get(key, shared = true) {
    try { const r = await window.storage.get(key, shared); return r ? JSON.parse(r.value) : null; } catch { return null; }
  },
  async set(key, val, shared = true) {
    try { await window.storage.set(key, JSON.stringify(val), shared); return true; } catch { return false; }
  },
  async list(prefix, shared = true) {
    try { const r = await window.storage.list(prefix, shared); return r?.keys || []; } catch { return []; }
  }
};

// ══════════════════════════════════════════
// CONFIRM MODAL (reemplaza window.confirm)
// ══════════════════════════════════════════
function ConfirmModal({ text, onYes, onNo }) {
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4" style={{ backgroundColor: "rgba(0,0,0,0.7)" }}>
      <div className="rounded-xl p-6 w-full max-w-xs" style={{ backgroundColor: "#1e293b" }}>
        <AlertTriangle size={32} style={{ color: "#f59e0b" }} className="mx-auto mb-3"/>
        <p className="text-white text-center mb-5 text-sm">{text}</p>
        <div className="flex gap-3">
          <button onClick={onNo} className="flex-1 py-3 rounded-lg font-bold text-white bg-gray-600">Cancelar</button>
          <button onClick={onYes} className="flex-1 py-3 rounded-lg font-bold text-white bg-red-600">Eliminar</button>
        </div>
      </div>
    </div>
  );
}

// ══════════════════════════════════════════
// FLASH MESSAGE
// ══════════════════════════════════════════
function FlashMsg({ msg }) {
  if (!msg.t) return null;
  return (
    <div className={`rounded-lg p-3 mb-4 flex items-start gap-2 ${msg.tipo === "ok" ? "bg-green-500/20 border-l-4 border-green-500" : "bg-red-500/20 border-l-4 border-red-500"}`}>
      {msg.tipo === "ok" ? <CheckCircle2 size={16} className="text-green-400 flex-shrink-0 mt-0.5"/> : <AlertCircle size={16} className="text-red-400 flex-shrink-0 mt-0.5"/>}
      <p className="text-sm whitespace-pre-line">{msg.t}</p>
    </div>
  );
}

// ══════════════════════════════════════════
// BACK BUTTON
// ══════════════════════════════════════════
function BackBtn({ onClick, label = "Volver" }) {
  return (
    <button onClick={onClick} className="mb-5 flex items-center gap-2 px-3 py-2 rounded-lg text-sm text-white" style={{ backgroundColor: "#334155" }}>
      <ChevronLeft size={15}/>{label}
    </button>
  );
}

// ══════════════════════════════════════════
// CAMERA VIEW (standalone component outside App)
// ══════════════════════════════════════════
function CamView({ color, videoRef, canvasRef, onStop }) {
  return (
    <div className="space-y-3">
      <div className="relative rounded-xl overflow-hidden bg-black" style={{ aspectRatio: "1" }}>
        <video ref={videoRef} autoPlay playsInline className="w-full h-full object-cover"/>
        <canvas ref={canvasRef} className="hidden"/>
        <div className="absolute inset-0 flex items-center justify-center">
          <div className="rounded-xl" style={{ width: "55%", height: "55%", border: `2px solid ${color}`, boxShadow: `0 0 20px ${color}55` }}/>
        </div>
        <p className="absolute bottom-3 w-full text-center text-xs" style={{ color: "#94a3b8" }}>Apunta el QR al cuadro</p>
      </div>
      <button onClick={onStop} className="w-full py-3 rounded-lg font-bold text-white bg-red-600">Cancelar Escaneo</button>
    </div>
  );
}

// ══════════════════════════════════════════
// MAIN APP
// ══════════════════════════════════════════
// ══════════════════════════════════════════
// CARD VIEW COMPONENT
// ══════════════════════════════════════════
function CardView({ emp, empresaNombre, empresaLogo, onClose, onPhotoUpload }) {
  const qrRef = useRef(null);
  const [qrReady, setQrReady] = useState(false);

  useEffect(() => {
    if (!qrRef.current) return;
    qrRef.current.innerHTML = "";
    const tryGen = () => {
      if (window.QRCode) {
        try {
          new window.QRCode(qrRef.current, {
            text: JSON.stringify({ n: emp.nombre, c: emp.clave }),
            width: 120, height: 120,
            colorDark: "#1e293b", colorLight: "#ffffff"
          });
          setTimeout(() => setQrReady(true), 300);
        } catch { qrRef.current.innerHTML = ""; }
      } else { setTimeout(tryGen, 400); }
    };
    tryGen();
  }, [emp]);

  const downloadCard = () => {
    const W = 900; const H = 520;
    const canvas = document.createElement("canvas");
    canvas.width = W; canvas.height = H;
    const ctx = canvas.getContext("2d");

    const drawAll = (photoImg, logoImg) => {
      // ── Background white ──
      ctx.fillStyle = "#ffffff";
      ctx.fillRect(0, 0, W, H);

      // ── Top header bar (company) ──
      ctx.fillStyle = "#1e293b";
      ctx.fillRect(0, 0, W, 90);

      // Company logo in header
      if (logoImg) {
        ctx.save();
        ctx.beginPath();
        ctx.arc(55, 45, 34, 0, Math.PI * 2);
        ctx.clip();
        ctx.drawImage(logoImg, 21, 11, 68, 68);
        ctx.restore();
      } else {
        // placeholder circle with initial
        ctx.fillStyle = "#334155";
        ctx.beginPath(); ctx.arc(55, 45, 34, 0, Math.PI * 2); ctx.fill();
        ctx.fillStyle = "#94a3b8"; ctx.font = "bold 28px Arial";
        ctx.textAlign = "center";
        ctx.fillText((empresaNombre || "E")[0].toUpperCase(), 55, 55);
        ctx.textAlign = "left";
      }

      // Company name in header
      ctx.fillStyle = "#ffffff";
      ctx.font = "bold 28px Arial Black, Arial";
      ctx.fillText(empresaNombre || "", 105, 40);
      ctx.fillStyle = "#94a3b8"; ctx.font = "13px Arial";
      ctx.fillText("Tarjeta de Acceso de Empleado", 105, 65);

      // Year right
      ctx.fillStyle = "#64748b"; ctx.font = "12px Arial";
      ctx.textAlign = "right";
      ctx.fillText(new Date().getFullYear(), W - 24, 50);
      ctx.textAlign = "left";

      // ── Accent left stripe ──
      ctx.fillStyle = "#3b82f6";
      ctx.fillRect(0, 90, 6, H - 90);

      // ── Body background ──
      ctx.fillStyle = "#f8fafc";
      ctx.fillRect(6, 90, W - 6, H - 150);

      // ── Photo circle ──
      const cx = 115; const cy = 300; const r = 88;
      // shadow
      ctx.shadowColor = "rgba(0,0,0,0.15)"; ctx.shadowBlur = 12;
      ctx.fillStyle = "#e2e8f0";
      ctx.beginPath(); ctx.arc(cx, cy, r, 0, Math.PI * 2); ctx.fill();
      ctx.shadowBlur = 0;
      // border
      ctx.strokeStyle = "#3b82f6"; ctx.lineWidth = 4;
      ctx.beginPath(); ctx.arc(cx, cy, r, 0, Math.PI * 2); ctx.stroke();

      if (photoImg) {
        ctx.save();
        ctx.beginPath(); ctx.arc(cx, cy, r - 3, 0, Math.PI * 2); ctx.clip();
        ctx.drawImage(photoImg, cx - r + 3, cy - r + 3, (r - 3) * 2, (r - 3) * 2);
        ctx.restore();
      } else {
        ctx.fillStyle = "#3b82f6"; ctx.font = "bold 70px Arial";
        ctx.textAlign = "center";
        ctx.fillText((emp.nombre || "?")[0].toUpperCase(), cx, cy + 24);
        ctx.textAlign = "left";
      }

      // ── Employee info (center) ──
      const iX = 240;
      // Name label
      ctx.fillStyle = "#94a3b8"; ctx.font = "11px Arial";
      ctx.fillText("NOMBRE COMPLETO", iX, 125);
      // Name value
      ctx.fillStyle = "#0f172a"; ctx.font = "bold 30px Arial Black, Arial";
      ctx.fillText(emp.nombre || "", iX, 158);
      // Divider
      ctx.fillStyle = "#e2e8f0"; ctx.fillRect(iX, 168, 370, 2);

      // Fields
      const fields = [
        ["CÉDULA / ID", emp.cedula || "No registrada"],
        ["TELÉFONO", emp.telefono || "No registrado"],
        ["CORREO", emp.correo || "No registrado"],
        ["DIRECCIÓN", emp.direccion || "No registrada"],
      ];
      fields.forEach(([label, val], i) => {
        const y = 195 + i * 58;
        ctx.fillStyle = "#64748b"; ctx.font = "10px Arial";
        ctx.fillText(label, iX, y);
        ctx.fillStyle = "#1e293b"; ctx.font = "15px Arial";
        ctx.fillText(val.length > 38 ? val.slice(0, 36) + "..." : val, iX, y + 20);
        ctx.fillStyle = "#f1f5f9"; ctx.fillRect(iX, y + 28, 370, 1);
      });

      // ── QR code (right) ──
      const qrCanvas = qrRef.current?.querySelector("canvas");
      const qX = W - 175; const qY = 110;
      // QR background card
      ctx.shadowColor = "rgba(0,0,0,0.10)"; ctx.shadowBlur = 10;
      ctx.fillStyle = "#ffffff";
      if (ctx.roundRect) { ctx.roundRect(qX - 14, qY - 14, 152, 180, 10); ctx.fill(); }
      else { ctx.fillRect(qX - 14, qY - 14, 152, 180); }
      ctx.shadowBlur = 0;
      if (qrCanvas) ctx.drawImage(qrCanvas, qX, qY, 124, 124);
      ctx.fillStyle = "#64748b"; ctx.font = "bold 11px Arial";
      ctx.textAlign = "center";
      ctx.fillText("CÓDIGO DE ACCESO", qX + 62, qY + 140);
      ctx.font = "10px Arial";
      ctx.fillText("Escanear en tablet", qX + 62, qY + 156);
      ctx.textAlign = "left";

      // ── Bottom footer ──
      ctx.fillStyle = "#1e293b"; ctx.fillRect(0, H - 60, W, 60);
      ctx.fillStyle = "#3b82f6"; ctx.font = "bold 12px Arial";
      ctx.fillText("Powered by Automatiza Pro · Nómina Activa", 24, H - 32);
      ctx.fillStyle = "#475569"; ctx.font = "11px Arial";
      ctx.fillText(`Emitida: ${new Date().toLocaleDateString("es-ES")}`, 24, H - 14);
      // ID right
      ctx.textAlign = "right";
      ctx.fillStyle = "#334155"; ctx.font = "10px Arial";
      ctx.fillText(`ID: ${emp.id}`, W - 20, H - 20);
      ctx.textAlign = "left";

      // download
      const link = document.createElement("a");
      link.download = `carnet_${(emp.nombre || "empleado").replace(/\s+/g, "_")}.png`;
      link.href = canvas.toDataURL("image/png");
      link.click();
    };

    // Load photo and logo, then draw
    const loadImg = (src) => new Promise((res) => {
      if (!src) return res(null);
      const img = new Image();
      img.onload = () => res(img);
      img.onerror = () => res(null);
      img.src = src;
    });

    Promise.all([loadImg(emp.foto), loadImg(empresaLogo)]).then(([p, l]) => drawAll(p, l));
  };

  return (
    <div style={{ backgroundColor: "#f1f5f9", minHeight: "100vh" }} className="p-4 flex items-start justify-center">
      <div className="w-full max-w-lg mt-4">
        <button onClick={onClose} className="mb-4 flex items-center gap-2 px-3 py-2 rounded-lg text-sm text-white" style={{ backgroundColor: "#334155" }}>
          <ChevronLeft size={15}/>Cerrar
        </button>

        {/* ── CARD PREVIEW ── */}
        <div className="rounded-2xl overflow-hidden shadow-2xl mb-4 bg-white" style={{ border: "1px solid #e2e8f0" }}>

          {/* Top header — Company */}
          <div className="flex items-center gap-3 px-5 py-4" style={{ backgroundColor: "#1e293b" }}>
            {/* Logo */}
            <div className="rounded-full overflow-hidden flex-shrink-0 flex items-center justify-center" style={{ width: 52, height: 52, backgroundColor: "#334155", border: "2px solid #3b82f6" }}>
              {empresaLogo
                ? <img src={empresaLogo} alt="logo" className="w-full h-full object-cover"/>
                : <span className="font-black text-2xl text-white">{(empresaNombre || "E")[0].toUpperCase()}</span>
              }
            </div>
            <div className="flex-1">
              <p className="font-black text-white text-lg leading-tight">{empresaNombre}</p>
              <p className="text-xs" style={{ color: "#94a3b8" }}>Tarjeta de Acceso de Empleado</p>
            </div>
            <p className="text-xs" style={{ color: "#64748b" }}>{new Date().getFullYear()}</p>
          </div>

          {/* Body */}
          <div className="flex gap-4 p-5" style={{ backgroundColor: "#f8fafc" }}>

            {/* Photo column */}
            <div className="flex flex-col items-center gap-2 flex-shrink-0">
              <div className="rounded-full overflow-hidden flex items-center justify-center" style={{ width: 100, height: 100, border: "3px solid #3b82f6", backgroundColor: "#e2e8f0" }}>
                {emp.foto
                  ? <img src={emp.foto} alt="foto" className="w-full h-full object-cover"/>
                  : <span className="font-black text-4xl" style={{ color: "#3b82f6" }}>{(emp.nombre || "?")[0].toUpperCase()}</span>
                }
              </div>
              {onPhotoUpload && (
                <label className="cursor-pointer text-xs px-2 py-1 rounded font-bold" style={{ backgroundColor: "#dbeafe", color: "#3b82f6" }}>
                  📷 Foto
                  <input type="file" accept="image/*" className="hidden" onChange={onPhotoUpload}/>
                </label>
              )}
            </div>

            {/* Info column */}
            <div className="flex-1 min-w-0">
              <p className="text-xs mb-0.5" style={{ color: "#94a3b8" }}>NOMBRE COMPLETO</p>
              <p className="font-black text-xl leading-tight mb-2" style={{ color: "#0f172a" }}>{emp.nombre}</p>
              <div className="space-y-1" style={{ borderTop: "2px solid #e2e8f0", paddingTop: 8 }}>
                {[
                  ["Cédula", emp.cedula || "No registrada"],
                  ["Teléfono", emp.telefono || "No registrado"],
                  ["Correo", emp.correo || "No registrado"],
                  ["Dirección", emp.direccion || "No registrada"],
                ].map(([l, v]) => (
                  <div key={l} className="flex gap-1">
                    <span className="text-xs w-16 flex-shrink-0 font-bold" style={{ color: "#64748b" }}>{l}:</span>
                    <span className="text-xs truncate" style={{ color: "#475569" }}>{v}</span>
                  </div>
                ))}
              </div>
            </div>

            {/* QR column */}
            <div className="flex flex-col items-center gap-1 flex-shrink-0">
              <div className="bg-white p-2 rounded-xl shadow-sm" style={{ border: "1px solid #e2e8f0" }}>
                <div ref={qrRef}/>
              </div>
              <p className="text-xs font-bold text-center" style={{ color: "#64748b" }}>Acceso</p>
            </div>
          </div>

          {/* Footer */}
          <div className="px-5 py-2 flex justify-between items-center" style={{ backgroundColor: "#1e293b" }}>
            <p className="text-xs" style={{ color: "#3b82f6" }}>Powered by Automatiza Pro · Nómina Activa</p>
            <p className="text-xs" style={{ color: "#475569" }}>{new Date().toLocaleDateString("es-ES")}</p>
          </div>
        </div>

        {/* Download */}
        <button
          onClick={downloadCard}
          disabled={!qrReady}
          className="w-full py-4 rounded-xl font-bold text-white text-lg disabled:opacity-50"
          style={{ backgroundColor: "#1e293b" }}
        >
          {qrReady ? "⬇️ Descargar Carnet PNG" : "Generando QR..."}
        </button>
        <p className="text-center text-xs mt-2" style={{ color: "#94a3b8" }}>El empleado guarda esta imagen en su teléfono para acceder</p>
      </div>
    </div>
  );
}

export default function App() {
  const [screen, setScreen] = useState("loading");
  const [empresa, setEmpresa] = useState(null);
  const [empId, setEmpId] = useState(null);
  const [msg, setMsg] = useState({ t: "", tipo: "" });
  const [confirm, setConfirm] = useState(null); // { text, onYes }

  // employee flow
  const [empleadoId, setEmpleadoId] = useState(null); // store ID only, look up fresh data
  const [esNuevo, setEsNuevo] = useState(false);
  const [nombreInput, setNombreInput] = useState("");
  const [claveInput, setClaveInput] = useState("");
  const [telefonoInput, setTelefonoInput] = useState("");

  // camera
  const [camActive, setCamActive] = useState(false);
  const videoRef = useRef(null);
  const canvasRef = useRef(null);
  const rafRef = useRef(null);

  // loan flow
  const [loanStep, setLoanStep] = useState(1);
  const [loanMonto, setLoanMonto] = useState("");
  const [loanPlazo, setLoanPlazo] = useState(1);
  const [loanFreq, setLoanFreq] = useState("mensual");
  const [loanTel, setLoanTel] = useState("");
  const [loanId, setLoanId] = useState("");

  // jefe
  const [jefePass, setJefePass] = useState("");
  const [jefeSub, setJefeSub] = useState("menu");
  const [dashPeriodo, setDashPeriodo] = useState("hoy");
  const [permisoModal, setPermisoModal] = useState(null); // { empId, fecha, nombre }
  const [permisoNota, setPermisoNota] = useState("");
  const [alertaModal, setAlertaModal] = useState(null); // { empId, registroIdx, nombre }
  const [horaSalidaManual, setHoraSalidaManual] = useState("");
  const [horaEntradaEdit, setHoraEntradaEdit] = useState(false);
  const [logoEdit, setLogoEdit] = useState(false);
  // password change
  const [claveActual, setClaveActual] = useState("");
  const [claveNueva, setClaveNueva] = useState("");
  const [claveConf, setClaveConf] = useState("");
  const [editEmpId, setEditEmpId] = useState(null);
  const [fNombre, setFNombre] = useState("");
  const [fClave, setFClave] = useState("");
  const [fTel, setFTel] = useState("");
  const [fCed, setFCed] = useState("");
  const [fCorreo, setFCorreo] = useState("");
  const [fDir, setFDir] = useState("");
  const [fSal, setFSal] = useState("");
  const [fDias, setFDias] = useState("20");
  const [fFoto, setFFoto] = useState("");
  const [qrEmp, setQrEmp] = useState(null);
  const qrRef = useRef(null);
  const [showPass, setShowPass] = useState(false);

  // setup
  const [sNombre, setSNombre] = useState("");
  const [sClave, setSClave] = useState("");
  const [sConf, setSConf] = useState("");

  // nomina activa
  const [naPass, setNaPass] = useState("");
  const [naView, setNaView] = useState("empresas");
  const [allEmpresas, setAllEmpresas] = useState([]);
  const [selEmpresa, setSelEmpresa] = useState(null);
  const [selEmpId2, setSelEmpId2] = useState(null);

  // helpers: get current employee fresh from empresa state
  const getEmpleado = useCallback((id, emp = empresa) => emp?.empleados?.find(e => e.id === id) || null, [empresa]);

  useEffect(() => {
    ["https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js",
     "https://cdnjs.cloudflare.com/ajax/libs/jsQR/1.4.0/jsQR.min.js"].forEach(src => {
      if (!document.querySelector(`script[src="${src}"]`)) {
        const s = document.createElement("script"); s.src = src; document.head.appendChild(s);
      }
    });
  }, []);

  useEffect(() => { init(); }, []);

  const init = async () => {
    const id = await st.get("tablet_empresa_id", false);
    if (id) {
      const datos = await st.get(`empresa:${id}`);
      if (datos) { setEmpId(id); setEmpresa(datos); setScreen("principal"); return; }
    }
    setScreen("setup");
  };

  const save = async (nueva) => {
    setEmpresa(nueva);
    if (empId) await st.set(`empresa:${empId}`, nueva);
  };

  const flash = (t, tipo = "ok") => { setMsg({ t, tipo }); setTimeout(() => setMsg({ t: "", tipo: "" }), 3500); };

  const resetFlow = useCallback(() => {
    setEmpleadoId(null); setEsNuevo(false);
    setNombreInput(""); setClaveInput(""); setTelefonoInput("");
    setLoanStep(1); setLoanMonto(""); setLoanPlazo(1); setLoanFreq("mensual"); setLoanTel("");
    stopCam();
  }, []);

  const goHome = useCallback(() => { resetFlow(); setScreen("principal"); }, [resetFlow]);

  // ── calc helpers ──
  const getSal = (id, emp = empresa) => emp?.salarios?.[id];
  const getLimite = (id) => { const s = getSal(id); return s ? (s.salarioMensual * 0.4).toFixed(2) : "0.00"; };
  const hasLoan = (id) => empresa?.prestamos?.some(p => p.empId === id && p.estado === "activo");
  const calcLoan = (monto, plazo, freq) => {
    const m = parseFloat(monto) || 0;
    const int = m * 0.05 * plazo;
    const total = m + int;
    const n = freq === "mensual" ? plazo : freq === "quincenal" ? plazo * 2 : Math.round(plazo * 4.3);
    return { monto: m.toFixed(2), int: int.toFixed(2), total: total.toFixed(2), n, cuota: (total / n).toFixed(2) };
  };
  const hHoy = (emp) => {
    const hoy = new Date().toLocaleDateString("es-ES");
    return (emp?.registros || []).filter(r => r.fecha === hoy && r.hsTrab).reduce((s, r) => s + parseFloat(r.hsTrab), 0);
  };
  const hMes = (emp) => {
    const mes = new Date().toLocaleString("es-ES", { month: "2-digit", year: "numeric" });
    return (emp?.registros || []).filter(r => {
      try { return new Date(r.fecha).toLocaleString("es-ES", { month: "2-digit", year: "numeric" }) === mes && r.hsTrab; } catch { return false; }
    }).reduce((s, r) => s + parseFloat(r.hsTrab), 0);
  };

  // ── camera ──
  const stopCam = () => {
    if (videoRef.current?.srcObject) videoRef.current.srcObject.getTracks().forEach(t => t.stop());
    if (rafRef.current) cancelAnimationFrame(rafRef.current);
    setCamActive(false);
  };

  const startCam = async () => {
    try {
      const stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: "environment" } });
      videoRef.current.srcObject = stream;
      await videoRef.current.play();
      setCamActive(true);
      const loop = () => {
        const v = videoRef.current; const c = canvasRef.current;
        if (v?.videoWidth && c) {
          c.width = v.videoWidth; c.height = v.videoHeight;
          c.getContext("2d").drawImage(v, 0, 0);
          if (window.jsQR) {
            const img = c.getContext("2d").getImageData(0, 0, c.width, c.height);
            const code = window.jsQR(img.data, img.width, img.height);
            if (code?.data) {
              try {
                const d = JSON.parse(code.data);
                const emp = empresa?.empleados?.find(e => e.nombre === d.n && e.clave === d.c);
                if (emp) { stopCam(); setEmpleadoId(emp.id); return; }
              } catch {}
            }
          }
        }
        rafRef.current = requestAnimationFrame(loop);
      };
      rafRef.current = requestAnimationFrame(loop);
    } catch { flash("No se pudo acceder a la cámara. Usa entrada manual.", "err"); }
  };

  // ── QR ──
  const genQR = (emp) => {
    setQrEmp(emp);
    setTimeout(() => {
      if (!qrRef.current) return;
      qrRef.current.innerHTML = "";
      const tryGen = () => {
        if (window.QRCode) {
          try { new window.QRCode(qrRef.current, { text: JSON.stringify({ n: emp.nombre, c: emp.clave }), width: 180, height: 180 }); }
          catch { qrRef.current.innerHTML = `<p style="color:#94a3b8;font-size:11px;padding:12px">Datos: ${emp.nombre}</p>`; }
        } else setTimeout(tryGen, 400);
      };
      tryGen();
    }, 100);
  };

  // ── verify employee ──
  const verificar = () => {
    const nombre = nombreInput.trim();
    const clave = claveInput.trim();
    if (!nombre || !clave) { flash("Completa nombre y clave", "err"); return; }
    // Look for existing employee (match name + clave)
    let emp = empresa?.empleados?.find(e => e.nombre.toLowerCase() === nombre.toLowerCase() && e.clave === clave);
    if (!emp) {
      // New employee self-registration
      emp = {
        id: Date.now(), nombre, clave,
        telefono: telefonoInput || "", cedula: "", correo: "", direccion: "",
        perfilCompleto: false, horasTrabajadas: 0, registros: [],
        fechaCreacion: new Date().toISOString()
      };
      const nueva = {
        ...empresa,
        empleados: [...(empresa.empleados || []), emp],
        salarios: { ...empresa.salarios, [emp.id]: { salarioMensual: 250, diasLaborales: 20 } }
      };
      save(nueva);
      setEsNuevo(true);
    }
    setEmpleadoId(emp.id);
  };

  // ── ENTRADA (always reads fresh data) ──
  const entrada = () => {
    const ahora = new Date();
    const hoy = ahora.toLocaleDateString("es-ES");
    // Get FRESH employee data from empresa state
    const emp = empresa?.empleados?.find(e => e.id === empleadoId);
    if (!emp) return;
    const yaEntro = emp.registros?.some(r => r.tipo === "in" && r.fecha === hoy && !r.salida);
    if (yaEntro) {
      flash("⚠️ Ya tienes entrada registrada hoy.\nUsa SALIDA para registrar tu salida.", "err");
      setTimeout(goHome, 2800);
      return;
    }
    const nuevaEmpresa = {
      ...empresa,
      empleados: empresa.empleados.map(e => e.id === empleadoId
        ? { ...e, registros: [...(e.registros || []), { tipo: "in", hora: ahora.toISOString(), fecha: hoy, horaReg: ahora.toLocaleTimeString("es-ES") }] }
        : e)
    };
    save(nuevaEmpresa);
    flash(`✅ ENTRADA REGISTRADA · ${ahora.toLocaleTimeString("es-ES")}`);
    setTimeout(goHome, 2000);
  };

  // ── SALIDA ──
  const salida = () => {
    const ahora = new Date();
    const emp = empresa?.empleados?.find(e => e.id === empleadoId);
    if (!emp) return;
    const ult = emp.registros?.find(r => r.tipo === "in" && !r.salida);
    if (!ult) { flash("⚠️ No tienes entrada registrada hoy", "err"); setTimeout(goHome, 2000); return; }
    const hs = (ahora - new Date(ult.hora)) / 3600000;
    const s = getSal(empleadoId);
    const tarifa = s ? s.salarioMensual / s.diasLaborales / 8 : 0;
    const nuevaEmpresa = {
      ...empresa,
      empleados: empresa.empleados.map(e => e.id === empleadoId
        ? { ...e, horasTrabajadas: (e.horasTrabajadas || 0) + hs, registros: e.registros.map(r => r === ult ? { ...r, salida: ahora.toISOString(), hsTrab: hs.toFixed(2) } : r) }
        : e)
    };
    save(nuevaEmpresa);
    flash(`✅ SALIDA REGISTRADA · ${hs.toFixed(2)}h · $${(hs * tarifa).toFixed(2)}`);
    setTimeout(goHome, 2000);
  };

  // ── LOAN ──
  const confirmarLoan = () => {
    const emp = empresa?.empleados?.find(e => e.id === empleadoId);
    if (!emp) return;
    const loan = calcLoan(loanMonto, loanPlazo, loanFreq);
    const cuotas = Array.from({ length: loan.n }, (_, i) => {
      const f = new Date();
      if (loanFreq === "mensual") f.setMonth(f.getMonth() + i + 1);
      else if (loanFreq === "quincenal") f.setDate(f.getDate() + (i + 1) * 15);
      else f.setDate(f.getDate() + (i + 1) * 7);
      return { num: i + 1, monto: parseFloat(loan.cuota), vence: f.toLocaleDateString("es-ES"), estado: "pendiente", pagado: null };
    });
    const id = `AP-${Date.now()}`;
    save({
      ...empresa,
      prestamos: [...(empresa.prestamos || []), {
        id, empId: empleadoId, empNombre: emp.nombre,
        montoSol: parseFloat(loanMonto), plazo: loanPlazo, frecuencia: loanFreq,
        interes: parseFloat(loan.int), total: parseFloat(loan.total),
        nCuotas: loan.n, cuota: parseFloat(loan.cuota), cuotas,
        estado: "activo", fecha: new Date().toISOString(), telefono: loanTel,
        // snapshot employee data at time of loan
        empTelefono: emp.telefono, empCedula: emp.cedula, empCorreo: emp.correo, empDireccion: emp.direccion
      }]
    });
    setLoanId(id);
    setLoanStep(4);
  };

  // ── JEFE ──
  const loginJefe = () => {
    if (jefePass === empresa?.claveJefe) { setScreen("jefe"); setJefePass(""); setShowPass(false); }
    else flash("Clave incorrecta", "err");
  };

  const clearEmpForm = () => {
    setEditEmpId(null);
    setFNombre(""); setFClave(""); setFTel(""); setFCed(""); setFCorreo(""); setFDir(""); setFSal(""); setFDias("20"); setFFoto("");
  };

  const saveEmp = () => {
    if (!fNombre.trim() || !fClave.trim()) { flash("Nombre y clave son obligatorios", "err"); return; }
    if (editEmpId) {
      // update existing
      const nuevaEmpresa = {
        ...empresa,
        empleados: empresa.empleados.map(e => e.id === editEmpId
          ? { ...e, nombre: fNombre.trim(), clave: fClave.trim(), telefono: fTel, cedula: fCed, correo: fCorreo, direccion: fDir, foto: fFoto || e.foto, perfilCompleto: !!(fCed && fCorreo) }
          : e),
        salarios: { ...empresa.salarios, [editEmpId]: { salarioMensual: parseFloat(fSal) || empresa.salarios[editEmpId]?.salarioMensual || 250, diasLaborales: parseInt(fDias) || empresa.salarios[editEmpId]?.diasLaborales || 20 } }
      };
      save(nuevaEmpresa);
      flash("✅ Empleado actualizado");
    } else {
      // check if name already exists (avoid duplicates)
      const existe = empresa?.empleados?.find(e => e.nombre.toLowerCase() === fNombre.trim().toLowerCase());
      if (existe) { flash("⚠️ Ya existe un empleado con ese nombre. Edita el existente.", "err"); return; }
      const emp = {
        id: Date.now(), nombre: fNombre.trim(), clave: fClave.trim(),
        telefono: fTel, cedula: fCed, correo: fCorreo, direccion: fDir,
        foto: fFoto,
        perfilCompleto: !!(fCed && fCorreo),
        horasTrabajadas: 0, registros: [], fechaCreacion: new Date().toISOString()
      };
      save({
        ...empresa,
        empleados: [...(empresa.empleados || []), emp],
        salarios: { ...empresa.salarios, [emp.id]: { salarioMensual: parseFloat(fSal) || 250, diasLaborales: parseInt(fDias) || 20 } }
      });
      flash("✅ Empleado agregado");
    }
    clearEmpForm();
  };

  const delEmp = (id) => {
    const emp = empresa?.empleados?.find(e => e.id === id);
    setConfirm({
      text: `¿Eliminar a ${emp?.nombre}? Esta acción no se puede deshacer.`,
      onYes: () => {
        save({ ...empresa, empleados: empresa.empleados.filter(e => e.id !== id) });
        setConfirm(null);
        flash("Empleado eliminado");
      }
    });
  };

  // ── NOMINA ACTIVA ──
  const loginNA = () => {
    if (naPass === CLAVE_MAESTRA) { loadEmpresas(); setScreen("nomina"); setNaPass(""); setShowPass(false); }
    else flash("Clave incorrecta", "err");
  };

  const loadEmpresas = async () => {
    const lista = await st.get("empresas_lista") || [];
    const result = await Promise.all(lista.map(async ({ id, nombre }) => ({ id, nombre, datos: await st.get(`empresa:${id}`) })));
    setAllEmpresas(result.filter(e => e.datos));
  };

  const marcarPagado = async (pId, cIdx, datos, eId) => {
    const nuevos = datos.prestamos.map(p => {
      if (p.id !== pId) return p;
      const cuotas = p.cuotas.map((c, i) => i === cIdx ? { ...c, estado: "pagado", pagado: new Date().toLocaleDateString("es-ES") } : c);
      return { ...p, cuotas, estado: cuotas.every(c => c.estado === "pagado") ? "pagado" : "activo" };
    });
    const nueva = { ...datos, prestamos: nuevos };
    await st.set(`empresa:${eId}`, nueva);
    if (eId === empId) setEmpresa(nueva);
    setAllEmpresas(prev => prev.map(e => e.id === eId ? { ...e, datos: nueva } : e));
    if (selEmpresa?.id === eId) setSelEmpresa(prev => ({ ...prev, datos: nueva }));
    flash("✅ Cuota marcada como pagada");
  };

  // ── SETUP ──
  const configurar = async () => {
    if (!sNombre.trim() || !sClave.trim()) { flash("Completa todos los campos", "err"); return; }
    if (sClave !== sConf) { flash("Las claves no coinciden", "err"); return; }
    if (sClave.length < 4) { flash("La clave debe tener mínimo 4 caracteres", "err"); return; }
    const id = `emp_${Date.now()}`;
    const nueva = { nombre: sNombre.trim(), claveJefe: sClave, fechaCreacion: new Date().toISOString(), empleados: [], salarios: {}, prestamos: [] };
    await st.set("tablet_empresa_id", id, false);
    const lista = await st.get("empresas_lista") || [];
    await st.set("empresas_lista", [...lista.filter(e => e.id !== id), { id, nombre: sNombre.trim() }]);
    await st.set(`empresa:${id}`, nueva);
    setEmpId(id); setEmpresa(nueva); setScreen("principal");
  };

  // ── DASHBOARD HELPERS ──
  const getHoraEntrada = () => empresa?.configHorario?.horaEntrada || "08:00";

  const getDatesForPeriod = (periodo) => {
    const today = new Date();
    const dates = [];
    if (periodo === "hoy") {
      dates.push(new Date(today));
    } else if (periodo === "semana") {
      const dow = today.getDay(); const offset = dow === 0 ? -6 : 1 - dow;
      for (let i = 0; i <= 6; i++) { const d = new Date(today); d.setDate(today.getDate() + offset + i); if (d <= today) dates.push(d); }
    } else if (periodo === "quincenal") {
      for (let i = 14; i >= 0; i--) { const d = new Date(today); d.setDate(today.getDate() - i); dates.push(d); }
    } else {
      const first = new Date(today.getFullYear(), today.getMonth(), 1);
      const d = new Date(first);
      while (d <= today) { dates.push(new Date(d)); d.setDate(d.getDate() + 1); }
    }
    return dates;
  };

  const getEmpStatusForDate = (emp, date) => {
    const dateStr = date.toLocaleDateString("es-ES");
    const today = new Date().toLocaleDateString("es-ES");
    const registro = emp.registros?.find(r => r.fecha === dateStr && r.tipo === "in");
    const permiso = empresa?.permisos?.find(p => p.empId === emp.id && p.fecha === dateStr);
    if (permiso) return { status: "permiso", nota: permiso.nota, fecha: dateStr };
    if (!registro) return { status: dateStr === today ? "sin_registrar" : "falta", fecha: dateStr };
    const horaLlegada = new Date(registro.hora);
    const [hh, mm] = getHoraEntrada().split(":").map(Number);
    const horaConfig = new Date(horaLlegada); horaConfig.setHours(hh, mm, 0, 0);
    const tarde = horaLlegada > horaConfig;
    const minutosTarde = tarde ? Math.round((horaLlegada - horaConfig) / 60000) : 0;
    if (!registro.salida) {
      const horasActivas = (new Date() - horaLlegada) / 3600000;
      return { status: horasActivas >= 8 ? "horas_extra" : "en_turno", tarde, minutosTarde, horaLlegada: registro.horaReg, horasActivas: horasActivas.toFixed(1), registroRef: registro, fecha: dateStr };
    }
    const horaSalidaStr = new Date(registro.salida).toLocaleTimeString("es-ES", { hour: "2-digit", minute: "2-digit" });
    const hs = parseFloat(registro.hsTrab || 0);
    return { status: hs >= 8 ? "completoExtra" : "completo", tarde, minutosTarde, horaLlegada: registro.horaReg, horaSalida: horaSalidaStr, horasTrabajadas: hs.toFixed(2), registroRef: registro, fecha: dateStr };
  };

  const statusConfig = {
    en_turno:      { label: "En turno",        color: "#10b981", bg: "#10b98120" },
    completo:      { label: "Completó",         color: "#06b6d4", bg: "#06b6d420" },
    completoExtra: { label: "Completó + Extra", color: "#8b5cf6", bg: "#8b5cf620" },
    horas_extra:   { label: "⚡ Horas Extra",   color: "#f59e0b", bg: "#f59e0b20" },
    falta:         { label: "Falta",            color: "#ef4444", bg: "#ef444420" },
    permiso:       { label: "Permiso",          color: "#94a3b8", bg: "#94a3b820" },
    sin_registrar: { label: "Sin registrar",    color: "#475569", bg: "#47556920" },
  };

  const marcarPermiso = () => {
    if (!permisoModal) return;
    const nueva = { ...empresa, permisos: [...(empresa.permisos || []).filter(p => !(p.empId === permisoModal.empId && p.fecha === permisoModal.fecha)), { id: Date.now(), empId: permisoModal.empId, fecha: permisoModal.fecha, nota: permisoNota }] };
    save(nueva); setPermisoModal(null); setPermisoNota(""); flash("✅ Permiso registrado");
  };

  const quitarPermiso = (empId, fecha) => {
    save({ ...empresa, permisos: (empresa.permisos || []).filter(p => !(p.empId === empId && p.fecha === fecha)) });
    flash("Permiso eliminado");
  };

  const resolverAlerta = (tipo) => {
    if (!alertaModal) return;
    const { empId, registroRef } = alertaModal;
    let horaFin;
    if (tipo === "manual") {
      if (!horaSalidaManual) { flash("Ingresa la hora de salida", "err"); return; }
      const [h, m] = horaSalidaManual.split(":").map(Number);
      horaFin = new Date(registroRef.hora); horaFin.setHours(h, m, 0, 0);
      if (horaFin < new Date(registroRef.hora)) horaFin.setDate(horaFin.getDate() + 1);
    } else {
      horaFin = new Date();
    }
    const hs = (horaFin - new Date(registroRef.hora)) / 3600000;
    const nuevaEmpresa = { ...empresa, empleados: empresa.empleados.map(e => e.id === empId ? { ...e, horasTrabajadas: (e.horasTrabajadas || 0) + hs, registros: e.registros.map(r => r === registroRef ? { ...r, salida: horaFin.toISOString(), hsTrab: hs.toFixed(2), resolucion: tipo } : r) } : e) };
    save(nuevaEmpresa); setAlertaModal(null); setHoraSalidaManual(""); flash(`✅ Resuelto como ${tipo === "extra" ? "horas extra" : "salida manual"}`);
  };

  // ── UPCOMING DUES ──
  const getProximasCuotas = (datos, dias = 7) => {
    const hoy = new Date(); hoy.setHours(0,0,0,0);
    const limite = new Date(); limite.setDate(hoy.getDate() + dias);
    const result = [];
    (datos?.prestamos || []).filter(p => p.estado === "activo").forEach(p => {
      p.cuotas?.filter(c => c.estado === "pendiente").forEach(c => {
        try {
          const fv = new Date(c.vence.split("/").reverse().join("-"));
          if (fv >= hoy && fv <= limite) {
            const emp = datos.empleados?.find(e => e.id === p.empId);
            result.push({ cuota: c, prestamo: p, empleadoNombre: emp?.nombre || "Desconocido", empId: p.empId });
          }
        } catch {}
      });
    });
    return result.sort((a, b) => new Date(a.cuota.vence.split("/").reverse().join("-")) - new Date(b.cuota.vence.split("/").reverse().join("-")));
  };

  const getVencidasSinPagar = (datos) => {
    const hoy = new Date(); hoy.setHours(0,0,0,0);
    const result = [];
    (datos?.prestamos || []).filter(p => p.estado === "activo").forEach(p => {
      p.cuotas?.filter(c => c.estado === "pendiente").forEach(c => {
        try {
          const fv = new Date(c.vence.split("/").reverse().join("-"));
          if (fv < hoy) {
            const emp = datos.empleados?.find(e => e.id === p.empId);
            result.push({ cuota: c, prestamo: p, empleadoNombre: emp?.nombre || "Desconocido" });
          }
        } catch {}
      });
    });
    return result;
  };

  // ── EXPORTS ──
  const exportarAsistencia = (periodo) => {
    const dates = getDatesForPeriod(periodo);
    const rows = [];
    (empresa?.empleados || []).forEach(emp => {
      dates.forEach(d => {
        const st = getEmpStatusForDate(emp, d);
        const cfg = statusConfig[st.status] || statusConfig.sin_registrar;
        rows.push({
          "Empleado": emp.nombre,
          "Fecha": d.toLocaleDateString("es-ES"),
          "Estado": cfg.label,
          "Hora Entrada": st.horaLlegada || "-",
          "Hora Salida": st.horaSalida || "-",
          "Horas Trabajadas": st.horasTrabajadas || "-",
          "Llegó Tarde": st.tarde ? `Sí (+${st.minutosTarde} min)` : "No",
          "Nota": st.nota || "-",
        });
      });
    });
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Asistencia");
    XLSX.writeFile(wb, `asistencia_${empresa?.nombre}_${new Date().toLocaleDateString("es-ES").replace(/\//g,"-")}.xlsx`);
  };

  const exportarReporteNomina = () => {
    // Sheet 1: Resumen por empleado
    const resumen = [];
    allEmpresas.forEach(e => {
      (e.datos?.empleados || []).forEach(emp => {
        const ps = e.datos?.prestamos?.filter(p => p.empId === emp.id) || [];
        const cs = calcCreditScore(emp.id, e.datos);
        const activos = ps.filter(p => p.estado === "activo");
        const porCobrar = activos.reduce((s, p) => s + p.cuotas.filter(c => c.estado === "pendiente").reduce((a, c) => a + c.monto, 0), 0);
        resumen.push({
          "Empresa": e.nombre,
          "Empleado": emp.nombre,
          "Cédula": emp.cedula || "-",
          "Teléfono": emp.telefono || "-",
          "Salario": e.datos?.salarios?.[emp.id]?.salarioMensual || "-",
          "Perfil Crediticio": `${cs.icon} ${cs.label}`,
          "Total Préstamos": ps.length,
          "Préstamos Activos": activos.length,
          "Monto por Cobrar": porCobrar.toFixed(2),
          "Estado": activos.length > 0 ? "Con deuda" : "Solvente",
        });
      });
    });

    // Sheet 2: Detalle préstamos
    const prestamos = [];
    allEmpresas.forEach(e => {
      (e.datos?.prestamos || []).forEach(p => {
        const emp = e.datos?.empleados?.find(em => em.id === p.empId);
        p.cuotas?.forEach(c => {
          prestamos.push({
            "Empresa": e.nombre,
            "Empleado": emp?.nombre || "-",
            "Préstamo ID": p.id,
            "Monto Capital": p.montoSol,
            "Total a Pagar": p.total,
            "Cuota #": `${c.num}/${p.nCuotas}`,
            "Monto Cuota": c.monto,
            "Vencimiento": c.vence,
            "Estado Cuota": c.estado,
            "Fecha Pago": c.pagado || "-",
          });
        });
      });
    });

    // Sheet 3: Próximas a vencer (7 días)
    const proximas = [];
    allEmpresas.forEach(e => {
      getProximasCuotas(e.datos, 7).forEach(item => {
        proximas.push({
          "Empresa": e.nombre,
          "Empleado": item.empleadoNombre,
          "Monto Cuota": item.cuota.monto,
          "Vence": item.cuota.vence,
          "Cuota #": `${item.cuota.num}/${item.prestamo.nCuotas}`,
        });
      });
    });

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(resumen), "Resumen Empleados");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(prestamos), "Detalle Préstamos");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(proximas), "Próximas a Vencer");
    XLSX.writeFile(wb, `reporte_nomina_activa_${new Date().toLocaleDateString("es-ES").replace(/\//g,"-")}.xlsx`);
  };

  // ── CREDIT SCORE ──
  const calcCreditScore = (empId, datos) => {
    const ps = datos?.prestamos?.filter(p => p.empId === empId) || [];
    if (!ps.length) return { label: "Sin historial", color: "#64748b", bg: "#64748b22", icon: "⬜", score: -1, detalle: { total: 0, alDia: 0, tarde: 0, vencidas: 0, pagados: 0 } };
    let alDia = 0, tarde = 0, vencidas = 0, pagados = 0;
    ps.forEach(p => {
      if (p.estado === "pagado") pagados++;
      p.cuotas?.forEach(c => {
        if (c.estado === "pagado") {
          const fp = new Date((c.pagado || "").split("/").reverse().join("-") || Date.now());
          const fv = new Date((c.vence || "").split("/").reverse().join("-") || Date.now());
          if (fp <= fv) alDia++; else tarde++;
        } else {
          const hoy = new Date();
          const fv = new Date((c.vence || "").split("/").reverse().join("-") || Date.now());
          if (hoy > fv) vencidas++;
        }
      });
    });
    const total = alDia + tarde + vencidas;
    if (vencidas > 0) return { label: "No recomendado", color: "#ef4444", bg: "#ef444422", icon: "🔴", score: 0, detalle: { total, alDia, tarde, vencidas, pagados } };
    if (tarde > 0) return { label: "Regular", color: "#f59e0b", bg: "#f59e0b22", icon: "🟡", score: 1, detalle: { total, alDia, tarde, vencidas, pagados } };
    return { label: "Excelente", color: "#10b981", bg: "#10b98122", icon: "⭐", score: 2, detalle: { total, alDia, tarde, vencidas, pagados } };
  };

  // ══════════════════════════════════════════
  // SCREENS
  // ══════════════════════════════════════════

  // LOADING
  if (screen === "loading") return (
    <div style={{ backgroundColor: "#0f172a", minHeight: "100vh" }} className="flex items-center justify-center">
      <div className="text-center"><LogoAP size={80}/><p className="mt-4 text-sm" style={{ color: "#94a3b8" }}>Iniciando...</p></div>
    </div>
  );

  // SETUP
  if (screen === "setup") return (
    <div style={{ backgroundColor: "#0f172a", minHeight: "100vh" }} className="text-white flex items-center justify-center p-4">
      <div className="w-full max-w-md">
        <div className="text-center mb-8"><div className="flex justify-center mb-4"><LogoAP size={90}/></div>
          <h1 className="text-3xl font-bold" style={{ color: "#10b981" }}>Automatiza Pro</h1>
          <p className="text-sm mt-1" style={{ color: "#94a3b8" }}>Configuración inicial</p>
        </div>
        <div className="rounded-xl p-6" style={{ backgroundColor: "#1e293b" }}>
          <FlashMsg msg={msg}/>
          <input type="text" placeholder="Nombre de la empresa *" value={sNombre} onChange={e => setSNombre(e.target.value)} className="w-full px-4 py-3 rounded-lg text-black text-sm mb-3" style={{ backgroundColor: "#e2e8f0" }}/>
          <input type="password" placeholder="Clave del jefe *" value={sClave} onChange={e => setSClave(e.target.value)} className="w-full px-4 py-3 rounded-lg text-black text-sm mb-3" style={{ backgroundColor: "#e2e8f0" }}/>
          <input type="password" placeholder="Confirmar clave *" value={sConf} onChange={e => setSConf(e.target.value)} className="w-full px-4 py-3 rounded-lg text-black text-sm mb-4" style={{ backgroundColor: "#e2e8f0" }}/>
          <button onClick={configurar} className="w-full py-4 rounded-xl font-bold text-white" style={{ backgroundColor: "#10b981" }}>CONFIGURAR SISTEMA</button>
        </div>
      </div>
    </div>
  );

  // PRINCIPAL
  if (screen === "principal") return (
    <div style={{ backgroundColor: "#0f172a", minHeight: "100vh" }} className="text-white flex items-center justify-center p-4">
      <div className="w-full max-w-md">
        <div className="text-center mb-8">
          <div className="flex justify-center mb-3"><LogoAP size={100}/></div>
          <h1 className="text-4xl font-bold" style={{ color: "#10b981" }}>Automatiza Pro</h1>
          <p className="text-sm mt-1" style={{ color: "#94a3b8" }}>{empresa?.nombre}</p>
          <p style={{ color: "#10b981", fontSize: "11px" }} className="mt-1">by Nómina Activa</p>
        </div>
        <div className="space-y-4 mb-5">
          <button onClick={() => setScreen("entrada")} className="w-full py-6 rounded-xl font-bold text-white text-lg" style={{ backgroundColor: "#10b981" }}>
            <LogIn size={28} className="inline mr-3"/>ENTRADA
          </button>
          <button onClick={() => setScreen("salida")} className="w-full py-6 rounded-xl font-bold text-white text-lg" style={{ backgroundColor: "#06b6d4" }}>
            <LogOut size={28} className="inline mr-3"/>SALIDA
          </button>
          <div>
            <button onClick={() => { resetFlow(); setScreen("adelanto"); }} className="w-full py-6 rounded-xl font-bold text-white text-lg" style={{ backgroundColor: "#f59e0b" }}>
              <CreditCard size={28} className="inline mr-3"/>ADELANTO DE NÓMINA
            </button>
            <p className="text-center text-xs mt-2" style={{ color: "#94a3b8" }}>
              Servicio financiero privado · <span style={{ color: "#f59e0b" }}>5% mensual</span> sobre el monto solicitado
            </p>
          </div>
        </div>
        <div className="flex gap-3">
          <button onClick={() => setScreen("jefe_login")} className="flex-1 py-3 rounded-lg font-bold text-white text-sm" style={{ backgroundColor: "#334155" }}>
            <Settings size={14} className="inline mr-2"/>Panel Jefe
          </button>
          <button onClick={() => setScreen("nomina_login")} className="py-3 px-4 rounded-lg" style={{ backgroundColor: "#1e293b" }} title="Nómina Activa">
            <Shield size={16} style={{ color: "#10b981" }}/>
          </button>
        </div>
      </div>
    </div>
  );

  // ── VERIFY BLOCK (inlined, not a sub-component, to avoid focus loss) ──
  const empActual = empleadoId ? empresa?.empleados?.find(e => e.id === empleadoId) : null;
  const existeNombre = nombreInput.length > 1 && empresa?.empleados?.find(e => e.nombre.toLowerCase() === nombreInput.trim().toLowerCase());

  const verifyBlock = (color) => (
    <div className="space-y-3">
      <button onClick={startCam} className="w-full py-4 rounded-lg font-bold text-white flex items-center justify-center gap-2" style={{ backgroundColor: color }}>
        <QrCode size={20}/>ESCANEAR TARJETA QR
      </button>
      <p className="text-center text-xs" style={{ color: "#475569" }}>──── O INGRESAR MANUALMENTE ────</p>
      <input
        type="text"
        placeholder="Tu nombre completo"
        value={nombreInput}
        onChange={e => setNombreInput(e.target.value)}
        className="w-full px-4 py-3 rounded-lg text-black text-sm"
        style={{ backgroundColor: "#e2e8f0" }}
      />
      <input
        type="password"
        placeholder="Tu clave personal"
        value={claveInput}
        onChange={e => setClaveInput(e.target.value)}
        className="w-full px-4 py-3 rounded-lg text-black text-sm"
        style={{ backgroundColor: "#e2e8f0" }}
      />
      {!existeNombre && nombreInput.length > 1 && (
        <input
          type="tel"
          placeholder="Tu teléfono (primer registro)"
          value={telefonoInput}
          onChange={e => setTelefonoInput(e.target.value)}
          className="w-full px-4 py-3 rounded-lg text-black text-sm"
          style={{ backgroundColor: "#e2e8f0" }}
        />
      )}
      <button
        onClick={verificar}
        disabled={!nombreInput.trim() || !claveInput.trim()}
        className="w-full py-3 rounded-lg font-bold text-white disabled:opacity-50"
        style={{ backgroundColor: color }}
      >
        VERIFICAR IDENTIDAD
      </button>
    </div>
  );

  // ENTRADA
  if (screen === "entrada") return (
    <div style={{ backgroundColor: "#0f172a", minHeight: "100vh" }} className="text-white p-4">
      <div className="max-w-md mx-auto mt-6">
        <BackBtn onClick={goHome}/>
        <div className="rounded-xl p-6" style={{ backgroundColor: "#1e293b", borderTop: "4px solid #10b981" }}>
          <h2 className="text-2xl font-bold mb-1" style={{ color: "#10b981" }}>ENTRADA</h2>
          <p className="text-sm mb-5" style={{ color: "#94a3b8" }}>Registra tu hora de ingreso</p>
          <FlashMsg msg={msg}/>
          {empActual ? (
            <>
              <div className="rounded-lg p-4 mb-4" style={{ backgroundColor: "#0f172a" }}>
                {esNuevo && (
                  <div className="rounded-lg p-3 mb-3" style={{ backgroundColor: "#1a2e1a", border: "1px solid #10b98155" }}>
                    <p className="text-sm font-bold" style={{ color: "#10b981" }}>✅ Tarjeta digital creada</p>
                    <p className="text-xs mt-1" style={{ color: "#94a3b8" }}>Dirígete a Recursos Humanos para recibir tu código QR de acceso.</p>
                  </div>
                )}
                <p className="text-xs" style={{ color: "#94a3b8" }}>Empleado verificado</p>
                <p className="text-2xl font-bold">{empActual.nombre}</p>
              </div>
              <button onClick={entrada} className="w-full py-4 rounded-lg font-bold text-white text-lg" style={{ backgroundColor: "#10b981" }}>✅ CONFIRMAR ENTRADA</button>
              <button onClick={() => { setEmpleadoId(null); setEsNuevo(false); setNombreInput(""); setClaveInput(""); setTelefonoInput(""); }} className="w-full py-2 rounded-lg font-bold text-white mt-3 bg-gray-600">Cancelar</button>
            </>
          ) : camActive
            ? <CamView color="#10b981" videoRef={videoRef} canvasRef={canvasRef} onStop={stopCam}/>
            : verifyBlock("#10b981")
          }
        </div>
      </div>
    </div>
  );

  // SALIDA
  if (screen === "salida") return (
    <div style={{ backgroundColor: "#0f172a", minHeight: "100vh" }} className="text-white p-4">
      <div className="max-w-md mx-auto mt-6">
        <BackBtn onClick={goHome}/>
        <div className="rounded-xl p-6" style={{ backgroundColor: "#1e293b", borderTop: "4px solid #06b6d4" }}>
          <h2 className="text-2xl font-bold mb-1" style={{ color: "#06b6d4" }}>SALIDA</h2>
          <p className="text-sm mb-5" style={{ color: "#94a3b8" }}>Registra tu hora de salida</p>
          <FlashMsg msg={msg}/>
          {empActual ? (
            <>
              <div className="rounded-lg p-4 mb-4" style={{ backgroundColor: "#0f172a" }}>
                <p className="text-xs" style={{ color: "#94a3b8" }}>Empleado verificado</p>
                <p className="text-2xl font-bold">{empActual.nombre}</p>
              </div>
              <button onClick={salida} className="w-full py-4 rounded-lg font-bold text-white text-lg" style={{ backgroundColor: "#06b6d4" }}>✅ CONFIRMAR SALIDA</button>
              <button onClick={() => { setEmpleadoId(null); setNombreInput(""); setClaveInput(""); }} className="w-full py-2 rounded-lg font-bold text-white mt-3 bg-gray-600">Cancelar</button>
            </>
          ) : camActive
            ? <CamView color="#06b6d4" videoRef={videoRef} canvasRef={canvasRef} onStop={stopCam}/>
            : verifyBlock("#06b6d4")
          }
        </div>
      </div>
    </div>
  );

  // ADELANTO step 4
  if (screen === "adelanto" && loanStep === 4) return (
    <div style={{ backgroundColor: "#0f172a", minHeight: "100vh" }} className="text-white flex items-center justify-center p-4">
      <div className="max-w-md w-full text-center">
        <div className="text-6xl mb-4">✅</div>
        <h2 className="text-3xl font-bold mb-2" style={{ color: "#10b981" }}>PRE-APROBADO</h2>
        <div className="rounded-xl p-6 mb-4" style={{ backgroundColor: "#1e293b" }}>
          <p className="mb-3">Un asesor te contactará en las próximas</p>
          <p className="text-3xl font-bold mb-3" style={{ color: "#f59e0b" }}>24 horas</p>
          <p className="text-sm mb-2" style={{ color: "#94a3b8" }}>al número:</p>
          <p className="text-2xl font-bold" style={{ color: "#06b6d4" }}>{loanTel}</p>
          <p className="text-xs mt-4" style={{ color: "#475569" }}>¿Número incorrecto? Habla con tu jefe para actualizarlo.</p>
        </div>
        <p className="text-xs mb-5" style={{ color: "#334155" }}>ID: {loanId}</p>
        <button onClick={goHome} className="w-full py-4 rounded-xl font-bold text-white" style={{ backgroundColor: "#10b981" }}>Volver al Inicio</button>
      </div>
    </div>
  );

  // ADELANTO step 3 (terms)
  if (screen === "adelanto" && loanStep === 3 && empActual) {
    const loan = calcLoan(loanMonto, loanPlazo, loanFreq);
    return (
      <div style={{ backgroundColor: "#0f172a", minHeight: "100vh" }} className="text-white p-4">
        <div className="max-w-md mx-auto mt-6">
          <BackBtn onClick={() => setLoanStep(2)} label="Calculadora"/>
          <div className="rounded-xl p-6" style={{ backgroundColor: "#1e293b", borderTop: "4px solid #f59e0b" }}>
            <h2 className="text-xl font-bold mb-5" style={{ color: "#f59e0b" }}>RESUMEN Y TÉRMINOS</h2>
            <FlashMsg msg={msg}/>
            {[["Monto solicitado", `$${loan.monto}`, null], [`Costo (5% × ${loanPlazo} mes${loanPlazo > 1 ? "es" : ""})`, `$${loan.int}`, "#f59e0b"], ["Total a devolver", `$${loan.total}`, "#ef4444"], [`${loan.n} cuota${loan.n > 1 ? "s" : ""} ${loanFreq === "mensual" ? "mensuales" : loanFreq === "quincenal" ? "quincenales" : "semanales"}`, `$${loan.cuota} c/u`, "#10b981"]].map(([l, v, c]) => (
              <div key={l} className="flex justify-between py-2" style={{ borderBottom: "1px solid #334155" }}>
                <span className="text-sm" style={{ color: "#94a3b8" }}>{l}</span>
                <span className="font-bold" style={{ color: c || "white" }}>{v}</span>
              </div>
            ))}
            <div className="rounded-lg p-3 my-4" style={{ backgroundColor: "#0f172a", border: "1px solid #f59e0b44" }}>
              <p className="text-xs" style={{ color: "#94a3b8" }}>⚠️ Servicio financiero privado con costo del <strong style={{ color: "#f59e0b" }}>5% mensual</strong>. El total se descontará de tu nómina. Al confirmar aceptas estos términos.</p>
            </div>
            <input type="tel" placeholder="Confirma tu número de teléfono" value={loanTel} onChange={e => setLoanTel(e.target.value)} className="w-full px-4 py-3 rounded-lg text-black text-sm mb-4" style={{ backgroundColor: "#e2e8f0" }}/>
            <button onClick={confirmarLoan} disabled={!loanTel} className="w-full py-4 rounded-lg font-bold text-white text-lg disabled:opacity-50" style={{ backgroundColor: "#f59e0b" }}>💰 ACEPTO Y CONFIRMO</button>
            <button onClick={() => setLoanStep(2)} className="w-full py-2 rounded-lg font-bold text-white mt-3 bg-gray-600">Atrás</button>
          </div>
        </div>
      </div>
    );
  }

  // ADELANTO step 2
  if (screen === "adelanto" && loanStep === 2 && empActual) {
    const limite = getLimite(empleadoId);
    const loan = loanMonto ? calcLoan(loanMonto, loanPlazo, loanFreq) : null;
    return (
      <div style={{ backgroundColor: "#0f172a", minHeight: "100vh" }} className="text-white p-4">
        <div className="max-w-md mx-auto mt-6">
          <BackBtn onClick={() => setLoanStep(1)} label="Verificación"/>
          <div className="rounded-xl p-6" style={{ backgroundColor: "#1e293b", borderTop: "4px solid #f59e0b" }}>
            <h2 className="text-xl font-bold mb-1" style={{ color: "#f59e0b" }}>CALCULAR ADELANTO</h2>
            <p className="text-sm mb-4" style={{ color: "#10b981" }}>Disponible: <strong>${limite}</strong> · {empActual.nombre}</p>
            <FlashMsg msg={msg}/>
            <p className="text-xs mb-1" style={{ color: "#94a3b8" }}>Monto a solicitar ($)</p>
            <input type="number" placeholder={`Máximo $${limite}`} value={loanMonto} onChange={e => setLoanMonto(e.target.value)} className="w-full px-4 py-3 rounded-lg text-black text-sm mb-4" style={{ backgroundColor: "#e2e8f0" }}/>
            <p className="text-xs mb-2" style={{ color: "#94a3b8" }}>Plazo</p>
            <div className="grid grid-cols-3 gap-2 mb-4">
              {[1, 2, 3].map(p => (
                <button key={p} onClick={() => setLoanPlazo(p)} className="py-3 rounded-lg font-bold" style={{ backgroundColor: loanPlazo === p ? "#f59e0b" : "#334155", color: "white" }}>
                  {p} {p === 1 ? "mes" : "meses"}
                </button>
              ))}
            </div>
            <p className="text-xs mb-2" style={{ color: "#94a3b8" }}>Frecuencia de pago</p>
            <div className="grid grid-cols-3 gap-2 mb-5">
              {["mensual", "quincenal", "semanal"].map(f => (
                <button key={f} onClick={() => setLoanFreq(f)} className="py-2 rounded-lg font-bold text-xs capitalize" style={{ backgroundColor: loanFreq === f ? "#f59e0b" : "#334155", color: "white" }}>{f}</button>
              ))}
            </div>
            {loan && (
              <div className="rounded-lg p-4 mb-4" style={{ backgroundColor: "#0f172a" }}>
                <div className="grid grid-cols-2 gap-3">
                  <div><p className="text-xs" style={{ color: "#64748b" }}>Total a pagar</p><p className="text-2xl font-bold" style={{ color: "#f59e0b" }}>${loan.total}</p></div>
                  <div><p className="text-xs" style={{ color: "#64748b" }}>{loan.n} cuota{loan.n > 1 ? "s" : ""} de</p><p className="text-2xl font-bold" style={{ color: "#10b981" }}>${loan.cuota}</p></div>
                </div>
                <p className="text-xs mt-2" style={{ color: "#475569" }}>Interés: ${loan.int} ({loanPlazo * 5}% total)</p>
              </div>
            )}
            <button onClick={() => {
              if (!loanMonto || parseFloat(loanMonto) <= 0) { flash("Ingresa un monto válido", "err"); return; }
              if (parseFloat(loanMonto) > parseFloat(limite)) { flash(`Máximo disponible: $${limite}`, "err"); return; }
              setLoanStep(3);
            }} disabled={!loanMonto} className="w-full py-4 rounded-lg font-bold text-white disabled:opacity-50" style={{ backgroundColor: "#f59e0b" }}>
              VER RESUMEN →
            </button>
          </div>
        </div>
      </div>
    );
  }

  // ADELANTO step 1
  if (screen === "adelanto") return (
    <div style={{ backgroundColor: "#0f172a", minHeight: "100vh" }} className="text-white p-4">
      <div className="max-w-md mx-auto mt-6">
        <BackBtn onClick={goHome}/>
        <div className="rounded-xl p-6" style={{ backgroundColor: "#1e293b", borderTop: "4px solid #f59e0b" }}>
          <div className="flex justify-between items-start mb-5">
            <div>
              <h2 className="text-xl font-bold" style={{ color: "#f59e0b" }}>ADELANTO DE NÓMINA</h2>
              <p className="text-xs mt-1" style={{ color: "#f59e0b" }}>5% mensual · Máx. 3 meses</p>
            </div>
            <LogoNA h={50}/>
          </div>
          <FlashMsg msg={msg}/>
          {empActual ? (
            <>
              <div className="rounded-lg p-4 mb-4" style={{ backgroundColor: "#0f172a" }}>
                <p className="font-bold text-xl">{empActual.nombre}</p>
                <p className="text-sm mt-1" style={{ color: "#10b981" }}>Límite: <strong>${getLimite(empleadoId)}</strong></p>
              </div>
              {hasLoan(empleadoId) ? (
                <div className="rounded-lg p-4" style={{ backgroundColor: "#1a0f00", border: "1px solid #f59e0b44" }}>
                  <p className="font-bold" style={{ color: "#f59e0b" }}>❌ Préstamo Activo</p>
                  <p className="text-sm mt-1" style={{ color: "#94a3b8" }}>Tienes un préstamo activo. Debes estar 100% solvente para solicitar uno nuevo.</p>
                </div>
              ) : !empresa?.empleados?.find(e => e.id === empleadoId)?.perfilCompleto ? (
                <div className="rounded-lg p-4" style={{ backgroundColor: "#1a1000", border: "1px solid #f59e0b44" }}>
                  <p className="font-bold" style={{ color: "#f59e0b" }}>⚠️ Perfil Incompleto</p>
                  <p className="text-sm mt-1" style={{ color: "#94a3b8" }}>Tu perfil no está completo. Dirígete a Recursos Humanos para completar tu registro.</p>
                </div>
              ) : (
                <button onClick={() => setLoanStep(2)} className="w-full py-4 rounded-lg font-bold text-white text-lg" style={{ backgroundColor: "#f59e0b" }}>
                  💰 SOLICITAR ADELANTO
                </button>
              )}
              <button onClick={() => { setEmpleadoId(null); setNombreInput(""); setClaveInput(""); }} className="w-full py-2 rounded-lg font-bold text-white mt-3 bg-gray-600">Cancelar</button>
            </>
          ) : camActive
            ? <CamView color="#f59e0b" videoRef={videoRef} canvasRef={canvasRef} onStop={stopCam}/>
            : verifyBlock("#f59e0b")
          }
        </div>
      </div>
    </div>
  );

  // JEFE LOGIN
  if (screen === "jefe_login") return (
    <div style={{ backgroundColor: "#0f172a", minHeight: "100vh" }} className="text-white flex items-center justify-center p-4">
      <div className="w-full max-w-md">
        <BackBtn onClick={goHome}/>
        <div className="rounded-xl p-8" style={{ backgroundColor: "#1e293b" }}>
          <div className="text-center mb-6"><Shield size={44} style={{ color: "#10b981" }} className="mx-auto mb-3"/>
            <h2 className="text-2xl font-bold" style={{ color: "#10b981" }}>Panel Jefe</h2>
            <p className="text-sm" style={{ color: "#94a3b8" }}>{empresa?.nombre}</p>
          </div>
          <FlashMsg msg={msg}/>
          <div className="relative mb-4">
            <input type={showPass ? "text" : "password"} placeholder="Clave de acceso" value={jefePass}
              onChange={e => setJefePass(e.target.value)} onKeyDown={e => e.key === "Enter" && loginJefe()}
              className="w-full px-4 py-3 rounded-lg text-black pr-12" style={{ backgroundColor: "#e2e8f0" }}/>
            <button onClick={() => setShowPass(!showPass)} className="absolute right-3 top-1/2 -translate-y-1/2 text-gray-500">{showPass ? <EyeOff size={18}/> : <Eye size={18}/>}</button>
          </div>
          <button onClick={loginJefe} className="w-full py-4 rounded-lg font-bold text-white" style={{ backgroundColor: "#10b981" }}>INGRESAR</button>
        </div>
      </div>
    </div>
  );

  // JEFE PANEL
  if (screen === "jefe") {

    // QR
    if (qrEmp) return (
      <CardView
        emp={qrEmp}
        empresaNombre={empresa?.nombre}
        empresaLogo={empresa?.logo}
        onClose={() => setQrEmp(null)}
        onPhotoUpload={(e) => {
          const file = e.target.files[0];
          if (!file) return;
          const reader = new FileReader();
          reader.onload = (ev) => {
            const nuevaEmpresa = { ...empresa, empleados: empresa.empleados.map(em => em.id === qrEmp.id ? { ...em, foto: ev.target.result } : em) };
            save(nuevaEmpresa);
            setQrEmp({ ...qrEmp, foto: ev.target.result });
          };
          reader.readAsDataURL(file);
        }}
      />
    );

    // Dashboard
    if (jefeSub === "dashboard") {
      const dates = getDatesForPeriod(dashPeriodo);
      const empsList = empresa?.empleados || [];
      // Resumen del período
      let totalPresente = 0, totalFalta = 0, totalTarde = 0, totalPermiso = 0, totalExtra = 0, alertasSinSalida = [];
      empsList.forEach(emp => {
        dates.forEach(d => {
          const st = getEmpStatusForDate(emp, d);
          if (["en_turno","completo","completoExtra","horas_extra"].includes(st.status)) totalPresente++;
          if (st.status === "falta") totalFalta++;
          if (st.status === "permiso") totalPermiso++;
          if (st.tarde) totalTarde++;
          if (["completoExtra","horas_extra"].includes(st.status)) totalExtra++;
          if (st.status === "horas_extra" && st.registroRef) alertasSinSalida.push({ emp, st });
        });
      });

      return (
        <div style={{ backgroundColor: "#0f172a", minHeight: "100vh" }} className="text-white p-4">
          {/* Permiso Modal */}
          {permisoModal && (
            <div className="fixed inset-0 z-50 flex items-center justify-center p-4" style={{ backgroundColor: "rgba(0,0,0,0.8)" }}>
              <div className="rounded-xl p-6 w-full max-w-xs" style={{ backgroundColor: "#1e293b" }}>
                <h3 className="font-bold mb-1" style={{ color: "#94a3b8" }}>Marcar Permiso</h3>
                <p className="text-sm mb-3" style={{ color: "#64748b" }}>{permisoModal.nombre} · {permisoModal.fecha}</p>
                <input type="text" placeholder="Motivo (opcional)" value={permisoNota} onChange={e => setPermisoNota(e.target.value)} className="w-full px-3 py-2 rounded-lg text-black text-sm mb-4" style={{ backgroundColor: "#e2e8f0" }}/>
                <div className="flex gap-2">
                  <button onClick={() => { setPermisoModal(null); setPermisoNota(""); }} className="flex-1 py-2 rounded-lg font-bold text-white bg-gray-600 text-sm">Cancelar</button>
                  <button onClick={marcarPermiso} className="flex-1 py-2 rounded-lg font-bold text-white text-sm" style={{ backgroundColor: "#94a3b8" }}>Marcar</button>
                </div>
              </div>
            </div>
          )}
          {/* Alerta Modal */}
          {alertaModal && (
            <div className="fixed inset-0 z-50 flex items-center justify-center p-4" style={{ backgroundColor: "rgba(0,0,0,0.8)" }}>
              <div className="rounded-xl p-6 w-full max-w-xs" style={{ backgroundColor: "#1e293b" }}>
                <p className="text-sm font-bold mb-1" style={{ color: "#f59e0b" }}>⚠️ Salida no marcada</p>
                <p className="text-sm mb-4" style={{ color: "#94a3b8" }}>{alertaModal.nombre} · Entrada: {alertaModal.registroRef?.horaReg}</p>
                <p className="text-xs mb-3" style={{ color: "#64748b" }}>¿Cómo registrar la salida?</p>
                <button onClick={() => resolverAlerta("extra")} className="w-full py-3 rounded-lg font-bold text-white mb-2 text-sm" style={{ backgroundColor: "#8b5cf6" }}>⚡ Marcar como Horas Extra</button>
                <div className="flex gap-2 mb-2">
                  <input type="time" value={horaSalidaManual} onChange={e => setHoraSalidaManual(e.target.value)} className="flex-1 px-3 py-2 rounded-lg text-black text-sm" style={{ backgroundColor: "#e2e8f0" }}/>
                  <button onClick={() => resolverAlerta("manual")} className="px-3 py-2 rounded-lg font-bold text-white text-sm" style={{ backgroundColor: "#06b6d4" }}>Guardar</button>
                </div>
                <button onClick={() => { setAlertaModal(null); setHoraSalidaManual(""); }} className="w-full py-2 rounded-lg font-bold text-white bg-gray-600 text-sm">Cancelar</button>
              </div>
            </div>
          )}

          <div className="max-w-2xl mx-auto mt-4">
            <BackBtn onClick={() => setJefeSub("menu")} label="Panel Jefe"/>

            {/* Header */}
            <div className="flex justify-between items-start mb-4">
              <div>
                <h2 className="text-xl font-bold" style={{ color: "#06b6d4" }}>Dashboard · {empresa?.nombre}</h2>
                <p className="text-xs mt-1" style={{ color: "#64748b" }}>Hora entrada: <strong style={{ color: "white" }}>{getHoraEntrada()}</strong>
                  <button onClick={() => setHoraEntradaEdit(!horaEntradaEdit)} className="ml-2 text-xs underline" style={{ color: "#06b6d4" }}>cambiar</button>
                </p>
                {horaEntradaEdit && (
                  <div className="flex items-center gap-2 mt-2">
                    <input type="time" defaultValue={getHoraEntrada()} id="horaEntradaInput" className="px-2 py-1 rounded text-black text-xs" style={{ backgroundColor: "#e2e8f0" }}/>
                    <button onClick={() => {
                      const val = document.getElementById("horaEntradaInput").value;
                      if (val) { save({ ...empresa, configHorario: { ...empresa.configHorario, horaEntrada: val } }); setHoraEntradaEdit(false); flash("✅ Hora actualizada"); }
                    }} className="px-3 py-1 rounded text-white text-xs font-bold" style={{ backgroundColor: "#10b981" }}>OK</button>
                    <button onClick={() => setHoraEntradaEdit(false)} className="px-2 py-1 rounded text-white text-xs bg-gray-600">✕</button>
                  </div>
                )}
              </div>
            </div>

            {/* Próximos descuentos en nómina */}
            {(() => {
              const proximas = getProximasCuotas(empresa, 7);
              const vencidas = getVencidasSinPagar(empresa);
              if (!proximas.length && !vencidas.length) return null;
              return (
                <div className="rounded-xl p-4 mb-4" style={{ backgroundColor: "#0f1a0f", border: "1px solid #10b98144" }}>
                  <p className="text-sm font-bold mb-2 flex items-center gap-2" style={{ color: "#10b981" }}>
                    <Bell size={14}/>Descuentos en Nómina
                  </p>
                  {vencidas.length > 0 && <p className="text-xs mb-1" style={{ color: "#ef4444" }}>⚠️ {vencidas.length} cuota{vencidas.length > 1 ? "s" : ""} vencida{vencidas.length > 1 ? "s" : ""} sin cobrar — aplicar descuento urgente</p>}
                  {proximas.map((item, i) => (
                    <div key={i} className="flex justify-between items-center py-1">
                      <span className="text-xs" style={{ color: "#94a3b8" }}>{item.empleadoNombre} · Cuota {item.cuota.num}/{item.prestamo.nCuotas}</span>
                      <span className="text-xs font-bold" style={{ color: "#10b981" }}>${item.cuota.monto} · {item.cuota.vence}</span>
                    </div>
                  ))}
                </div>
              );
            })()}

            {/* Alertas destacadas */}
            {alertasSinSalida.length > 0 && (
              <div className="rounded-xl p-4 mb-4" style={{ backgroundColor: "#1a0f00", border: "1px solid #f59e0b55" }}>
                <p className="text-sm font-bold mb-2" style={{ color: "#f59e0b" }}>⚠️ Alertas — Sin salida registrada</p>
                {alertasSinSalida.map(({ emp, st }) => (
                  <div key={emp.id} className="flex justify-between items-center py-2" style={{ borderBottom: "1px solid #334155" }}>
                    <div>
                      <p className="text-sm font-bold">{emp.nombre}</p>
                      <p className="text-xs" style={{ color: "#94a3b8" }}>Entrada: {st.horaLlegada} · {st.horasActivas}h activo</p>
                    </div>
                    <button onClick={() => setAlertaModal({ empId: emp.id, nombre: emp.nombre, registroRef: st.registroRef })} className="px-3 py-1 rounded text-xs font-bold text-white" style={{ backgroundColor: "#f59e0b" }}>Resolver</button>
                  </div>
                ))}
              </div>
            )}

            {/* Período selector */}
            <div className="flex gap-2 mb-4">
              {[["hoy","Hoy"],["semana","Semana"],["quincenal","15 días"],["mes","Mes"]].map(([val, label]) => (
                <button key={val} onClick={() => setDashPeriodo(val)} className="flex-1 py-2 rounded-lg text-xs font-bold transition-all" style={{ backgroundColor: dashPeriodo === val ? "#06b6d4" : "#1e293b", color: dashPeriodo === val ? "white" : "#64748b" }}>{label}</button>
              ))}
            </div>

            {/* Resumen del período */}
            <div className="grid grid-cols-5 gap-2 mb-5">
              {[["Presentes", totalPresente, "#10b981"], ["Faltas", totalFalta, "#ef4444"], ["Tardanzas", totalTarde, "#f59e0b"], ["Permisos", totalPermiso, "#94a3b8"], ["H. Extra", totalExtra, "#8b5cf6"]].map(([l, v, c]) => (
                <div key={l} className="rounded-lg p-2 text-center" style={{ backgroundColor: "#1e293b" }}>
                  <p className="text-xs mb-1" style={{ color: "#475569" }}>{l}</p>
                  <p className="text-xl font-bold" style={{ color: c }}>{v}</p>
                </div>
              ))}
            </div>

            {/* Empleados */}
            <div className="space-y-4">
              {empsList.map(emp => {
                const s = getSal(emp.id); const tarifa = s ? s.salarioMensual / s.diasLaborales / 8 : 0;
                const ps = empresa.prestamos?.filter(p => p.empId === emp.id && p.estado === "activo") || [];
                const statuses = dates.map(d => getEmpStatusForDate(emp, d));
                const horasPeriodo = statuses.reduce((sum, st) => sum + (st.horasTrabajadas ? parseFloat(st.horasTrabajadas) : 0), 0);
                const faltasPeriodo = statuses.filter(st => st.status === "falta").length;
                const tardanzasPeriodo = statuses.filter(st => st.tarde).length;
                const permisosPeriodo = statuses.filter(st => st.status === "permiso").length;

                return (
                  <div key={emp.id} className="rounded-xl p-4" style={{ backgroundColor: "#1e293b" }}>
                    {/* Employee header */}
                    <div className="flex justify-between items-start mb-3">
                      <div>
                        <p className="font-bold text-lg">{emp.nombre}</p>
                        <p className="text-xs" style={{ color: "#64748b" }}>{s ? `$${s.salarioMensual}/mes` : "Sin salario"}{emp.telefono ? ` · ${emp.telefono}` : ""}</p>
                      </div>
                      <div className="flex flex-wrap gap-1 justify-end max-w-40">
                        {ps.length > 0 && <span className="px-2 py-0.5 rounded text-xs bg-yellow-500/20 text-yellow-400">Préstamo</span>}
                        {faltasPeriodo > 0 && <span className="px-2 py-0.5 rounded text-xs" style={{ backgroundColor: "#ef444420", color: "#ef4444" }}>{faltasPeriodo} falta{faltasPeriodo > 1 ? "s" : ""}</span>}
                        {tardanzasPeriodo > 0 && <span className="px-2 py-0.5 rounded text-xs" style={{ backgroundColor: "#f59e0b20", color: "#f59e0b" }}>{tardanzasPeriodo} tarde</span>}
                      </div>
                    </div>

                    {/* Period summary mini */}
                    <div className="grid grid-cols-3 gap-2 mb-3">
                      <div className="rounded p-2" style={{ backgroundColor: "#0f172a" }}>
                        <p className="text-xs" style={{ color: "#475569" }}>Horas período</p>
                        <p className="font-bold" style={{ color: "#06b6d4" }}>{horasPeriodo.toFixed(1)}h</p>
                        <p className="text-xs" style={{ color: "#10b981" }}>${(horasPeriodo * tarifa).toFixed(2)}</p>
                      </div>
                      <div className="rounded p-2" style={{ backgroundColor: "#0f172a" }}>
                        <p className="text-xs" style={{ color: "#475569" }}>Faltas / Permisos</p>
                        <p className="font-bold" style={{ color: faltasPeriodo > 0 ? "#ef4444" : "#10b981" }}>{faltasPeriodo} / {permisosPeriodo}</p>
                      </div>
                      <div className="rounded p-2" style={{ backgroundColor: "#0f172a" }}>
                        <p className="text-xs" style={{ color: "#475569" }}>Total horas</p>
                        <p className="font-bold" style={{ color: "#94a3b8" }}>{(emp.horasTrabajadas || 0).toFixed(1)}h</p>
                      </div>
                    </div>

                    {/* Day by day */}
                    <div className="space-y-1">
                      {statuses.map((st, i) => {
                        const cfg = statusConfig[st.status] || statusConfig.sin_registrar;
                        const dateLabel = dates[i].toLocaleDateString("es-ES", dashPeriodo === "hoy" ? { weekday: "long" } : { weekday: "short", day: "numeric" });
                        return (
                          <div key={i} className="flex items-center justify-between p-2 rounded-lg" style={{ backgroundColor: "#0f172a" }}>
                            <div className="flex items-center gap-2 flex-1">
                              <span className="text-xs w-20 flex-shrink-0 capitalize" style={{ color: "#475569" }}>{dateLabel}</span>
                              <span className="text-xs px-2 py-0.5 rounded font-bold" style={{ backgroundColor: cfg.bg, color: cfg.color }}>{cfg.label}</span>
                              {st.tarde && <span className="text-xs" style={{ color: "#f59e0b" }}>+{st.minutosTarde}min</span>}
                              {st.status === "permiso" && st.nota && <span className="text-xs" style={{ color: "#64748b" }}>· {st.nota}</span>}
                            </div>
                            <div className="flex items-center gap-3">
                              {st.horaLlegada && <span className="text-xs" style={{ color: "#10b981" }}>▶ {st.horaLlegada}</span>}
                              {st.horaSalida && <span className="text-xs" style={{ color: "#06b6d4" }}>◀ {st.horaSalida}</span>}
                              {st.horasActivas && <span className="text-xs" style={{ color: "#f59e0b" }}>{st.horasActivas}h</span>}
                              {/* Actions */}
                              {st.status === "falta" && (
                                <button onClick={() => setPermisoModal({ empId: emp.id, fecha: st.fecha, nombre: emp.nombre })} className="text-xs px-2 py-0.5 rounded" style={{ backgroundColor: "#94a3b820", color: "#94a3b8" }}>Permiso</button>
                              )}
                              {st.status === "permiso" && (
                                <button onClick={() => quitarPermiso(emp.id, st.fecha)} className="text-xs px-2 py-0.5 rounded" style={{ backgroundColor: "#ef444420", color: "#ef4444" }}>✕</button>
                              )}
                              {st.status === "horas_extra" && st.registroRef && (
                                <button onClick={() => setAlertaModal({ empId: emp.id, nombre: emp.nombre, registroRef: st.registroRef })} className="text-xs px-2 py-0.5 rounded font-bold" style={{ backgroundColor: "#f59e0b20", color: "#f59e0b" }}>Resolver</button>
                              )}
                            </div>
                          </div>
                        );
                      })}
                    </div>

                    {/* Descuentos nómina */}
                    {ps.length > 0 && (
                      <div className="mt-3 pt-2" style={{ borderTop: "1px solid #334155" }}>
                        <p className="text-xs font-bold mb-1" style={{ color: "#f59e0b" }}>💰 Descuentos en nómina:</p>
                        {ps.map(p => { const c = p.cuotas.find(c => c.estado === "pendiente"); return c ? <p key={p.id} className="text-xs" style={{ color: "#94a3b8" }}>Cuota {c.num}/{p.nCuotas}: <strong style={{ color: "white" }}>${c.monto}</strong> · vence {c.vence}</p> : null; })}
                      </div>
                    )}
                  </div>
                );
              })}
              {!empsList.length && <p className="text-center py-6" style={{ color: "#475569" }}>No hay empleados registrados</p>}
            </div>
          </div>
        </div>
      );
    }

    // Gestionar Empleados
    if (jefeSub === "empleados") return (
      <div style={{ backgroundColor: "#0f172a", minHeight: "100vh" }} className="text-white p-4">
        {confirm && <ConfirmModal text={confirm.text} onYes={confirm.onYes} onNo={() => setConfirm(null)}/>}
        <div className="max-w-md mx-auto mt-4">
          <BackBtn onClick={() => { setJefeSub("menu"); clearEmpForm(); }} label="Panel Jefe"/>
          <h2 className="text-xl font-bold mb-4" style={{ color: "#10b981" }}>Gestionar Empleados</h2>
          <FlashMsg msg={msg}/>
          <div className="rounded-xl p-4 mb-4" style={{ backgroundColor: "#1e293b" }}>
            <h3 className="font-bold mb-3 text-sm" style={{ color: editEmpId ? "#f59e0b" : "#10b981" }}>{editEmpId ? "✏️ Editar Empleado" : "➕ Nuevo Empleado"}</h3>
            <div className="grid grid-cols-2 gap-2">
              <input type="text" placeholder="Nombre *" value={fNombre} onChange={e => setFNombre(e.target.value)} className="px-3 py-2 rounded-lg text-black text-xs" style={{ backgroundColor: "#e2e8f0" }}/>
              <input type="password" placeholder="Clave *" value={fClave} onChange={e => setFClave(e.target.value)} className="px-3 py-2 rounded-lg text-black text-xs" style={{ backgroundColor: "#e2e8f0" }}/>
              <input type="tel" placeholder="Teléfono" value={fTel} onChange={e => setFTel(e.target.value)} className="px-3 py-2 rounded-lg text-black text-xs" style={{ backgroundColor: "#e2e8f0" }}/>
              <input type="text" placeholder="Cédula" value={fCed} onChange={e => setFCed(e.target.value)} className="px-3 py-2 rounded-lg text-black text-xs" style={{ backgroundColor: "#e2e8f0" }}/>
              <input type="email" placeholder="Correo" value={fCorreo} onChange={e => setFCorreo(e.target.value)} className="px-3 py-2 rounded-lg text-black text-xs" style={{ backgroundColor: "#e2e8f0" }}/>
              <input type="text" placeholder="Dirección" value={fDir} onChange={e => setFDir(e.target.value)} className="px-3 py-2 rounded-lg text-black text-xs" style={{ backgroundColor: "#e2e8f0" }}/>
              <input type="number" placeholder="Salario $/mes" value={fSal} onChange={e => setFSal(e.target.value)} className="px-3 py-2 rounded-lg text-black text-xs" style={{ backgroundColor: "#e2e8f0" }}/>
              <input type="number" placeholder="Días laborales" value={fDias} onChange={e => setFDias(e.target.value)} className="px-3 py-2 rounded-lg text-black text-xs" style={{ backgroundColor: "#e2e8f0" }}/>
            </div>
            <div className="flex gap-2 mt-3">
              <button onClick={saveEmp} className="flex-1 py-2 rounded-lg font-bold text-white text-sm" style={{ backgroundColor: editEmpId ? "#f59e0b" : "#10b981" }}>{editEmpId ? "ACTUALIZAR" : "AGREGAR"}</button>
              {editEmpId && <button onClick={clearEmpForm} className="py-2 px-3 rounded-lg text-white bg-gray-600"><X size={14}/></button>}
            </div>
          </div>
          <div className="space-y-2">
            {empresa?.empleados?.map(emp => {
              const s = getSal(emp.id);
              return (
                <div key={emp.id} className="rounded-lg p-4" style={{ backgroundColor: "#1e293b" }}>
                  <div className="flex justify-between items-start">
                    <div className="flex-1 mr-2">
                      <p className="font-bold">{emp.nombre}</p>
                      <p className="text-xs" style={{ color: "#64748b" }}>{emp.telefono || "Sin tel"} · {emp.cedula || "Sin cédula"}</p>
                      <p className="text-xs mt-0.5" style={{ color: emp.perfilCompleto ? "#10b981" : "#f59e0b" }}>
                        {s ? `$${s.salarioMensual}/mes` : "Sin salario"} · {emp.perfilCompleto ? "✅ Perfil completo" : "⚠️ Incompleto"}
                      </p>
                    </div>
                    <div className="flex gap-1">
                      <button onClick={() => genQR(emp)} className="p-2 rounded-lg" style={{ backgroundColor: "#334155" }}><QrCode size={13} style={{ color: "#10b981" }}/></button>
                      <button onClick={() => {
                        setEditEmpId(emp.id);
                        setFNombre(emp.nombre); setFClave(emp.clave); setFTel(emp.telefono || "");
                        setFCed(emp.cedula || ""); setFCorreo(emp.correo || ""); setFDir(emp.direccion || "");
                        setFSal(s?.salarioMensual || ""); setFDias(s?.diasLaborales || "20"); setFFoto(emp.foto || "");
                      }} className="p-2 rounded-lg" style={{ backgroundColor: "#334155" }}><Edit2 size={13} style={{ color: "#06b6d4" }}/></button>
                      <button onClick={() => delEmp(emp.id)} className="p-2 rounded-lg" style={{ backgroundColor: "#334155" }}><Trash2 size={13} style={{ color: "#ef4444" }}/></button>
                    </div>
                  </div>
                </div>
              );
            })}
            {!empresa?.empleados?.length && <p className="text-center py-6 text-sm" style={{ color: "#475569" }}>No hay empleados. Agrega el primero arriba.</p>}
          </div>
        </div>
      </div>
    );

    // Cambiar Clave
    if (jefeSub === "clave") return (
      <div style={{ backgroundColor: "#0f172a", minHeight: "100vh" }} className="text-white p-4">
        <div className="max-w-md mx-auto mt-6">
          <BackBtn onClick={() => { setJefeSub("menu"); setClaveActual(""); setClaveNueva(""); setClaveConf(""); }} label="Panel Jefe"/>
          <div className="rounded-xl p-6" style={{ backgroundColor: "#1e293b" }}>
            <div className="flex items-center gap-3 mb-6">
              <Key size={28} style={{ color: "#10b981" }}/>
              <h2 className="text-xl font-bold" style={{ color: "#10b981" }}>Cambiar Clave</h2>
            </div>
            <FlashMsg msg={msg}/>
            {[["Clave actual", claveActual, setClaveActual], ["Nueva clave", claveNueva, setClaveNueva], ["Confirmar nueva clave", claveConf, setClaveConf]].map(([ph, val, set]) => (
              <input key={ph} type="password" placeholder={ph} value={val} onChange={e => set(e.target.value)}
                className="w-full px-4 py-3 rounded-lg text-black mb-3 text-sm" style={{ backgroundColor: "#e2e8f0" }}/>
            ))}
            <button onClick={() => {
              if (!claveActual || !claveNueva || !claveConf) { flash("Completa todos los campos", "err"); return; }
              if (claveActual !== empresa?.claveJefe) { flash("La clave actual es incorrecta", "err"); return; }
              if (claveNueva !== claveConf) { flash("Las claves nuevas no coinciden", "err"); return; }
              if (claveNueva.length < 4) { flash("La clave debe tener mínimo 4 caracteres", "err"); return; }
              save({ ...empresa, claveJefe: claveNueva });
              setClaveActual(""); setClaveNueva(""); setClaveConf("");
              flash("✅ Clave actualizada correctamente");
              setTimeout(() => setJefeSub("menu"), 1500);
            }} className="w-full py-4 rounded-xl font-bold text-white" style={{ backgroundColor: "#10b981" }}>
              ACTUALIZAR CLAVE
            </button>
          </div>
        </div>
      </div>
    );

    // Jefe menu
    return (
      <div style={{ backgroundColor: "#0f172a", minHeight: "100vh" }} className="text-white p-4">
        <div className="max-w-md mx-auto mt-6">
          <button onClick={() => setScreen("principal")} className="mb-6 flex items-center gap-2 px-3 py-2 rounded-lg text-sm" style={{ backgroundColor: "#334155" }}><Home size={15}/>Inicio</button>

          {/* Upcoming dues alert */}
          {(() => {
            const proximas = getProximasCuotas(empresa, 7);
            const vencidas = getVencidasSinPagar(empresa);
            if (!proximas.length && !vencidas.length) return null;
            return (
              <div className="rounded-xl p-4 mb-4" style={{ backgroundColor: "#1a0f00", border: "1px solid #f59e0b44" }}>
                <p className="font-bold text-sm mb-2 flex items-center gap-2" style={{ color: "#f59e0b" }}>
                  <Bell size={15}/>Alertas de Pagos
                </p>
                {vencidas.length > 0 && (
                  <p className="text-xs mb-1" style={{ color: "#ef4444" }}>⚠️ {vencidas.length} cuota{vencidas.length > 1 ? "s" : ""} vencida{vencidas.length > 1 ? "s" : ""} sin cobrar</p>
                )}
                {proximas.length > 0 && (
                  <p className="text-xs" style={{ color: "#f59e0b" }}>🔔 {proximas.length} cuota{proximas.length > 1 ? "s" : ""} vence{proximas.length === 1 ? "" : "n"} en los próximos 7 días</p>
                )}
                {proximas.slice(0, 3).map((item, i) => (
                  <div key={i} className="flex justify-between mt-1">
                    <span className="text-xs" style={{ color: "#94a3b8" }}>{item.empleadoNombre}</span>
                    <span className="text-xs font-bold" style={{ color: "#f59e0b" }}>${item.cuota.monto} · {item.cuota.vence}</span>
                  </div>
                ))}
              </div>
            );
          })()}

          <div className="rounded-xl p-6" style={{ backgroundColor: "#1e293b" }}>
            <div className="flex items-center gap-3 mb-6">
              <Shield size={32} style={{ color: "#10b981" }}/>
              <div className="flex-1">
                <h2 className="text-xl font-bold" style={{ color: "#10b981" }}>Panel Jefe</h2>
                <p className="text-sm" style={{ color: "#94a3b8" }}>{empresa?.nombre}</p>
              </div>
              <label className="cursor-pointer" title="Subir logo de empresa">
                <div className="rounded-full overflow-hidden flex items-center justify-center" style={{ width: 52, height: 52, backgroundColor: "#334155", border: empresa?.logo ? "2px solid #10b981" : "2px dashed #475569" }}>
                  {empresa?.logo
                    ? <img src={empresa.logo} alt="logo" className="w-full h-full object-cover"/>
                    : <span className="text-xs text-center leading-tight" style={{ color: "#64748b" }}>+Logo</span>
                  }
                </div>
                <input type="file" accept="image/*" className="hidden" onChange={(e) => {
                  const file = e.target.files[0]; if (!file) return;
                  const reader = new FileReader();
                  reader.onload = (ev) => save({ ...empresa, logo: ev.target.result });
                  reader.readAsDataURL(file);
                }}/>
              </label>
            </div>
            <div className="space-y-2">
              {[
                ["Gestionar Empleados", Users, "#10b981", () => setJefeSub("empleados")],
                ["Ver Dashboard", BarChart3, "#06b6d4", () => setJefeSub("dashboard")],
                ["Cambiar Clave", Key, "#8b5cf6", () => setJefeSub("clave")],
              ].map(([l, Icon, c, fn]) => (
                <button key={l} onClick={fn} className="w-full py-4 rounded-xl text-white flex items-center gap-3 px-5 font-bold" style={{ backgroundColor: "#0f172a" }}>
                  <Icon size={20} style={{ color: c }}/>{l}<ChevronRight size={16} className="ml-auto" style={{ color: "#475569" }}/>
                </button>
              ))}
              {/* Export attendance */}
              <button onClick={() => exportarAsistencia(dashPeriodo)} className="w-full py-4 rounded-xl text-white flex items-center gap-3 px-5 font-bold" style={{ backgroundColor: "#0f172a" }}>
                <FileSpreadsheet size={20} style={{ color: "#10b981" }}/>Exportar Asistencia Excel<ChevronRight size={16} className="ml-auto" style={{ color: "#475569" }}/>
              </button>
            </div>
          </div>
        </div>
      </div>
    );
  }

  // NOMINA LOGIN
  if (screen === "nomina_login") return (
    <div style={{ backgroundColor: "#0f172a", minHeight: "100vh" }} className="text-white flex items-center justify-center p-4">
      <div className="w-full max-w-md">
        <BackBtn onClick={goHome}/>
        <div className="rounded-xl p-8" style={{ backgroundColor: "#1e293b" }}>
          <div className="text-center mb-6"><LogoNA h={60}/><p className="text-sm mt-3" style={{ color: "#94a3b8" }}>Panel Administrativo · Acceso Restringido</p></div>
          <FlashMsg msg={msg}/>
          <div className="relative mb-4">
            <input type={showPass ? "text" : "password"} placeholder="Clave maestra" value={naPass}
              onChange={e => setNaPass(e.target.value)} onKeyDown={e => e.key === "Enter" && loginNA()}
              className="w-full px-4 py-3 rounded-lg text-black pr-12" style={{ backgroundColor: "#e2e8f0" }}/>
            <button onClick={() => setShowPass(!showPass)} className="absolute right-3 top-1/2 -translate-y-1/2 text-gray-500">{showPass ? <EyeOff size={18}/> : <Eye size={18}/>}</button>
          </div>
          <button onClick={loginNA} className="w-full py-4 rounded-lg font-bold text-white" style={{ backgroundColor: "#10b981" }}>INGRESAR</button>
        </div>
      </div>
    </div>
  );

  // NOMINA PANEL
  if (screen === "nomina") {

    // Employee detail
    // ── Tarjeta NA (must be before empleado check) ──
    if (naView === "tarjeta_na" && selEmpId2 && selEmpresa) {
      const empData = selEmpresa.datos?.empleados?.find(e => e.id === selEmpId2);
      return <CardView emp={empData} empresaNombre={selEmpresa.nombre} empresaLogo={selEmpresa.datos?.logo} onClose={() => setNaView("empleado")}/>;
    }

    if (naView === "empleado" && selEmpId2 && selEmpresa) {
      const empData = selEmpresa.datos?.empleados?.find(e => e.id === selEmpId2);
      const ps = selEmpresa.datos?.prestamos?.filter(p => p.empId === selEmpId2) || [];
      const sData = selEmpresa.datos?.salarios?.[selEmpId2];
      const totalPrestado = ps.filter(p => p.estado === "activo").reduce((s, p) => s + p.montoSol, 0);
      const porCobrar = ps.filter(p => p.estado === "activo").reduce((s, p) => s + p.cuotas.filter(c => c.estado === "pendiente").reduce((a, c) => a + c.monto, 0), 0);
      return (
        <div style={{ backgroundColor: "#0f172a", minHeight: "100vh" }} className="text-white p-4">
          <div className="max-w-md mx-auto mt-4">
            <div className="flex justify-between items-center mb-0">
              <BackBtn onClick={() => { setNaView("empresa"); setSelEmpId2(null); }} label={selEmpresa.nombre}/>
              <button onClick={() => setNaView("tarjeta_na")} className="mb-5 flex items-center gap-2 px-3 py-2 rounded-lg text-sm font-bold" style={{ backgroundColor: "#10b981", color: "white" }}>
                🪪 Ver Tarjeta
              </button>
            </div>
            <FlashMsg msg={msg}/>

            {/* Employee profile card */}
            <div className="rounded-xl p-5 mb-4" style={{ backgroundColor: "#1e293b" }}>
              <h2 className="text-xl font-bold mb-3" style={{ color: "#10b981" }}>{empData?.nombre}</h2>
              <div className="grid grid-cols-1 gap-2 mb-3">
                {[
                  [Phone, "Teléfono", empData?.telefono || "No registrado"],
                  [Mail, "Correo", empData?.correo || "No registrado"],
                  ["CI", "Cédula", empData?.cedula || "No registrada"],
                  [MapPin, "Dirección", empData?.direccion || "No registrada"],
                ].map(([Icon, label, val], i) => (
                  <div key={i} className="flex items-start gap-2 py-1" style={{ borderBottom: "1px solid #334155" }}>
                    <div className="flex items-center gap-1 w-24 flex-shrink-0">
                      {typeof Icon === "string" ? <span className="text-xs font-bold" style={{ color: "#475569" }}>{Icon}</span> : <Icon size={13} style={{ color: "#475569" }}/>}
                      <span className="text-xs" style={{ color: "#475569" }}>{label}</span>
                    </div>
                    <span className="text-sm" style={{ color: val.includes("No") ? "#475569" : "white" }}>{val}</span>
                  </div>
                ))}
                {sData && (
                  <div className="flex items-center gap-2 py-1">
                    <span className="text-xs w-24 flex-shrink-0" style={{ color: "#475569" }}>💰 Salario</span>
                    <span className="text-sm font-bold" style={{ color: "#10b981" }}>${sData.salarioMensual}/mes · {sData.diasLaborales} días</span>
                  </div>
                )}
                <div className="flex items-center gap-2 py-1">
                  <span className="text-xs w-24 flex-shrink-0" style={{ color: "#475569" }}>📋 Perfil</span>
                  <span className="text-sm" style={{ color: empData?.perfilCompleto ? "#10b981" : "#f59e0b" }}>{empData?.perfilCompleto ? "✅ Completo" : "⚠️ Incompleto"}</span>
                </div>
                <div className="flex items-center gap-2 py-1">
                  <span className="text-xs w-24 flex-shrink-0" style={{ color: "#475569" }}>📅 Registro</span>
                  <span className="text-sm" style={{ color: "#94a3b8" }}>{empData?.fechaCreacion ? new Date(empData.fechaCreacion).toLocaleDateString("es-ES") : "N/A"}</span>
                </div>
              </div>

              {/* Financial summary */}
              <div className="grid grid-cols-2 gap-2 mt-3">
                <div className="rounded-lg p-3" style={{ backgroundColor: "#0f172a" }}>
                  <p className="text-xs" style={{ color: "#64748b" }}>Total prestado</p>
                  <p className="text-xl font-bold" style={{ color: "#f59e0b" }}>${totalPrestado.toFixed(2)}</p>
                </div>
                <div className="rounded-lg p-3" style={{ backgroundColor: "#0f172a" }}>
                  <p className="text-xs" style={{ color: "#64748b" }}>Por cobrar</p>
                  <p className="text-xl font-bold" style={{ color: porCobrar > 0 ? "#ef4444" : "#10b981" }}>${porCobrar.toFixed(2)}</p>
                </div>
              </div>
            </div>

            {/* Loans */}
            {/* Credit Score Card */}
            {(() => {
              const cs = calcCreditScore(selEmpId2, selEmpresa.datos);
              return (
                <div className="rounded-xl p-4 mb-4" style={{ backgroundColor: "#1e293b", border: `1px solid ${cs.color}44` }}>
                  <div className="flex items-center justify-between mb-3">
                    <p className="font-bold text-sm" style={{ color: "#94a3b8" }}>Perfil Crediticio</p>
                    <span className="text-lg px-3 py-1 rounded-full font-bold" style={{ backgroundColor: cs.bg, color: cs.color }}>{cs.icon} {cs.label}</span>
                  </div>
                  {cs.score >= 0 ? (
                    <div className="grid grid-cols-4 gap-2">
                      {[["Al día", cs.detalle.alDia, "#10b981"], ["Con retraso", cs.detalle.tarde, "#f59e0b"], ["Vencidas", cs.detalle.vencidas, "#ef4444"], ["Préstamos", cs.detalle.pagados + (selEmpresa.datos?.prestamos?.filter(p => p.empId === selEmpId2 && p.estado === "activo").length || 0), "#8b5cf6"]].map(([l, v, c]) => (
                        <div key={l} className="text-center rounded-lg p-2" style={{ backgroundColor: "#0f172a" }}>
                          <p className="text-lg font-bold" style={{ color: c }}>{v}</p>
                          <p className="text-xs" style={{ color: "#475569" }}>{l}</p>
                        </div>
                      ))}
                    </div>
                  ) : (
                    <p className="text-sm" style={{ color: "#475569" }}>Este empleado no tiene historial de préstamos aún.</p>
                  )}
                </div>
              );
            })()}

            <h3 className="font-bold mb-3" style={{ color: "#f59e0b" }}>Historial de Préstamos ({ps.length})</h3>
            {!ps.length
              ? <p className="text-center py-6 text-sm" style={{ color: "#475569" }}>Sin préstamos registrados</p>
              : (
                <div className="space-y-4">
                  {ps.map(p => (
                    <div key={p.id} className="rounded-xl p-4" style={{ backgroundColor: "#1e293b" }}>
                      {/* Loan header */}
                      <div className="flex justify-between items-start mb-3">
                        <div>
                          <span className="text-xs px-2 py-1 rounded font-bold" style={{ backgroundColor: p.estado === "pagado" ? "#10b98122" : "#f59e0b22", color: p.estado === "pagado" ? "#10b981" : "#f59e0b" }}>
                            {p.estado === "pagado" ? "✅ Pagado" : "🟡 Activo"}
                          </span>
                          <p className="text-xs mt-1" style={{ color: "#475569" }}>{p.id}</p>
                          <p className="text-xs" style={{ color: "#475569" }}>Solicitado: {new Date(p.fecha).toLocaleDateString("es-ES")}</p>
                        </div>
                        <div className="text-right">
                          <p className="font-bold" style={{ color: "#f59e0b" }}>${p.total}</p>
                          <p className="text-xs" style={{ color: "#64748b" }}>Capital: ${p.montoSol}</p>
                          <p className="text-xs" style={{ color: "#64748b" }}>Interés: ${p.interes}</p>
                        </div>
                      </div>
                      {/* Loan details */}
                      <div className="grid grid-cols-3 gap-2 mb-3 p-2 rounded" style={{ backgroundColor: "#0f172a" }}>
                        <div><p className="text-xs" style={{ color: "#475569" }}>Plazo</p><p className="text-xs font-bold">{p.plazo} mes{p.plazo > 1 ? "es" : ""}</p></div>
                        <div><p className="text-xs" style={{ color: "#475569" }}>Frecuencia</p><p className="text-xs font-bold capitalize">{p.frecuencia}</p></div>
                        <div><p className="text-xs" style={{ color: "#475569" }}>Cuota</p><p className="text-xs font-bold" style={{ color: "#10b981" }}>${p.cuota}</p></div>
                      </div>
                      {/* Tel at time of loan */}
                      {p.telefono && <p className="text-xs mb-3" style={{ color: "#64748b" }}>📱 Tel al solicitar: <span style={{ color: "#94a3b8" }}>{p.telefono}</span></p>}
                      {/* Cuotas */}
                      <div className="space-y-2">
                        {p.cuotas.map((c, idx) => (
                          <div key={idx} className="flex justify-between items-center p-2 rounded-lg" style={{ backgroundColor: "#0f172a" }}>
                            <div>
                              <p className="text-xs font-bold" style={{ color: c.estado === "pagado" ? "#10b981" : "#94a3b8" }}>Cuota {c.num}/{p.nCuotas}</p>
                              <p className="text-xs" style={{ color: "#475569" }}>Vence: {c.vence}</p>
                              {c.pagado && <p className="text-xs" style={{ color: "#10b981" }}>Cobrado: {c.pagado}</p>}
                            </div>
                            <div className="flex items-center gap-2">
                              <span className="font-bold text-sm">${c.monto}</span>
                              {c.estado !== "pagado"
                                ? <button onClick={() => marcarPagado(p.id, idx, selEmpresa.datos, selEmpresa.id)} className="px-2 py-1 rounded text-xs font-bold text-white" style={{ backgroundColor: "#10b981" }}>✓ Cobrado</button>
                                : <Check size={14} style={{ color: "#10b981" }}/>
                              }
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>
                  ))}
                </div>
              )
            }
          </div>
        </div>
      );
    }

    // Company detail
    if (naView === "empresa" && selEmpresa) {
      const d = selEmpresa.datos;
      const porCobrar = d?.prestamos?.filter(p => p.estado === "activo").reduce((s, p) => s + p.cuotas.filter(c => c.estado === "pendiente").reduce((a, c) => a + c.monto, 0), 0) || 0;
      return (
        <div style={{ backgroundColor: "#0f172a", minHeight: "100vh" }} className="text-white p-4">
          <div className="max-w-md mx-auto mt-4">
            <BackBtn onClick={() => setNaView("empresas")} label="Todas las Empresas"/>
            <FlashMsg msg={msg}/>
            <h2 className="text-2xl font-bold mb-1" style={{ color: "#10b981" }}>{selEmpresa.nombre}</h2>
            <p className="text-xs mb-4" style={{ color: "#64748b" }}>Desde: {d?.fechaCreacion ? new Date(d.fechaCreacion).toLocaleDateString("es-ES") : "N/A"}</p>
            <div className="grid grid-cols-2 gap-3 mb-5">
              {[
                ["Empleados", d?.empleados?.length || 0, "#10b981"],
                ["Préstamos activos", d?.prestamos?.filter(p => p.estado === "activo").length || 0, "#f59e0b"],
                ["Total prestado", `$${(d?.prestamos?.filter(p => p.estado === "activo").reduce((s, p) => s + p.montoSol, 0) || 0).toFixed(0)}`, "#06b6d4"],
                ["Por cobrar", `$${porCobrar.toFixed(0)}`, "#ef4444"]
              ].map(([l, v, c]) => (
                <div key={l} className="rounded-lg p-3" style={{ backgroundColor: "#1e293b" }}>
                  <p className="text-xs" style={{ color: "#64748b" }}>{l}</p>
                  <p className="text-2xl font-bold" style={{ color: c }}>{v}</p>
                </div>
              ))}
            </div>
            <h3 className="font-bold mb-3" style={{ color: "#94a3b8" }}>Empleados</h3>
            <div className="space-y-2">
              {d?.empleados?.map(emp => {
                const pa = d.prestamos?.filter(p => p.empId === emp.id && p.estado === "activo") || [];
                const pc = pa.reduce((s, p) => s + p.cuotas.filter(c => c.estado === "pendiente").reduce((a, c) => a + c.monto, 0), 0);
                const cs = calcCreditScore(emp.id, d);
                return (
                  <button key={emp.id} onClick={() => { setSelEmpId2(emp.id); setNaView("empleado"); }} className="w-full rounded-lg p-4 text-left" style={{ backgroundColor: "#1e293b" }}>
                    <div className="flex justify-between items-center">
                      <div>
                        <div className="flex items-center gap-2 mb-0.5">
                          <p className="font-bold">{emp.nombre}</p>
                          <span className="text-xs px-2 py-0.5 rounded-full font-bold" style={{ backgroundColor: cs.bg, color: cs.color }}>{cs.icon} {cs.label}</span>
                        </div>
                        <p className="text-xs" style={{ color: "#64748b" }}>{emp.telefono || "Sin tel"} · {emp.cedula || "Sin cédula"}</p>
                      </div>
                      <div className="text-right">
                        {pa.length > 0 ? (<><p className="text-sm font-bold" style={{ color: "#f59e0b" }}>Deuda activa</p><p className="text-xs" style={{ color: "#ef4444" }}>${pc.toFixed(2)}</p></>) : <p className="text-sm" style={{ color: "#10b981" }}>Solvente ✅</p>}
                        <ChevronRight size={14} className="mt-1" style={{ color: "#475569" }}/>
                      </div>
                    </div>
                  </button>
                );
              })}
              {!d?.empleados?.length && <p className="text-center py-4 text-sm" style={{ color: "#475569" }}>Sin empleados</p>}
            </div>
          </div>
        </div>
      );
    }

    // All companies
    const gTotal = allEmpresas.reduce((s, e) => s + (e.datos?.prestamos?.reduce((a, p) => a + p.montoSol, 0) || 0), 0);
    const gActivos = allEmpresas.reduce((s, e) => s + (e.datos?.prestamos?.filter(p => p.estado === "activo").length || 0), 0);
    const gPagados = allEmpresas.reduce((s, e) => s + (e.datos?.prestamos?.filter(p => p.estado === "pagado").length || 0), 0);
    const gCobrar = allEmpresas.reduce((s, e) => s + (e.datos?.prestamos?.filter(p => p.estado === "activo").reduce((a, p) => a + p.cuotas.filter(c => c.estado === "pendiente").reduce((b, c) => b + c.monto, 0), 0) || 0), 0);
    const gVencidas = allEmpresas.reduce((s, e) => s + (e.datos?.prestamos?.reduce((a, p) => a + p.cuotas.filter(c => c.estado === "pendiente" && new Date() > new Date((c.vence || "").split("/").reverse().join("-"))).length, 0) || 0), 0);
    const gEmpleados = allEmpresas.reduce((s, e) => s + (e.datos?.empleados?.length || 0), 0);
    const totalLoans = gActivos + gPagados;

    // Chart helpers
    const pct = (v, t) => t > 0 ? Math.round((v / t) * 100) : 0;
    const activosPct = pct(gActivos, totalLoans);
    const pagadosPct = pct(gPagados, totalLoans);
    const cobradoPct = gTotal > 0 ? Math.round(((gTotal - gCobrar) / gTotal) * 100) : 0;

    return (
      <div style={{ backgroundColor: "#0f172a", minHeight: "100vh" }} className="text-white p-4">
        <div className="max-w-lg mx-auto mt-4">
          <div className="flex justify-between items-center mb-5">
            <LogoNA h={50}/>
            <div className="flex gap-2">
              <button onClick={loadEmpresas} className="p-2 rounded-lg" style={{ backgroundColor: "#1e293b" }}><RefreshCw size={15} style={{ color: "#10b981" }}/></button>
              <button onClick={() => setScreen("principal")} className="p-2 rounded-lg" style={{ backgroundColor: "#1e293b" }}><Home size={15} style={{ color: "#94a3b8" }}/></button>
            </div>
          </div>
          <FlashMsg msg={msg}/>

          {/* ── KPI Cards ── */}
          <div className="grid grid-cols-2 gap-3 mb-5">
            {[
              ["Empresas activas", allEmpresas.length, "#10b981"],
              ["Empleados totales", gEmpleados, "#06b6d4"],
              ["Total desembolsado", `$${gTotal.toFixed(0)}`, "#f59e0b"],
              ["Por cobrar", `$${gCobrar.toFixed(0)}`, "#ef4444"],
              ["Préstamos activos", gActivos, "#8b5cf6"],
              ["Cuotas vencidas", gVencidas, gVencidas > 0 ? "#ef4444" : "#10b981"],
            ].map(([l, v, c]) => (
              <div key={l} className="rounded-xl p-3" style={{ backgroundColor: "#1e293b" }}>
                <p className="text-xs mb-1" style={{ color: "#64748b" }}>{l}</p>
                <p className="text-2xl font-bold" style={{ color: c }}>{v}</p>
              </div>
            ))}
          </div>

          {/* ── CHARTS ── */}
          {totalLoans > 0 && (
            <div className="rounded-xl p-4 mb-5" style={{ backgroundColor: "#1e293b" }}>
              <p className="font-bold mb-4" style={{ color: "#06b6d4" }}>📊 Resumen del Negocio</p>

              {/* Loan status donut */}
              <p className="text-xs mb-2" style={{ color: "#64748b" }}>Estado de Préstamos</p>
              <div className="flex items-center gap-4 mb-4">
                <svg width="90" height="90" viewBox="0 0 90 90">
                  <circle cx="45" cy="45" r="35" fill="none" stroke="#334155" strokeWidth="14"/>
                  <circle cx="45" cy="45" r="35" fill="none" stroke="#10b981" strokeWidth="14"
                    strokeDasharray={`${pagadosPct * 2.199} ${219.9 - pagadosPct * 2.199}`}
                    strokeDashoffset="54.975" strokeLinecap="round"/>
                  <circle cx="45" cy="45" r="35" fill="none" stroke="#8b5cf6" strokeWidth="14"
                    strokeDasharray={`${activosPct * 2.199} ${219.9 - activosPct * 2.199}`}
                    strokeDashoffset={`${54.975 - pagadosPct * 2.199}`} strokeLinecap="round"/>
                  <text x="45" y="49" textAnchor="middle" fill="white" fontSize="13" fontWeight="bold">{totalLoans}</text>
                </svg>
                <div className="space-y-2">
                  {[["Pagados", gPagados, "#10b981"], ["Activos", gActivos, "#8b5cf6"], ["Vencidas", gVencidas, "#ef4444"]].map(([l, v, c]) => (
                    <div key={l} className="flex items-center gap-2">
                      <div className="rounded-full w-3 h-3 flex-shrink-0" style={{ backgroundColor: c }}/>
                      <span className="text-xs" style={{ color: "#94a3b8" }}>{l}:</span>
                      <span className="text-xs font-bold" style={{ color: c }}>{v}</span>
                    </div>
                  ))}
                </div>
              </div>

              {/* Recovery bar */}
              <p className="text-xs mb-1" style={{ color: "#64748b" }}>Recuperación de Capital</p>
              <div className="rounded-full overflow-hidden mb-1" style={{ backgroundColor: "#334155", height: 20 }}>
                <div className="h-full flex items-center justify-center text-xs font-bold text-white rounded-full transition-all"
                  style={{ width: `${cobradoPct}%`, minWidth: cobradoPct > 0 ? 40 : 0, backgroundColor: cobradoPct >= 80 ? "#10b981" : cobradoPct >= 50 ? "#f59e0b" : "#ef4444" }}>
                  {cobradoPct > 5 ? `${cobradoPct}%` : ""}
                </div>
              </div>
              <div className="flex justify-between text-xs" style={{ color: "#475569" }}>
                <span>Cobrado: ${(gTotal - gCobrar).toFixed(0)}</span>
                <span>Pendiente: ${gCobrar.toFixed(0)}</span>
              </div>

              {/* Per company bars */}
              {allEmpresas.length > 1 && (
                <>
                  <p className="text-xs mt-4 mb-2" style={{ color: "#64748b" }}>Cartera por Empresa</p>
                  {allEmpresas.map(e => {
                    const tot = e.datos?.prestamos?.reduce((s, p) => s + p.montoSol, 0) || 0;
                    const pct2 = gTotal > 0 ? Math.round((tot / gTotal) * 100) : 0;
                    return (
                      <div key={e.id} className="mb-2">
                        <div className="flex justify-between text-xs mb-0.5">
                          <span style={{ color: "#94a3b8" }}>{e.nombre}</span>
                          <span style={{ color: "#f59e0b" }}>${tot.toFixed(0)}</span>
                        </div>
                        <div className="rounded-full overflow-hidden" style={{ backgroundColor: "#334155", height: 8 }}>
                          <div className="h-full rounded-full" style={{ width: `${pct2}%`, backgroundColor: "#f59e0b" }}/>
                        </div>
                      </div>
                    );
                  })}
                </>
              )}
            </div>
          )}

          {/* ── Companies list ── */}
          {/* Upcoming dues across all companies */}
          {(() => {
            const todasProximas = allEmpresas.flatMap(e => getProximasCuotas(e.datos, 7).map(i => ({ ...i, empresa: e.nombre })));
            const todasVencidas = allEmpresas.flatMap(e => getVencidasSinPagar(e.datos).map(i => ({ ...i, empresa: e.nombre })));
            if (!todasProximas.length && !todasVencidas.length) return null;
            return (
              <div className="rounded-xl p-4 mb-4" style={{ backgroundColor: "#1a0f00", border: "1px solid #f59e0b44" }}>
                <p className="font-bold text-sm mb-2 flex items-center gap-2" style={{ color: "#f59e0b" }}>
                  <Bell size={15}/>Alertas Globales
                </p>
                {todasVencidas.length > 0 && (
                  <div className="mb-2">
                    <p className="text-xs font-bold mb-1" style={{ color: "#ef4444" }}>⚠️ {todasVencidas.length} cuota{todasVencidas.length > 1 ? "s" : ""} vencida{todasVencidas.length > 1 ? "s" : ""} sin cobrar</p>
                    {todasVencidas.slice(0, 3).map((item, i) => (
                      <div key={i} className="flex justify-between">
                        <span className="text-xs" style={{ color: "#94a3b8" }}>{item.empresa} · {item.empleadoNombre}</span>
                        <span className="text-xs font-bold" style={{ color: "#ef4444" }}>${item.cuota.monto}</span>
                      </div>
                    ))}
                  </div>
                )}
                {todasProximas.length > 0 && (
                  <div>
                    <p className="text-xs font-bold mb-1" style={{ color: "#f59e0b" }}>🔔 Vencen en 7 días</p>
                    {todasProximas.slice(0, 4).map((item, i) => (
                      <div key={i} className="flex justify-between">
                        <span className="text-xs" style={{ color: "#94a3b8" }}>{item.empresa} · {item.empleadoNombre}</span>
                        <span className="text-xs font-bold" style={{ color: "#f59e0b" }}>${item.cuota.monto} · {item.cuota.vence}</span>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            );
          })()}

          {/* Export button */}
          {allEmpresas.length > 0 && (
            <button onClick={exportarReporteNomina} className="w-full py-3 rounded-xl font-bold text-white flex items-center justify-center gap-2 mb-4" style={{ backgroundColor: "#1e293b", border: "1px solid #10b98133" }}>
              <FileSpreadsheet size={18} style={{ color: "#10b981" }}/>
              Exportar Reporte Completo Excel
            </button>
          )}

          <h3 className="font-bold mb-3" style={{ color: "#94a3b8" }}>Empresas</h3>
          {!allEmpresas.length ? (
            <div className="text-center py-10" style={{ color: "#334155" }}>
              <Building2 size={44} className="mx-auto mb-3 opacity-30"/>
              <p className="text-sm">No hay empresas registradas aún.</p>
            </div>
          ) : (
            <div className="space-y-3">
              {allEmpresas.map(e => {
                const act = e.datos?.prestamos?.filter(p => p.estado === "activo").length || 0;
                const cob = e.datos?.prestamos?.filter(p => p.estado === "activo").reduce((s, p) => s + p.cuotas.filter(c => c.estado === "pendiente").reduce((a, c) => a + c.monto, 0), 0) || 0;
                const venc = e.datos?.prestamos?.reduce((s, p) => s + p.cuotas.filter(c => c.estado === "pendiente" && new Date() > new Date((c.vence || "").split("/").reverse().join("-"))).length, 0) || 0;
                const emps = e.datos?.empleados || [];
                const scores = emps.map(emp => calcCreditScore(emp.id, e.datos).score);
                const noRec = scores.filter(s => s === 0).length;
                return (
                  <button key={e.id} onClick={() => { setSelEmpresa(e); setNaView("empresa"); }} className="w-full rounded-xl p-4 text-left" style={{ backgroundColor: "#1e293b" }}>
                    <div className="flex justify-between items-start">
                      <div className="flex-1">
                        <div className="flex items-center gap-2 mb-1">
                          {e.datos?.logo && <img src={e.datos.logo} alt="" className="rounded-full object-cover" style={{ width: 24, height: 24 }}/>}
                          <p className="font-bold">{e.nombre}</p>
                        </div>
                        <p className="text-xs" style={{ color: "#64748b" }}>{emps.length} empleados · {act} préstamos activos</p>
                        {venc > 0 && <p className="text-xs mt-0.5" style={{ color: "#ef4444" }}>⚠️ {venc} cuota{venc > 1 ? "s" : ""} vencida{venc > 1 ? "s" : ""}</p>}
                        {noRec > 0 && <p className="text-xs mt-0.5" style={{ color: "#f59e0b" }}>🔴 {noRec} no recomendado{noRec > 1 ? "s" : ""}</p>}
                      </div>
                      <div className="text-right flex flex-col items-end gap-1">
                        {cob > 0 && <p className="text-sm font-bold" style={{ color: "#ef4444" }}>${cob.toFixed(0)}</p>}
                        <ChevronRight size={15} style={{ color: "#475569" }}/>
                      </div>
                    </div>
                  </button>
                );
              })}
            </div>
          )}
        </div>
      </div>
    );
  }

  return null;
}


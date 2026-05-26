import { useState, useEffect, useCallback, useMemo, useRef } from "react";
import { LayoutDashboard, CalendarDays, Users, Building2, TrendingUp, FileText, Settings, Plus, Pencil, Trash2, Printer, Eye, X, Check, AlertCircle, Search, Download, CreditCard, Plane, Hotel, Bus, Map, Shield, Ship, Package, FileCheck, ArrowRight, Upload, BarChart2, BookOpen } from "lucide-react";

// ══════════════════════════════════════════════════
//  UTILS
// ══════════════════════════════════════════════════
const genId = () => `${Date.now()}_${Math.random().toString(36).substr(2,5)}`;
const fmt$ = (n, currency) => {
  const sym = currency === "USD" ? "US$" : currency === "EUR" ? "€" : "$";
  return `${sym}${Number(n||0).toLocaleString("es-AR",{minimumFractionDigits:2,maximumFractionDigits:2})}`;
};
const fmtDate = d => d ? new Date(d+"T12:00:00").toLocaleDateString("es-AR") : "—";
const today = () => new Date().toISOString().split("T")[0];
const sum = arr => arr.reduce((a,b)=>a+(Number(b)||0),0);
const paidSum = arr => sum((arr||[]).map(p=>p.amount||0));

const STATUS = {
  cotizacion: {label:"Cotización",   c:"#6366F1",bg:"#EEF2FF"},
  confirmada: {label:"Confirmada",    c:"#0EA5E9",bg:"#F0F9FF"},
  en_viaje:   {label:"En viaje",      c:"#10B981",bg:"#ECFDF5"},
  completada: {label:"Completada",    c:"#8B5CF6",bg:"#F5F3FF"},
  cancelada:  {label:"Cancelada",     c:"#EF4444",bg:"#FEF2F2"},
};
const SVC_TYPES = [
  {v:"vuelo",l:"Vuelo",Icon:Plane},
  {v:"hotel",l:"Hotel",Icon:Hotel},
  {v:"traslado",l:"Traslado",Icon:Bus},
  {v:"excursion",l:"Excursión",Icon:Map},
  {v:"seguro",l:"Seguro",Icon:Shield},
  {v:"crucero",l:"Crucero",Icon:Ship},
  {v:"paquete",l:"Paquete",Icon:Package},
  {v:"otro",l:"Otro",Icon:FileText},
];
const PAY_METHODS = ["Efectivo","Transferencia","Tarjeta crédito","Tarjeta débito","Cheque","Otro"];
const PROV_CATS = ["Aerolínea","Hotel","Agencia receptiva","Seguro","Crucero","Transfer","Tour operador","Otro"];
const CURRENCIES = ["ARS","USD","EUR"];

// ══════════════════════════════════════════════════
//  STORAGE (persistent via window.storage API)
// ══════════════════════════════════════════════════
const SK = "travelmanager_v4";
const INIT = {
  settings:{
    agencyName:"Mi Agencia de Viajes",
    address:"",phone:"",email:"",cuit:"",
    logo:null,defaultCommission:15,currency:"ARS",
    voucherColor:"#1a56db",receiptColor:"#1a56db",
    footerText:"Gracias por elegirnos.",
    headerNote:""
  },
  clients:[],providers:[],reservations:[],nextFile:1
};

async function loadDB(){
  try{
    const r = await window.storage.get(SK);
    return r ? JSON.parse(r.value) : INIT;
  } catch { return INIT; }
}
async function saveDB(d){
  try{ await window.storage.set(SK, JSON.stringify(d)); } catch{}
}

// ══════════════════════════════════════════════════
//  EXCEL EXPORT
// ══════════════════════════════════════════════════
function buildCSV(rows, headers){
  const esc = v => `"${String(v??'').replace(/"/g,'""')}"`;
  return [headers.map(esc).join(','), ...rows.map(r=>r.map(esc).join(','))].join('\n');
}
function downloadCSV(csv, filename){
  const blob = new Blob(['\uFEFF'+csv], {type:'text/csv;charset=utf-8;'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a'); a.href=url; a.download=filename; a.click();
  URL.revokeObjectURL(url);
}

function exportClientStatement(client, reservations, settings){
  const clientRes = reservations.filter(r =>
    r.passengers.some(p => p.clientId === client.id) || r.clientId === client.id
  );
  const headers = ['File','Destino','Fecha Salida','Estado','Precio Venta','Total Cobrado','Saldo Pendiente','Forma de Pago'];
  const rows = clientRes.map(r => {
    const cobrado = paidSum(r.paymentsReceived);
    const saldo = (r.salePrice||0) - cobrado;
    const metodos = [...new Set((r.paymentsReceived||[]).map(p=>p.method))].join(', ');
    return [r.fileNumber, r.destination||'', fmtDate(r.departureDate), STATUS[r.status]?.label||r.status,
      r.salePrice||0, cobrado, saldo, metodos];
  });
  const totalVenta = sum(clientRes.map(r=>r.salePrice||0));
  const totalCobrado = sum(clientRes.map(r=>paidSum(r.paymentsReceived)));
  rows.push(['','','','TOTAL',totalVenta,totalCobrado,totalVenta-totalCobrado,'']);
  downloadCSV(buildCSV(rows,headers), `estado_cuenta_${client.name.replace(/\s/g,'_')}.csv`);
}

function exportProviderStatement(provider, reservations, settings){
  const headers = ['File','Destino','Servicio','Descripción','Costo Neto','Pagado al Proveedor','Pendiente al Proveedor'];
  const rows = [];
  reservations.forEach(r => {
    r.services.filter(s=>s.providerId===provider.id).forEach(s=>{
      const pagado = sum((s.paymentsDue||[]).filter(p=>p.paid).map(p=>p.amount||0));
      const pendiente = (s.costPrice||0) - pagado;
      rows.push([r.fileNumber, r.destination||'',
        SVC_TYPES.find(t=>t.v===s.type)?.l||s.type,
        s.description||'', s.costPrice||0, pagado, pendiente]);
    });
  });
  const totalCosto = sum(rows.map(r=>Number(r[4])));
  const totalPagado = sum(rows.map(r=>Number(r[5])));
  rows.push(['','','','TOTAL',totalCosto,totalPagado,totalCosto-totalPagado]);
  downloadCSV(buildCSV(rows,headers), `estado_cuenta_${provider.name.replace(/\s/g,'_')}.csv`);
}

function exportReservations(reservations, settings){
  const headers = ['File','Estado','Destino','Fecha Salida','Fecha Regreso','Pasajero Principal','Precio Venta','Total Cobrado','Saldo','Comisión'];
  const rows = reservations.map(r => {
    const cobrado = paidSum(r.paymentsReceived);
    const costo = sum(r.services.map(s=>s.costPrice||0));
    return [r.fileNumber, STATUS[r.status]?.label||r.status,
      r.destination||'', fmtDate(r.departureDate), fmtDate(r.returnDate),
      r.passengers[0]?.name||'', r.salePrice||0, cobrado, (r.salePrice||0)-cobrado, (r.salePrice||0)-costo];
  });
  downloadCSV(buildCSV(rows,headers), 'reporte_reservas.csv');
}

function exportCommissions(reservations, settings){
  const headers = ['File','Destino','Pasajero','Fecha Salida','Precio Venta','Costo Total','Comisión Neta','% Comisión'];
  const rows = reservations.filter(r=>r.status!=='cancelada').map(r => {
    const costo = sum(r.services.map(s=>s.costPrice||0));
    const comision = (r.salePrice||0) - costo;
    const pct = r.salePrice ? ((comision/r.salePrice)*100).toFixed(1) : 0;
    return [r.fileNumber, r.destination||'', r.passengers[0]?.name||'',
      fmtDate(r.departureDate), r.salePrice||0, costo, comision, pct+'%'];
  });
  downloadCSV(buildCSV(rows,headers), 'reporte_comisiones.csv');
}

// ══════════════════════════════════════════════════
//  PRINT HELPERS
// ══════════════════════════════════════════════════
function getLogoHtml(settings){
  if(!settings.logo) return `<div style="font-size:22px;font-weight:700;color:${settings.voucherColor||'#1a56db'}">${settings.agencyName}</div>`;
  return `<img src="${settings.logo}" style="max-height:60px;max-width:180px;object-fit:contain;" alt="Logo"/>`;
}

function printReceipt(res, settings, providers){
  const received = paidSum(res.paymentsReceived);
  const pending = (res.salePrice||0)-received;
  const color = settings.receiptColor || '#1a56db';
  const cur = settings.currency || 'ARS';
  const html=`<!DOCTYPE html><html><head><meta charset="utf-8"><title>Recibo #${res.fileNumber}</title>
<style>
  *{margin:0;padding:0;box-sizing:border-box;}
  body{font-family:'Helvetica Neue',Arial,sans-serif;padding:40px;color:#1a1a1a;font-size:13px;}
  .header{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:32px;padding-bottom:20px;border-bottom:2px solid ${color};}
  .agency-info{font-size:12px;color:#555;line-height:1.6;margin-top:4px;}
  .recibo-title{text-align:right;}
  .recibo-title h2{font-size:20px;font-weight:700;color:#1a1a1a;}
  .recibo-title p{font-size:12px;color:#666;margin-top:4px;}
  .section{margin-bottom:24px;}
  .section-title{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:1px;color:#888;margin-bottom:10px;padding-bottom:6px;border-bottom:1px solid #eee;}
  .info-grid{display:grid;grid-template-columns:1fr 1fr;gap:8px;}
  .info-item label{display:block;font-size:11px;color:#888;margin-bottom:2px;}
  .info-item span{font-size:13px;font-weight:500;}
  table{width:100%;border-collapse:collapse;font-size:12px;}
  th{background:#f5f7fa;padding:8px 10px;text-align:left;font-weight:600;font-size:11px;text-transform:uppercase;color:#555;}
  td{padding:8px 10px;border-bottom:1px solid #eee;}
  .total-box{margin-top:20px;padding:16px;background:#f5f7fa;border-radius:8px;}
  .total-row{display:flex;justify-content:space-between;padding:4px 0;font-size:13px;}
  .total-row.highlight{font-weight:700;font-size:15px;color:${color};border-top:2px solid #ddd;margin-top:8px;padding-top:10px;}
  .footer{margin-top:40px;padding-top:16px;border-top:1px solid #ddd;text-align:center;font-size:11px;color:#888;}
  .header-note{background:#f0f4ff;border-left:3px solid ${color};padding:8px 12px;margin-bottom:20px;font-size:12px;color:#444;}
  @media print{body{padding:20px;}}
</style></head><body>
<div class="header">
  <div>
    ${getLogoHtml(settings)}
    <div class="agency-info">
      ${settings.address?settings.address+"<br>":""}
      ${settings.phone?"Tel: "+settings.phone+"  ":""} ${settings.email?"Email: "+settings.email+"<br>":""}
      ${settings.cuit?"CUIT: "+settings.cuit:""}
    </div>
  </div>
  <div class="recibo-title">
    <h2>RECIBO DE PAGO</h2>
    <p>File N° <strong>${res.fileNumber}</strong></p>
    <p>Fecha: ${fmtDate(today())}</p>
  </div>
</div>

${settings.headerNote?`<div class="header-note">${settings.headerNote}</div>`:""}

<div class="section">
  <div class="section-title">Datos del pasajero</div>
  <div class="info-grid">
    ${res.passengers.map(p=>`<div class="info-item"><label>Pasajero</label><span>${p.name}</span></div>
    ${p.dni?`<div class="info-item"><label>DNI/Pasaporte</label><span>${p.dni}</span></div>`:""}`).join("")}
  </div>
</div>

<div class="section">
  <div class="section-title">Detalle del viaje</div>
  <div class="info-grid">
    <div class="info-item"><label>Destino</label><span>${res.destination||"—"}</span></div>
    <div class="info-item"><label>Fecha de salida</label><span>${fmtDate(res.departureDate)}</span></div>
    <div class="info-item"><label>Fecha de regreso</label><span>${fmtDate(res.returnDate)}</span></div>
    <div class="info-item"><label>Estado</label><span>${STATUS[res.status]?.label||res.status}</span></div>
  </div>
  ${res.description?`<div style="margin-top:10px;padding:10px;background:#f9fafb;border-radius:6px;font-size:12px;color:#555;">${res.description}</div>`:""}
</div>

<div class="section">
  <div class="section-title">Servicios incluidos</div>
  <table><thead><tr><th>Servicio</th><th>Proveedor</th><th>File/Localizador</th><th style="text-align:right">Precio</th></tr></thead>
  <tbody>${res.services.map(s=>{const prov=providers.find(p=>p.id===s.providerId);return`<tr><td>${SVC_TYPES.find(t=>t.v===s.type)?.l||s.type} — ${s.description||""}</td><td>${prov?.name||"—"}</td><td>${s.providerFileNumber||"—"}</td><td style="text-align:right">${fmt$(s.salePrice,cur)}</td></tr>`;}).join("")}
  </tbody></table>
</div>

<div class="section">
  <div class="section-title">Historial de pagos recibidos</div>
  <table><thead><tr><th>Fecha</th><th>Forma de pago</th><th>Notas</th><th style="text-align:right">Importe</th></tr></thead>
  <tbody>${(res.paymentsReceived||[]).map(p=>`<tr><td>${fmtDate(p.date)}</td><td>${p.method||"—"}</td><td>${p.notes||""}</td><td style="text-align:right">${fmt$(p.amount,cur)}</td></tr>`).join("")}
  </tbody></table>
</div>

<div class="total-box">
  <div class="total-row"><span>Precio total del viaje</span><span>${fmt$(res.salePrice,cur)}</span></div>
  <div class="total-row"><span>Total abonado</span><span>${fmt$(received,cur)}</span></div>
  <div class="total-row highlight"><span>Saldo pendiente</span><span>${fmt$(pending,cur)}</span></div>
</div>

<div class="footer">
  ${settings.agencyName}${settings.footerText?" — "+settings.footerText:""}
</div>
<script>window.print();window.onafterprint=()=>window.close();</script>
</body></html>`;
  const w=window.open("","_blank","width=800,height=900");
  w.document.write(html);w.document.close();
}

function printVoucher(res, settings, providers, type){
  const isAereo = type==="aereo";
  const color = settings.voucherColor || '#1a56db';
  const cur = settings.currency || 'ARS';
  const filteredServices = res.services.filter(s=>isAereo ? s.type==="vuelo" : s.type!=="vuelo");
  const html=`<!DOCTYPE html><html><head><meta charset="utf-8"><title>Voucher ${isAereo?"Aéreo":"Terrestre"} #${res.fileNumber}</title>
<style>
  *{margin:0;padding:0;box-sizing:border-box;}
  body{font-family:'Helvetica Neue',Arial,sans-serif;color:#1a1a1a;font-size:13px;}
  .voucher{padding:30px;max-width:700px;margin:0 auto;}
  .header{background:${color};color:white;padding:20px 24px;border-radius:12px 12px 0 0;display:flex;justify-content:space-between;align-items:center;}
  .agency-name{font-size:18px;font-weight:700;}
  .voucher-type{font-size:12px;opacity:.85;text-transform:uppercase;letter-spacing:1px;}
  .content{border:1px solid #dde1e9;border-top:none;border-radius:0 0 12px 12px;padding:24px;}
  .file-badge{display:inline-block;background:#EFF6FF;color:#1d4ed8;font-weight:700;font-size:12px;padding:4px 12px;border-radius:20px;margin-bottom:20px;}
  .grid{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:20px;}
  .field label{font-size:10px;text-transform:uppercase;letter-spacing:.5px;color:#888;display:block;margin-bottom:3px;}
  .field span{font-size:13px;font-weight:500;color:#1a1a1a;}
  .section-title{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:1px;color:#888;margin-bottom:12px;padding-bottom:6px;border-bottom:1px solid #eee;}
  .service-card{background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:14px;margin-bottom:10px;}
  .service-header{font-weight:600;font-size:13px;margin-bottom:8px;color:${color};}
  .service-detail{font-size:12px;color:#555;line-height:1.7;}
  .passengers-list{display:flex;flex-wrap:wrap;gap:10px;}
  .passenger-badge{background:#EFF6FF;color:#1d4ed8;border:1px solid #BFDBFE;border-radius:8px;padding:8px 14px;font-size:12px;font-weight:500;}
  .important-box{background:#FFF7ED;border:1px solid #FED7AA;border-radius:8px;padding:14px;margin-top:20px;font-size:12px;color:#92400E;}
  .footer{margin-top:24px;text-align:center;font-size:11px;color:#aaa;padding-top:16px;border-top:1px solid #eee;}
  .logo-header{display:flex;align-items:center;gap:12px;}
  @media print{body{-webkit-print-color-adjust:exact;print-color-adjust:exact;}}
</style></head><body>
<div class="voucher">
  <div class="header">
    <div class="logo-header">
      ${settings.logo?`<img src="${settings.logo}" style="max-height:40px;max-width:100px;object-fit:contain;background:white;padding:4px;border-radius:6px;" alt="Logo"/>`:""}
      <div>
        <div class="agency-name">${settings.agencyName}</div>
        <div class="voucher-type">Voucher ${isAereo?"Aéreo":"Terrestre"}</div>
      </div>
    </div>
    <div style="text-align:right;font-size:12px;opacity:.9;">
      <div>File N° <strong>${res.fileNumber}</strong></div>
      <div>Emitido: ${fmtDate(today())}</div>
    </div>
  </div>
  <div class="content">
    <div class="file-badge">${STATUS[res.status]?.label||res.status} — ${res.destination||"Sin destino"}</div>
    
    <div class="section-title">Pasajeros</div>
    <div class="passengers-list" style="margin-bottom:20px;">
      ${res.passengers.map(p=>`<div class="passenger-badge">${p.name}${p.dni?" · "+p.dni:""}</div>`).join("")}
    </div>

    <div class="grid" style="margin-bottom:20px;">
      <div class="field"><label>Fecha de salida</label><span>${fmtDate(res.departureDate)}</span></div>
      <div class="field"><label>Fecha de regreso</label><span>${fmtDate(res.returnDate)}</span></div>
      <div class="field"><label>Destino</label><span>${res.destination||"—"}</span></div>
      <div class="field"><label>Días / Noches</label><span>${res.departureDate&&res.returnDate?Math.ceil((new Date(res.returnDate)-new Date(res.departureDate))/(1000*60*60*24))+" días":"—"}</span></div>
    </div>

    ${filteredServices.length===0?`<p style="color:#94a3b8;font-size:13px;">No hay servicios ${isAereo?"aéreos":"terrestres"} registrados para este file.</p>`:""}
    <div class="section-title">${filteredServices.length>0?"Servicios incluidos":""}</div>
    ${filteredServices.map(s=>{
      const prov=providers.find(p=>p.id===s.providerId);
      return`<div class="service-card">
        <div class="service-header">${SVC_TYPES.find(t=>t.v===s.type)?.l||s.type}</div>
        <div class="service-detail">
          ${s.description?"<strong>Descripción:</strong> "+s.description+"<br>":""}
          ${prov?"<strong>Proveedor:</strong> "+prov.name+"<br>":""}
          ${s.providerFileNumber?"<strong>Localizador/File:</strong> "+s.providerFileNumber+"<br>":""}
          ${s.salePrice?"<strong>Precio:</strong> "+fmt$(s.salePrice,cur)+"<br>":""}
        </div>
      </div>`;}).join("")}
    
    ${res.notes?`<div class="important-box"><strong>Observaciones:</strong> ${res.notes}</div>`:""}

    <div class="footer">
      ${settings.agencyName}${settings.phone?" · "+settings.phone:""}${settings.email?" · "+settings.email:""}
      <br>${settings.footerText||"Gracias por elegir nuestros servicios."}
    </div>
  </div>
</div>
<script>window.print();window.onafterprint=()=>window.close();</script>
</body></html>`;
  const w=window.open("","_blank","width=750,height=900");
  w.document.write(html);w.document.close();
}

// ══════════════════════════════════════════════════
//  UI PRIMITIVES
// ══════════════════════════════════════════════════
const S = {
  primaryBtn:{background:"#2563EB",color:"#fff",border:"none",padding:"7px 16px",borderRadius:8,fontWeight:600,fontSize:13,cursor:"pointer",display:"inline-flex",alignItems:"center",gap:6},
  secondaryBtn:{background:"#F1F5F9",color:"#475569",border:"1px solid #E2E8F0",padding:"7px 16px",borderRadius:8,fontWeight:600,fontSize:13,cursor:"pointer",display:"inline-flex",alignItems:"center",gap:6},
  dangerBtn:{background:"#EF4444",color:"#fff",border:"none",padding:"7px 16px",borderRadius:8,fontWeight:600,fontSize:13,cursor:"pointer",display:"inline-flex",alignItems:"center",gap:6},
  ghostBtn:{background:"transparent",color:"#64748B",border:"1px solid #CBD5E1",padding:"6px 12px",borderRadius:8,fontWeight:500,fontSize:12,cursor:"pointer",display:"inline-flex",alignItems:"center",gap:4},
  successBtn:{background:"#059669",color:"#fff",border:"none",padding:"7px 16px",borderRadius:8,fontWeight:600,fontSize:13,cursor:"pointer",display:"inline-flex",alignItems:"center",gap:6},
  input:{padding:"8px 12px",border:"1px solid #E2E8F0",borderRadius:8,fontSize:13,color:"#1E293B",background:"#fff",outline:"none",width:"100%",boxSizing:"border-box"},
  label:{fontSize:11,fontWeight:700,color:"#475569",textTransform:"uppercase",letterSpacing:.5,display:"block",marginBottom:4},
  card:{background:"#fff",borderRadius:12,boxShadow:"0 1px 3px rgba(0,0,0,.07)",padding:20},
  sectionTitle:{margin:"0 0 16px",fontSize:14,fontWeight:700,color:"#0F172A",display:"flex",alignItems:"center",gap:8},
};

const Btn=({children,onClick,variant="primary",size="md",disabled,style:sx})=>{
  const base=variant==="primary"?S.primaryBtn:variant==="danger"?S.dangerBtn:variant==="ghost"?S.ghostBtn:variant==="success"?S.successBtn:S.secondaryBtn;
  const sz=size==="sm"?{padding:"4px 10px",fontSize:12}:size==="lg"?{padding:"10px 20px",fontSize:14}:{};
  return <button onClick={onClick} disabled={disabled} style={{...base,...sz,opacity:disabled?.5:1,...sx}}>{children}</button>;
};

const Field=({label,children,col=1,style:sx})=>(
  <div style={{gridColumn:`span ${col}`,...sx}}>
    {label&&<label style={S.label}>{label}</label>}
    {children}
  </div>
);

const Inp=({label,value,onChange,type="text",placeholder,required,col,style:sx})=>(
  <Field label={label} col={col} style={sx}>
    <input type={type} value={value||""} onChange={e=>onChange(e.target.value)} placeholder={placeholder}
      style={S.input} required={required}/>
  </Field>
);

const Sel=({label,value,onChange,options,col,style:sx})=>(
  <Field label={label} col={col} style={sx}>
    <select value={value||""} onChange={e=>onChange(e.target.value)} style={{...S.input,appearance:"auto"}}>
      <option value="">Seleccionar…</option>
      {options.map(o=><option key={o.v||o} value={o.v||o}>{o.l||o}</option>)}
    </select>
  </Field>
);

const Txta=({label,value,onChange,rows=3,placeholder,col,style:sx})=>(
  <Field label={label} col={col} style={sx}>
    <textarea value={value||""} onChange={e=>onChange(e.target.value)} rows={rows} placeholder={placeholder}
      style={{...S.input,resize:"vertical"}}/>
  </Field>
);

const Badge=({status})=>{
  const s=STATUS[status]||STATUS.cotizacion;
  return <span style={{background:s.bg,color:s.c,padding:"3px 10px",borderRadius:20,fontSize:11,fontWeight:700,whiteSpace:"nowrap"}}>{s.label}</span>;
};

const Card=({children,style:sx,onClick})=>(
  <div style={{...S.card,...sx,cursor:onClick?"pointer":"default"}} onClick={onClick}>{children}</div>
);

const Modal=({title,onClose,children,width=720,footer})=>(
  <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.55)",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
    <div style={{background:"#fff",borderRadius:16,width:"100%",maxWidth:width,maxHeight:"92vh",display:"flex",flexDirection:"column",overflow:"hidden"}}>
      <div style={{padding:"16px 20px",borderBottom:"1px solid #E2E8F0",display:"flex",alignItems:"center",justifyContent:"space-between",flexShrink:0}}>
        <h3 style={{margin:0,fontSize:16,fontWeight:700,color:"#0F172A"}}>{title}</h3>
        <button onClick={onClose} style={{background:"none",border:"none",cursor:"pointer",color:"#94A3B8",display:"flex",alignItems:"center"}}><X size={20}/></button>
      </div>
      <div style={{overflowY:"auto",flex:1,padding:20}}>{children}</div>
      {footer&&<div style={{padding:"12px 20px",borderTop:"1px solid #E2E8F0",display:"flex",gap:8,justifyContent:"flex-end",flexShrink:0}}>{footer}</div>}
    </div>
  </div>
);

const Tabs=({tabs,active,setActive})=>(
  <div style={{display:"flex",gap:2,borderBottom:"2px solid #E2E8F0",marginBottom:16}}>
    {tabs.map(t=>(
      <button key={t.id} onClick={()=>setActive(t.id)} style={{
        padding:"8px 14px",border:"none",background:"none",cursor:"pointer",
        fontSize:13,fontWeight:600,
        color:active===t.id?"#2563EB":"#64748B",
        borderBottom:active===t.id?"2px solid #2563EB":"2px solid transparent",
        marginBottom:-2,transition:"color .15s"
      }}>{t.label}</button>
    ))}
  </div>
);

const EmptyState=({icon:Icon,title,sub,action})=>(
  <div style={{textAlign:"center",padding:"48px 24px",color:"#94A3B8"}}>
    <Icon size={40} style={{marginBottom:12,opacity:.4}}/>
    <div style={{fontSize:15,fontWeight:600,color:"#64748B",marginBottom:6}}>{title}</div>
    {sub&&<div style={{fontSize:13,marginBottom:16}}>{sub}</div>}
    {action}
  </div>
);

const Grid=({cols=2,gap=16,children,style:sx})=>(
  <div style={{display:"grid",gridTemplateColumns:`repeat(${cols},1fr)`,gap,...sx}}>{children}</div>
);

const Divider=({label})=>(
  <div style={{display:"flex",alignItems:"center",gap:8,margin:"16px 0"}}>
    <div style={{flex:1,height:1,background:"#E2E8F0"}}/>
    {label&&<span style={{fontSize:11,fontWeight:700,color:"#94A3B8",textTransform:"uppercase",letterSpacing:.5}}>{label}</span>}
    <div style={{flex:1,height:1,background:"#E2E8F0"}}/>
  </div>
);

// ══════════════════════════════════════════════════
//  SIDEBAR
// ══════════════════════════════════════════════════
const NAV=[
  {id:"dashboard",Icon:LayoutDashboard,l:"Dashboard"},
  {id:"reservations",Icon:CalendarDays,l:"Reservas"},
  {id:"clients",Icon:Users,l:"Clientes"},
  {id:"providers",Icon:Building2,l:"Proveedores"},
  {id:"commissions",Icon:TrendingUp,l:"Comisiones"},
  {id:"reports",Icon:BarChart2,l:"Reportes"},
  {id:"settings",Icon:Settings,l:"Configuración"},
];

function Sidebar({section,setSection,settings}){
  return (
    <aside style={{width:210,background:"#0F172A",display:"flex",flexDirection:"column",flexShrink:0,userSelect:"none"}}>
      <div style={{padding:"18px 16px",borderBottom:"1px solid #1E293B",display:"flex",alignItems:"center",gap:10}}>
        {settings.logo
          ? <img src={settings.logo} style={{height:36,width:36,objectFit:"contain",borderRadius:6,background:"#fff",padding:2}} alt="Logo"/>
          : <div style={{width:36,height:36,borderRadius:8,background:"#2563EB",display:"flex",alignItems:"center",justifyContent:"center",color:"#fff",fontSize:14,fontWeight:800}}>
              {settings.agencyName?.charAt(0)||"A"}
            </div>
        }
        <div>
          <div style={{fontSize:10,color:"#475569",fontWeight:700,textTransform:"uppercase",letterSpacing:1,marginBottom:2}}>Agencia</div>
          <div style={{fontSize:12,color:"#F8FAFC",fontWeight:700,lineHeight:1.3}}>{settings.agencyName}</div>
        </div>
      </div>
      <nav style={{flex:1,padding:"8px 0",overflowY:"auto"}}>
        {NAV.map(({id,Icon,l})=>(
          <button key={id} onClick={()=>setSection(id)} style={{
            display:"flex",alignItems:"center",gap:10,width:"100%",padding:"10px 16px",
            background:section===id?"rgba(59,130,246,.15)":"transparent",
            border:"none",cursor:"pointer",textAlign:"left",
            color:section===id?"#60A5FA":"#94A3B8",
            fontSize:13,fontWeight:section===id?600:400,
            borderLeft:section===id?"3px solid #3B82F6":"3px solid transparent",
          }}>
            <Icon size={15}/>{l}
          </button>
        ))}
      </nav>
      <div style={{padding:"10px 16px",borderTop:"1px solid #1E293B",fontSize:10,color:"#334155",textAlign:"center"}}>
        TravelManager v2.0
      </div>
    </aside>
  );
}

// ══════════════════════════════════════════════════
//  DASHBOARD
// ══════════════════════════════════════════════════
function Dashboard({data,setSection}){
  const {reservations,clients,providers,settings}=data;
  const n=today();
  const cur = settings.currency||'ARS';

  const stats=useMemo(()=>{
    const notCancelled=reservations.filter(r=>r.status!=="cancelada");
    const totalSale=sum(notCancelled.map(r=>r.salePrice||0));
    const totalReceived=sum(reservations.map(r=>paidSum(r.paymentsReceived)));
    const totalPending=totalSale-totalReceived;
    const totalCost=sum(notCancelled.map(r=>sum(r.services.map(s=>s.costPrice||0))));
    const totalCommission=totalSale-totalCost;
    const active=reservations.filter(r=>["confirmada","en_viaje"].includes(r.status)).length;

    const upcoming=reservations
      .filter(r=>r.departureDate>=n&&r.status!=="cancelada")
      .sort((a,b)=>a.departureDate.localeCompare(b.departureDate))
      .slice(0,6);

    const alerts=[];
    reservations.forEach(res=>{
      res.services.forEach(svc=>{
        (svc.paymentsDue||[]).filter(p=>!p.paid).forEach(p=>{
          const prov=providers.find(pr=>pr.id===svc.providerId);
          const overdue=p.dueDate&&p.dueDate<n;
          alerts.push({res,svc,p,providerName:prov?.name||"Sin proveedor",overdue});
        });
      });
    });
    alerts.sort((a,b)=>(a.p.dueDate||"").localeCompare(b.p.dueDate||""));

    return{totalSale,totalReceived,totalPending,totalCommission,active,upcoming,alerts};
  },[reservations]);

  const cards=[
    {l:"Reservas activas",v:stats.active,Icon:CalendarDays,c:"#2563EB",bg:"#EFF6FF"},
    {l:"Ingresos totales",v:fmt$(stats.totalSale,cur),Icon:TrendingUp,c:"#059669",bg:"#ECFDF5"},
    {l:"Total cobrado",v:fmt$(stats.totalReceived,cur),Icon:Check,c:"#7C3AED",bg:"#F5F3FF"},
    {l:"Por cobrar",v:fmt$(stats.totalPending,cur),Icon:AlertCircle,c:"#D97706",bg:"#FFFBEB"},
    {l:"Comisiones",v:fmt$(stats.totalCommission,cur),Icon:TrendingUp,c:"#DC2626",bg:"#FEF2F2"},
    {l:"Total clientes",v:clients.length,Icon:Users,c:"#0891B2",bg:"#F0F9FF"},
  ];

  return(
    <div style={{padding:28,overflowY:"auto",flex:1}}>
      <div style={{marginBottom:24}}>
        <h1 style={{margin:0,fontSize:22,fontWeight:800,color:"#0F172A"}}>Dashboard</h1>
        <p style={{margin:"4px 0 0",fontSize:13,color:"#64748B"}}>Resumen general de la agencia</p>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(175px,1fr))",gap:14,marginBottom:28}}>
        {cards.map(x=>(
          <div key={x.l} style={{background:x.bg,borderRadius:12,padding:16}}>
            <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:10}}>
              <x.Icon size={16} style={{color:x.c}}/>
              <span style={{fontSize:11,fontWeight:700,color:x.c,textTransform:"uppercase",letterSpacing:.5}}>{x.l}</span>
            </div>
            <div style={{fontSize:22,fontWeight:800,color:x.c}}>{x.v}</div>
          </div>
        ))}
      </div>
      <Grid cols={2} gap={16}>
        <Card>
          <h3 style={S.sectionTitle}><CalendarDays size={16} style={{color:"#2563EB"}}/>Próximas salidas</h3>
          {stats.upcoming.length===0
            ?<p style={{color:"#94A3B8",fontSize:13}}>No hay salidas próximas.</p>
            :stats.upcoming.map(r=>(
              <div key={r.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"8px 0",borderBottom:"1px solid #F1F5F9"}}>
                <div>
                  <div style={{fontSize:13,fontWeight:600,color:"#1E293B"}}>{r.destination||"Sin destino"}</div>
                  <div style={{fontSize:11,color:"#64748B"}}>{r.passengers[0]?.name||"Sin pasajero"} — File #{r.fileNumber}</div>
                </div>
                <div style={{display:"flex",flexDirection:"column",alignItems:"flex-end",gap:3}}>
                  <Badge status={r.status}/>
                  <span style={{fontSize:11,color:"#2563EB",fontWeight:700}}>{fmtDate(r.departureDate)}</span>
                </div>
              </div>
            ))
          }
        </Card>
        <Card>
          <h3 style={S.sectionTitle}><AlertCircle size={16} style={{color:"#EF4444"}}/>Pagos a proveedores pendientes</h3>
          {stats.alerts.length===0
            ?<p style={{color:"#94A3B8",fontSize:13}}>No hay pagos pendientes.</p>
            :stats.alerts.slice(0,7).map((a,i)=>(
              <div key={i} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"7px 0",borderBottom:"1px solid #F1F5F9"}}>
                <div>
                  <div style={{fontSize:13,fontWeight:600,color:"#1E293B"}}>{a.providerName}</div>
                  <div style={{fontSize:11,color:"#64748B"}}>File #{a.res.fileNumber} — {a.svc.description||a.svc.type}</div>
                </div>
                <div style={{textAlign:"right"}}>
                  <div style={{fontSize:12,fontWeight:700,color:"#EF4444"}}>{fmt$(a.p.amount,cur)}</div>
                  <div style={{fontSize:11,color:a.overdue?"#EF4444":"#64748B",fontWeight:a.overdue?700:400}}>
                    {a.overdue?"⚠ ":""}{fmtDate(a.p.dueDate)}
                  </div>
                </div>
              </div>
            ))
          }
        </Card>
      </Grid>
    </div>
  );
}

// ══════════════════════════════════════════════════
//  RESERVATION MODAL
// ══════════════════════════════════════════════════
const newRes=(fileNum,defComm)=>({
  id:genId(),fileNumber:String(fileNum).padStart(5,"0"),
  status:"cotizacion",destination:"",departureDate:"",returnDate:"",description:"",
  salePrice:0,commissionPercent:defComm||15,
  passengers:[],services:[],paymentsReceived:[],notes:"",createdAt:today()
});
const newPass=()=>({id:genId(),name:"",dni:"",email:"",phone:"",birthDate:""});
const newSvc=()=>({id:genId(),type:"vuelo",description:"",providerId:"",providerFileNumber:"",costPrice:0,salePrice:0,paymentsDue:[]});
const newPay=()=>({id:genId(),date:today(),amount:0,method:"Transferencia",notes:""});
const newDue=()=>({id:genId(),dueDate:"",amount:0,paid:false,paidDate:null,notes:""});

function ReservationModal({initial,providers,settings,onSave,onClose}){
  const [r,setR]=useState(initial||newRes(1,settings.defaultCommission));
  const [tab,setTab]=useState("general");
  const cur = settings.currency||'ARS';

  const upd=(k,v)=>setR(p=>({...p,[k]:v}));
  const arrUpd=(k,id,fn)=>setR(p=>({...p,[k]:p[k].map(x=>x.id===id?fn(x):x)}));
  const arrDel=(k,id)=>setR(p=>({...p,[k]:p[k].filter(x=>x.id!==id)}));
  const arrAdd=(k,v)=>setR(p=>({...p,[k]:[...p[k],v]}));

  const received=paidSum(r.paymentsReceived);
  const pending=(r.salePrice||0)-received;
  const totalCost=sum(r.services.map(s=>s.costPrice||0));
  const commission=(r.salePrice||0)-totalCost;

  const tabs=[
    {id:"general",label:"✈ General"},
    {id:"passengers",label:`👥 Pasajeros (${r.passengers.length})`},
    {id:"services",label:`🛎 Servicios (${r.services.length})`},
    {id:"payments",label:`💳 Cobros (${r.paymentsReceived.length})`},
  ];

  return(
    <Modal title={`${initial?"Editar":"Nueva"} Reserva — File #${r.fileNumber}`} onClose={onClose} width={780}
      footer={<>
        <Btn variant="secondary" onClick={onClose}>Cancelar</Btn>
        <Btn onClick={()=>onSave(r)}>Guardar reserva</Btn>
      </>}>
      <Tabs tabs={tabs} active={tab} setActive={setTab}/>

      {tab==="general"&&(
        <div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:12,marginBottom:12}}>
            <Inp label="N° de file" value={r.fileNumber} onChange={v=>upd("fileNumber",v)} required/>
            <Sel label="Estado" value={r.status} onChange={v=>upd("status",v)}
              options={Object.entries(STATUS).map(([v,s])=>({v,l:s.label}))}/>
            <Inp label="Destino" value={r.destination} onChange={v=>upd("destination",v)} placeholder="Ej: Cancún, México"/>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:12,marginBottom:12}}>
            <Inp label="Fecha de salida" value={r.departureDate} onChange={v=>upd("departureDate",v)} type="date"/>
            <Inp label="Fecha de regreso" value={r.returnDate} onChange={v=>upd("returnDate",v)} type="date"/>
            <Field label="Días de viaje">
              <div style={{...S.input,background:"#F8FAFC",color:"#64748B",display:"flex",alignItems:"center"}}>
                {r.departureDate&&r.returnDate?Math.ceil((new Date(r.returnDate)-new Date(r.departureDate))/(86400000))+" días":"—"}
              </div>
            </Field>
          </div>
          <Txta label="Descripción del viaje" value={r.description} onChange={v=>upd("description",v)}
            placeholder="Descripción general del servicio contratado..." style={{marginBottom:12}}/>
          <Divider label="Resumen financiero"/>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:12,marginBottom:12}}>
            <Field label={`Precio de venta (${cur})`}>
              <input type="number" value={r.salePrice||""} onChange={e=>upd("salePrice",+e.target.value)} style={S.input} placeholder="0.00"/>
            </Field>
            <Field label="% Comisión">
              <input type="number" value={r.commissionPercent||""} onChange={e=>upd("commissionPercent",+e.target.value)} style={S.input} placeholder="15"/>
            </Field>
            <Field label="Costo total (servicios)">
              <div style={{...S.input,background:"#F8FAFC",color:"#64748B",display:"flex",alignItems:"center"}}>{fmt$(totalCost,cur)}</div>
            </Field>
          </div>
          <div style={{background:"#EFF6FF",border:"1px solid #BFDBFE",borderRadius:10,padding:14,display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12}}>
            {[
              {l:"Precio de venta",v:fmt$(r.salePrice,cur),c:"#1E40AF"},
              {l:"Total cobrado",v:fmt$(received,cur),c:"#059669"},
              {l:"Saldo pendiente",v:fmt$(pending,cur),c:pending>0?"#D97706":"#059669"},
              {l:"Comisión neta",v:fmt$(commission,cur),c:"#7C3AED"},
            ].map(x=>(
              <div key={x.l} style={{textAlign:"center"}}>
                <div style={{fontSize:10,fontWeight:700,textTransform:"uppercase",letterSpacing:.5,color:"#64748B",marginBottom:4}}>{x.l}</div>
                <div style={{fontSize:18,fontWeight:800,color:x.c}}>{x.v}</div>
              </div>
            ))}
          </div>
          <Txta label="Notas internas" value={r.notes} onChange={v=>upd("notes",v)} rows={2}
            placeholder="Observaciones internas..." style={{marginTop:12}}/>
        </div>
      )}

      {tab==="passengers"&&(
        <div>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
            <span style={{fontSize:13,color:"#64748B"}}>Pasajeros de la reserva.</span>
            <Btn size="sm" onClick={()=>arrAdd("passengers",newPass())}><Plus size={14}/>Agregar pasajero</Btn>
          </div>
          {r.passengers.length===0&&<EmptyState icon={Users} title="Sin pasajeros" sub="Agregá al menos un pasajero."/>}
          {r.passengers.map((p,i)=>(
            <div key={p.id} style={{border:"1px solid #E2E8F0",borderRadius:10,padding:14,marginBottom:12}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
                <span style={{fontSize:13,fontWeight:700,color:"#1E293B"}}>Pasajero {i+1}</span>
                <button onClick={()=>arrDel("passengers",p.id)} style={{...S.ghostBtn,color:"#EF4444",borderColor:"#FCA5A5"}}><Trash2 size={13}/>Eliminar</button>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10,marginBottom:10}}>
                <Inp label="Nombre completo" value={p.name} onChange={v=>arrUpd("passengers",p.id,x=>({...x,name:v}))} required/>
                <Inp label="DNI / Pasaporte" value={p.dni} onChange={v=>arrUpd("passengers",p.id,x=>({...x,dni:v}))}/>
                <Inp label="Fecha de nacimiento" value={p.birthDate} onChange={v=>arrUpd("passengers",p.id,x=>({...x,birthDate:v}))} type="date"/>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
                <Inp label="Email" value={p.email} onChange={v=>arrUpd("passengers",p.id,x=>({...x,email:v}))} type="email"/>
                <Inp label="Teléfono" value={p.phone} onChange={v=>arrUpd("passengers",p.id,x=>({...x,phone:v}))}/>
              </div>
            </div>
          ))}
        </div>
      )}

      {tab==="services"&&(
        <div>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
            <span style={{fontSize:13,color:"#64748B"}}>Servicios contratados con proveedores.</span>
            <Btn size="sm" onClick={()=>arrAdd("services",newSvc())}><Plus size={14}/>Agregar servicio</Btn>
          </div>
          {r.services.length===0&&<EmptyState icon={FileCheck} title="Sin servicios" sub="Agregá los servicios del viaje."/>}
          {r.services.map((s,i)=>{
            const SvcIcon=SVC_TYPES.find(t=>t.v===s.type)?.Icon||FileText;
            return(
              <div key={s.id} style={{border:"1px solid #E2E8F0",borderRadius:10,padding:14,marginBottom:12}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
                  <div style={{display:"flex",alignItems:"center",gap:8}}>
                    <SvcIcon size={15} style={{color:"#2563EB"}}/>
                    <span style={{fontSize:13,fontWeight:700,color:"#1E293B"}}>Servicio {i+1}</span>
                  </div>
                  <button onClick={()=>arrDel("services",s.id)} style={{...S.ghostBtn,color:"#EF4444",borderColor:"#FCA5A5"}}><Trash2 size={13}/>Eliminar</button>
                </div>
                <div style={{display:"grid",gridTemplateColumns:"1fr 2fr",gap:10,marginBottom:10}}>
                  <Sel label="Tipo" value={s.type} onChange={v=>arrUpd("services",s.id,x=>({...x,type:v}))}
                    options={SVC_TYPES.map(t=>({v:t.v,l:t.l}))}/>
                  <Inp label="Descripción" value={s.description} onChange={v=>arrUpd("services",s.id,x=>({...x,description:v}))}
                    placeholder="Ej: Vuelo BUE-CUN ida y vuelta"/>
                </div>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",gap:10,marginBottom:10}}>
                  <Sel label="Proveedor" value={s.providerId} onChange={v=>arrUpd("services",s.id,x=>({...x,providerId:v}))}
                    options={providers.map(p=>({v:p.id,l:p.name}))}/>
                  <Inp label="File proveedor" value={s.providerFileNumber} onChange={v=>arrUpd("services",s.id,x=>({...x,providerFileNumber:v}))} placeholder="N° file/localizador"/>
                  <Field label={`Costo neto (${cur})`}>
                    <input type="number" value={s.costPrice||""} onChange={e=>arrUpd("services",s.id,x=>({...x,costPrice:+e.target.value}))} style={S.input} placeholder="0.00"/>
                  </Field>
                  <Field label={`Precio venta (${cur})`}>
                    <input type="number" value={s.salePrice||""} onChange={e=>arrUpd("services",s.id,x=>({...x,salePrice:+e.target.value}))} style={S.input} placeholder="0.00"/>
                  </Field>
                </div>
                <div style={{background:"#F8FAFC",borderRadius:8,padding:12}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
                    <span style={{fontSize:11,fontWeight:700,color:"#475569",textTransform:"uppercase",letterSpacing:.5}}>Pagos al proveedor</span>
                    <button onClick={()=>arrUpd("services",s.id,x=>({...x,paymentsDue:[...(x.paymentsDue||[]),newDue()]}))}
                      style={S.ghostBtn}><Plus size={12}/>Agregar pago</button>
                  </div>
                  {(s.paymentsDue||[]).length===0&&<p style={{fontSize:12,color:"#94A3B8",margin:0}}>Sin pagos programados.</p>}
                  {(s.paymentsDue||[]).map((pd,j)=>(
                    <div key={pd.id} style={{display:"grid",gridTemplateColumns:"1fr 1fr 2fr auto auto",gap:8,alignItems:"end",marginBottom:8}}>
                      <Field label={j===0?"Vencimiento":""}>
                        <input type="date" value={pd.dueDate||""} onChange={e=>arrUpd("services",s.id,x=>({...x,paymentsDue:x.paymentsDue.map(p=>p.id===pd.id?{...p,dueDate:e.target.value}:p)}))} style={S.input}/>
                      </Field>
                      <Field label={j===0?"Importe":""}>
                        <input type="number" value={pd.amount||""} onChange={e=>arrUpd("services",s.id,x=>({...x,paymentsDue:x.paymentsDue.map(p=>p.id===pd.id?{...p,amount:+e.target.value}:p)}))} style={S.input} placeholder="0.00"/>
                      </Field>
                      <Field label={j===0?"Notas":""}>
                        <input type="text" value={pd.notes||""} onChange={e=>arrUpd("services",s.id,x=>({...x,paymentsDue:x.paymentsDue.map(p=>p.id===pd.id?{...p,notes:e.target.value}:p)}))} style={S.input} placeholder="Obs…"/>
                      </Field>
                      <Field label={j===0?"Pagado":""}>
                        <div style={{display:"flex",alignItems:"center",height:36}}>
                          <input type="checkbox" checked={pd.paid||false}
                            onChange={e=>arrUpd("services",s.id,x=>({...x,paymentsDue:x.paymentsDue.map(p=>p.id===pd.id?{...p,paid:e.target.checked,paidDate:e.target.checked?today():null}:p)}))}
                            style={{width:16,height:16,cursor:"pointer"}}/>
                        </div>
                      </Field>
                      <Field label={j===0?" ":""}>
                        <button onClick={()=>arrUpd("services",s.id,x=>({...x,paymentsDue:x.paymentsDue.filter(p=>p.id!==pd.id)}))}
                          style={{...S.ghostBtn,color:"#EF4444",height:36}}><Trash2 size={13}/></button>
                      </Field>
                    </div>
                  ))}
                  {(s.paymentsDue||[]).length>0&&(
                    <div style={{display:"flex",justifyContent:"flex-end",gap:16,fontSize:12,marginTop:8,paddingTop:8,borderTop:"1px solid #E2E8F0"}}>
                      <span style={{color:"#64748B"}}>Total: <strong>{fmt$(sum(s.paymentsDue.map(p=>p.amount)),cur)}</strong></span>
                      <span style={{color:"#059669"}}>Pagado: <strong>{fmt$(sum(s.paymentsDue.filter(p=>p.paid).map(p=>p.amount)),cur)}</strong></span>
                      <span style={{color:"#EF4444"}}>Pendiente: <strong>{fmt$(sum(s.paymentsDue.filter(p=>!p.paid).map(p=>p.amount)),cur)}</strong></span>
                    </div>
                  )}
                </div>
              </div>
            );
          })}
        </div>
      )}

      {tab==="payments"&&(
        <div>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
            <span style={{fontSize:13,color:"#64748B"}}>Pagos recibidos del cliente.</span>
            <Btn size="sm" onClick={()=>arrAdd("paymentsReceived",newPay())}><Plus size={14}/>Registrar pago</Btn>
          </div>
          <div style={{background:"#ECFDF5",border:"1px solid #A7F3D0",borderRadius:10,padding:14,display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:12,marginBottom:16}}>
            {[
              {l:"Precio total",v:fmt$(r.salePrice,cur),c:"#1E293B"},
              {l:"Total cobrado",v:fmt$(received,cur),c:"#059669"},
              {l:"Saldo pendiente",v:fmt$(pending,cur),c:pending>0?"#D97706":"#059669"},
            ].map(x=>(
              <div key={x.l} style={{textAlign:"center"}}>
                <div style={{fontSize:10,fontWeight:700,textTransform:"uppercase",letterSpacing:.5,color:"#64748B",marginBottom:3}}>{x.l}</div>
                <div style={{fontSize:20,fontWeight:800,color:x.c}}>{x.v}</div>
              </div>
            ))}
          </div>
          {r.paymentsReceived.length===0&&<EmptyState icon={CreditCard} title="Sin cobros registrados" sub="Registrá los pagos recibidos."/>}
          {r.paymentsReceived.map((p,i)=>(
            <div key={p.id} style={{display:"grid",gridTemplateColumns:"1fr 1fr 2fr 1fr auto",gap:8,alignItems:"end",marginBottom:8}}>
              <Field label={i===0?"Fecha":""}>
                <input type="date" value={p.date||""} onChange={e=>arrUpd("paymentsReceived",p.id,x=>({...x,date:e.target.value}))} style={S.input}/>
              </Field>
              <Field label={i===0?"Forma de pago":""}>
                <select value={p.method||""} onChange={e=>arrUpd("paymentsReceived",p.id,x=>({...x,method:e.target.value}))} style={{...S.input,appearance:"auto"}}>
                  <option value="">—</option>
                  {PAY_METHODS.map(m=><option key={m} value={m}>{m}</option>)}
                </select>
              </Field>
              <Field label={i===0?"Notas":""}>
                <input type="text" value={p.notes||""} onChange={e=>arrUpd("paymentsReceived",p.id,x=>({...x,notes:e.target.value}))} style={S.input} placeholder="Obs…"/>
              </Field>
              <Field label={i===0?`Importe (${cur})`:""}>
                <input type="number" value={p.amount||""} onChange={e=>arrUpd("paymentsReceived",p.id,x=>({...x,amount:+e.target.value}))} style={S.input} placeholder="0.00"/>
              </Field>
              <Field label={i===0?" ":""}>
                <button onClick={()=>arrDel("paymentsReceived",p.id)} style={{...S.ghostBtn,color:"#EF4444",height:36}}><Trash2 size={13}/></button>
              </Field>
            </div>
          ))}
        </div>
      )}
    </Modal>
  );
}

// ══════════════════════════════════════════════════
//  RESERVATIONS PAGE
// ══════════════════════════════════════════════════
function ReservationsPage({data,update}){
  const {reservations,providers,settings}=data;
  const [search,setSearch]=useState("");
  const [filter,setFilter]=useState("all");
  const [modal,setModal]=useState(null);
  const [viewModal,setViewModal]=useState(null);
  const cur = settings.currency||'ARS';

  const filtered=useMemo(()=>{
    return reservations.filter(r=>{
      const q=search.toLowerCase();
      const matchQ=!q||(r.destination||"").toLowerCase().includes(q)||
        String(r.fileNumber).includes(q)||
        r.passengers.some(p=>p.name.toLowerCase().includes(q));
      const matchS=filter==="all"||r.status===filter;
      return matchQ&&matchS;
    }).sort((a,b)=>b.createdAt?.localeCompare(a.createdAt||"")||0);
  },[reservations,search,filter]);

  const saveRes=r=>{
    update(d=>{
      const idx=d.reservations.findIndex(x=>x.id===r.id);
      const updated=idx>=0
        ?d.reservations.map(x=>x.id===r.id?r:x)
        :[...d.reservations,r];
      return{...d,reservations:updated,nextFile:Math.max(d.nextFile,Number(r.fileNumber.replace(/^0+/,""))+1)};
    });
    setModal(null);
  };

  const delRes=id=>{
    if(!window.confirm("¿Eliminar esta reserva?")) return;
    update(d=>({...d,reservations:d.reservations.filter(r=>r.id!==id)}));
  };

  const newModal=()=>setModal({isNew:true,data:newRes(data.nextFile,settings.defaultCommission)});

  return(
    <div style={{padding:28,overflowY:"auto",flex:1}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20}}>
        <div>
          <h1 style={{margin:0,fontSize:22,fontWeight:800,color:"#0F172A"}}>Reservas</h1>
          <p style={{margin:"4px 0 0",fontSize:13,color:"#64748B"}}>{filtered.length} reserva(s) encontradas</p>
        </div>
        <div style={{display:"flex",gap:8}}>
          <Btn variant="success" onClick={()=>exportReservations(reservations,settings)}><Download size={14}/>Exportar Excel</Btn>
          <Btn onClick={newModal}><Plus size={14}/>Nueva reserva</Btn>
        </div>
      </div>

      <Card style={{marginBottom:16}}>
        <div style={{display:"flex",gap:10,alignItems:"center"}}>
          <div style={{flex:1,position:"relative"}}>
            <Search size={14} style={{position:"absolute",left:10,top:"50%",transform:"translateY(-50%)",color:"#94A3B8"}}/>
            <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Buscar por file, destino o pasajero…"
              style={{...S.input,paddingLeft:32}}/>
          </div>
          <select value={filter} onChange={e=>setFilter(e.target.value)} style={{...S.input,width:"auto",appearance:"auto"}}>
            <option value="all">Todos los estados</option>
            {Object.entries(STATUS).map(([v,s])=><option key={v} value={v}>{s.label}</option>)}
          </select>
        </div>
      </Card>

      {filtered.length===0
        ?<EmptyState icon={CalendarDays} title="Sin reservas" sub="Creá tu primera reserva." action={<Btn onClick={newModal}><Plus size={14}/>Nueva reserva</Btn>}/>
        :<Card style={{padding:0}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
            <thead>
              <tr style={{background:"#F8FAFC",borderBottom:"1px solid #E2E8F0"}}>
                {["File","Destino","Pasajero","Salida","Estado","Venta","Cobrado","Saldo","Acciones"].map(h=>(
                  <th key={h} style={{padding:"10px 12px",textAlign:"left",fontWeight:700,fontSize:11,color:"#64748B",textTransform:"uppercase"}}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.map(r=>{
                const rec=paidSum(r.paymentsReceived);
                const pend=(r.salePrice||0)-rec;
                return(
                  <tr key={r.id} style={{borderBottom:"1px solid #F1F5F9"}}
                    onMouseEnter={e=>e.currentTarget.style.background="#F8FAFC"}
                    onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
                    <td style={{padding:"10px 12px",fontWeight:700,color:"#2563EB"}}>#{r.fileNumber}</td>
                    <td style={{padding:"10px 12px",fontWeight:500}}>{r.destination||"—"}</td>
                    <td style={{padding:"10px 12px",color:"#64748B"}}>{r.passengers[0]?.name||"—"}</td>
                    <td style={{padding:"10px 12px",color:"#64748B"}}>{fmtDate(r.departureDate)}</td>
                    <td style={{padding:"10px 12px"}}><Badge status={r.status}/></td>
                    <td style={{padding:"10px 12px",fontWeight:600}}>{fmt$(r.salePrice,cur)}</td>
                    <td style={{padding:"10px 12px",color:"#059669",fontWeight:600}}>{fmt$(rec,cur)}</td>
                    <td style={{padding:"10px 12px",color:pend>0?"#D97706":"#059669",fontWeight:700}}>{fmt$(pend,cur)}</td>
                    <td style={{padding:"10px 12px"}}>
                      <div style={{display:"flex",gap:4}}>
                        <button title="Ver" onClick={()=>setViewModal(r)} style={S.ghostBtn}><Eye size={13}/></button>
                        <button title="Editar" onClick={()=>setModal({isNew:false,data:r})} style={S.ghostBtn}><Pencil size={13}/></button>
                        <button title="Voucher aéreo" onClick={()=>printVoucher(r,settings,providers,"aereo")} style={S.ghostBtn}><Plane size={13}/></button>
                        <button title="Voucher terrestre" onClick={()=>printVoucher(r,settings,providers,"terrestre")} style={S.ghostBtn}><Bus size={13}/></button>
                        <button title="Recibo" onClick={()=>printReceipt(r,settings,providers)} style={S.ghostBtn}><Printer size={13}/></button>
                        <button title="Eliminar" onClick={()=>delRes(r.id)} style={{...S.ghostBtn,color:"#EF4444",borderColor:"#FCA5A5"}}><Trash2 size={13}/></button>
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </Card>
      }

      {modal&&<ReservationModal
        initial={modal.isNew?null:modal.data}
        providers={providers}
        settings={settings}
        onSave={saveRes}
        onClose={()=>setModal(null)}/>}

      {viewModal&&(
        <Modal title={`Reserva File #${viewModal.fileNumber} — ${viewModal.destination||"Sin destino"}`} onClose={()=>setViewModal(null)} width={680}
          footer={<>
            <Btn variant="secondary" onClick={()=>{setViewModal(null);setModal({isNew:false,data:viewModal});}}>Editar</Btn>
            <Btn variant="secondary" onClick={()=>printVoucher(viewModal,settings,providers,"aereo")}><Plane size={13}/>Voucher Aéreo</Btn>
            <Btn variant="secondary" onClick={()=>printVoucher(viewModal,settings,providers,"terrestre")}><Bus size={13}/>Voucher Terrestre</Btn>
            <Btn onClick={()=>printReceipt(viewModal,settings,providers)}><Printer size={13}/>Imprimir Recibo</Btn>
          </>}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:16}}>
            {[
              {l:"Estado",v:<Badge status={viewModal.status}/>},
              {l:"Destino",v:viewModal.destination||"—"},
              {l:"Salida",v:fmtDate(viewModal.departureDate)},
              {l:"Regreso",v:fmtDate(viewModal.returnDate)},
              {l:"Precio venta",v:fmt$(viewModal.salePrice,cur)},
              {l:"Saldo pendiente",v:fmt$((viewModal.salePrice||0)-paidSum(viewModal.paymentsReceived),cur)},
            ].map(x=>(
              <div key={x.l}>
                <div style={S.label}>{x.l}</div>
                <div style={{fontSize:14,fontWeight:500,color:"#1E293B"}}>{x.v}</div>
              </div>
            ))}
          </div>
          {viewModal.description&&<div style={{background:"#F8FAFC",borderRadius:8,padding:12,marginBottom:16,fontSize:13,color:"#64748B"}}>{viewModal.description}</div>}
          <Divider label="Pasajeros"/>
          <div style={{display:"flex",flexWrap:"wrap",gap:8,marginBottom:16}}>
            {viewModal.passengers.map(p=>(
              <div key={p.id} style={{background:"#EFF6FF",color:"#1D4ED8",border:"1px solid #BFDBFE",borderRadius:8,padding:"6px 12px",fontSize:12,fontWeight:500}}>
                {p.name}{p.dni?" · "+p.dni:""}
              </div>
            ))}
          </div>
          <Divider label="Servicios"/>
          {viewModal.services.map(s=>{
            const prov=providers.find(p=>p.id===s.providerId);
            return(
              <div key={s.id} style={{background:"#F8FAFC",borderRadius:8,padding:12,marginBottom:8,display:"flex",justifyContent:"space-between"}}>
                <div>
                  <div style={{fontSize:12,fontWeight:700,color:"#2563EB"}}>{SVC_TYPES.find(t=>t.v===s.type)?.l||s.type}</div>
                  <div style={{fontSize:12,color:"#64748B"}}>{s.description||"—"}</div>
                  {prov&&<div style={{fontSize:11,color:"#94A3B8"}}>Prov: {prov.name}</div>}
                </div>
                <div style={{textAlign:"right"}}>
                  <div style={{fontSize:12,fontWeight:600}}>{fmt$(s.salePrice,cur)}</div>
                  {s.providerFileNumber&&<div style={{fontSize:11,color:"#94A3B8"}}>File: {s.providerFileNumber}</div>}
                </div>
              </div>
            );
          })}
        </Modal>
      )}
    </div>
  );
}

// ══════════════════════════════════════════════════
//  CLIENTS PAGE
// ══════════════════════════════════════════════════
function ClientsPage({data,update}){
  const {clients,reservations,settings}=data;
  const [search,setSearch]=useState("");
  const [modal,setModal]=useState(null);
  const cur = settings.currency||'ARS';

  const filtered=clients.filter(c=>{
    const q=search.toLowerCase();
    return !q||c.name.toLowerCase().includes(q)||(c.email||"").toLowerCase().includes(q)||(c.phone||"").includes(q);
  });

  const saveClient=c=>{
    update(d=>{
      const idx=d.clients.findIndex(x=>x.id===c.id);
      return{...d,clients:idx>=0?d.clients.map(x=>x.id===c.id?c:x):[...d.clients,c]};
    });
    setModal(null);
  };

  const delClient=id=>{
    if(!window.confirm("¿Eliminar este cliente?")) return;
    update(d=>({...d,clients:d.clients.filter(c=>c.id!==id)}));
  };

  const newClient=()=>setModal({id:genId(),name:"",email:"",phone:"",dni:"",address:"",notes:""});

  const clientReservations = c => reservations.filter(r =>
    r.passengers.some(p => p.name && c.name && p.name.toLowerCase()===c.name.toLowerCase())
  );

  return(
    <div style={{padding:28,overflowY:"auto",flex:1}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20}}>
        <div>
          <h1 style={{margin:0,fontSize:22,fontWeight:800,color:"#0F172A"}}>Clientes</h1>
          <p style={{margin:"4px 0 0",fontSize:13,color:"#64748B"}}>{filtered.length} cliente(s)</p>
        </div>
        <Btn onClick={newClient}><Plus size={14}/>Nuevo cliente</Btn>
      </div>
      <Card style={{marginBottom:16}}>
        <div style={{position:"relative"}}>
          <Search size={14} style={{position:"absolute",left:10,top:"50%",transform:"translateY(-50%)",color:"#94A3B8"}}/>
          <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Buscar clientes…" style={{...S.input,paddingLeft:32}}/>
        </div>
      </Card>
      {filtered.length===0
        ?<EmptyState icon={Users} title="Sin clientes" sub="Agregá tu primer cliente." action={<Btn onClick={newClient}><Plus size={14}/>Nuevo cliente</Btn>}/>
        :<Card style={{padding:0}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
            <thead>
              <tr style={{background:"#F8FAFC",borderBottom:"1px solid #E2E8F0"}}>
                {["Nombre","Email","Teléfono","DNI","Reservas","Acciones"].map(h=>(
                  <th key={h} style={{padding:"10px 12px",textAlign:"left",fontWeight:700,fontSize:11,color:"#64748B",textTransform:"uppercase"}}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.map(c=>{
                const cRes = clientReservations(c);
                return(
                  <tr key={c.id} style={{borderBottom:"1px solid #F1F5F9"}}
                    onMouseEnter={e=>e.currentTarget.style.background="#F8FAFC"}
                    onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
                    <td style={{padding:"10px 12px",fontWeight:600,color:"#0F172A"}}>{c.name}</td>
                    <td style={{padding:"10px 12px",color:"#64748B"}}>{c.email||"—"}</td>
                    <td style={{padding:"10px 12px",color:"#64748B"}}>{c.phone||"—"}</td>
                    <td style={{padding:"10px 12px",color:"#64748B"}}>{c.dni||"—"}</td>
                    <td style={{padding:"10px 12px"}}>
                      <span style={{background:"#EFF6FF",color:"#1D4ED8",padding:"2px 8px",borderRadius:12,fontSize:11,fontWeight:700}}>{cRes.length}</span>
                    </td>
                    <td style={{padding:"10px 12px"}}>
                      <div style={{display:"flex",gap:4}}>
                        <button title="Editar" onClick={()=>setModal(c)} style={S.ghostBtn}><Pencil size={13}/></button>
                        <button title="Estado de cuenta" onClick={()=>exportClientStatement(c,reservations,settings)} style={{...S.ghostBtn,color:"#059669",borderColor:"#A7F3D0"}}>
                          <Download size={13}/>Excel
                        </button>
                        <button title="Eliminar" onClick={()=>delClient(c.id)} style={{...S.ghostBtn,color:"#EF4444",borderColor:"#FCA5A5"}}><Trash2 size={13}/></button>
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </Card>
      }
      {modal&&(
        <Modal title={modal.name?"Editar cliente":"Nuevo cliente"} onClose={()=>setModal(null)}
          footer={<>
            <Btn variant="secondary" onClick={()=>setModal(null)}>Cancelar</Btn>
            <Btn onClick={()=>saveClient(modal)}>Guardar</Btn>
          </>}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
            <Inp label="Nombre completo" value={modal.name} onChange={v=>setModal(p=>({...p,name:v}))} col={2} style={{gridColumn:"span 2"}}/>
            <Inp label="Email" value={modal.email} onChange={v=>setModal(p=>({...p,email:v}))} type="email"/>
            <Inp label="Teléfono" value={modal.phone} onChange={v=>setModal(p=>({...p,phone:v}))}/>
            <Inp label="DNI / Pasaporte" value={modal.dni} onChange={v=>setModal(p=>({...p,dni:v}))}/>
            <Inp label="Dirección" value={modal.address} onChange={v=>setModal(p=>({...p,address:v}))}/>
            <Txta label="Notas" value={modal.notes} onChange={v=>setModal(p=>({...p,notes:v}))} col={2} style={{gridColumn:"span 2"}}/>
          </div>
        </Modal>
      )}
    </div>
  );
}

// ══════════════════════════════════════════════════
//  PROVIDERS PAGE
// ══════════════════════════════════════════════════
function ProvidersPage({data,update}){
  const {providers,reservations,settings}=data;
  const [search,setSearch]=useState("");
  const [modal,setModal]=useState(null);
  const cur = settings.currency||'ARS';

  const filtered=providers.filter(p=>{
    const q=search.toLowerCase();
    return !q||p.name.toLowerCase().includes(q)||(p.category||"").toLowerCase().includes(q);
  });

  const saveProv=p=>{
    update(d=>{
      const idx=d.providers.findIndex(x=>x.id===p.id);
      return{...d,providers:idx>=0?d.providers.map(x=>x.id===p.id?p:x):[...d.providers,p]};
    });
    setModal(null);
  };

  const delProv=id=>{
    if(!window.confirm("¿Eliminar este proveedor?")) return;
    update(d=>({...d,providers:d.providers.filter(p=>p.id!==id)}));
  };

  const newProv=()=>setModal({id:genId(),name:"",category:"",email:"",phone:"",address:"",notes:"",cuit:""});

  const provDebt = p => {
    let total=0, pagado=0;
    reservations.forEach(r => r.services.filter(s=>s.providerId===p.id).forEach(s=>{
      total += s.costPrice||0;
      pagado += sum((s.paymentsDue||[]).filter(x=>x.paid).map(x=>x.amount||0));
    }));
    return {total, pagado, pendiente: total-pagado};
  };

  return(
    <div style={{padding:28,overflowY:"auto",flex:1}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20}}>
        <div>
          <h1 style={{margin:0,fontSize:22,fontWeight:800,color:"#0F172A"}}>Proveedores</h1>
          <p style={{margin:"4px 0 0",fontSize:13,color:"#64748B"}}>{filtered.length} proveedor(es)</p>
        </div>
        <Btn onClick={newProv}><Plus size={14}/>Nuevo proveedor</Btn>
      </div>
      <Card style={{marginBottom:16}}>
        <div style={{position:"relative"}}>
          <Search size={14} style={{position:"absolute",left:10,top:"50%",transform:"translateY(-50%)",color:"#94A3B8"}}/>
          <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Buscar proveedores…" style={{...S.input,paddingLeft:32}}/>
        </div>
      </Card>
      {filtered.length===0
        ?<EmptyState icon={Building2} title="Sin proveedores" sub="Agregá tu primer proveedor." action={<Btn onClick={newProv}><Plus size={14}/>Nuevo proveedor</Btn>}/>
        :<Card style={{padding:0}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
            <thead>
              <tr style={{background:"#F8FAFC",borderBottom:"1px solid #E2E8F0"}}>
                {["Nombre","Categoría","Contacto","Deuda Total","Pagado","Pendiente","Acciones"].map(h=>(
                  <th key={h} style={{padding:"10px 12px",textAlign:"left",fontWeight:700,fontSize:11,color:"#64748B",textTransform:"uppercase"}}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.map(p=>{
                const d = provDebt(p);
                return(
                  <tr key={p.id} style={{borderBottom:"1px solid #F1F5F9"}}
                    onMouseEnter={e=>e.currentTarget.style.background="#F8FAFC"}
                    onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
                    <td style={{padding:"10px 12px",fontWeight:600,color:"#0F172A"}}>{p.name}</td>
                    <td style={{padding:"10px 12px"}}>
                      {p.category&&<span style={{background:"#F1F5F9",color:"#475569",padding:"2px 8px",borderRadius:12,fontSize:11}}>{p.category}</span>}
                    </td>
                    <td style={{padding:"10px 12px",color:"#64748B",fontSize:12}}>{p.email||p.phone||"—"}</td>
                    <td style={{padding:"10px 12px",fontWeight:600}}>{fmt$(d.total,cur)}</td>
                    <td style={{padding:"10px 12px",color:"#059669",fontWeight:600}}>{fmt$(d.pagado,cur)}</td>
                    <td style={{padding:"10px 12px",color:d.pendiente>0?"#EF4444":"#059669",fontWeight:700}}>{fmt$(d.pendiente,cur)}</td>
                    <td style={{padding:"10px 12px"}}>
                      <div style={{display:"flex",gap:4}}>
                        <button title="Editar" onClick={()=>setModal(p)} style={S.ghostBtn}><Pencil size={13}/></button>
                        <button title="Estado de cuenta" onClick={()=>exportProviderStatement(p,reservations,settings)} style={{...S.ghostBtn,color:"#059669",borderColor:"#A7F3D0"}}>
                          <Download size={13}/>Excel
                        </button>
                        <button title="Eliminar" onClick={()=>delProv(p.id)} style={{...S.ghostBtn,color:"#EF4444",borderColor:"#FCA5A5"}}><Trash2 size={13}/></button>
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </Card>
      }
      {modal&&(
        <Modal title={modal.name?"Editar proveedor":"Nuevo proveedor"} onClose={()=>setModal(null)}
          footer={<>
            <Btn variant="secondary" onClick={()=>setModal(null)}>Cancelar</Btn>
            <Btn onClick={()=>saveProv(modal)}>Guardar</Btn>
          </>}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
            <Inp label="Nombre" value={modal.name} onChange={v=>setModal(p=>({...p,name:v}))} col={2} style={{gridColumn:"span 2"}}/>
            <Sel label="Categoría" value={modal.category} onChange={v=>setModal(p=>({...p,category:v}))} options={PROV_CATS}/>
            <Inp label="CUIT" value={modal.cuit} onChange={v=>setModal(p=>({...p,cuit:v}))}/>
            <Inp label="Email" value={modal.email} onChange={v=>setModal(p=>({...p,email:v}))} type="email"/>
            <Inp label="Teléfono" value={modal.phone} onChange={v=>setModal(p=>({...p,phone:v}))}/>
            <Inp label="Dirección" value={modal.address} onChange={v=>setModal(p=>({...p,address:v}))} col={2} style={{gridColumn:"span 2"}}/>
            <Txta label="Notas" value={modal.notes} onChange={v=>setModal(p=>({...p,notes:v}))} col={2} style={{gridColumn:"span 2"}}/>
          </div>
        </Modal>
      )}
    </div>
  );
}

// ══════════════════════════════════════════════════
//  COMMISSIONS PAGE
// ══════════════════════════════════════════════════
function CommissionsPage({data}){
  const {reservations,settings}=data;
  const [tab,setTab]=useState("summary");
  const cur = settings.currency||'ARS';

  const notCancelled=reservations.filter(r=>r.status!=="cancelada");
  const byStatus={};
  Object.keys(STATUS).forEach(k=>{ byStatus[k]=reservations.filter(r=>r.status===k).length; });

  const totalSale=sum(notCancelled.map(r=>r.salePrice||0));
  const totalCost=sum(notCancelled.map(r=>sum(r.services.map(s=>s.costPrice||0))));
  const totalComm=totalSale-totalCost;
  const totalReceived=sum(reservations.map(r=>paidSum(r.paymentsReceived)));

  return(
    <div style={{padding:28,overflowY:"auto",flex:1}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20}}>
        <div>
          <h1 style={{margin:0,fontSize:22,fontWeight:800,color:"#0F172A"}}>Comisiones</h1>
          <p style={{margin:"4px 0 0",fontSize:13,color:"#64748B"}}>Análisis financiero de la agencia</p>
        </div>
        <Btn variant="success" onClick={()=>exportCommissions(reservations,settings)}><Download size={14}/>Exportar Excel</Btn>
      </div>
      <Grid cols={4} gap={14} style={{marginBottom:20}}>
        {[
          {l:"Ingresos totales",v:fmt$(totalSale,cur),c:"#059669",bg:"#ECFDF5"},
          {l:"Costos totales",v:fmt$(totalCost,cur),c:"#EF4444",bg:"#FEF2F2"},
          {l:"Comisión neta",v:fmt$(totalComm,cur),c:"#7C3AED",bg:"#F5F3FF"},
          {l:"Total cobrado",v:fmt$(totalReceived,cur),c:"#2563EB",bg:"#EFF6FF"},
        ].map(x=>(
          <div key={x.l} style={{background:x.bg,borderRadius:12,padding:16}}>
            <div style={{fontSize:11,fontWeight:700,color:x.c,textTransform:"uppercase",letterSpacing:.5,marginBottom:8}}>{x.l}</div>
            <div style={{fontSize:22,fontWeight:800,color:x.c}}>{x.v}</div>
          </div>
        ))}
      </Grid>
      <Tabs tabs={[{id:"summary",label:"Resumen"},{id:"detail",label:"Detalle"},{id:"pending",label:"Saldos pendientes"}]} active={tab} setActive={setTab}/>
      {tab==="summary"&&(
        <Card>
          <h3 style={S.sectionTitle}>Distribución por estado</h3>
          <Grid cols={3} gap={14}>
            {Object.entries(STATUS).map(([k,s])=>(
              <div key={k} style={{background:s.bg,borderRadius:10,padding:16,textAlign:"center"}}>
                <div style={{fontSize:11,fontWeight:700,textTransform:"uppercase",letterSpacing:.5,color:s.c,marginBottom:8}}>{s.label}</div>
                <div style={{fontSize:28,fontWeight:800,color:s.c}}>{byStatus[k]||0}</div>
                <div style={{fontSize:11,color:"#94A3B8",marginTop:4}}>reservas</div>
              </div>
            ))}
          </Grid>
        </Card>
      )}
      {tab==="detail"&&(
        <Card style={{padding:0}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
            <thead>
              <tr style={{background:"#F8FAFC",borderBottom:"1px solid #E2E8F0"}}>
                {["File","Destino","Pasajero","Estado","Venta","Costo","Comisión","% Comisión"].map(h=>(
                  <th key={h} style={{padding:"9px 12px",textAlign:"left",fontWeight:700,fontSize:11,color:"#64748B",textTransform:"uppercase"}}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {notCancelled.map(r=>{
                const costo=sum(r.services.map(s=>s.costPrice||0));
                const comision=(r.salePrice||0)-costo;
                const pct=r.salePrice?((comision/r.salePrice)*100).toFixed(1):0;
                return(
                  <tr key={r.id} style={{borderBottom:"1px solid #F1F5F9"}}
                    onMouseEnter={e=>e.currentTarget.style.background="#F8FAFC"}
                    onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
                    <td style={{padding:"9px 12px",fontWeight:700,color:"#2563EB"}}>#{r.fileNumber}</td>
                    <td style={{padding:"9px 12px"}}>{r.destination||"—"}</td>
                    <td style={{padding:"9px 12px",color:"#64748B"}}>{r.passengers[0]?.name||"—"}</td>
                    <td style={{padding:"9px 12px"}}><Badge status={r.status}/></td>
                    <td style={{padding:"9px 12px",fontWeight:600}}>{fmt$(r.salePrice,cur)}</td>
                    <td style={{padding:"9px 12px",color:"#EF4444"}}>{fmt$(costo,cur)}</td>
                    <td style={{padding:"9px 12px",color:"#7C3AED",fontWeight:700}}>{fmt$(comision,cur)}</td>
                    <td style={{padding:"9px 12px"}}><span style={{background:"#F5F3FF",color:"#7C3AED",padding:"2px 8px",borderRadius:12,fontSize:11,fontWeight:700}}>{pct}%</span></td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </Card>
      )}
      {tab==="pending"&&(
        <Card style={{padding:0}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
            <thead>
              <tr style={{background:"#F8FAFC",borderBottom:"1px solid #E2E8F0"}}>
                {["File","Pasajero","Destino","Salida","Precio venta","Cobrado","Saldo"].map(h=>(
                  <th key={h} style={{padding:"9px 12px",textAlign:"left",fontWeight:700,fontSize:11,color:"#64748B",textTransform:"uppercase"}}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {reservations.filter(r=>r.status!=="cancelada"&&(r.salePrice||0)-paidSum(r.paymentsReceived)>0)
                .sort((a,b)=>a.departureDate?.localeCompare(b.departureDate||"")||0)
                .map(r=>{
                  const rec=paidSum(r.paymentsReceived);
                  const pend=(r.salePrice||0)-rec;
                  return(
                    <tr key={r.id} style={{borderBottom:"1px solid #F1F5F9"}}
                      onMouseEnter={e=>e.currentTarget.style.background="#F8FAFC"}
                      onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
                      <td style={{padding:"9px 12px",fontWeight:700,color:"#2563EB"}}>#{r.fileNumber}</td>
                      <td style={{padding:"9px 12px"}}>{r.passengers[0]?.name||"—"}</td>
                      <td style={{padding:"9px 12px"}}>{r.destination||"—"}</td>
                      <td style={{padding:"9px 12px",color:"#64748B"}}>{fmtDate(r.departureDate)}</td>
                      <td style={{padding:"9px 12px",fontWeight:600}}>{fmt$(r.salePrice,cur)}</td>
                      <td style={{padding:"9px 12px",color:"#059669",fontWeight:600}}>{fmt$(rec,cur)}</td>
                      <td style={{padding:"9px 12px",color:"#D97706",fontWeight:700}}>{fmt$(pend,cur)}</td>
                    </tr>
                  );
                })
              }
            </tbody>
          </table>
        </Card>
      )}
    </div>
  );
}

// ══════════════════════════════════════════════════
//  REPORTS PAGE (NEW)
// ══════════════════════════════════════════════════
function ReportsPage({data}){
  const {reservations,clients,providers,settings}=data;
  const [selectedClient,setSelectedClient]=useState("");
  const [selectedProvider,setSelectedProvider]=useState("");
  const [dateFrom,setDateFrom]=useState("");
  const [dateTo,setDateTo]=useState("");
  const cur = settings.currency||'ARS';

  const filteredRes = useMemo(()=>{
    return reservations.filter(r=>{
      const af = !dateFrom || (r.departureDate&&r.departureDate>=dateFrom);
      const at = !dateTo || (r.departureDate&&r.departureDate<=dateTo);
      return af&&at;
    });
  },[reservations,dateFrom,dateTo]);

  const totalVenta = sum(filteredRes.filter(r=>r.status!=='cancelada').map(r=>r.salePrice||0));
  const totalCobrado = sum(filteredRes.map(r=>paidSum(r.paymentsReceived)));
  const totalCosto = sum(filteredRes.filter(r=>r.status!=='cancelada').map(r=>sum(r.services.map(s=>s.costPrice||0))));
  const totalComision = totalVenta - totalCosto;

  const handleExportFiltered = () => {
    exportReservations(filteredRes, settings);
  };

  const clientObj = clients.find(c=>c.id===selectedClient);
  const provObj = providers.find(p=>p.id===selectedProvider);

  return(
    <div style={{padding:28,overflowY:"auto",flex:1}}>
      <div style={{marginBottom:24}}>
        <h1 style={{margin:0,fontSize:22,fontWeight:800,color:"#0F172A"}}>Reportes</h1>
        <p style={{margin:"4px 0 0",fontSize:13,color:"#64748B"}}>Generá estados de cuenta y reportes descargables en Excel</p>
      </div>

      {/* Summary Cards */}
      <Card style={{marginBottom:20}}>
        <h3 style={{...S.sectionTitle,marginBottom:16}}>Filtro por período</h3>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr auto",gap:12,alignItems:"end"}}>
          <Field label="Fecha desde">
            <input type="date" value={dateFrom} onChange={e=>setDateFrom(e.target.value)} style={S.input}/>
          </Field>
          <Field label="Fecha hasta">
            <input type="date" value={dateTo} onChange={e=>setDateTo(e.target.value)} style={S.input}/>
          </Field>
          <div style={{display:"flex",gap:8}}>
            <Btn variant="secondary" onClick={()=>{setDateFrom("");setDateTo("");}}>Limpiar</Btn>
            <Btn variant="success" onClick={handleExportFiltered}><Download size={14}/>Exportar Excel</Btn>
          </div>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12,marginTop:16}}>
          {[
            {l:"Reservas",v:filteredRes.length,c:"#2563EB",bg:"#EFF6FF"},
            {l:"Ventas",v:fmt$(totalVenta,cur),c:"#059669",bg:"#ECFDF5"},
            {l:"Cobrado",v:fmt$(totalCobrado,cur),c:"#7C3AED",bg:"#F5F3FF"},
            {l:"Comisión",v:fmt$(totalComision,cur),c:"#D97706",bg:"#FFFBEB"},
          ].map(x=>(
            <div key={x.l} style={{background:x.bg,borderRadius:10,padding:14,textAlign:"center"}}>
              <div style={{fontSize:11,fontWeight:700,color:x.c,textTransform:"uppercase",marginBottom:6}}>{x.l}</div>
              <div style={{fontSize:18,fontWeight:800,color:x.c}}>{x.v}</div>
            </div>
          ))}
        </div>
      </Card>

      <Grid cols={2} gap={16}>
        {/* Estado de cuenta clientes */}
        <Card>
          <h3 style={{...S.sectionTitle,marginBottom:16}}><Users size={16} style={{color:"#2563EB"}}/>Estado de cuenta — Clientes</h3>
          <p style={{fontSize:12,color:"#64748B",marginBottom:12}}>Descargá el historial completo de reservas y pagos de un cliente en Excel.</p>
          <Field label="Seleccioná un cliente">
            <select value={selectedClient} onChange={e=>setSelectedClient(e.target.value)} style={{...S.input,appearance:"auto",marginBottom:12}}>
              <option value="">— Elegir cliente —</option>
              {clients.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}
            </select>
          </Field>
          {clientObj&&(
            <div style={{background:"#F8FAFC",borderRadius:8,padding:12,marginBottom:12,fontSize:12}}>
              <div style={{fontWeight:700,color:"#0F172A",marginBottom:4}}>{clientObj.name}</div>
              {clientObj.email&&<div style={{color:"#64748B"}}>{clientObj.email}</div>}
              {clientObj.phone&&<div style={{color:"#64748B"}}>{clientObj.phone}</div>}
            </div>
          )}
          <Btn variant="success" disabled={!selectedClient} onClick={()=>clientObj&&exportClientStatement(clientObj,reservations,settings)}>
            <Download size={14}/>Descargar estado de cuenta
          </Btn>
        </Card>

        {/* Estado de cuenta proveedores */}
        <Card>
          <h3 style={{...S.sectionTitle,marginBottom:16}}><Building2 size={16} style={{color:"#7C3AED"}}/>Estado de cuenta — Proveedores</h3>
          <p style={{fontSize:12,color:"#64748B",marginBottom:12}}>Descargá el detalle de servicios y pagos realizados a un proveedor en Excel.</p>
          <Field label="Seleccioná un proveedor">
            <select value={selectedProvider} onChange={e=>setSelectedProvider(e.target.value)} style={{...S.input,appearance:"auto",marginBottom:12}}>
              <option value="">— Elegir proveedor —</option>
              {providers.map(p=><option key={p.id} value={p.id}>{p.name}</option>)}
            </select>
          </Field>
          {provObj&&(
            <div style={{background:"#F8FAFC",borderRadius:8,padding:12,marginBottom:12,fontSize:12}}>
              <div style={{fontWeight:700,color:"#0F172A",marginBottom:4}}>{provObj.name}</div>
              {provObj.category&&<div style={{color:"#64748B"}}>{provObj.category}</div>}
              {provObj.email&&<div style={{color:"#64748B"}}>{provObj.email}</div>}
            </div>
          )}
          <Btn style={{background:"#7C3AED"}} disabled={!selectedProvider} onClick={()=>provObj&&exportProviderStatement(provObj,reservations,settings)}>
            <Download size={14}/>Descargar estado de cuenta
          </Btn>
        </Card>
      </Grid>

      <Card style={{marginTop:16}}>
        <h3 style={{...S.sectionTitle,marginBottom:16}}><BarChart2 size={16} style={{color:"#D97706"}}/>Reportes generales</h3>
        <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:12}}>
          {[
            {title:"Todas las reservas",desc:"Listado completo con fechas, estados, ventas y saldos.",color:"#2563EB",onClick:()=>exportReservations(reservations,settings)},
            {title:"Comisiones",desc:"Detalle de comisión neta por reserva (excluye canceladas).",color:"#7C3AED",onClick:()=>exportCommissions(reservations,settings)},
            {title:"Saldos pendientes",desc:"Reservas con saldo por cobrar ordenadas por fecha.",color:"#D97706",onClick:()=>{
              const pending = reservations.filter(r=>r.status!=='cancelada'&&(r.salePrice||0)-paidSum(r.paymentsReceived)>0);
              exportReservations(pending,settings);
            }},
          ].map(x=>(
            <div key={x.title} style={{border:"1px solid #E2E8F0",borderRadius:10,padding:16}}>
              <div style={{fontSize:13,fontWeight:700,color:x.color,marginBottom:6}}>{x.title}</div>
              <div style={{fontSize:12,color:"#64748B",marginBottom:12}}>{x.desc}</div>
              <Btn variant="secondary" size="sm" onClick={x.onClick}><Download size={12}/>Descargar Excel</Btn>
            </div>
          ))}
        </div>
      </Card>
    </div>
  );
}

// ══════════════════════════════════════════════════
//  SETTINGS PAGE
// ══════════════════════════════════════════════════
function SettingsPage({data,update}){
  const [s,setS]=useState(data.settings);
  const [saved,setSaved]=useState(false);
  const fileRef=useRef();

  const save=()=>{
    update(d=>({...d,settings:s}));
    setSaved(true);
    setTimeout(()=>setSaved(false),2000);
  };

  const handleLogo=e=>{
    const file=e.target.files[0];
    if(!file) return;
    const reader=new FileReader();
    reader.onload=ev=>setS(p=>({...p,logo:ev.target.result}));
    reader.readAsDataURL(file);
  };

  return(
    <div style={{padding:28,maxWidth:700,overflowY:"auto",flex:1}}>
      <div style={{marginBottom:24}}>
        <h1 style={{margin:0,fontSize:22,fontWeight:800,color:"#0F172A"}}>Configuración</h1>
        <p style={{margin:"4px 0 0",fontSize:13,color:"#64748B"}}>Datos de tu agencia y personalización de documentos</p>
      </div>

      <Card style={{marginBottom:16}}>
        <h3 style={{...S.sectionTitle,marginBottom:20}}>Datos de la agencia</h3>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
          <Inp label="Nombre de la agencia" value={s.agencyName} onChange={v=>setS(p=>({...p,agencyName:v}))} col={2} style={{gridColumn:"span 2"}}/>
          <Inp label="Dirección" value={s.address} onChange={v=>setS(p=>({...p,address:v}))} col={2} style={{gridColumn:"span 2"}}/>
          <Inp label="Teléfono" value={s.phone} onChange={v=>setS(p=>({...p,phone:v}))}/>
          <Inp label="Email" value={s.email} onChange={v=>setS(p=>({...p,email:v}))} type="email"/>
          <Inp label="CUIT" value={s.cuit} onChange={v=>setS(p=>({...p,cuit:v}))}/>
          <Field label="% Comisión por defecto">
            <input type="number" value={s.defaultCommission||""} onChange={e=>setS(p=>({...p,defaultCommission:+e.target.value}))} style={S.input} placeholder="15"/>
          </Field>
          <Sel label="Moneda principal" value={s.currency||"ARS"} onChange={v=>setS(p=>({...p,currency:v}))} options={CURRENCIES}/>
        </div>
      </Card>

      <Card style={{marginBottom:16}}>
        <h3 style={{...S.sectionTitle,marginBottom:16}}>Logo de la agencia</h3>
        <div style={{display:"flex",alignItems:"center",gap:16,marginBottom:12}}>
          {s.logo
            ?<img src={s.logo} style={{height:60,maxWidth:160,objectFit:"contain",border:"1px solid #E2E8F0",borderRadius:8,padding:8}} alt="Logo"/>
            :<div style={{width:80,height:60,background:"#F1F5F9",borderRadius:8,display:"flex",alignItems:"center",justifyContent:"center",color:"#94A3B8",fontSize:12}}>Sin logo</div>
          }
          <div>
            <Btn variant="secondary" onClick={()=>fileRef.current.click()}><Upload size={14}/>Subir logo</Btn>
            {s.logo&&<Btn variant="ghost" style={{marginLeft:8}} onClick={()=>setS(p=>({...p,logo:null}))}>Quitar</Btn>}
            <input ref={fileRef} type="file" accept="image/*" onChange={handleLogo} style={{display:"none"}}/>
            <p style={{fontSize:11,color:"#94A3B8",marginTop:6}}>PNG, JPG o SVG. Recomendado: fondo transparente.</p>
          </div>
        </div>
      </Card>

      <Card style={{marginBottom:16}}>
        <h3 style={{...S.sectionTitle,marginBottom:16}}>Personalización de documentos</h3>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
          <Field label="Color de vouchers">
            <div style={{display:"flex",gap:8,alignItems:"center"}}>
              <input type="color" value={s.voucherColor||"#1a56db"} onChange={e=>setS(p=>({...p,voucherColor:e.target.value}))}
                style={{width:40,height:36,border:"1px solid #E2E8F0",borderRadius:6,cursor:"pointer",padding:2}}/>
              <span style={{fontSize:12,color:"#64748B"}}>Color del encabezado de vouchers</span>
            </div>
          </Field>
          <Field label="Color de recibos">
            <div style={{display:"flex",gap:8,alignItems:"center"}}>
              <input type="color" value={s.receiptColor||"#1a56db"} onChange={e=>setS(p=>({...p,receiptColor:e.target.value}))}
                style={{width:40,height:36,border:"1px solid #E2E8F0",borderRadius:6,cursor:"pointer",padding:2}}/>
              <span style={{fontSize:12,color:"#64748B"}}>Color de líneas y totales en recibos</span>
            </div>
          </Field>
          <Txta label="Nota en encabezado de recibos" value={s.headerNote} onChange={v=>setS(p=>({...p,headerNote:v}))} rows={2}
            placeholder="Ej: Condiciones de cancelación, políticas..." col={2} style={{gridColumn:"span 2"}}/>
          <Inp label="Texto del pie de página" value={s.footerText} onChange={v=>setS(p=>({...p,footerText:v}))}
            placeholder="Gracias por elegirnos." col={2} style={{gridColumn:"span 2"}}/>
        </div>
        <div style={{marginTop:12,background:"#F8FAFC",borderRadius:8,padding:12}}>
          <div style={{fontSize:11,fontWeight:700,color:"#475569",marginBottom:6,textTransform:"uppercase",letterSpacing:.5}}>Vista previa del encabezado</div>
          <div style={{background:s.voucherColor||"#1a56db",color:"white",padding:"12px 16px",borderRadius:8,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <div>
              {s.logo
                ?<img src={s.logo} style={{height:28,objectFit:"contain",background:"white",padding:2,borderRadius:4}} alt="Logo"/>
                :<div style={{fontWeight:700,fontSize:14}}>{s.agencyName}</div>
              }
              <div style={{fontSize:10,opacity:.8,marginTop:2}}>Voucher Aéreo / Terrestre</div>
            </div>
            <div style={{fontSize:11,textAlign:"right",opacity:.9}}>File N° 00001<br/>Emitido: {fmtDate(today())}</div>
          </div>
        </div>
      </Card>

      <Card style={{marginBottom:16}}>
        <h3 style={{...S.sectionTitle,marginBottom:14}}>Estadísticas del sistema</h3>
        <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10}}>
          {[
            {l:"Total reservas",v:data.reservations.length},
            {l:"Total clientes",v:data.clients.length},
            {l:"Total proveedores",v:data.providers.length},
          ].map(x=>(
            <div key={x.l} style={{background:"#F8FAFC",borderRadius:8,padding:12,textAlign:"center"}}>
              <div style={{fontSize:10,fontWeight:700,textTransform:"uppercase",letterSpacing:.5,color:"#94A3B8",marginBottom:4}}>{x.l}</div>
              <div style={{fontSize:22,fontWeight:800,color:"#1E293B"}}>{x.v}</div>
            </div>
          ))}
        </div>
      </Card>

      <div style={{display:"flex",alignItems:"center",gap:12}}>
        <Btn size="lg" onClick={save}>Guardar configuración</Btn>
        {saved&&<span style={{color:"#059669",fontSize:13,fontWeight:600,display:"flex",alignItems:"center",gap:4}}><Check size={14}/>¡Guardado!</span>}
      </div>
    </div>
  );
}

// ══════════════════════════════════════════════════
//  APP ROOT
// ══════════════════════════════════════════════════
export default function App(){
  const [db,setDb]=useState(null);
  const [section,setSection]=useState("dashboard");

  useEffect(()=>{loadDB().then(setDb);},[]);

  const update=useCallback(fn=>{
    setDb(prev=>{
      const next=typeof fn==="function"?fn(prev):{...prev,...fn};
      saveDB(next);
      return next;
    });
  },[]);

  if(!db)return(
    <div style={{display:"flex",alignItems:"center",justifyContent:"center",height:"100vh",background:"#0F172A",color:"#94A3B8",fontFamily:"system-ui",gap:10,fontSize:14}}>
      <span style={{width:20,height:20,border:"2px solid #3B82F6",borderTopColor:"transparent",borderRadius:"50%",display:"inline-block",animation:"spin 1s linear infinite"}}/>
      Cargando TravelManager…
      <style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style>
    </div>
  );

  return(
    <div style={{display:"flex",height:"100vh",fontFamily:"'Segoe UI',system-ui,sans-serif",background:"#F1F5F9",overflow:"hidden"}}>
      <Sidebar section={section} setSection={setSection} settings={db.settings}/>
      <main style={{flex:1,overflow:"hidden",display:"flex",flexDirection:"column"}}>
        {section==="dashboard"&&<Dashboard data={db} setSection={setSection}/>}
        {section==="reservations"&&<ReservationsPage data={db} update={update}/>}
        {section==="clients"&&<ClientsPage data={db} update={update}/>}
        {section==="providers"&&<ProvidersPage data={db} update={update}/>}
        {section==="commissions"&&<CommissionsPage data={db}/>}
        {section==="reports"&&<ReportsPage data={db}/>}
        {section==="settings"&&<SettingsPage data={db} update={update}/>}
      </main>
    </div>
  );
}

import { useState, useEffect, useCallback, useMemo, useRef } from "react";
import { LayoutDashboard, CalendarDays, Users, Building2, TrendingUp, FileText, Settings, Plus, Pencil, Trash2, Printer, Eye, X, Check, AlertCircle, Search, Download, CreditCard, Plane, Hotel, Bus, Map, Shield, Ship, Package, FileCheck, Upload, BarChart2, ClipboardList, ArrowRight } from "lucide-react";

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
  cotizacion:{label:"Cotización",c:"#6366F1",bg:"#EEF2FF"},
  confirmada:{label:"Confirmada",c:"#0EA5E9",bg:"#F0F9FF"},
  en_viaje:{label:"En viaje",c:"#10B981",bg:"#ECFDF5"},
  completada:{label:"Completada",c:"#8B5CF6",bg:"#F5F3FF"},
  cancelada:{label:"Cancelada",c:"#EF4444",bg:"#FEF2F2"},
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
const REGIMENES = ["Solo habitación","BB - Bed & Breakfast","HB - Media Pensión","FB - Pensión Completa","AI - Todo Incluido"];

// ── FIREBASE CONFIG ───────────────────────────────
const FB_CONFIG = {
  apiKey: "AIzaSyCNOXMel7aiTJCXK9nbX2Q8aKUf_o96WE8",
  authDomain: "travel-manager-b413f.firebaseapp.com",
  databaseURL: "https://travel-manager-b413f-default-rtdb.firebaseio.com",
  projectId: "travel-manager-b413f",
  storageBucket: "travel-manager-b413f.firebasestorage.app",
  messagingSenderId: "678056777556",
  appId: "1:678056777556:web:fc0f052c34860950cb5ac1"
};

const DEFAULT_DOC_CONFIG = {
  recibo:{showServiceDetail:false,showProviders:false,showProviderRef:false,showServicePrices:false,showPaymentHistory:true,showTotals:true},
  voucherAereo:{showPassengerDNI:true,showFlightNumbers:true,showBaggageInfo:true,showProviderRef:true},
  voucherTerrestre:{showPassengerDNI:true,showRoomType:true,showRegimen:true,showHotelPhone:true,showHotelAddress:true,showProviderRef:false,showImportantInfo:true},
};
const INIT = {
  settings:{agencyName:"Mi Agencia de Viajes",address:"",phone:"",email:"",cuit:"",emergencyPhone:"",logo:null,defaultCommission:15,currency:"ARS",voucherColor:"#1a56db",receiptColor:"#1a56db",footerText:"Gracias por elegirnos.",headerNote:"",docConfig:DEFAULT_DOC_CONFIG},
  clients:[],providers:[],reservations:[],nextFile:1
};

// Firebase REST API helpers
const FB_URL = FB_CONFIG.databaseURL;
async function loadDB(){
  try{
    const r = await fetch(`${FB_URL}/travelmanager.json`);
    const data = await r.json();
    if(!data) return INIT;
    // Convertir arrays almacenados como objetos de Firebase
    const clients = data.clients ? Object.values(data.clients).filter(Boolean) : [];
    const providers = data.providers ? Object.values(data.providers).filter(Boolean) : [];
    const reservations = data.reservations ? Object.values(data.reservations).filter(Boolean).map(r=>({
      ...r,
      passengers: r.passengers ? (Array.isArray(r.passengers) ? r.passengers : Object.values(r.passengers)).filter(Boolean) : [],
      services: r.services ? (Array.isArray(r.services) ? r.services : Object.values(r.services)).filter(Boolean).map(s=>({
        ...s,
        paymentsDue: s.paymentsDue ? (Array.isArray(s.paymentsDue) ? s.paymentsDue : Object.values(s.paymentsDue)).filter(Boolean) : [],
      })) : [],
      paymentsReceived: r.paymentsReceived ? (Array.isArray(r.paymentsReceived) ? r.paymentsReceived : Object.values(r.paymentsReceived)).filter(Boolean) : [],
    })) : [];
    return {
      ...INIT,
      ...data,
      clients,
      providers,
      reservations,
    };
  } catch(e){ console.error("Firebase load error:",e); return INIT; }
}
async function saveDB(d){
  try{
    // Convertir arrays a objetos con ID como clave para Firebase
    const toObj = arr => arr.reduce((acc,item)=>{acc[item.id]=item;return acc;},{});
    const payload = {
      ...d,
      clients: toObj(d.clients||[]),
      providers: toObj(d.providers||[]),
      reservations: toObj(d.reservations||[]),
    };
    await fetch(`${FB_URL}/travelmanager.json`,{
      method:"PUT",
      headers:{"Content-Type":"application/json"},
      body: JSON.stringify(payload)
    });
  } catch(e){ console.error("Firebase save error:",e); }
}

// ── EXCEL ─────────────────────────────────────────
function buildCSV(rows,headers){const esc=v=>`"${String(v??'').replace(/"/g,'""')}"`;return[headers.map(esc).join(','),...rows.map(r=>r.map(esc).join(','))].join('\n');}
function downloadCSV(csv,filename){const blob=new Blob(['\uFEFF'+csv],{type:'text/csv;charset=utf-8;'});const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=filename;a.click();URL.revokeObjectURL(url);}
function exportClientStatement(client,reservations,settings){
  const clientRes=reservations.filter(r=>r.passengers.some(p=>p.name&&client.name&&p.name.toLowerCase()===client.name.toLowerCase()));
  const headers=['File','Destino','Fecha Salida','Estado','Precio Venta','Total Cobrado','Saldo Pendiente','Forma de Pago'];
  const rows=clientRes.map(r=>{const cobrado=paidSum(r.paymentsReceived);const metodos=[...new Set((r.paymentsReceived||[]).map(p=>p.method))].join(', ');return[r.fileNumber,r.destination||'',fmtDate(r.departureDate),STATUS[r.status]?.label||r.status,r.salePrice||0,cobrado,(r.salePrice||0)-cobrado,metodos];});
  const tv=sum(clientRes.map(r=>r.salePrice||0));const tc=sum(clientRes.map(r=>paidSum(r.paymentsReceived)));
  rows.push(['','','','TOTAL',tv,tc,tv-tc,'']);
  downloadCSV(buildCSV(rows,headers),`estado_cuenta_${client.name.replace(/\s/g,'_')}.csv`);
}
function exportProviderStatement(provider,reservations,settings){
  const headers=['File','Destino','Servicio','Descripción','Costo Neto','Pagado','Pendiente'];
  const rows=[];
  reservations.forEach(r=>{r.services.filter(s=>s.providerId===provider.id).forEach(s=>{const pagado=sum((s.paymentsDue||[]).filter(p=>p.paid).map(p=>p.amount||0));rows.push([r.fileNumber,r.destination||'',SVC_TYPES.find(t=>t.v===s.type)?.l||s.type,s.description||'',s.costPrice||0,pagado,(s.costPrice||0)-pagado]);});});
  const tc=sum(rows.map(r=>Number(r[4])));const tp=sum(rows.map(r=>Number(r[5])));
  rows.push(['','','','TOTAL',tc,tp,tc-tp]);
  downloadCSV(buildCSV(rows,headers),`estado_cuenta_${provider.name.replace(/\s/g,'_')}.csv`);
}
function exportReservations(reservations,settings){
  const headers=['File','Estado','Destino','Fecha Salida','Fecha Regreso','Pasajeros','Precio Venta','Total Cobrado','Saldo','Comisión'];
  const rows=reservations.map(r=>{const cobrado=paidSum(r.paymentsReceived);const costo=sum(r.services.map(s=>s.costPrice||0));return[r.fileNumber,STATUS[r.status]?.label||r.status,r.destination||'',fmtDate(r.departureDate),fmtDate(r.returnDate),r.passengers.map(p=>p.name).join(' / '),r.salePrice||0,cobrado,(r.salePrice||0)-cobrado,(r.salePrice||0)-costo];});
  downloadCSV(buildCSV(rows,headers),'reporte_reservas.csv');
}
function exportCommissions(reservations,settings){
  const headers=['File','Destino','Pasajero','Fecha Salida','Precio Venta','Costo Total','Comisión Neta','% Comisión'];
  const rows=reservations.filter(r=>r.status!=='cancelada').map(r=>{const costo=sum(r.services.map(s=>s.costPrice||0));const comision=(r.salePrice||0)-costo;const pct=r.salePrice?((comision/r.salePrice)*100).toFixed(1):0;return[r.fileNumber,r.destination||'',r.passengers[0]?.name||'',fmtDate(r.departureDate),r.salePrice||0,costo,comision,pct+'%'];});
  downloadCSV(buildCSV(rows,headers),'reporte_comisiones.csv');
}

// ── PRINT: RECIBO ─────────────────────────────────
function printReceipt(res,settings,providers){
  const dc=settings.docConfig?.recibo||{};
  const received=paidSum(res.paymentsReceived);
  const pending=(res.salePrice||0)-received;
  const color=settings.receiptColor||'#1a56db';
  const cur=settings.currency||'ARS';
  const logoHtml=settings.logo?`<img src="${settings.logo}" style="max-height:60px;max-width:180px;object-fit:contain;" alt="Logo"/>`:`<div style="font-size:20px;font-weight:800;color:${color}">${settings.agencyName}</div>`;
  const passengerBlock=res.passengers.map(p=>`<div style="display:flex;gap:24px;padding:6px 0;border-bottom:1px solid #e2e8f0;font-size:12px;"><span style="font-weight:600;min-width:200px;">${p.name}</span>${p.dni?`<span style="color:#64748b;">DNI/Pasaporte: <strong>${p.dni}</strong></span>`:''}</div>`).join('');

  const html=`<!DOCTYPE html><html><head><meta charset="utf-8"><title>Recibo #${res.fileNumber}</title>
<style>
*{margin:0;padding:0;box-sizing:border-box;}
body{font-family:'Helvetica Neue',Arial,sans-serif;padding:36px;color:#1a1a1a;font-size:13px;}
.hdr{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:28px;padding-bottom:18px;border-bottom:2px solid ${color};}
.agency-info{font-size:11px;color:#555;line-height:1.7;margin-top:6px;}
.doc-title h2{font-size:22px;font-weight:800;color:#1a1a1a;text-align:right;}
.doc-title p{font-size:12px;color:#666;text-align:right;margin-top:3px;}
.sec{margin-bottom:20px;}
.sec-title{font-size:10px;font-weight:800;text-transform:uppercase;letter-spacing:1px;color:#888;margin-bottom:8px;padding-bottom:5px;border-bottom:1px solid #eee;}
table{width:100%;border-collapse:collapse;font-size:12px;}
th{background:#f5f7fa;padding:7px 10px;text-align:left;font-weight:700;font-size:10px;text-transform:uppercase;color:#555;border-bottom:1px solid #e2e8f0;}
td{padding:7px 10px;border-bottom:1px solid #f1f5f9;}
.total-box{background:#f8fafc;border:1px solid #e2e8f0;border-radius:10px;padding:16px;margin-top:16px;}
.tr{display:flex;justify-content:space-between;padding:4px 0;font-size:13px;}
.tr.big{font-weight:800;font-size:16px;color:${color};border-top:2px solid #ddd;margin-top:8px;padding-top:10px;}
.footer{margin-top:36px;padding-top:14px;border-top:1px solid #ddd;text-align:center;font-size:11px;color:#aaa;}
.hnote{background:#f0f4ff;border-left:3px solid ${color};padding:8px 12px;margin-bottom:16px;font-size:12px;color:#444;}
.trip{background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:14px;display:grid;grid-template-columns:1fr 1fr;gap:10px;}
.tf label{display:block;font-size:10px;font-weight:700;text-transform:uppercase;color:#94a3b8;margin-bottom:2px;}
.tf span{font-size:13px;font-weight:600;color:#1e293b;}
@media print{body{padding:20px;-webkit-print-color-adjust:exact;print-color-adjust:exact;}}
</style></head><body>
<div class="hdr">
  <div>${logoHtml}<div class="agency-info">${settings.address?settings.address+'<br>':''}${settings.phone?'Tel: '+settings.phone+'&nbsp;&nbsp;':''}${settings.email?'Email: '+settings.email+'<br>':''}${settings.cuit?'CUIT: '+settings.cuit:''}</div></div>
  <div class="doc-title"><h2>RECIBO DE PAGO</h2><p>File N° <strong>${res.fileNumber}</strong></p><p>Fecha: ${fmtDate(today())}</p></div>
</div>
${settings.headerNote?`<div class="hnote">${settings.headerNote}</div>`:''}
<div class="sec"><div class="sec-title">Pasajeros</div><div style="margin-top:6px;">${passengerBlock}</div></div>
<div class="sec">
  <div class="sec-title">Detalle del viaje</div>
  <div class="trip">
    <div class="tf"><label>Destino</label><span>${res.destination||'—'}</span></div>
    <div class="tf"><label>Estado</label><span>${STATUS[res.status]?.label||res.status}</span></div>
    <div class="tf"><label>Fecha de salida</label><span>${fmtDate(res.departureDate)}</span></div>
    <div class="tf"><label>Fecha de regreso</label><span>${fmtDate(res.returnDate)}</span></div>
  </div>
  ${res.description?`<div style="margin-top:10px;padding:10px;background:#f9fafb;border-radius:6px;font-size:12px;color:#555;">${res.description}</div>`:''}
</div>
${dc.showServiceDetail?`<div class="sec"><div class="sec-title">Servicios incluidos</div><table><thead><tr><th>Servicio</th>${dc.showProviders?'<th>Proveedor</th>':''}${dc.showProviderRef?'<th>Referencia</th>':''}${dc.showServicePrices?'<th style="text-align:right">Precio</th>':''}</tr></thead><tbody>${res.services.map(s=>{const prov=providers.find(p=>p.id===s.providerId);const svcLabel=SVC_TYPES.find(t=>t.v===s.type)?.l||s.type;const iataLabel=s.type==='vuelo'&&(s.originCode||s.destinationCode)?` (${s.originCode||'?'} → ${s.destinationCode||'?'})${s.flightNumber?' · '+s.flightNumber:''}`:s.description?' — '+s.description:'';const hotelLabel=s.type==='hotel'?[s._extractedProviderName||s.description||'',s.nights?s.nights+' noches':'',s.destination||''].filter(Boolean).join(' · '):'';return`<tr><td>${svcLabel}${s.type==='hotel'?(hotelLabel?' — '+hotelLabel:''):iataLabel}</td>${dc.showProviders?`<td>${prov?.name||'—'}</td>`:''}${dc.showProviderRef?`<td>${s.providerFileNumber||'—'}</td>`:''}${dc.showServicePrices?`<td style="text-align:right">${fmt$(s.salePrice,cur)}</td>`:''}</tr>`;}).join('')}</tbody></table></div>`:''}
${dc.showPaymentHistory?`<div class="sec"><div class="sec-title">Historial de pagos recibidos</div><table><thead><tr><th>Fecha</th><th>Forma de pago</th><th>Notas</th><th style="text-align:right">Importe</th></tr></thead><tbody>${(res.paymentsReceived||[]).map(p=>`<tr><td>${fmtDate(p.date)}</td><td>${p.method||'—'}</td><td>${p.notes||''}</td><td style="text-align:right">${fmt$(p.amount,cur)}</td></tr>`).join('')}</tbody></table></div>`:''}
${dc.showTotals?`<div class="total-box"><div class="tr"><span>Precio total del viaje</span><span>${fmt$(res.salePrice,cur)}</span></div><div class="tr"><span>Total abonado</span><span style="color:#059669;font-weight:700;">${fmt$(received,cur)}</span></div><div class="tr big"><span>Saldo pendiente</span><span>${fmt$(pending,cur)}</span></div></div>`:''}
<div class="footer">${settings.agencyName}${settings.footerText?' — '+settings.footerText:''}${settings.emergencyPhone?'<br>Tel. emergencias: '+settings.emergencyPhone:''}</div>
<script>window.print();window.onafterprint=()=>window.close();<\/script>
</body></html>`;
  const w=window.open("","_blank","width=800,height=900");w.document.write(html);w.document.close();
}

// ── PRINT: VOUCHER AÉREO ──────────────────────────
function printVoucherAereo(res,settings,providers){
  const dc=settings.docConfig?.voucherAereo||{};
  const color=settings.voucherColor||'#1a56db';
  const cur=settings.currency||'ARS';
  const vuelos=res.services.filter(s=>s.type==='vuelo').sort((a,b)=>{
    const da=(a.departureDate||"")+(a.departureTime||"");
    const db=(b.departureDate||"")+(b.departureTime||"");
    return da.localeCompare(db);
  });
  const logoHtml=settings.logo?`<img src="${settings.logo}" style="max-height:40px;max-width:100px;object-fit:contain;background:white;padding:4px;border-radius:6px;" alt="Logo"/>`:'' ;
  const passengerBlock=res.passengers.map(p=>`<div style="display:flex;gap:24px;padding:6px 0;border-bottom:1px solid #e2e8f0;font-size:12px;"><span style="font-weight:600;min-width:200px;">${p.name}</span>${dc.showPassengerDNI&&p.dni?`<span style="color:#64748b;">DNI/Pasaporte: <strong>${p.dni}</strong></span>`:''}${p.birthDate?`<span style="color:#64748b;">Nac: ${fmtDate(p.birthDate)}</span>`:''}</div>`).join('');
  const vueloCards=vuelos.map(s=>{
    const prov=providers.find(p=>p.id===s.providerId);
    return`<div style="border:1px solid #e2e8f0;border-radius:10px;overflow:hidden;margin-bottom:14px;">
      <div style="background:${color};color:white;padding:10px 16px;display:flex;justify-content:space-between;align-items:center;">
        <span style="font-weight:700;font-size:14px;">✈ ${s.flightType||'Vuelo'}</span>
        <div style="display:flex;gap:16px;font-size:12px;opacity:.95;">
          ${dc.showFlightNumbers&&s.flightNumber?`<span>Vuelo: <strong>${s.flightNumber}</strong></span>`:''}
          ${dc.showProviderRef&&s.providerFileNumber?`<span>Localizador: <strong>${s.providerFileNumber}</strong></span>`:''}
        </div>
      </div>
      <div style="padding:16px;display:grid;grid-template-columns:1fr auto 1fr;gap:12px;align-items:center;">
        <div style="text-align:center;">
          <div style="font-size:28px;font-weight:800;color:#1e293b;letter-spacing:2px;">${s.originCode||'—'}</div>
          <div style="font-size:12px;color:#64748b;font-weight:600;margin-top:2px;">${s.origin||''}</div>
          <div style="font-size:11px;color:#94a3b8;margin-top:2px;">${fmtDate(s.departureDate||res.departureDate)}</div>
          ${s.departureTime?`<div style="font-size:18px;font-weight:800;color:${color};margin-top:4px;">${s.departureTime}</div>`:''}
        </div>
        <div style="text-align:center;color:#94a3b8;padding:0 12px;">
          <div style="font-size:10px;text-transform:uppercase;letter-spacing:.5px;margin-bottom:6px;">${s.stops||'Directo'}</div>
          <div style="font-size:24px;line-height:1;">→</div>
          ${s.duration?`<div style="font-size:10px;margin-top:6px;">${s.duration}</div>`:''}
        </div>
        <div style="text-align:center;">
          <div style="font-size:28px;font-weight:800;color:#1e293b;letter-spacing:2px;">${s.destinationCode||'—'}</div>
          <div style="font-size:12px;color:#64748b;font-weight:600;margin-top:2px;">${s.destination||''}</div>
          <div style="font-size:11px;color:#94a3b8;margin-top:2px;">${fmtDate(s.arrivalDate||res.returnDate)}</div>
          ${s.arrivalTime?`<div style="font-size:18px;font-weight:800;color:${color};margin-top:4px;">${s.arrivalTime}</div>`:''}
        </div>
      </div>
      <div style="padding:10px 16px;background:#f8fafc;border-top:1px solid #e2e8f0;display:flex;flex-wrap:wrap;gap:16px;font-size:12px;">
        ${(prov?.name||s.airline)?`<span>✈ <strong>${prov?.name||s.airline}</strong></span>`:''}
        ${s.flightClass?`<span>Clase: <strong>${s.flightClass}</strong></span>`:''}
        ${dc.showBaggageInfo&&s.baggage?`<span>🧳 ${s.baggage}</span>`:''}
        ${s.terminal?`<span>Terminal: <strong>${s.terminal}</strong></span>`:''}
        ${s.description?`<span>Obs: ${s.description}</span>`:''}
      </div>
    </div>`;
  }).join('');

  const html=`<!DOCTYPE html><html><head><meta charset="utf-8"><title>Voucher Aéreo #${res.fileNumber}</title>
<style>*{margin:0;padding:0;box-sizing:border-box;}body{font-family:'Helvetica Neue',Arial,sans-serif;color:#1a1a1a;font-size:13px;}.voucher{padding:30px;max-width:720px;margin:0 auto;}.hdr{background:${color};color:white;padding:18px 22px;border-radius:12px 12px 0 0;display:flex;justify-content:space-between;align-items:center;}.content{border:1px solid #dde1e9;border-top:none;border-radius:0 0 12px 12px;padding:22px;}.sec-title{font-size:10px;font-weight:800;text-transform:uppercase;letter-spacing:1px;color:#888;margin-bottom:10px;padding-bottom:6px;border-bottom:1px solid #eee;}.footer{margin-top:22px;text-align:center;font-size:11px;color:#aaa;padding-top:14px;border-top:1px solid #eee;}@media print{body{-webkit-print-color-adjust:exact;print-color-adjust:exact;}}</style>
</head><body>
<div class="voucher">
  <div class="hdr">
    <div style="display:flex;align-items:center;gap:12px;">${logoHtml}<div><div style="font-size:17px;font-weight:800;">${settings.agencyName}</div><div style="font-size:11px;opacity:.85;text-transform:uppercase;letter-spacing:1px;">Voucher Aéreo</div></div></div>
    <div style="text-align:right;font-size:12px;opacity:.9;"><div>File N° <strong>${res.fileNumber}</strong></div>${settings.emergencyPhone?`<div style="margin-top:4px;">Emergencias: ${settings.emergencyPhone}</div>`:''}</div>
  </div>
  <div class="content">
    <div style="margin-bottom:18px;"><div class="sec-title">Pasajeros</div>${passengerBlock}</div>
    <div style="margin-bottom:4px;"><div class="sec-title">Itinerario de vuelos</div>${vuelos.length===0?'<p style="color:#94a3b8;font-size:13px;">No hay vuelos registrados.</p>':vueloCards}</div>
    ${res.notes?`<div style="background:#FFF7ED;border:1px solid #FED7AA;border-radius:8px;padding:12px;margin-top:14px;font-size:12px;color:#92400E;"><strong>Observaciones:</strong> ${res.notes}</div>`:''}
    <div class="footer">${settings.agencyName}${settings.phone?' · '+settings.phone:''}${settings.email?' · '+settings.email:''}<br>${settings.footerText||'Gracias por elegirnos.'}</div>
  </div>
</div>
<script>window.print();window.onafterprint=()=>window.close();<\/script>
</body></html>`;
  const w=window.open("","_blank","width=760,height=900");w.document.write(html);w.document.close();
}

// ── PRINT: VOUCHER TERRESTRE ──────────────────────
function printVoucherTerrestre(res,settings,providers,svcIndex){
  const dc=settings.docConfig?.voucherTerrestre||{};
  const color=settings.voucherColor||'#1a56db';
  const cur=settings.currency||'ARS';
  const terrestres=res.services.filter(s=>s.type!=='vuelo').sort((a,b)=>{
    const da=a.checkIn||a.serviceDate||"";
    const db=b.checkIn||b.serviceDate||"";
    return da.localeCompare(db);
  });
  const servicesToPrint=svcIndex!==undefined?[terrestres[svcIndex]]:terrestres;
  const logoHtml=settings.logo?`<img src="${settings.logo}" style="max-height:40px;max-width:100px;object-fit:contain;background:white;padding:4px;border-radius:6px;" alt="Logo"/>`:'' ;
  const passengerBlock=res.passengers.map(p=>`<div style="display:flex;gap:24px;padding:6px 0;border-bottom:1px solid #e2e8f0;font-size:12px;"><span style="font-weight:600;min-width:200px;">${p.name}</span>${dc.showPassengerDNI&&p.dni?`<span style="color:#64748b;">DNI/Pasaporte: <strong>${p.dni}</strong></span>`:''}${p.birthDate?`<span style="color:#64748b;">Nac: ${fmtDate(p.birthDate)}</span>`:''}</div>`).join('');
  const svcCards=servicesToPrint.filter(Boolean).map(s=>{
    const prov=providers.find(p=>p.id===s.providerId);
    const isHotel=s.type==='hotel';
    return`<div style="border:1px solid #e2e8f0;border-radius:10px;overflow:hidden;margin-bottom:16px;">
      <div style="background:${color};color:white;padding:10px 16px;display:flex;justify-content:space-between;align-items:center;">
        <span style="font-weight:700;font-size:14px;">${SVC_TYPES.find(t=>t.v===s.type)?.l||s.type}</span>
        ${dc.showProviderRef&&s.providerFileNumber?`<span style="font-size:12px;opacity:.9;">Ref: ${s.providerFileNumber}</span>`:''}
      </div>
      <div style="padding:16px 16px 8px;">
        <div style="font-size:17px;font-weight:800;color:#1e293b;margin-bottom:12px;">${prov?.name||s._extractedProviderName||s.description||'Sin nombre'}</div>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px 16px;font-size:12px;margin-bottom:8px;">
          ${dc.showHotelAddress&&(prov?.address||s._extractedProviderAddress)?`<div style="grid-column:span 2"><strong>Dirección:</strong> ${prov?.address||s._extractedProviderAddress}</div>`:''}
          ${dc.showHotelPhone&&(prov?.phone||s._extractedProviderPhone)?`<div style="grid-column:span 2"><strong>Teléfono:</strong> ${prov?.phone||s._extractedProviderPhone}</div>`:''}
          ${isHotel&&s.checkIn?`<div><strong>Check-in:</strong> ${fmtDate(s.checkIn)}</div>`:''}
          ${isHotel&&s.checkOut?`<div><strong>Check-out:</strong> ${fmtDate(s.checkOut)}</div>`:''}
          ${isHotel&&s.nights?`<div><strong>Noches:</strong> ${s.nights}</div>`:''}
          ${isHotel&&s.rooms?`<div><strong>Habitaciones:</strong> ${s.rooms}</div>`:''}
          ${dc.showRoomType&&s.roomType?`<div style="grid-column:span 2"><strong>Tipo de habitación:</strong> ${s.roomType}</div>`:''}
          ${dc.showRegimen&&s.regimen?`<div><strong>Régimen:</strong> ${s.regimen}</div>`:''}
          ${!isHotel&&s.serviceDate?`<div><strong>Fecha:</strong> ${fmtDate(s.serviceDate)}</div>`:''}
          ${!isHotel&&s.serviceTime?`<div><strong>Hora:</strong> ${s.serviceTime}</div>`:''}
          ${s.description&&isHotel&&s.description!==s._extractedProviderName?`<div style="grid-column:span 2"><strong>Descripción:</strong> ${s.description}</div>`:''}
        </div>
      </div>
      ${dc.showImportantInfo&&s.importantInfo?`<div style="background:#FFF7ED;border-top:1px solid #FED7AA;padding:12px 16px;font-size:11px;color:#92400E;"><strong>Información importante:</strong><br>${s.importantInfo}</div>`:''}
    </div>`;
  }).join('');

  const html=`<!DOCTYPE html><html><head><meta charset="utf-8"><title>Voucher Terrestre #${res.fileNumber}</title>
<style>*{margin:0;padding:0;box-sizing:border-box;}body{font-family:'Helvetica Neue',Arial,sans-serif;color:#1a1a1a;font-size:13px;}.voucher{padding:30px;max-width:720px;margin:0 auto;}.hdr{background:${color};color:white;padding:18px 22px;border-radius:12px 12px 0 0;display:flex;justify-content:space-between;align-items:center;}.content{border:1px solid #dde1e9;border-top:none;border-radius:0 0 12px 12px;padding:22px;}.sec-title{font-size:10px;font-weight:800;text-transform:uppercase;letter-spacing:1px;color:#888;margin-bottom:10px;padding-bottom:6px;border-bottom:1px solid #eee;}.footer{margin-top:22px;text-align:center;font-size:11px;color:#aaa;padding-top:14px;border-top:1px solid #eee;}@media print{body{-webkit-print-color-adjust:exact;print-color-adjust:exact;}}</style>
</head><body>
<div class="voucher">
  <div class="hdr">
    <div style="display:flex;align-items:center;gap:12px;">${logoHtml}<div><div style="font-size:17px;font-weight:800;">${settings.agencyName}</div><div style="font-size:11px;opacity:.85;text-transform:uppercase;letter-spacing:1px;">Voucher Terrestre</div></div></div>
    <div style="text-align:right;font-size:12px;opacity:.9;"><div>File N° <strong>${res.fileNumber}</strong></div>${settings.emergencyPhone?`<div style="margin-top:4px;">Emergencias: ${settings.emergencyPhone}</div>`:''}</div>
  </div>
  <div class="content">
    <div style="margin-bottom:18px;"><div class="sec-title">Pasajeros</div>${passengerBlock}</div>
    <div><div class="sec-title">Servicios</div>${servicesToPrint.length===0?'<p style="color:#94a3b8;font-size:13px;">No hay servicios terrestres.</p>':svcCards}</div>
    ${res.notes?`<div style="background:#FFF7ED;border:1px solid #FED7AA;border-radius:8px;padding:12px;font-size:12px;color:#92400E;margin-top:12px;"><strong>Observaciones:</strong> ${res.notes}</div>`:''}
    <div class="footer">${settings.agencyName}${settings.phone?' · '+settings.phone:''}${settings.email?' · '+settings.email:''}<br>${settings.footerText||'Gracias por elegirnos.'}</div>
  </div>
</div>
<script>window.print();window.onafterprint=()=>window.close();<\/script>
</body></html>`;
  const w=window.open("","_blank","width=760,height=900");w.document.write(html);w.document.close();
}

// ── UI PRIMITIVES ─────────────────────────────────
const S={
  primaryBtn:{background:"#2563EB",color:"#fff",border:"none",padding:"7px 16px",borderRadius:8,fontWeight:600,fontSize:13,cursor:"pointer",display:"inline-flex",alignItems:"center",gap:6},
  secondaryBtn:{background:"#F1F5F9",color:"#475569",border:"1px solid #E2E8F0",padding:"7px 16px",borderRadius:8,fontWeight:600,fontSize:13,cursor:"pointer",display:"inline-flex",alignItems:"center",gap:6},
  ghostBtn:{background:"transparent",color:"#64748B",border:"1px solid #CBD5E1",padding:"6px 10px",borderRadius:8,fontWeight:500,fontSize:12,cursor:"pointer",display:"inline-flex",alignItems:"center",gap:4},
  successBtn:{background:"#059669",color:"#fff",border:"none",padding:"7px 16px",borderRadius:8,fontWeight:600,fontSize:13,cursor:"pointer",display:"inline-flex",alignItems:"center",gap:6},
  input:{padding:"8px 12px",border:"1px solid #E2E8F0",borderRadius:8,fontSize:13,color:"#1E293B",background:"#fff",outline:"none",width:"100%",boxSizing:"border-box"},
  label:{fontSize:11,fontWeight:700,color:"#475569",textTransform:"uppercase",letterSpacing:.5,display:"block",marginBottom:4},
  card:{background:"#fff",borderRadius:12,boxShadow:"0 1px 3px rgba(0,0,0,.07)",padding:20},
  sectionTitle:{margin:"0 0 16px",fontSize:14,fontWeight:700,color:"#0F172A",display:"flex",alignItems:"center",gap:8},
};
const Btn=({children,onClick,variant="primary",size="md",disabled,style:sx})=>{
  const base=variant==="primary"?S.primaryBtn:variant==="ghost"?S.ghostBtn:variant==="success"?S.successBtn:S.secondaryBtn;
  const sz=size==="sm"?{padding:"4px 10px",fontSize:12}:size==="lg"?{padding:"10px 20px",fontSize:14}:{};
  return<button onClick={onClick} disabled={disabled} style={{...base,...sz,opacity:disabled?.5:1,...sx}}>{children}</button>;
};
const Field=({label,children,col,style:sx})=>(<div style={{gridColumn:col?`span ${col}`:undefined,...sx}}>{label&&<label style={S.label}>{label}</label>}{children}</div>);
const Inp=({label,value,onChange,type="text",placeholder,col,style:sx})=>(<Field label={label} col={col} style={sx}><input type={type} value={value||""} onChange={e=>onChange(e.target.value)} placeholder={placeholder} style={S.input}/></Field>);
const Sel=({label,value,onChange,options,col,style:sx})=>(<Field label={label} col={col} style={sx}><select value={value||""} onChange={e=>onChange(e.target.value)} style={{...S.input,appearance:"auto"}}><option value="">Seleccionar…</option>{options.map(o=><option key={o.v||o} value={o.v||o}>{o.l||o}</option>)}</select></Field>);
const Txta=({label,value,onChange,rows=3,placeholder,col,style:sx})=>(<Field label={label} col={col} style={sx}><textarea value={value||""} onChange={e=>onChange(e.target.value)} rows={rows} placeholder={placeholder} style={{...S.input,resize:"vertical"}}/></Field>);
const Badge=({status})=>{const s=STATUS[status]||STATUS.cotizacion;return<span style={{background:s.bg,color:s.c,padding:"3px 10px",borderRadius:20,fontSize:11,fontWeight:700,whiteSpace:"nowrap"}}>{s.label}</span>;};
const Card=({children,style:sx})=>(<div style={{...S.card,...sx}}>{children}</div>);
const Modal=({title,onClose,children,width=720,footer})=>(
  <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.55)",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
    <div style={{background:"#fff",borderRadius:16,width:"100%",maxWidth:width,maxHeight:"92vh",display:"flex",flexDirection:"column",overflow:"hidden"}}>
      <div style={{padding:"16px 20px",borderBottom:"1px solid #E2E8F0",display:"flex",alignItems:"center",justifyContent:"space-between",flexShrink:0}}>
        <h3 style={{margin:0,fontSize:16,fontWeight:700,color:"#0F172A"}}>{title}</h3>
        <button onClick={onClose} style={{background:"none",border:"none",cursor:"pointer",color:"#94A3B8"}}><X size={20}/></button>
      </div>
      <div style={{overflowY:"auto",flex:1,padding:20}}>{children}</div>
      {footer&&<div style={{padding:"12px 20px",borderTop:"1px solid #E2E8F0",display:"flex",gap:8,justifyContent:"flex-end",flexShrink:0}}>{footer}</div>}
    </div>
  </div>
);
const Tabs=({tabs,active,setActive})=>(<div style={{display:"flex",gap:2,borderBottom:"2px solid #E2E8F0",marginBottom:16}}>{tabs.map(t=>(<button key={t.id} onClick={()=>setActive(t.id)} style={{padding:"8px 14px",border:"none",background:"none",cursor:"pointer",fontSize:13,fontWeight:600,color:active===t.id?"#2563EB":"#64748B",borderBottom:active===t.id?"2px solid #2563EB":"2px solid transparent",marginBottom:-2}}>{t.label}</button>))}</div>);
const EmptyState=({icon:Icon,title,sub,action})=>(<div style={{textAlign:"center",padding:"48px 24px",color:"#94A3B8"}}><Icon size={40} style={{marginBottom:12,opacity:.4}}/><div style={{fontSize:15,fontWeight:600,color:"#64748B",marginBottom:6}}>{title}</div>{sub&&<div style={{fontSize:13,marginBottom:16}}>{sub}</div>}{action}</div>);
const Grid=({cols=2,gap=16,children,style:sx})=>(<div style={{display:"grid",gridTemplateColumns:`repeat(${cols},1fr)`,gap,...sx}}>{children}</div>);
const Divider=({label})=>(<div style={{display:"flex",alignItems:"center",gap:8,margin:"16px 0"}}><div style={{flex:1,height:1,background:"#E2E8F0"}}/>{label&&<span style={{fontSize:11,fontWeight:700,color:"#94A3B8",textTransform:"uppercase",letterSpacing:.5}}>{label}</span>}<div style={{flex:1,height:1,background:"#E2E8F0"}}/></div>);
const Toggle=({label,checked,onChange,desc})=>(<div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",padding:"10px 0",borderBottom:"1px solid #F1F5F9"}}><div><div style={{fontSize:13,fontWeight:600,color:"#1E293B"}}>{label}</div>{desc&&<div style={{fontSize:11,color:"#94A3B8",marginTop:2}}>{desc}</div>}</div><div onClick={()=>onChange(!checked)} style={{width:42,height:24,borderRadius:12,background:checked?"#2563EB":"#CBD5E1",cursor:"pointer",position:"relative",flexShrink:0,transition:"background .2s",marginLeft:12}}><div style={{position:"absolute",top:3,left:checked?19:3,width:18,height:18,borderRadius:9,background:"white",boxShadow:"0 1px 3px rgba(0,0,0,.2)",transition:"left .2s"}}/></div></div>);

// ── SIDEBAR ───────────────────────────────────────
const NAV=[{id:"dashboard",Icon:LayoutDashboard,l:"Dashboard"},{id:"quotations",Icon:ClipboardList,l:"Cotizaciones"},{id:"reservations",Icon:CalendarDays,l:"Reservas"},{id:"clients",Icon:Users,l:"Clientes"},{id:"providers",Icon:Building2,l:"Proveedores"},{id:"commissions",Icon:TrendingUp,l:"Comisiones"},{id:"reports",Icon:BarChart2,l:"Reportes"},{id:"settings",Icon:Settings,l:"Configuración"}];
function Sidebar({section,setSection,settings}){
  return(<aside style={{width:210,background:"#0F172A",display:"flex",flexDirection:"column",flexShrink:0,userSelect:"none"}}>
    <div style={{padding:"18px 16px",borderBottom:"1px solid #1E293B",display:"flex",alignItems:"center",gap:10}}>
      {settings.logo?<img src={settings.logo} style={{height:36,width:36,objectFit:"contain",borderRadius:6,background:"#fff",padding:2}} alt="Logo"/>:<div style={{width:36,height:36,borderRadius:8,background:"#2563EB",display:"flex",alignItems:"center",justifyContent:"center",color:"#fff",fontSize:14,fontWeight:800}}>{settings.agencyName?.charAt(0)||"A"}</div>}
      <div><div style={{fontSize:10,color:"#475569",fontWeight:700,textTransform:"uppercase",letterSpacing:1,marginBottom:2}}>Agencia</div><div style={{fontSize:12,color:"#F8FAFC",fontWeight:700,lineHeight:1.3}}>{settings.agencyName}</div></div>
    </div>
    <nav style={{flex:1,padding:"8px 0",overflowY:"auto"}}>{NAV.map(({id,Icon,l})=>(<button key={id} onClick={()=>setSection(id)} style={{display:"flex",alignItems:"center",gap:10,width:"100%",padding:"10px 16px",background:section===id?"rgba(59,130,246,.15)":"transparent",border:"none",cursor:"pointer",textAlign:"left",color:section===id?"#60A5FA":"#94A3B8",fontSize:13,fontWeight:section===id?600:400,borderLeft:section===id?"3px solid #3B82F6":"3px solid transparent"}}><Icon size={15}/>{l}</button>))}</nav>
    <div style={{padding:"10px 16px",borderTop:"1px solid #1E293B",fontSize:10,color:"#334155",textAlign:"center"}}>TravelManager v3.0</div>
  </aside>);
}

// ── DASHBOARD ─────────────────────────────────────
function Dashboard({data}){
  const {reservations,clients,providers,settings}=data;
  const n=today();const cur=settings.currency||'ARS';
  const stats=useMemo(()=>{
    const nc=reservations.filter(r=>r.status!=="cancelada");
    const totalSale=sum(nc.map(r=>r.salePrice||0));
    const totalReceived=sum(reservations.map(r=>paidSum(r.paymentsReceived)));
    const totalCost=sum(nc.map(r=>sum(r.services.map(s=>s.costPrice||0))));
    const active=reservations.filter(r=>["confirmada","en_viaje"].includes(r.status)).length;
    const upcoming=reservations.filter(r=>r.departureDate>=n&&r.status!=="cancelada").sort((a,b)=>a.departureDate.localeCompare(b.departureDate)).slice(0,6);
    const alerts=[];
    reservations.forEach(res=>res.services.forEach(svc=>(svc.paymentsDue||[]).filter(p=>!p.paid).forEach(p=>{const prov=providers.find(pr=>pr.id===svc.providerId);alerts.push({res,svc,p,providerName:prov?.name||"Sin proveedor",overdue:p.dueDate&&p.dueDate<n});})));
    return{totalSale,totalReceived,totalPending:totalSale-totalReceived,totalCommission:totalSale-totalCost,active,upcoming,alerts};
  },[reservations]);
  const cards=[{l:"Reservas activas",v:stats.active,Icon:CalendarDays,c:"#2563EB",bg:"#EFF6FF"},{l:"Ingresos totales",v:fmt$(stats.totalSale,cur),Icon:TrendingUp,c:"#059669",bg:"#ECFDF5"},{l:"Total cobrado",v:fmt$(stats.totalReceived,cur),Icon:Check,c:"#7C3AED",bg:"#F5F3FF"},{l:"Por cobrar",v:fmt$(stats.totalPending,cur),Icon:AlertCircle,c:"#D97706",bg:"#FFFBEB"},{l:"Comisiones",v:fmt$(stats.totalCommission,cur),Icon:TrendingUp,c:"#DC2626",bg:"#FEF2F2"},{l:"Clientes",v:clients.length,Icon:Users,c:"#0891B2",bg:"#F0F9FF"}];
  return(<div style={{padding:28,overflowY:"auto",flex:1}}>
    <h1 style={{margin:"0 0 20px",fontSize:22,fontWeight:800,color:"#0F172A"}}>Dashboard</h1>
    <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(175px,1fr))",gap:14,marginBottom:28}}>{cards.map(x=>(<div key={x.l} style={{background:x.bg,borderRadius:12,padding:16}}><div style={{display:"flex",alignItems:"center",gap:8,marginBottom:10}}><x.Icon size={16} style={{color:x.c}}/><span style={{fontSize:11,fontWeight:700,color:x.c,textTransform:"uppercase",letterSpacing:.5}}>{x.l}</span></div><div style={{fontSize:22,fontWeight:800,color:x.c}}>{x.v}</div></div>))}</div>
    <Grid cols={2} gap={16}>
      <Card><h3 style={S.sectionTitle}><CalendarDays size={16} style={{color:"#2563EB"}}/>Próximas salidas</h3>{stats.upcoming.length===0?<p style={{color:"#94A3B8",fontSize:13}}>No hay salidas próximas.</p>:stats.upcoming.map(r=>(<div key={r.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"8px 0",borderBottom:"1px solid #F1F5F9"}}><div><div style={{fontSize:13,fontWeight:600}}>{r.destination||"Sin destino"}</div><div style={{fontSize:11,color:"#64748B"}}>{r.passengers.map(p=>p.name).filter(Boolean).join(", ")||"—"} — File #{r.fileNumber}</div></div><div style={{display:"flex",flexDirection:"column",alignItems:"flex-end",gap:3}}><Badge status={r.status}/><span style={{fontSize:11,color:"#2563EB",fontWeight:700}}>{fmtDate(r.departureDate)}</span></div></div>))}</Card>
      <Card><h3 style={S.sectionTitle}><AlertCircle size={16} style={{color:"#EF4444"}}/>Pagos a proveedores pendientes</h3>{stats.alerts.length===0?<p style={{color:"#94A3B8",fontSize:13}}>No hay pagos pendientes.</p>:stats.alerts.slice(0,7).map((a,i)=>(<div key={i} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"7px 0",borderBottom:"1px solid #F1F5F9"}}><div><div style={{fontSize:13,fontWeight:600}}>{a.providerName}</div><div style={{fontSize:11,color:"#64748B"}}>File #{a.res.fileNumber}</div></div><div style={{textAlign:"right"}}><div style={{fontSize:12,fontWeight:700,color:"#EF4444"}}>{fmt$(a.p.amount,cur)}</div><div style={{fontSize:11,color:a.overdue?"#EF4444":"#64748B",fontWeight:a.overdue?700:400}}>{a.overdue?"⚠ ":""}{fmtDate(a.p.dueDate)}</div></div></div>))}</Card>
    </Grid>
  </div>);
}

// ── RESERVATION MODAL ─────────────────────────────
const newRes=(fileNum,defComm)=>({id:genId(),fileNumber:String(fileNum).padStart(5,"0"),status:"cotizacion",destination:"",departureDate:"",returnDate:"",description:"",salePrice:0,commissionPercent:defComm||15,passengers:[],services:[],paymentsReceived:[],notes:"",createdAt:today()});
const newPass=()=>({id:genId(),name:"",dni:"",email:"",phone:"",birthDate:""});
const newSvc=()=>({id:genId(),type:"vuelo",description:"",providerId:"",providerFileNumber:"",costPrice:0,salePrice:0,paymentsDue:[],flightType:"",flightNumber:"",airline:"",origin:"",originCode:"",destination:"",destinationCode:"",departureDate:"",departureTime:"",arrivalDate:"",arrivalTime:"",stops:"Directo",duration:"",flightClass:"",baggage:"",terminal:"",checkIn:"",checkOut:"",nights:"",rooms:"1",roomType:"",regimen:"",importantInfo:"",serviceDate:"",serviceTime:""});
const newPay=()=>({id:genId(),date:today(),amount:0,method:"Transferencia",notes:""});
const newDue=()=>({id:genId(),dueDate:"",amount:0,paid:false,paidDate:null,notes:""});

// ── VOUCHER READER (pegar texto) ─────────────────
function VoucherReader({onApply}){
  const [open, setOpen] = useState(false);
  const [text, setText] = useState("");
  const [preview, setPreview] = useState(null);

  const extractAndPreview = () => {
    const t = text;
    const iataToCity={"EZE":"Buenos Aires","AEP":"Buenos Aires","COR":"Cordoba","MDZ":"Mendoza","BRC":"Bariloche","IGR":"Iguazu","TUC":"Tucuman","USH":"Ushuaia","GRU":"Sao Paulo","GIG":"Rio de Janeiro","CUN":"Cancun","MEX":"Ciudad de Mexico","BOG":"Bogota","SCL":"Santiago","LIM":"Lima","CUZ":"Cusco","UIO":"Quito","GYE":"Guayaquil","MVD":"Montevideo","ASU":"Asuncion","MIA":"Miami","JFK":"Nueva York","LAX":"Los Angeles","ORD":"Chicago","ATL":"Atlanta","SFO":"San Francisco","BOS":"Boston","LHR":"Londres","CDG":"Paris","AMS":"Amsterdam","FRA":"Frankfurt","MUC":"Munich","FCO":"Roma","BCN":"Barcelona","MAD":"Madrid","LIS":"Lisboa","ZRH":"Zurich","VIE":"Viena","JMK":"Mykonos","ATH":"Atenas","HER":"Heraklion","RHO":"Rodas","PMI":"Palma de Mallorca","IBZ":"Ibiza","TFS":"Tenerife","LPA":"Gran Canaria","AGP":"Malaga","ALC":"Alicante","VLC":"Valencia","DXB":"Dubai","DOH":"Doha","SIN":"Singapur","BKK":"Bangkok","NRT":"Tokio","SYD":"Sidney","YYZ":"Toronto"};
    const g = patterns => { for(const p of patterns){const m=t.match(p);if(m&&m[1])return m[1].trim();} return ""; };
    const allDates=[...t.matchAll(/(\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4})/g)].map(m=>m[1]);
    const allTimes=[...t.matchAll(/(\d{1,2}:\d{2})(?:\s*hs\.?)?/g)].map(m=>m[1]);
    const allIATA=[...t.matchAll(/\b([A-Z]{3})\b/g)].map(m=>m[1]).filter(c=>c.length===3&&!/^(IB|LA|AR|AA|UA|DL|LH|AF|BA|EK|hs|PM|AM|SR|MR|MS|BB|HB|FB|AI)$/.test(c));
    const pd = s => {
      if(!s)return "";
      const m=s.match(/(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})/);
      if(m){const y=m[3].length===2?"20"+m[3]:m[3];return `${y}-${m[2].padStart(2,"0")}-${m[1].padStart(2,"0")}`;}
      return "";
    };
    const flightMatch = t.match(/\b([A-Z]{2})\s*(\d{3,4})\b/);
    const hasVuelo = !!flightMatch || /itinerario de vuelo|e-ticket|boarding/i.test(t);
    const hasHotel = /check.?in|check.?out|habitaci|room type|hotel confirmation|night|noche|alojamiento/i.test(t);
    const type = hasVuelo?"vuelo":hasHotel?"hotel":"otro";

    // Detectar nombre del hotel — buscar antes de "Check In" o después de ciudad
    const hotelName = g([
      /\n([^\n]*(?:ibis|Hilton|Marriott|Hyatt|Sheraton|Radisson|Novotel|Mercure|Holiday Inn|Best Western|NH |Meli[aá]|Riu |Barcelo|Occidental|Wyndham|Crowne|InterContinental|Courtyard|Hampton|Doubletree|Renaissance|W Hotel|Sofitel|Pullman|MGallery|Swissotel|Fairmont|Accor|Eurostars|Tryp|Hesperia|AC Hotel|Premier Inn|Travelodge|Comfort Inn|Quality Inn|Days Inn|Ramada|Howard Johnson|La Quinta)[^\n]*)\n/i,
      /\n([^\n]*(?:Hotel|Inn|Suites?|Resort|Apart(?:ment)?|Hostel|Lodge|Palace|Plaza|Grand)[^\n]{3,50})\n/i,
    ]);

    // Detectar referencia/confirmación del hotel
    const hotelRef = g([
      /Hotel Confirmation No[:\s]*([\w]+)/i,
      /Confirmation[:\s]*([\w]{4,12})/i,
      /Agency Ref[:\s]*([\w-]+)/i,
      /TBOH Confirmation[:\s]*([\w]+)/i,
      /Voucher Details.*?([A-Z0-9]{5,10})/i,
    ]);

    // Detectar fechas en inglés: "30 May 2026", "Jun 04 2026"
    const engMonths={jan:"01",feb:"02",mar:"03",apr:"04",may:"05",jun:"06",jul:"07",aug:"08",sep:"09",oct:"10",nov:"11",dec:"12"};
    const pdEng = s => {
      if(!s)return "";
      // DD Month YYYY o Month DD YYYY
      const m1=s.match(/(\d{1,2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+(\d{4})/i);
      if(m1)return `${m1[3]}-${engMonths[m1[2].toLowerCase().slice(0,3)]}-${m1[1].padStart(2,"0")}`;
      const m2=s.match(/(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+(\d{1,2}),?\s+(\d{4})/i);
      if(m2)return `${m2[3]}-${engMonths[m2[1].toLowerCase().slice(0,3)]}-${m2[2].padStart(2,"0")}`;
      return pd(s);
    };

    // Extraer check-in y check-out con soporte inglés/español
    const checkInRaw = g([/Check\s*In[:\s]+([^\n]+)/i, /check.in[:\s]+([^\n]+)/i, /fecha de llegada[:\s]+([^\n]+)/i, /entrada[:\s]+([^\n]+)/i]);
    const checkOutRaw = g([/Check\s*Out[:\s]+([^\n]+)/i, /check.out[:\s]+([^\n]+)/i, /fecha de salida[:\s]+([^\n]+)/i, /salida[:\s]+([^\n]+)/i]);
    const checkInDate = pdEng(checkInRaw)||pd(allDates[0]);
    const checkOutDate = pdEng(checkOutRaw)||pd(allDates[1]);

    // Calcular noches
    let nights = g([/(\d+)\s*noche/i, /noches?[:\s]+(\d+)/i, /(\d+)\s*night/i, /nights?[:\s]+(\d+)/i]);
    if(!nights && checkInDate && checkOutDate){
      const d1=new Date(checkInDate), d2=new Date(checkOutDate);
      const diff=Math.round((d2-d1)/(1000*60*60*24));
      if(diff>0) nights=String(diff);
    }

    // Tipo de habitación
    const roomTypeRaw = g([
      /Room\s*\d+\s*([^\n]*(?:Double|Single|Twin|Suite|Studio|Queen|King|Deluxe|Superior|Standard|Triple|Family)[^\n]*)/i,
      /Room type[:\s]+([^\n]+)/i,
      /tipo de habitaci[oó]n[:\s]+([^\n]+)/i,
      /categor[ií]a[:\s]+([^\n]+)/i,
    ]);
    // Limpiar "Incl:..." y "NonSmoking" del tipo de habitación
    const roomType = roomTypeRaw.replace(/\s*Incl[:\s]+[^,]*/gi,"").replace(/,?\s*NonSmoking/gi,"").trim();

    // Régimen
    const regimenRaw = g([/Incl[:\s]+([^\n]+)/i, /Inclusion[:\s]+([^\n]+)/i, /r[eé]gimen[:\s]+([^\n]+)/i, /basis[:\s]+([^\n]+)/i, /\b(BB|HB|FB|AI|Room Only|Bed and Breakfast|Half Board|Full Board|All Inclusive|Solo habitaci[oó]n)\b/i]);
    const regimenMap={"room only":"Solo habitación","solo habitacion":"Solo habitación","bb":"BB - Bed & Breakfast","bed and breakfast":"BB - Bed & Breakfast","hb":"HB - Media Pensión","half board":"HB - Media Pensión","fb":"FB - Pensión Completa","full board":"FB - Pensión Completa","ai":"AI - Todo Incluido","all inclusive":"AI - Todo Incluido"};
    const regimenKey=Object.keys(regimenMap).find(k=>regimenRaw.toLowerCase().includes(k));
    const regimen = regimenKey?regimenMap[regimenKey]:regimenRaw;

    // Dirección del hotel
    const hotelAddress = g([/([^\n]*(?:\d+[^\n]*(?:Street|St\.|Avenue|Ave\.|Road|Rd\.|Calle|C\.|Plaza|Blvd)[^\n]*))/i, /([^\n]*,\s*\d{4,6}[^\n]*)/]);

    // Pasajeros del hotel
    const hotelPassengers=[...t.matchAll(/(?:Mrs?|Mr|Ms|Miss|Sr|Sra)\.?\s+([A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑA-Za-z\s]+)/g)].map(m=>m[1].trim()).filter(p=>p.length>3);

    // Calcular fecha de llegada — si la hora de llegada es menor que la de salida, es día siguiente
    const depTime = allTimes[0]||"";
    const arrTime = allTimes[1]||"";
    const depDate = pd(allDates[0]);
    let arrDate = pd(allDates[1]||allDates[0]);
    if(depDate && depTime && arrTime && arrDate===depDate){
      const [dh,dm]= depTime.split(":").map(Number);
      const [ah,am]= arrTime.split(":").map(Number);
      const depMins = dh*60+dm;
      const arrMins = ah*60+am;
      if(arrMins <= depMins){
        // Llega al día siguiente
        const d = new Date(depDate+"T12:00:00");
        d.setDate(d.getDate()+1);
        arrDate = d.toISOString().split("T")[0];
      }
    }
    // También calcular desde duración si está disponible
    const durationMatch = t.match(/[Dd]uraci[oó]n[:\s]+(\d+)h\s*(\d+)m/);
    if(durationMatch && depDate && depTime && !arrDate){
      const dh=parseInt(durationMatch[1]), dm=parseInt(durationMatch[2]);
      const [h,m]= depTime.split(":").map(Number);
      const totalMins = h*60+m+dh*60+dm;
      const d = new Date(depDate+"T12:00:00");
      if(totalMins>=1440) d.setDate(d.getDate()+1);
      arrDate = d.toISOString().split("T")[0];
    }
    // Teléfono del hotel
    const hotelPhone = g([/Phone[:\s]+([+\d\s()-]{6,20})/i, /Tel[eé]fono[:\s]+([+\d\s()-]{6,20})/i, /\+?\d[\d\s()-]{8,18}\d/]);

    // Código de reserva aerolínea (diferente al código general)
    const airlineBookingRef = g([/c[oó]digo de reserva de la aerol[ií]nea[:\s]*([\w]{5,8})/i, /Airline.*?Ref[:\s]*([\w]{5,8})/i, /\b([A-Z]{6})\b(?=\s*\()/]);

    // E-Ticket numbers
    const etickets = [...t.matchAll(/(?:E-TICKET|E\.TICKET|TICKET)[:\s]*([\d-]{10,20})/gi)].map(m=>m[1]).join(", ");

    // Horarios de check-in/out del hotel
    const checkInTime = g([/CheckIn Time[^:]*:[:\s]*(\d{1,2}:\d{2}\s*(?:AM|PM)?)/i, /Check.?in.*?Time[:\s]*(\d{1,2}:\d{2}\s*(?:AM|PM)?)/i]);
    const checkOutTime = g([/CheckOut Time[^:]*:[:\s]*(\d{1,2}:\d{2}\s*(?:AM|PM)?)/i, /Check.?out.*?Time[:\s]*(\d{1,2}:\d{2}\s*(?:AM|PM)?)/i]);

    const data = {
      type,
      flightNumber: flightMatch?flightMatch[1]+" "+flightMatch[2]:"",
      airline: g([/\b(IBERIA|LATAM|AEROL[IÍ]NEAS ARGENTINAS|AMERICAN AIRLINES|UNITED|DELTA|LUFTHANSA|AIR FRANCE|BRITISH AIRWAYS|EMIRATES|FLYBONDI|JETSMART|RYANAIR|VUELING)\b/i]),
      providerFileNumber: hotelRef||airlineBookingRef||g([/c[oó]digo de reserva[:\s]*(\w{5,8})/i, /localizador[:\s]*(\w{5,8})/i]),
      _extractedProviderName: hotelName||"",
      _extractedProviderAddress: hotelAddress||"",
      _extractedProviderPhone: hotelPhone||"",
      _etickets: etickets||"",
      _airlineBookingRef: airlineBookingRef||"",
      _checkInTime: checkInTime||"",
      _checkOutTime: checkOutTime||"",
      originCode: hasVuelo?(allIATA[0]||""):"",
      origin: hasVuelo?(iataToCity[allIATA[0]]||""):"",
      destinationCode: hasVuelo?(allIATA[1]||""):"",
      destination: hasVuelo?(iataToCity[allIATA[1]]||""):"",
      departureDate: hasVuelo?depDate:"",
      departureTime: hasVuelo?depTime:"",
      arrivalDate: hasVuelo?arrDate:"",
      arrivalTime: hasVuelo?arrTime:"",
      checkIn: checkInDate,
      checkOut: checkOutDate,
      nights,
      rooms: g([/habitaciones?[:\s]+(\d+)/i, /Room\s+(\d+):/i, /(\d+)\s+(?:room|habitaci)/i])||"1",
      roomType,
      regimen,
      baggage: g([/equipaje[:\s]+([^\n]+)/i, /\d+\s*(?:kg|piezas?)/i]),
      importantInfo: g([/Special Instructions?[:\s]+([^\n]{10,200})/i, /Important[:\s]+([^\n]{10,200})/i]),
    };
    // Limpiar vacíos
    Object.keys(data).forEach(k=>{ if(!data[k])delete data[k]; });
    setPreview(data);
  };

  const apply = () => {
    const toApply = {...preview};
    // Para hoteles, usar nombre como descripción si no hay descripción
    // Para hoteles NO usar el nombre como descripción — ya aparece como nombre del servicio
    if(toApply.type==="hotel"){
      toApply.description = "";
    }
    // Agregar horarios de check-in/out a la info importante
    if(toApply.type==="hotel" && (toApply._checkInTime||toApply._checkOutTime)){
      const tiempos = [
        toApply._checkInTime?`Check-in: ${toApply._checkInTime}`:"",
        toApply._checkOutTime?`Check-out: ${toApply._checkOutTime}`:"",
      ].filter(Boolean).join(" | ");
      toApply.importantInfo = tiempos + (toApply.importantInfo?"\n"+toApply.importantInfo:"");
    }
    // Agregar e-tickets y ref aerolínea a info importante del vuelo
    if(toApply.type==="vuelo"){
      const extras = [
        toApply._airlineBookingRef?`Ref. aerolínea: ${toApply._airlineBookingRef}`:"",
        toApply._etickets?`E-Tickets: ${toApply._etickets}`:"",
      ].filter(Boolean).join(" | ");
      if(extras) toApply.importantInfo = extras;
    }
    // NO borrar campos internos — los necesita el voucher para mostrar dirección y teléfono
    // delete toApply._extractedProviderPhone; -- comentado intencional
    delete toApply._etickets;
    delete toApply._airlineBookingRef;
    delete toApply._checkInTime;
    delete toApply._checkOutTime;
    onApply(toApply);
    setPreview(null);
    setOpen(false);
    setText("");
  };

  if(!open) return(
    <button onClick={()=>setOpen(true)} style={{...S.ghostBtn,background:"#F0F9FF",borderColor:"#BAE6FD",color:"#0369A1",fontWeight:600,marginBottom:10}}>
      <Upload size={13}/>📋 Autocompletar desde voucher del proveedor
    </button>
  );

  if(preview) return(
    <div style={{background:"#F0FDF4",border:"1px solid #86EFAC",borderRadius:10,padding:14,marginBottom:10}}>
      <div style={{fontSize:12,fontWeight:700,color:"#166534",marginBottom:10}}>✅ Datos detectados — revisá y confirmá:</div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:6,fontSize:12,marginBottom:12}}>
        {[
          ["Tipo", preview.type==="vuelo"?"✈ Vuelo":preview.type==="hotel"?"🏨 Hotel":preview.type==="traslado"?"🚌 Traslado":"Otro"],
          ["Hotel/Proveedor", preview._extractedProviderName],
          ["Dirección", preview._extractedProviderAddress],
          ["Teléfono hotel", preview._extractedProviderPhone],
          ["Referencia/Confirmación", preview.providerFileNumber],
          ["Ref. aerolínea", preview._airlineBookingRef],
          ["E-Tickets", preview._etickets],
          ["Check-in / Salida", preview.checkIn?fmtDate(preview.checkIn):(preview.departureDate?fmtDate(preview.departureDate):"")],
          ["Check-out / Llegada", preview.checkOut?fmtDate(preview.checkOut):(preview.arrivalDate?fmtDate(preview.arrivalDate):"")],
          ["Horario check-in", preview._checkInTime],
          ["Horario check-out", preview._checkOutTime],
          ["Noches", preview.nights],
          ["Habitaciones", preview.rooms],
          ["Tipo de habitación", preview.roomType],
          ["Régimen", preview.regimen],
          ["Vuelo", preview.flightNumber],
          ["Aerolínea", preview.airline],
          ["Origen", preview.originCode?`${preview.origin} (${preview.originCode})`:preview.origin],
          ["Destino", preview.destinationCode?`${preview.destination} (${preview.destinationCode})`:preview.destination],
          ["Hora salida", preview.departureTime],
          ["Hora llegada", preview.arrivalTime],
          ["Equipaje", preview.baggage],
        ].filter(([,v])=>v).map(([k,v])=>(
          <div key={k} style={{background:"white",borderRadius:6,padding:"5px 8px",border:"1px solid #D1FAE5",fontSize:11}}>
            <span style={{color:"#64748B",fontWeight:700}}>{k}:</span> {v}
          </div>
        ))}
      </div>
      {preview.importantInfo&&<div style={{background:"#FFF7ED",border:"1px solid #FED7AA",borderRadius:6,padding:"6px 10px",fontSize:11,color:"#92400E",marginBottom:10}}><strong>Info importante:</strong> {preview.importantInfo}</div>}
      <div style={{display:"flex",gap:8}}>
        <Btn onClick={apply}><Check size={13}/>Aplicar datos</Btn>
        <Btn variant="secondary" onClick={()=>setPreview(null)}>Volver</Btn>
        <Btn variant="secondary" onClick={()=>{setPreview(null);setOpen(false);setText("");}}>Cancelar</Btn>
      </div>
    </div>
  );

  return(
    <div style={{background:"#F0F9FF",border:"1px solid #BAE6FD",borderRadius:10,padding:14,marginBottom:10}}>
      <div style={{fontSize:12,fontWeight:700,color:"#0369A1",marginBottom:4}}>📋 Pegá el texto del voucher</div>
      <div style={{fontSize:11,color:"#64748B",marginBottom:8}}>Abrí el PDF → Ctrl+A → Ctrl+C → pegalo acá abajo → Extraer</div>
      <textarea value={text} onChange={e=>setText(e.target.value)} rows={5}
        placeholder="Pegá el texto acá..." style={{...S.input,resize:"vertical",marginBottom:8,fontSize:12}}/>
      <div style={{display:"flex",gap:8}}>
        <Btn onClick={extractAndPreview} disabled={!text.trim()}>Extraer datos</Btn>
        <Btn variant="secondary" onClick={()=>{setOpen(false);setText("");}}>Cancelar</Btn>
      </div>
    </div>
  );
}

function ServiceFields({s,arrUpd}){
  const up=(k,v)=>arrUpd("services",s.id,x=>({...x,[k]:v}));
  if(s.type==="vuelo") return(<div>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10,marginBottom:10}}>
      <Field label="Tipo de vuelo"><select value={s.flightType||""} onChange={e=>up("flightType",e.target.value)} style={{...S.input,appearance:"auto"}}><option value="">—</option><option>Ida</option><option>Regreso</option><option>Ida y vuelta</option><option>Escala</option></select></Field>
      <Inp label="N° de vuelo" value={s.flightNumber} onChange={v=>up("flightNumber",v)} placeholder="AR1234"/>
      <Inp label="Aerolínea" value={s.airline} onChange={v=>up("airline",v)} placeholder="Aerolíneas Argentinas"/>
    </div>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",gap:10,marginBottom:10}}>
      <Inp label="Ciudad origen" value={s.origin} onChange={v=>up("origin",v)} placeholder="Buenos Aires"/>
      <Inp label="Código IATA" value={s.originCode} onChange={v=>up("originCode",v.toUpperCase())} placeholder="EZE"/>
      <Inp label="Ciudad destino" value={s.destination} onChange={v=>up("destination",v)} placeholder="Cancún"/>
      <Inp label="Código IATA" value={s.destinationCode} onChange={v=>up("destinationCode",v.toUpperCase())} placeholder="CUN"/>
    </div>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",gap:10,marginBottom:10}}>
      <Inp label="Fecha salida" value={s.departureDate} onChange={v=>up("departureDate",v)} type="date"/>
      <Inp label="Hora salida" value={s.departureTime} onChange={v=>up("departureTime",v)} placeholder="08:30"/>
      <Inp label="Fecha llegada" value={s.arrivalDate} onChange={v=>up("arrivalDate",v)} type="date"/>
      <Inp label="Hora llegada" value={s.arrivalTime} onChange={v=>up("arrivalTime",v)} placeholder="14:45"/>
    </div>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",gap:10,marginBottom:10}}>
      <Field label="Escalas"><select value={s.stops||"Directo"} onChange={e=>up("stops",e.target.value)} style={{...S.input,appearance:"auto"}}><option>Directo</option><option>1 escala</option><option>2 escalas</option><option>3 escalas</option></select></Field>
      <Inp label="Duración" value={s.duration} onChange={v=>up("duration",v)} placeholder="5h 30m"/>
      <Field label="Clase"><select value={s.flightClass||""} onChange={e=>up("flightClass",e.target.value)} style={{...S.input,appearance:"auto"}}><option value="">—</option><option>Económica</option><option>Premium Economy</option><option>Business</option><option>Primera clase</option></select></Field>
      <Inp label="Terminal" value={s.terminal} onChange={v=>up("terminal",v)} placeholder="Terminal A"/>
    </div>
    <Field label="Equipaje incluido">
      <select value={s.baggage||""} onChange={e=>up("baggage",e.target.value)} style={{...S.input,appearance:"auto"}}>
        <option value="">-- Sin especificar --</option>
        <option value="Solo articulo personal">Solo articulo personal</option>
        <option value="Articulo personal + equipaje de mano">Articulo personal + equipaje de mano</option>
        <option value="1 pieza 23kg + equipaje de mano">1 pieza 23kg + equipaje de mano</option>
        <option value="2 piezas 23kg + equipaje de mano">2 piezas 23kg + equipaje de mano</option>
      </select>
    </Field>
  </div>);
  if(s.type==="hotel") return(<div>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",gap:10,marginBottom:10}}>
      <Inp label="Check-in" value={s.checkIn} onChange={v=>up("checkIn",v)} type="date"/>
      <Inp label="Check-out" value={s.checkOut} onChange={v=>up("checkOut",v)} type="date"/>
      <Field label="Noches"><input type="number" value={s.nights||""} onChange={e=>up("nights",e.target.value)} style={S.input} placeholder="0"/></Field>
      <Field label="Habitaciones"><input type="number" value={s.rooms||""} onChange={e=>up("rooms",e.target.value)} style={S.input} placeholder="1"/></Field>
    </div>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}>
      <Inp label="Tipo de habitación" value={s.roomType} onChange={v=>up("roomType",v)} placeholder="Suite doble con vista al mar"/>
      <Sel label="Régimen" value={s.regimen} onChange={v=>up("regimen",v)} options={REGIMENES}/>
    </div>
    <Txta label="Información importante" value={s.importantInfo} onChange={v=>up("importantInfo",v)} rows={2} placeholder="Políticas del hotel, requisitos de check-in, etc."/>
  </div>);
  return(<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
    <Inp label="Fecha del servicio" value={s.serviceDate} onChange={v=>up("serviceDate",v)} type="date"/>
    <Inp label="Hora del servicio" value={s.serviceTime} onChange={v=>up("serviceTime",v)} placeholder="10:00"/>
  </div>);
}

function ReservationModal({initial,providers,settings,onSave,onClose}){
  const [r,setR]=useState(initial||newRes(1,settings.defaultCommission));
  const [tab,setTab]=useState("general");
  const cur=settings.currency||'ARS';
  const upd=(k,v)=>setR(p=>({...p,[k]:v}));
  const arrUpd=(k,id,fn)=>setR(p=>({...p,[k]:p[k].map(x=>x.id===id?fn(x):x)}));
  const arrDel=(k,id)=>setR(p=>({...p,[k]:p[k].filter(x=>x.id!==id)}));
  const arrAdd=(k,v)=>setR(p=>({...p,[k]:[...p[k],v]}));
  const received=paidSum(r.paymentsReceived);
  const pending=(r.salePrice||0)-received;
  const totalCost=sum(r.services.map(s=>s.costPrice||0));
  const commission=(r.salePrice||0)-totalCost;
  const tabs=[{id:"general",label:"✈ General"},{id:"passengers",label:`👥 Pasajeros (${r.passengers.length})`},{id:"services",label:`🛎 Servicios (${r.services.length})`},{id:"payments",label:`💳 Cobros (${r.paymentsReceived.length})`}];
  return(<Modal title={`${initial?"Editar":"Nueva"} Reserva — File #${r.fileNumber}`} onClose={onClose} width={840} footer={<><Btn variant="secondary" onClick={onClose}>Cancelar</Btn><Btn onClick={()=>onSave(r)}>Guardar reserva</Btn></>}>
    <Tabs tabs={tabs} active={tab} setActive={setTab}/>
    {tab==="general"&&(<div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:12,marginBottom:12}}>
        <Inp label="N° de file" value={r.fileNumber} onChange={v=>upd("fileNumber",v)}/>
        <Sel label="Estado" value={r.status} onChange={v=>upd("status",v)} options={Object.entries(STATUS).map(([v,s])=>({v,l:s.label}))}/>
        <Inp label="Destino" value={r.destination} onChange={v=>upd("destination",v)} placeholder="Ej: Cancún, México"/>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:12,marginBottom:12}}>
        <Inp label="Fecha de salida" value={r.departureDate} onChange={v=>upd("departureDate",v)} type="date"/>
        <Inp label="Fecha de regreso" value={r.returnDate} onChange={v=>upd("returnDate",v)} type="date"/>
        <Field label="Días de viaje"><div style={{...S.input,background:"#F8FAFC",color:"#64748B",display:"flex",alignItems:"center"}}>{r.departureDate&&r.returnDate?Math.ceil((new Date(r.returnDate)-new Date(r.departureDate))/86400000)+" días":"—"}</div></Field>
      </div>
      <Txta label="Descripción" value={r.description} onChange={v=>upd("description",v)} placeholder="Descripción general..." style={{marginBottom:12}}/>
      <Divider label="Resumen financiero"/>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:12,marginBottom:12}}>
        <Field label={`Precio de venta (${cur})`}><input type="number" value={r.salePrice||""} onChange={e=>upd("salePrice",+e.target.value)} style={S.input} placeholder="0.00"/></Field>
        <Field label="% Comisión"><input type="number" value={r.commissionPercent||""} onChange={e=>upd("commissionPercent",+e.target.value)} style={S.input} placeholder="15"/></Field>
        <Field label="Costo total"><div style={{...S.input,background:"#F8FAFC",color:"#64748B",display:"flex",alignItems:"center"}}>{fmt$(totalCost,cur)}</div></Field>
      </div>
      <div style={{background:"#EFF6FF",border:"1px solid #BFDBFE",borderRadius:10,padding:14,display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12}}>
        {[{l:"Precio venta",v:fmt$(r.salePrice,cur),c:"#1E40AF"},{l:"Total cobrado",v:fmt$(received,cur),c:"#059669"},{l:"Saldo pendiente",v:fmt$(pending,cur),c:pending>0?"#D97706":"#059669"},{l:"Comisión neta",v:fmt$(commission,cur),c:"#7C3AED"}].map(x=>(<div key={x.l} style={{textAlign:"center"}}><div style={{fontSize:10,fontWeight:700,textTransform:"uppercase",letterSpacing:.5,color:"#64748B",marginBottom:4}}>{x.l}</div><div style={{fontSize:18,fontWeight:800,color:x.c}}>{x.v}</div></div>))}
      </div>
      <Txta label="Notas internas" value={r.notes} onChange={v=>upd("notes",v)} rows={2} placeholder="Observaciones internas..." style={{marginTop:12}}/>
    </div>)}
    {tab==="passengers"&&(<div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}><span style={{fontSize:13,color:"#64748B"}}>Todos los pasajeros de la reserva.</span><Btn size="sm" onClick={()=>arrAdd("passengers",newPass())}><Plus size={14}/>Agregar pasajero</Btn></div>
      {r.passengers.length===0&&<EmptyState icon={Users} title="Sin pasajeros" sub="Agregá al menos un pasajero."/>}
      {r.passengers.map((p,i)=>(<div key={p.id} style={{border:"1px solid #E2E8F0",borderRadius:10,padding:14,marginBottom:12}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}><span style={{fontSize:13,fontWeight:700}}>Pasajero {i+1}</span><button onClick={()=>arrDel("passengers",p.id)} style={{...S.ghostBtn,color:"#EF4444",borderColor:"#FCA5A5"}}><Trash2 size={13}/>Eliminar</button></div>
        <div style={{display:"grid",gridTemplateColumns:"2fr 1fr 1fr",gap:10,marginBottom:10}}>
          <Inp label="Nombre completo" value={p.name} onChange={v=>arrUpd("passengers",p.id,x=>({...x,name:v}))}/>
          <Inp label="DNI / Pasaporte" value={p.dni} onChange={v=>arrUpd("passengers",p.id,x=>({...x,dni:v}))}/>
          <Inp label="Fecha de nacimiento" value={p.birthDate} onChange={v=>arrUpd("passengers",p.id,x=>({...x,birthDate:v}))} type="date"/>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
          <Inp label="Email" value={p.email} onChange={v=>arrUpd("passengers",p.id,x=>({...x,email:v}))} type="email"/>
          <Inp label="Teléfono" value={p.phone} onChange={v=>arrUpd("passengers",p.id,x=>({...x,phone:v}))}/>
        </div>
      </div>))}
    </div>)}
    {tab==="services"&&(<div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}><span style={{fontSize:13,color:"#64748B"}}>Servicios contratados.</span><Btn size="sm" onClick={()=>arrAdd("services",newSvc())}><Plus size={14}/>Agregar servicio</Btn></div>
      {r.services.length===0&&<EmptyState icon={FileCheck} title="Sin servicios" sub="Agregá los servicios del viaje."/>}
      {r.services.map((s,i)=>{const SvcIcon=SVC_TYPES.find(t=>t.v===s.type)?.Icon||FileText;return(<div key={s.id} style={{border:"1px solid #E2E8F0",borderRadius:10,padding:14,marginBottom:12}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}><div style={{display:"flex",alignItems:"center",gap:8}}><SvcIcon size={15} style={{color:"#2563EB"}}/><span style={{fontSize:13,fontWeight:700}}>Servicio {i+1}</span></div><button onClick={()=>arrDel("services",s.id)} style={{...S.ghostBtn,color:"#EF4444",borderColor:"#FCA5A5"}}><Trash2 size={13}/>Eliminar</button></div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 2fr 1fr 1fr",gap:10,marginBottom:10}}>
          <Sel label="Tipo" value={s.type} onChange={v=>arrUpd("services",s.id,x=>({...x,type:v}))} options={SVC_TYPES.map(t=>({v:t.v,l:t.l}))}/>
          <Inp label="Descripción" value={s.description} onChange={v=>arrUpd("services",s.id,x=>({...x,description:v}))} placeholder="Descripción general"/>
          <Field label={`Costo neto (${cur})`}><input type="number" value={s.costPrice||""} onChange={e=>arrUpd("services",s.id,x=>({...x,costPrice:+e.target.value}))} style={S.input} placeholder="0.00"/></Field>
          <Field label={`Precio venta (${cur})`}><input type="number" value={s.salePrice||""} onChange={e=>arrUpd("services",s.id,x=>({...x,salePrice:+e.target.value}))} style={S.input} placeholder="0.00"/></Field>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:12}}>
          <Sel label="Proveedor" value={s.providerId} onChange={v=>arrUpd("services",s.id,x=>({...x,providerId:v}))} options={providers.map(p=>({v:p.id,l:p.name}))}/>
          <Inp label="File/Localizador proveedor" value={s.providerFileNumber} onChange={v=>arrUpd("services",s.id,x=>({...x,providerFileNumber:v}))} placeholder="N° file del proveedor"/>
        </div>
        <div style={{background:"#F8FAFC",borderRadius:8,padding:12,marginBottom:10}}>
          <VoucherReader svcId={s.id} svcType={s.type} providers={providers}
            onApply={data=>setR(prev=>({...prev,services:prev.services.map(sv=>sv.id===s.id?{...sv,...data}:sv)}))}/>
          <ServiceFields s={s} arrUpd={arrUpd}/>
        </div>
        <div style={{background:"#F0FDF4",borderRadius:8,padding:12}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}><span style={{fontSize:11,fontWeight:700,color:"#475569",textTransform:"uppercase",letterSpacing:.5}}>Pagos al proveedor</span><button onClick={()=>arrUpd("services",s.id,x=>({...x,paymentsDue:[...(x.paymentsDue||[]),newDue()]}))} style={S.ghostBtn}><Plus size={12}/>Agregar</button></div>
          {(s.paymentsDue||[]).length===0&&<p style={{fontSize:12,color:"#94A3B8",margin:0}}>Sin pagos programados.</p>}
          {(s.paymentsDue||[]).map((pd,j)=>(<div key={pd.id} style={{display:"grid",gridTemplateColumns:"1fr 1fr 2fr auto auto",gap:8,alignItems:"end",marginBottom:8}}>
            <Field label={j===0?"Vencimiento":""}><input type="date" value={pd.dueDate||""} onChange={e=>arrUpd("services",s.id,x=>({...x,paymentsDue:x.paymentsDue.map(p=>p.id===pd.id?{...p,dueDate:e.target.value}:p)}))} style={S.input}/></Field>
            <Field label={j===0?"Importe":""}><input type="number" value={pd.amount||""} onChange={e=>arrUpd("services",s.id,x=>({...x,paymentsDue:x.paymentsDue.map(p=>p.id===pd.id?{...p,amount:+e.target.value}:p)}))} style={S.input} placeholder="0.00"/></Field>
            <Field label={j===0?"Notas":""}><input type="text" value={pd.notes||""} onChange={e=>arrUpd("services",s.id,x=>({...x,paymentsDue:x.paymentsDue.map(p=>p.id===pd.id?{...p,notes:e.target.value}:p)}))} style={S.input} placeholder="Obs…"/></Field>
            <Field label={j===0?"Pagado":""}><div style={{display:"flex",alignItems:"center",height:36}}><input type="checkbox" checked={pd.paid||false} onChange={e=>arrUpd("services",s.id,x=>({...x,paymentsDue:x.paymentsDue.map(p=>p.id===pd.id?{...p,paid:e.target.checked,paidDate:e.target.checked?today():null}:p)}))} style={{width:16,height:16,cursor:"pointer"}}/></div></Field>
            <Field label={j===0?" ":""}><button onClick={()=>arrUpd("services",s.id,x=>({...x,paymentsDue:x.paymentsDue.filter(p=>p.id!==pd.id)}))} style={{...S.ghostBtn,color:"#EF4444",height:36}}><Trash2 size={13}/></button></Field>
          </div>))}
        </div>
      </div>);})}
    </div>)}
    {tab==="payments"&&(<div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}><span style={{fontSize:13,color:"#64748B"}}>Pagos recibidos del cliente.</span><Btn size="sm" onClick={()=>arrAdd("paymentsReceived",newPay())}><Plus size={14}/>Registrar pago</Btn></div>
      <div style={{background:"#ECFDF5",border:"1px solid #A7F3D0",borderRadius:10,padding:14,display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:12,marginBottom:16}}>
        {[{l:"Precio total",v:fmt$(r.salePrice,cur),c:"#1E293B"},{l:"Total cobrado",v:fmt$(received,cur),c:"#059669"},{l:"Saldo pendiente",v:fmt$(pending,cur),c:pending>0?"#D97706":"#059669"}].map(x=>(<div key={x.l} style={{textAlign:"center"}}><div style={{fontSize:10,fontWeight:700,textTransform:"uppercase",letterSpacing:.5,color:"#64748B",marginBottom:3}}>{x.l}</div><div style={{fontSize:20,fontWeight:800,color:x.c}}>{x.v}</div></div>))}
      </div>
      {r.paymentsReceived.length===0&&<EmptyState icon={CreditCard} title="Sin cobros" sub="Registrá los pagos recibidos."/>}
      {r.paymentsReceived.map((p,i)=>(<div key={p.id} style={{display:"grid",gridTemplateColumns:"1fr 1fr 2fr 1fr auto",gap:8,alignItems:"end",marginBottom:8}}>
        <Field label={i===0?"Fecha":""}><input type="date" value={p.date||""} onChange={e=>arrUpd("paymentsReceived",p.id,x=>({...x,date:e.target.value}))} style={S.input}/></Field>
        <Field label={i===0?"Forma de pago":""}><select value={p.method||""} onChange={e=>arrUpd("paymentsReceived",p.id,x=>({...x,method:e.target.value}))} style={{...S.input,appearance:"auto"}}><option value="">—</option>{PAY_METHODS.map(m=><option key={m} value={m}>{m}</option>)}</select></Field>
        <Field label={i===0?"Notas":""}><input type="text" value={p.notes||""} onChange={e=>arrUpd("paymentsReceived",p.id,x=>({...x,notes:e.target.value}))} style={S.input} placeholder="Obs…"/></Field>
        <Field label={i===0?`Importe (${cur})`:""} ><input type="number" value={p.amount||""} onChange={e=>arrUpd("paymentsReceived",p.id,x=>({...x,amount:+e.target.value}))} style={S.input} placeholder="0.00"/></Field>
        <Field label={i===0?" ":""}><button onClick={()=>arrDel("paymentsReceived",p.id)} style={{...S.ghostBtn,color:"#EF4444",height:36}}><Trash2 size={13}/></button></Field>
      </div>))}
    </div>)}
  </Modal>);
}

// ── RESERVATIONS PAGE ─────────────────────────────
function ReservationsPage({data,update}){
  const {reservations,providers,settings}=data;
  const [search,setSearch]=useState("");
  const [filter,setFilter]=useState("all");
  const [modal,setModal]=useState(null);
  const [viewModal,setViewModal]=useState(null);
  const cur=settings.currency||'ARS';
  const filtered=useMemo(()=>reservations.filter(r=>{const q=search.toLowerCase();const mQ=!q||(r.destination||"").toLowerCase().includes(q)||String(r.fileNumber).includes(q)||r.passengers.some(p=>p.name.toLowerCase().includes(q));return mQ&&r.status!=="cotizacion"&&(filter==="all"||r.status===filter);}).sort((a,b)=>b.createdAt?.localeCompare(a.createdAt||"")||0),[reservations,search,filter]);
  const saveRes=r=>{update(d=>{const idx=d.reservations.findIndex(x=>x.id===r.id);const updated=idx>=0?d.reservations.map(x=>x.id===r.id?r:x):[...d.reservations,r];return{...d,reservations:updated,nextFile:Math.max(d.nextFile,Number(r.fileNumber.replace(/^0+/,""))+1)};});setModal(null);};
  const terr=r=>r.services.filter(s=>s.type!=='vuelo');
  return(<div style={{padding:28,overflowY:"auto",flex:1}}>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20}}>
      <div><h1 style={{margin:0,fontSize:22,fontWeight:800,color:"#0F172A"}}>Reservas</h1><p style={{margin:"4px 0 0",fontSize:13,color:"#64748B"}}>{filtered.length} reservas</p></div>
      <div style={{display:"flex",gap:8}}><Btn variant="success" onClick={()=>exportReservations(reservations,settings)}><Download size={14}/>Excel</Btn><Btn onClick={()=>setModal({isNew:true,data:newRes(data.nextFile,settings.defaultCommission)})}><Plus size={14}/>Nueva reserva</Btn></div>
    </div>
    <Card style={{marginBottom:16}}><div style={{display:"flex",gap:10}}><div style={{flex:1,position:"relative"}}><Search size={14} style={{position:"absolute",left:10,top:"50%",transform:"translateY(-50%)",color:"#94A3B8"}}/><input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Buscar por file, destino o pasajero…" style={{...S.input,paddingLeft:32}}/></div><select value={filter} onChange={e=>setFilter(e.target.value)} style={{...S.input,width:"auto",appearance:"auto"}}><option value="all">Todos los estados</option>{Object.entries(STATUS).filter(([v])=>v!=="cotizacion").map(([v,s])=><option key={v} value={v}>{s.label}</option>)}</select></div></Card>
    {filtered.length===0?<EmptyState icon={CalendarDays} title="Sin reservas" sub="Creá tu primera reserva." action={<Btn onClick={()=>setModal({isNew:true,data:newRes(data.nextFile,settings.defaultCommission)})}><Plus size={14}/>Nueva reserva</Btn>}/>
      :<Card style={{padding:0}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
        <thead><tr style={{background:"#F8FAFC",borderBottom:"1px solid #E2E8F0"}}>{["File","Destino","Pasajeros","Salida","Estado","Venta","Cobrado","Saldo","Acciones"].map(h=>(<th key={h} style={{padding:"10px 12px",textAlign:"left",fontWeight:700,fontSize:11,color:"#64748B",textTransform:"uppercase"}}>{h}</th>))}</tr></thead>
        <tbody>{filtered.map(r=>{const rec=paidSum(r.paymentsReceived);const pend=(r.salePrice||0)-rec;const t=terr(r);return(<tr key={r.id} style={{borderBottom:"1px solid #F1F5F9"}} onMouseEnter={e=>e.currentTarget.style.background="#F8FAFC"} onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
          <td style={{padding:"10px 12px",fontWeight:700,color:"#2563EB"}}>#{r.fileNumber}</td>
          <td style={{padding:"10px 12px",fontWeight:500}}>{r.destination||"—"}</td>
          <td style={{padding:"10px 12px",color:"#64748B",fontSize:12,maxWidth:160}}>{r.passengers.map(p=>p.name).filter(Boolean).join(", ")||"—"}</td>
          <td style={{padding:"10px 12px",color:"#64748B"}}>{fmtDate(r.departureDate)}</td>
          <td style={{padding:"10px 12px"}}><Badge status={r.status}/></td>
          <td style={{padding:"10px 12px",fontWeight:600}}>{fmt$(r.salePrice,cur)}</td>
          <td style={{padding:"10px 12px",color:"#059669",fontWeight:600}}>{fmt$(rec,cur)}</td>
          <td style={{padding:"10px 12px",color:pend>0?"#D97706":"#059669",fontWeight:700}}>{fmt$(pend,cur)}</td>
          <td style={{padding:"10px 12px"}}><div style={{display:"flex",gap:3,flexWrap:"wrap"}}>
            <button title="Ver" onClick={()=>setViewModal(r)} style={S.ghostBtn}><Eye size={13}/></button>
            <button title="Editar" onClick={()=>setModal({isNew:false,data:r})} style={S.ghostBtn}><Pencil size={13}/></button>
            <button title="Voucher Aéreo" onClick={()=>printVoucherAereo(r,settings,providers)} style={{...S.ghostBtn,color:"#2563EB",borderColor:"#BFDBFE"}}><Plane size={13}/></button>
            {t.length===0&&<button title="Voucher Terrestre" onClick={()=>printVoucherTerrestre(r,settings,providers)} style={{...S.ghostBtn,color:"#059669",borderColor:"#A7F3D0"}}><Hotel size={13}/></button>}
            {t.length===1&&<button title={`Voucher: ${t[0].description||t[0].type}`} onClick={()=>printVoucherTerrestre(r,settings,providers)} style={{...S.ghostBtn,color:"#059669",borderColor:"#A7F3D0"}}><Hotel size={13}/></button>}
            {t.length>1&&t.map((s,i)=>(<button key={s.id} title={`Voucher ${i+1}: ${s.description||s.type}`} onClick={()=>printVoucherTerrestre(r,settings,providers,i)} style={{...S.ghostBtn,color:"#059669",borderColor:"#A7F3D0"}}><Hotel size={13}/>{i+1}</button>))}
            <button title="Recibo" onClick={()=>printReceipt(r,settings,providers)} style={S.ghostBtn}><Printer size={13}/></button>
            <button title="Eliminar" onClick={()=>{if(window.confirm("¿Eliminar esta reserva?"))update(d=>({...d,reservations:d.reservations.filter(x=>x.id!==r.id)}))}} style={{...S.ghostBtn,color:"#EF4444",borderColor:"#FCA5A5"}}><Trash2 size={13}/></button>
          </div></td>
        </tr>);})}</tbody>
      </table></Card>}
    {modal&&<ReservationModal initial={modal.isNew?null:modal.data} providers={providers} settings={settings} onSave={saveRes} onClose={()=>setModal(null)}/>}
    {viewModal&&(<Modal title={`File #${viewModal.fileNumber} — ${viewModal.destination||"Sin destino"}`} onClose={()=>setViewModal(null)} width={640} footer={<><Btn variant="secondary" onClick={()=>{setViewModal(null);setModal({isNew:false,data:viewModal});}}>Editar</Btn><Btn variant="secondary" onClick={()=>printVoucherAereo(viewModal,settings,providers)}><Plane size={13}/>V. Aéreo</Btn><Btn variant="secondary" onClick={()=>printVoucherTerrestre(viewModal,settings,providers)}><Hotel size={13}/>V. Terrestre</Btn><Btn onClick={()=>printReceipt(viewModal,settings,providers)}><Printer size={13}/>Recibo</Btn></>}>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:16}}>{[{l:"Estado",v:<Badge status={viewModal.status}/>},{l:"Destino",v:viewModal.destination||"—"},{l:"Salida",v:fmtDate(viewModal.departureDate)},{l:"Regreso",v:fmtDate(viewModal.returnDate)},{l:"Precio venta",v:fmt$(viewModal.salePrice,cur)},{l:"Saldo",v:fmt$((viewModal.salePrice||0)-paidSum(viewModal.paymentsReceived),cur)}].map(x=>(<div key={x.l}><div style={S.label}>{x.l}</div><div style={{fontSize:14,fontWeight:500}}>{x.v}</div></div>))}</div>
      <Divider label="Pasajeros"/>
      <div style={{display:"flex",flexWrap:"wrap",gap:8}}>{viewModal.passengers.map(p=>(<div key={p.id} style={{background:"#EFF6FF",color:"#1D4ED8",border:"1px solid #BFDBFE",borderRadius:8,padding:"6px 12px",fontSize:12,fontWeight:600}}>{p.name}{p.dni?" · "+p.dni:""}</div>))}</div>
    </Modal>)}
  </div>);
}

// ── CLIENTS ───────────────────────────────────────
function ClientsPage({data,update}){
  const {clients,reservations,settings}=data;
  const [search,setSearch]=useState("");
  const [modal,setModal]=useState(null);
  const filtered=clients.filter(c=>{const q=search.toLowerCase();return !q||c.name.toLowerCase().includes(q)||(c.email||"").toLowerCase().includes(q);});
  const saveClient=c=>{update(d=>{const idx=d.clients.findIndex(x=>x.id===c.id);return{...d,clients:idx>=0?d.clients.map(x=>x.id===c.id?c:x):[...d.clients,c]};});setModal(null);};
  const clientRes=c=>reservations.filter(r=>r.passengers.some(p=>p.name&&c.name&&p.name.toLowerCase()===c.name.toLowerCase()));
  return(<div style={{padding:28,overflowY:"auto",flex:1}}>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20}}><div><h1 style={{margin:0,fontSize:22,fontWeight:800,color:"#0F172A"}}>Clientes</h1><p style={{margin:"4px 0 0",fontSize:13,color:"#64748B"}}>{filtered.length} clientes</p></div><Btn onClick={()=>setModal({id:genId(),name:"",email:"",phone:"",dni:"",address:"",notes:""})}><Plus size={14}/>Nuevo cliente</Btn></div>
    <Card style={{marginBottom:16}}><div style={{position:"relative"}}><Search size={14} style={{position:"absolute",left:10,top:"50%",transform:"translateY(-50%)",color:"#94A3B8"}}/><input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Buscar clientes…" style={{...S.input,paddingLeft:32}}/></div></Card>
    {filtered.length===0?<EmptyState icon={Users} title="Sin clientes" sub="Agregá tu primer cliente."/>
      :<Card style={{padding:0}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
        <thead><tr style={{background:"#F8FAFC",borderBottom:"1px solid #E2E8F0"}}>{["Nombre","Email","Teléfono","DNI","Reservas","Acciones"].map(h=>(<th key={h} style={{padding:"10px 12px",textAlign:"left",fontWeight:700,fontSize:11,color:"#64748B",textTransform:"uppercase"}}>{h}</th>))}</tr></thead>
        <tbody>{filtered.map(c=>(<tr key={c.id} style={{borderBottom:"1px solid #F1F5F9"}} onMouseEnter={e=>e.currentTarget.style.background="#F8FAFC"} onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
          <td style={{padding:"10px 12px",fontWeight:600}}>{c.name}</td><td style={{padding:"10px 12px",color:"#64748B"}}>{c.email||"—"}</td><td style={{padding:"10px 12px",color:"#64748B"}}>{c.phone||"—"}</td><td style={{padding:"10px 12px",color:"#64748B"}}>{c.dni||"—"}</td>
          <td style={{padding:"10px 12px"}}><span style={{background:"#EFF6FF",color:"#1D4ED8",padding:"2px 8px",borderRadius:12,fontSize:11,fontWeight:700}}>{clientRes(c).length}</span></td>
          <td style={{padding:"10px 12px"}}><div style={{display:"flex",gap:4}}><button onClick={()=>setModal(c)} style={S.ghostBtn}><Pencil size={13}/></button><button onClick={()=>exportClientStatement(c,reservations,settings)} style={{...S.ghostBtn,color:"#059669",borderColor:"#A7F3D0"}}><Download size={13}/>Excel</button><button onClick={()=>{if(window.confirm("¿Eliminar?"))update(d=>({...d,clients:d.clients.filter(x=>x.id!==c.id)}))}} style={{...S.ghostBtn,color:"#EF4444",borderColor:"#FCA5A5"}}><Trash2 size={13}/></button></div></td>
        </tr>))}</tbody>
      </table></Card>}
    {modal&&(<Modal title={modal.name?"Editar cliente":"Nuevo cliente"} onClose={()=>setModal(null)} footer={<><Btn variant="secondary" onClick={()=>setModal(null)}>Cancelar</Btn><Btn onClick={()=>saveClient(modal)}>Guardar</Btn></>}>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
        <Inp label="Nombre completo" value={modal.name} onChange={v=>setModal(p=>({...p,name:v}))} col={2} style={{gridColumn:"span 2"}}/>
        <Inp label="Email" value={modal.email} onChange={v=>setModal(p=>({...p,email:v}))} type="email"/>
        <Inp label="Teléfono" value={modal.phone} onChange={v=>setModal(p=>({...p,phone:v}))}/>
        <Inp label="DNI / Pasaporte" value={modal.dni} onChange={v=>setModal(p=>({...p,dni:v}))}/>
        <Inp label="Dirección" value={modal.address} onChange={v=>setModal(p=>({...p,address:v}))}/>
        <Txta label="Notas" value={modal.notes} onChange={v=>setModal(p=>({...p,notes:v}))} col={2} style={{gridColumn:"span 2"}}/>
      </div>
    </Modal>)}
  </div>);
}

// ── PROVIDERS ─────────────────────────────────────
function ProvidersPage({data,update}){
  const {providers,reservations,settings}=data;
  const [search,setSearch]=useState("");
  const [modal,setModal]=useState(null);
  const cur=settings.currency||'ARS';
  const filtered=providers.filter(p=>{const q=search.toLowerCase();return !q||p.name.toLowerCase().includes(q)||(p.category||"").toLowerCase().includes(q);});
  const saveProv=p=>{update(d=>{const idx=d.providers.findIndex(x=>x.id===p.id);return{...d,providers:idx>=0?d.providers.map(x=>x.id===p.id?p:x):[...d.providers,p]};});setModal(null);};
  const provDebt=p=>{let total=0,pagado=0;reservations.forEach(r=>r.services.filter(s=>s.providerId===p.id).forEach(s=>{total+=s.costPrice||0;pagado+=sum((s.paymentsDue||[]).filter(x=>x.paid).map(x=>x.amount||0));}));return{total,pagado,pendiente:total-pagado};};
  return(<div style={{padding:28,overflowY:"auto",flex:1}}>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20}}><div><h1 style={{margin:0,fontSize:22,fontWeight:800,color:"#0F172A"}}>Proveedores</h1><p style={{margin:"4px 0 0",fontSize:13,color:"#64748B"}}>{filtered.length} proveedores</p></div><Btn onClick={()=>setModal({id:genId(),name:"",category:"",email:"",phone:"",address:"",cuit:"",notes:""})}><Plus size={14}/>Nuevo proveedor</Btn></div>
    <Card style={{marginBottom:16}}><div style={{position:"relative"}}><Search size={14} style={{position:"absolute",left:10,top:"50%",transform:"translateY(-50%)",color:"#94A3B8"}}/><input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Buscar proveedores…" style={{...S.input,paddingLeft:32}}/></div></Card>
    {filtered.length===0?<EmptyState icon={Building2} title="Sin proveedores" sub="Agregá tu primer proveedor."/>
      :<Card style={{padding:0}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
        <thead><tr style={{background:"#F8FAFC",borderBottom:"1px solid #E2E8F0"}}>{["Nombre","Categoría","Contacto","Deuda Total","Pagado","Pendiente","Acciones"].map(h=>(<th key={h} style={{padding:"10px 12px",textAlign:"left",fontWeight:700,fontSize:11,color:"#64748B",textTransform:"uppercase"}}>{h}</th>))}</tr></thead>
        <tbody>{filtered.map(p=>{const d=provDebt(p);return(<tr key={p.id} style={{borderBottom:"1px solid #F1F5F9"}} onMouseEnter={e=>e.currentTarget.style.background="#F8FAFC"} onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
          <td style={{padding:"10px 12px",fontWeight:600}}>{p.name}</td><td style={{padding:"10px 12px"}}>{p.category&&<span style={{background:"#F1F5F9",color:"#475569",padding:"2px 8px",borderRadius:12,fontSize:11}}>{p.category}</span>}</td><td style={{padding:"10px 12px",color:"#64748B",fontSize:12}}>{p.email||p.phone||"—"}</td>
          <td style={{padding:"10px 12px",fontWeight:600}}>{fmt$(d.total,cur)}</td><td style={{padding:"10px 12px",color:"#059669",fontWeight:600}}>{fmt$(d.pagado,cur)}</td><td style={{padding:"10px 12px",color:d.pendiente>0?"#EF4444":"#059669",fontWeight:700}}>{fmt$(d.pendiente,cur)}</td>
          <td style={{padding:"10px 12px"}}><div style={{display:"flex",gap:4}}><button onClick={()=>setModal(p)} style={S.ghostBtn}><Pencil size={13}/></button><button onClick={()=>exportProviderStatement(p,reservations,settings)} style={{...S.ghostBtn,color:"#059669",borderColor:"#A7F3D0"}}><Download size={13}/>Excel</button><button onClick={()=>{if(window.confirm("¿Eliminar?"))update(d=>({...d,providers:d.providers.filter(x=>x.id!==p.id)}))}} style={{...S.ghostBtn,color:"#EF4444",borderColor:"#FCA5A5"}}><Trash2 size={13}/></button></div></td>
        </tr>);})}</tbody>
      </table></Card>}
    {modal&&(<Modal title={modal.name?"Editar proveedor":"Nuevo proveedor"} onClose={()=>setModal(null)} footer={<><Btn variant="secondary" onClick={()=>setModal(null)}>Cancelar</Btn><Btn onClick={()=>saveProv(modal)}>Guardar</Btn></>}>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
        <Inp label="Nombre" value={modal.name} onChange={v=>setModal(p=>({...p,name:v}))} col={2} style={{gridColumn:"span 2"}}/>
        <Sel label="Categoría" value={modal.category} onChange={v=>setModal(p=>({...p,category:v}))} options={PROV_CATS}/>
        <Inp label="CUIT" value={modal.cuit} onChange={v=>setModal(p=>({...p,cuit:v}))}/>
        <Inp label="Email" value={modal.email} onChange={v=>setModal(p=>({...p,email:v}))} type="email"/>
        <Inp label="Teléfono" value={modal.phone} onChange={v=>setModal(p=>({...p,phone:v}))}/>
        <Inp label="Dirección" value={modal.address} onChange={v=>setModal(p=>({...p,address:v}))} col={2} style={{gridColumn:"span 2"}}/>
        <Txta label="Notas" value={modal.notes} onChange={v=>setModal(p=>({...p,notes:v}))} col={2} style={{gridColumn:"span 2"}}/>
      </div>
    </Modal>)}
  </div>);
}

// ── COMMISSIONS ───────────────────────────────────
function CommissionsPage({data}){
  const {reservations,settings}=data;
  const [tab,setTab]=useState("summary");
  const cur=settings.currency||'ARS';
  const nc=reservations.filter(r=>r.status!=="cancelada"&&r.status!=="cotizacion");
  const totalCost=sum(nc.map(r=>sum(r.services.map(s=>s.costPrice||0))));
  const totalComm=totalSale-totalCost;
  const totalReceived=sum(nc.map(r=>paidSum(r.paymentsReceived)));
  return(<div style={{padding:28,overflowY:"auto",flex:1}}>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20}}><h1 style={{margin:0,fontSize:22,fontWeight:800,color:"#0F172A"}}>Comisiones</h1><Btn variant="success" onClick={()=>exportCommissions(reservations,settings)}><Download size={14}/>Exportar Excel</Btn></div>
    <Grid cols={4} gap={14} style={{marginBottom:20}}>{[{l:"Ingresos",v:fmt$(totalSale,cur),c:"#059669",bg:"#ECFDF5"},{l:"Costos",v:fmt$(totalCost,cur),c:"#EF4444",bg:"#FEF2F2"},{l:"Comisión neta",v:fmt$(totalComm,cur),c:"#7C3AED",bg:"#F5F3FF"},{l:"Cobrado",v:fmt$(totalReceived,cur),c:"#2563EB",bg:"#EFF6FF"}].map(x=>(<div key={x.l} style={{background:x.bg,borderRadius:12,padding:16}}><div style={{fontSize:11,fontWeight:700,color:x.c,textTransform:"uppercase",letterSpacing:.5,marginBottom:8}}>{x.l}</div><div style={{fontSize:22,fontWeight:800,color:x.c}}>{x.v}</div></div>))}</Grid>
    <Tabs tabs={[{id:"summary",label:"Resumen"},{id:"detail",label:"Detalle"},{id:"pending",label:"Saldos pendientes"}]} active={tab} setActive={setTab}/>
    {tab==="summary"&&<Card><Grid cols={3} gap={14}>{Object.entries(STATUS).map(([k,s])=>(<div key={k} style={{background:s.bg,borderRadius:10,padding:16,textAlign:"center"}}><div style={{fontSize:11,fontWeight:700,textTransform:"uppercase",color:s.c,marginBottom:8}}>{s.label}</div><div style={{fontSize:28,fontWeight:800,color:s.c}}>{reservations.filter(r=>r.status===k).length}</div></div>))}</Grid></Card>}
    {tab==="detail"&&<Card style={{padding:0}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}><thead><tr style={{background:"#F8FAFC",borderBottom:"1px solid #E2E8F0"}}>{["File","Destino","Pasajero","Estado","Venta","Costo","Comisión","% Comisión"].map(h=>(<th key={h} style={{padding:"9px 12px",textAlign:"left",fontWeight:700,fontSize:11,color:"#64748B",textTransform:"uppercase"}}>{h}</th>))}</tr></thead><tbody>{nc.map(r=>{const costo=sum(r.services.map(s=>s.costPrice||0));const comision=(r.salePrice||0)-costo;const pct=r.salePrice?((comision/r.salePrice)*100).toFixed(1):0;return(<tr key={r.id} style={{borderBottom:"1px solid #F1F5F9"}} onMouseEnter={e=>e.currentTarget.style.background="#F8FAFC"} onMouseLeave={e=>e.currentTarget.style.background="transparent"}><td style={{padding:"9px 12px",fontWeight:700,color:"#2563EB"}}>#{r.fileNumber}</td><td style={{padding:"9px 12px"}}>{r.destination||"—"}</td><td style={{padding:"9px 12px",color:"#64748B"}}>{r.passengers[0]?.name||"—"}</td><td style={{padding:"9px 12px"}}><Badge status={r.status}/></td><td style={{padding:"9px 12px",fontWeight:600}}>{fmt$(r.salePrice,cur)}</td><td style={{padding:"9px 12px",color:"#EF4444"}}>{fmt$(costo,cur)}</td><td style={{padding:"9px 12px",color:"#7C3AED",fontWeight:700}}>{fmt$(comision,cur)}</td><td style={{padding:"9px 12px"}}><span style={{background:"#F5F3FF",color:"#7C3AED",padding:"2px 8px",borderRadius:12,fontSize:11,fontWeight:700}}>{pct}%</span></td></tr>);})}</tbody></table></Card>}
    {tab==="pending"&&<Card style={{padding:0}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}><thead><tr style={{background:"#F8FAFC",borderBottom:"1px solid #E2E8F0"}}>{["File","Pasajero","Destino","Salida","Venta","Cobrado","Saldo"].map(h=>(<th key={h} style={{padding:"9px 12px",textAlign:"left",fontWeight:700,fontSize:11,color:"#64748B",textTransform:"uppercase"}}>{h}</th>))}</tr></thead><tbody>{reservations.filter(r=>r.status!=="cancelada"&&r.status!=="cotizacion"&&(r.salePrice||0)-paidSum(r.paymentsReceived)>0).map(r=>{const rec=paidSum(r.paymentsReceived);const pend=(r.salePrice||0)-rec;return(<tr key={r.id} style={{borderBottom:"1px solid #F1F5F9"}} onMouseEnter={e=>e.currentTarget.style.background="#F8FAFC"} onMouseLeave={e=>e.currentTarget.style.background="transparent"}><td style={{padding:"9px 12px",fontWeight:700,color:"#2563EB"}}>#{r.fileNumber}</td><td style={{padding:"9px 12px"}}>{r.passengers[0]?.name||"—"}</td><td style={{padding:"9px 12px"}}>{r.destination||"—"}</td><td style={{padding:"9px 12px",color:"#64748B"}}>{fmtDate(r.departureDate)}</td><td style={{padding:"9px 12px",fontWeight:600}}>{fmt$(r.salePrice,cur)}</td><td style={{padding:"9px 12px",color:"#059669",fontWeight:600}}>{fmt$(rec,cur)}</td><td style={{padding:"9px 12px",color:"#D97706",fontWeight:700}}>{fmt$(pend,cur)}</td></tr>);})}</tbody></table></Card>}
  </div>);
}

// ── REPORTS ───────────────────────────────────────
function ReportsPage({data}){
  const {reservations,clients,providers,settings}=data;
  const [selectedClient,setSelectedClient]=useState("");
  const [selectedProvider,setSelectedProvider]=useState("");
  const [dateFrom,setDateFrom]=useState("");
  const [dateTo,setDateTo]=useState("");
  const cur=settings.currency||'ARS';
  const filteredRes=useMemo(()=>reservations.filter(r=>{const af=!dateFrom||(r.departureDate&&r.departureDate>=dateFrom);const at=!dateTo||(r.departureDate&&r.departureDate<=dateTo);return af&&at;}),[reservations,dateFrom,dateTo]);
  const totalVenta=sum(filteredRes.filter(r=>r.status!=='cancelada').map(r=>r.salePrice||0));
  const totalCobrado=sum(filteredRes.map(r=>paidSum(r.paymentsReceived)));
  const totalCosto=sum(filteredRes.filter(r=>r.status!=='cancelada').map(r=>sum(r.services.map(s=>s.costPrice||0))));
  const clientObj=clients.find(c=>c.id===selectedClient);
  const provObj=providers.find(p=>p.id===selectedProvider);
  return(<div style={{padding:28,overflowY:"auto",flex:1}}>
    <h1 style={{margin:"0 0 20px",fontSize:22,fontWeight:800,color:"#0F172A"}}>Reportes</h1>
    <Card style={{marginBottom:16}}>
      <h3 style={{...S.sectionTitle,marginBottom:14}}>Filtro por período</h3>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr auto auto",gap:12,alignItems:"end",marginBottom:16}}>
        <Field label="Desde"><input type="date" value={dateFrom} onChange={e=>setDateFrom(e.target.value)} style={S.input}/></Field>
        <Field label="Hasta"><input type="date" value={dateTo} onChange={e=>setDateTo(e.target.value)} style={S.input}/></Field>
        <Btn variant="secondary" onClick={()=>{setDateFrom("");setDateTo("");}}>Limpiar</Btn>
        <Btn variant="success" onClick={()=>exportReservations(filteredRes,settings)}><Download size={14}/>Exportar Excel</Btn>
      </div>
      <Grid cols={4} gap={12}>{[{l:"Reservas",v:filteredRes.length,c:"#2563EB",bg:"#EFF6FF"},{l:"Ventas",v:fmt$(totalVenta,cur),c:"#059669",bg:"#ECFDF5"},{l:"Cobrado",v:fmt$(totalCobrado,cur),c:"#7C3AED",bg:"#F5F3FF"},{l:"Comisión",v:fmt$(totalVenta-totalCosto,cur),c:"#D97706",bg:"#FFFBEB"}].map(x=>(<div key={x.l} style={{background:x.bg,borderRadius:10,padding:14,textAlign:"center"}}><div style={{fontSize:11,fontWeight:700,color:x.c,textTransform:"uppercase",marginBottom:6}}>{x.l}</div><div style={{fontSize:18,fontWeight:800,color:x.c}}>{x.v}</div></div>))}</Grid>
    </Card>
    <Grid cols={2} gap={16} style={{marginBottom:16}}>
      <Card><h3 style={{...S.sectionTitle,marginBottom:14}}><Users size={16} style={{color:"#2563EB"}}/>Estado de cuenta — Clientes</h3><Field label="Seleccioná un cliente"><select value={selectedClient} onChange={e=>setSelectedClient(e.target.value)} style={{...S.input,appearance:"auto",marginBottom:12}}><option value="">— Elegir cliente —</option>{clients.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}</select></Field><Btn variant="success" disabled={!selectedClient} onClick={()=>clientObj&&exportClientStatement(clientObj,reservations,settings)}><Download size={14}/>Descargar Excel</Btn></Card>
      <Card><h3 style={{...S.sectionTitle,marginBottom:14}}><Building2 size={16} style={{color:"#7C3AED"}}/>Estado de cuenta — Proveedores</h3><Field label="Seleccioná un proveedor"><select value={selectedProvider} onChange={e=>setSelectedProvider(e.target.value)} style={{...S.input,appearance:"auto",marginBottom:12}}><option value="">— Elegir proveedor —</option>{providers.map(p=><option key={p.id} value={p.id}>{p.name}</option>)}</select></Field><Btn style={{background:"#7C3AED"}} disabled={!selectedProvider} onClick={()=>provObj&&exportProviderStatement(provObj,reservations,settings)}><Download size={14}/>Descargar Excel</Btn></Card>
    </Grid>
    <Card><h3 style={{...S.sectionTitle,marginBottom:14}}><BarChart2 size={16} style={{color:"#D97706"}}/>Reportes generales</h3><Grid cols={3} gap={12}>{[{title:"Todas las reservas",desc:"Listado completo con fechas, estados y saldos.",color:"#2563EB",onClick:()=>exportReservations(reservations,settings)},{title:"Comisiones",desc:"Comisión neta por reserva (excluye canceladas).",color:"#7C3AED",onClick:()=>exportCommissions(reservations,settings)},{title:"Saldos pendientes",desc:"Reservas con saldo pendiente por cobrar.",color:"#D97706",onClick:()=>exportReservations(reservations.filter(r=>r.status!=='cancelada'&&(r.salePrice||0)-paidSum(r.paymentsReceived)>0),settings)}].map(x=>(<div key={x.title} style={{border:"1px solid #E2E8F0",borderRadius:10,padding:16}}><div style={{fontSize:13,fontWeight:700,color:x.color,marginBottom:6}}>{x.title}</div><div style={{fontSize:12,color:"#64748B",marginBottom:12}}>{x.desc}</div><Btn variant="secondary" size="sm" onClick={x.onClick}><Download size={12}/>Descargar Excel</Btn></div>))}</Grid></Card>
  </div>);
}

// ── SETTINGS ──────────────────────────────────────
function SettingsPage({data,update}){
  const [s,setS]=useState(data.settings);
  const [saved,setSaved]=useState(false);
  const [docTab,setDocTab]=useState("recibo");
  const fileRef=useRef();
  const save=()=>{update(d=>({...d,settings:s}));setSaved(true);setTimeout(()=>setSaved(false),2000);};
  const handleLogo=e=>{const file=e.target.files[0];if(!file)return;const reader=new FileReader();reader.onload=ev=>setS(p=>({...p,logo:ev.target.result}));reader.readAsDataURL(file);};
  const getDc=(doc,key)=>{const def=DEFAULT_DOC_CONFIG[doc]?.[key];const val=s.docConfig?.[doc]?.[key];return val===undefined?def:val;};
  const setDc=(doc,key,val)=>setS(p=>({...p,docConfig:{...p.docConfig,[doc]:{...p.docConfig?.[doc],[key]:val}}}));
  const docTabs=[{id:"recibo",label:"🧾 Recibo"},{id:"voucherAereo",label:"✈ V. Aéreo"},{id:"voucherTerrestre",label:"🏨 V. Terrestre"}];
  return(<div style={{padding:28,maxWidth:720,overflowY:"auto",flex:1}}>
    <h1 style={{margin:"0 0 20px",fontSize:22,fontWeight:800,color:"#0F172A"}}>Configuración</h1>
    <Card style={{marginBottom:16}}>
      <h3 style={{...S.sectionTitle,marginBottom:16}}>Datos de la agencia</h3>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
        <Inp label="Nombre de la agencia" value={s.agencyName} onChange={v=>setS(p=>({...p,agencyName:v}))} col={2} style={{gridColumn:"span 2"}}/>
        <Inp label="Dirección" value={s.address} onChange={v=>setS(p=>({...p,address:v}))} col={2} style={{gridColumn:"span 2"}}/>
        <Inp label="Teléfono" value={s.phone} onChange={v=>setS(p=>({...p,phone:v}))}/>
        <Inp label="Email" value={s.email} onChange={v=>setS(p=>({...p,email:v}))} type="email"/>
        <Inp label="CUIT" value={s.cuit} onChange={v=>setS(p=>({...p,cuit:v}))}/>
        <Inp label="Teléfono de emergencias" value={s.emergencyPhone} onChange={v=>setS(p=>({...p,emergencyPhone:v}))} placeholder="+54 9 11 0000-0000"/>
        <Field label="% Comisión por defecto"><input type="number" value={s.defaultCommission||""} onChange={e=>setS(p=>({...p,defaultCommission:+e.target.value}))} style={S.input} placeholder="15"/></Field>
        <Sel label="Moneda principal" value={s.currency||"ARS"} onChange={v=>setS(p=>({...p,currency:v}))} options={CURRENCIES}/>
      </div>
    </Card>
    <Card style={{marginBottom:16}}>
      <h3 style={{...S.sectionTitle,marginBottom:14}}>Logo de la agencia</h3>
      <div style={{display:"flex",alignItems:"center",gap:16}}>
        {s.logo?<img src={s.logo} style={{height:60,maxWidth:160,objectFit:"contain",border:"1px solid #E2E8F0",borderRadius:8,padding:8}} alt="Logo"/>:<div style={{width:80,height:60,background:"#F1F5F9",borderRadius:8,display:"flex",alignItems:"center",justifyContent:"center",color:"#94A3B8",fontSize:12}}>Sin logo</div>}
        <div><Btn variant="secondary" onClick={()=>fileRef.current.click()}><Upload size={14}/>Subir logo</Btn>{s.logo&&<Btn variant="ghost" style={{marginLeft:8}} onClick={()=>setS(p=>({...p,logo:null}))}>Quitar</Btn>}<input ref={fileRef} type="file" accept="image/*" onChange={handleLogo} style={{display:"none"}}/><p style={{fontSize:11,color:"#94A3B8",marginTop:6}}>PNG, JPG o SVG. Recomendado: fondo transparente.</p></div>
      </div>
    </Card>
    <Card style={{marginBottom:16}}>
      <h3 style={{...S.sectionTitle,marginBottom:14}}>Colores y pie de página</h3>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
        <Field label="Color de vouchers"><div style={{display:"flex",gap:8,alignItems:"center"}}><input type="color" value={s.voucherColor||"#1a56db"} onChange={e=>setS(p=>({...p,voucherColor:e.target.value}))} style={{width:40,height:36,border:"1px solid #E2E8F0",borderRadius:6,cursor:"pointer",padding:2}}/><span style={{fontSize:12,color:"#64748B"}}>Encabezado de vouchers</span></div></Field>
        <Field label="Color de recibos"><div style={{display:"flex",gap:8,alignItems:"center"}}><input type="color" value={s.receiptColor||"#1a56db"} onChange={e=>setS(p=>({...p,receiptColor:e.target.value}))} style={{width:40,height:36,border:"1px solid #E2E8F0",borderRadius:6,cursor:"pointer",padding:2}}/><span style={{fontSize:12,color:"#64748B"}}>Totales en recibos</span></div></Field>
        <Txta label="Nota en encabezado de recibos" value={s.headerNote} onChange={v=>setS(p=>({...p,headerNote:v}))} rows={2} placeholder="Políticas de cancelación..." col={2} style={{gridColumn:"span 2"}}/>
        <Inp label="Texto del pie de página" value={s.footerText} onChange={v=>setS(p=>({...p,footerText:v}))} placeholder="Gracias por elegirnos." col={2} style={{gridColumn:"span 2"}}/>
        <Txta label="Políticas por defecto (cotizaciones)" value={s.quotationPolicies} onChange={v=>setS(p=>({...p,quotationPolicies:v}))} rows={4} placeholder="• Precios sujetos a disponibilidad..." col={2} style={{gridColumn:"span 2"}}/>
      </div>
    </Card>
    <Card style={{marginBottom:16}}>
      <h3 style={{...S.sectionTitle,marginBottom:4}}>Configuración de documentos</h3>
      <p style={{fontSize:12,color:"#64748B",marginBottom:16}}>Controlá qué información aparece en cada documento impreso.</p>
      <Tabs tabs={docTabs} active={docTab} setActive={setDocTab}/>
      {docTab==="recibo"&&(<div>
        <Toggle label="Mostrar detalle de servicios" desc="Lista los servicios incluidos en el viaje" checked={getDc("recibo","showServiceDetail")} onChange={v=>setDc("recibo","showServiceDetail",v)}/>
        <Toggle label="Mostrar proveedor" desc="Nombre del proveedor por servicio" checked={getDc("recibo","showProviders")} onChange={v=>setDc("recibo","showProviders",v)}/>
        <Toggle label="Mostrar referencia del proveedor" desc="N° de file/localizador del proveedor" checked={getDc("recibo","showProviderRef")} onChange={v=>setDc("recibo","showProviderRef",v)}/>
        <Toggle label="Mostrar precio por servicio" checked={getDc("recibo","showServicePrices")} onChange={v=>setDc("recibo","showServicePrices",v)}/>
        <Toggle label="Mostrar historial de pagos" desc="Tabla con todos los pagos recibidos" checked={getDc("recibo","showPaymentHistory")} onChange={v=>setDc("recibo","showPaymentHistory",v)}/>
        <Toggle label="Mostrar totales" desc="Precio total, cobrado y saldo pendiente" checked={getDc("recibo","showTotals")} onChange={v=>setDc("recibo","showTotals",v)}/>
      </div>)}
      {docTab==="voucherAereo"&&(<div>
        <Toggle label="Mostrar DNI/Pasaporte de pasajeros" checked={getDc("voucherAereo","showPassengerDNI")} onChange={v=>setDc("voucherAereo","showPassengerDNI",v)}/>
        <Toggle label="Mostrar número de vuelo" checked={getDc("voucherAereo","showFlightNumbers")} onChange={v=>setDc("voucherAereo","showFlightNumbers",v)}/>
        <Toggle label="Mostrar información de equipaje" checked={getDc("voucherAereo","showBaggageInfo")} onChange={v=>setDc("voucherAereo","showBaggageInfo",v)}/>
        <Toggle label="Mostrar localizador del proveedor" checked={getDc("voucherAereo","showProviderRef")} onChange={v=>setDc("voucherAereo","showProviderRef",v)}/>
      </div>)}
      {docTab==="voucherTerrestre"&&(<div>
        <Toggle label="Mostrar DNI/Pasaporte de pasajeros" checked={getDc("voucherTerrestre","showPassengerDNI")} onChange={v=>setDc("voucherTerrestre","showPassengerDNI",v)}/>
        <Toggle label="Mostrar tipo de habitación" checked={getDc("voucherTerrestre","showRoomType")} onChange={v=>setDc("voucherTerrestre","showRoomType",v)}/>
        <Toggle label="Mostrar régimen de alojamiento" checked={getDc("voucherTerrestre","showRegimen")} onChange={v=>setDc("voucherTerrestre","showRegimen",v)}/>
        <Toggle label="Mostrar teléfono del hotel" checked={getDc("voucherTerrestre","showHotelPhone")} onChange={v=>setDc("voucherTerrestre","showHotelPhone",v)}/>
        <Toggle label="Mostrar dirección del hotel" checked={getDc("voucherTerrestre","showHotelAddress")} onChange={v=>setDc("voucherTerrestre","showHotelAddress",v)}/>
        <Toggle label="Mostrar localizador del proveedor" checked={getDc("voucherTerrestre","showProviderRef")} onChange={v=>setDc("voucherTerrestre","showProviderRef",v)}/>
        <Toggle label="Mostrar información importante" desc="Políticas y observaciones del servicio" checked={getDc("voucherTerrestre","showImportantInfo")} onChange={v=>setDc("voucherTerrestre","showImportantInfo",v)}/>
      </div>)}
    </Card>
    <div style={{display:"flex",alignItems:"center",gap:12}}>
      <Btn size="lg" onClick={save}>Guardar configuración</Btn>
      {saved&&<span style={{color:"#059669",fontSize:13,fontWeight:600,display:"flex",alignItems:"center",gap:4}}><Check size={14}/>¡Guardado!</span>}
    </div>
  </div>);
}


// ── PRINT: COTIZACIÓN ─────────────────────────────
function printQuotation(res, settings, depositPct, validDays, policies, serviceDesc, priceBasis){
  const color = settings.voucherColor||'#1a3a5c';
  const cur = settings.currency||'ARS';
  const logoHtml = settings.logo
    ? `<img src="${settings.logo}" style="height:60px;max-width:180px;object-fit:contain;" alt="Logo"/>`
    : `<div style="font-size:18px;font-weight:700;color:#1a1a1a;">${settings.agencyName}</div>`;

  const vuelos = res.services.filter(s=>s.type==='vuelo').sort((a,b)=>{
    const da=(a.departureDate||"")+(a.departureTime||"");
    const db=(b.departureDate||"")+(b.departureTime||"");
    return da.localeCompare(db);
  });
  const hoteles = res.services.filter(s=>s.type==='hotel').sort((a,b)=>(a.checkIn||"").localeCompare(b.checkIn||""));
  const otros = res.services.filter(s=>s.type!=='vuelo'&&s.type!=='hotel');

  const passengerBlock = res.passengers.map(p=>
    `<div style="display:flex;gap:24px;padding:5px 0;border-bottom:1px solid #f0f0f0;font-size:11px;">
      <span style="font-weight:600;min-width:180px;">${p.name}</span>
      ${p.dni?`<span style="color:#999;">DNI: ${p.dni}</span>`:''}
    </div>`
  ).join('');

  const vuelosBlock = vuelos.map(s=>`
    <div style="margin-bottom:10px;">
      <div style="display:flex;justify-content:space-between;margin-bottom:4px;">
        <span style="font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;color:#888;">${s.flightType||'Vuelo'} · ${s.departureDate?new Date(s.departureDate+'T12:00:00').toLocaleDateString('es-AR'):''}</span>
        <span style="font-size:10px;font-weight:600;color:#555;">${s.airline||''}</span>
      </div>
      <div style="display:flex;align-items:center;gap:8px;">
        <div style="min-width:80px;">
          <span style="font-size:13px;font-weight:700;color:#1a1a1a;">${s.departureTime||''}</span>
          <span style="font-size:11px;color:#555;margin-left:4px;">· ${s.origin||''} (${s.originCode||'?'})</span>
        </div>
        <div style="flex:1;text-align:center;font-size:9px;color:#bbb;font-style:italic;">${s.stops||'Directo'}${s.duration?' · '+s.duration:''}</div>
        <div style="min-width:80px;text-align:right;">
          <span style="font-size:13px;font-weight:700;color:#1a1a1a;">${s.arrivalTime||''}</span>
          <span style="font-size:11px;color:#555;margin-left:4px;">· ${s.destination||''} (${s.destinationCode||'?'})</span>
        </div>
      </div>
      ${s.baggage?`<div style="margin-top:5px;"><span style="background:#f2f2f0;color:#666;font-size:9px;padding:2px 8px;border-radius:20px;">${s.baggage}</span></div>`:''}
    </div>
    <div style="height:1px;background:#f0f0ee;margin:8px 0;"></div>
  `).join('');

  const hotelesBlock = hoteles.map(s=>`
    <div style="margin-bottom:10px;">
      <div style="font-size:13px;font-weight:700;color:#1a1a1a;margin-bottom:3px;">${s._extractedProviderName||s.description||'Hotel'}</div>
      <div style="font-size:11px;color:#888;margin-bottom:6px;">${s.regimen||''}${s.roomType?' | '+s.roomType:''}</div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:3px;font-size:11px;color:#666;">
        ${s.checkIn?`<div>Check-in: <strong>${new Date(s.checkIn+'T12:00:00').toLocaleDateString('es-AR')}</strong></div>`:''}
        ${s.checkOut?`<div>Check-out: <strong>${new Date(s.checkOut+'T12:00:00').toLocaleDateString('es-AR')}</strong></div>`:''}
        ${s.nights?`<div>Noches: <strong>${s.nights}</strong></div>`:''}
        ${s.rooms?`<div>Habitaciones: <strong>${s.rooms}</strong></div>`:''}
      </div>
    </div>
    <div style="height:1px;background:#f0f0ee;margin:8px 0;"></div>
  `).join('');

  const otrosBlock = otros.map(s=>`
    <div style="margin-bottom:8px;font-size:12px;">
      <span style="font-weight:600;color:#1a1a1a;">${s.type.charAt(0).toUpperCase()+s.type.slice(1)}</span>
      ${s.description?` — ${s.description}`:''}
      ${s.serviceDate?` · ${new Date(s.serviceDate+'T12:00:00').toLocaleDateString('es-AR')}`:''}
    </div>
  `).join('');

  const depositAmt = Math.round((res.salePrice||0) * (depositPct/100) * 100) / 100;
  const balanceAmt = (res.salePrice||0) - depositAmt;
  const validDate = validDays ? (() => { const d=new Date(); d.setDate(d.getDate()+validDays); return d.toLocaleDateString('es-AR'); })() : '';

  const html = `<!DOCTYPE html><html><head><meta charset="utf-8">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<title>Cotización #${res.fileNumber}</title>
<style>
*{margin:0;padding:0;box-sizing:border-box;}
body{font-family:'Inter',sans-serif;color:#1a1a1a;font-size:12px;background:white;}
.wrap{max-width:700px;margin:0 auto;border:1px solid #e2e2e2;border-radius:12px;overflow:hidden;}
.header{padding:1.2rem 1.8rem;display:flex;justify-content:space-between;align-items:center;border-bottom:1px solid #e8e8e8;}
.destino{font-size:24px;font-weight:700;letter-spacing:-0.5px;}
.sec{padding:1rem 1.8rem;border-bottom:1px solid #e8e8e8;}
.sec-title{display:flex;align-items:center;gap:7px;margin-bottom:12px;}
.sec-title-bar{width:3px;height:14px;background:${color};border-radius:2px;}
.sec-title-text{font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:1.2px;color:${color};}
.passengers-bg{background:#fafaf8;}
.price-box{background:${color};color:white;border-radius:8px;padding:12px 18px;display:flex;justify-content:space-between;align-items:center;margin-bottom:9px;}
.plazos{background:#f9f9f7;border-radius:8px;padding:11px 15px;}
.plazo-row{display:flex;justify-content:space-between;padding:5px 0;font-size:11px;}
.plazo-row+.plazo-row{border-top:1px solid #eee;}
.politicas{font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:1.2px;color:#444;margin-bottom:8px;}
ul{padding-left:15px;font-size:11px;color:#666;line-height:2;}
.footer{background:${color};padding:.8rem 1.8rem;text-align:center;font-size:10px;color:rgba(255,255,255,.7);}
@media print{body{-webkit-print-color-adjust:exact;print-color-adjust:exact;}.wrap{border:none;border-radius:0;}}
</style></head><body>
<div class="wrap">

  <div class="header">
    <div style="display:flex;align-items:center;">
      ${logoHtml}
    </div>
    <div style="text-align:right;">
      <div class="destino">${(res.destination||'').toUpperCase()}</div>
      <div style="font-size:11px;color:#666;">Salida ${res.departureDate?new Date(res.departureDate+'T12:00:00').toLocaleDateString('es-AR'):''} ${res.departureDate&&res.returnDate?'| '+Math.ceil((new Date(res.returnDate)-new Date(res.departureDate))/86400000)+' Noches':''}</div>
      ${validDate?`<div style="font-size:10px;color:#aaa;margin-top:2px;">Válida hasta: ${validDate}</div>`:''}
    </div>
  </div>

  ${res.passengers.length>0?`
  <div class="sec passengers-bg">
    <div style="font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:1.2px;color:#999;margin-bottom:6px;">Pasajeros</div>
    ${passengerBlock}
  </div>`:''}

  ${vuelos.length>0?`
  <div class="sec">
    <div class="sec-title"><div class="sec-title-bar"></div><div class="sec-title-text">Itinerario de vuelos</div></div>
    ${vuelosBlock}
  </div>`:''}

  ${hoteles.length>0?`
  <div class="sec">
    <div class="sec-title"><div class="sec-title-bar"></div><div class="sec-title-text">Alojamiento seleccionado</div></div>
    ${hotelesBlock}
  </div>`:''}

  ${otros.length>0?`
  <div class="sec">
    <div class="sec-title"><div class="sec-title-bar"></div><div class="sec-title-text">Servicios adicionales</div></div>
    ${otrosBlock}
  </div>`:''}

  ${serviceDesc?`
  <div class="sec">
    <div class="sec-title"><div class="sec-title-bar"></div><div class="sec-title-text">Servicios incluidos</div></div>
    <div style="font-size:11px;color:#555;line-height:1.8;">${serviceDesc}</div>
  </div>`:''}

  <div class="sec">
    <div class="price-box">
      <div style="font-size:11px;opacity:.8;">Precio ${priceBasis||'por persona'}</div>
      <div style="font-size:20px;font-weight:700;">${cur==='USD'?'US$':cur==='EUR'?'€':'$'} ${Number(res.salePrice||0).toLocaleString('es-AR',{minimumFractionDigits:2})}</div>
    </div>
    <div class="plazos">
      <div style="font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:1.2px;color:#999;margin-bottom:8px;">Plazos de pago</div>
      <div class="plazo-row">
        <span style="color:#666;">Seña al confirmar (${depositPct}%)</span>
        <span style="font-weight:700;">${cur==='USD'?'US$':cur==='EUR'?'€':'$'} ${depositAmt.toLocaleString('es-AR',{minimumFractionDigits:2})}</span>
      </div>
      <div class="plazo-row">
        <span style="color:#666;">Saldo hasta 45 días antes de la salida</span>
        <span style="font-weight:700;">${cur==='USD'?'US$':cur==='EUR'?'€':'$'} ${balanceAmt.toLocaleString('es-AR',{minimumFractionDigits:2})}</span>
      </div>
    </div>
  </div>

  ${policies?`
  <div class="sec">
    <div class="politicas">Políticas y observaciones</div>
    <ul>${policies.split('\n').filter(l=>l.trim()).map(l=>`<li>${l.replace(/^[•\-\*]\s*/,'')}</li>`).join('')}</ul>
  </div>`:''}

  <div class="footer">
    ${settings.agencyName}${settings.phone?' · '+settings.phone:''}${settings.email?' · '+settings.email:''}${settings.footerText?' · '+settings.footerText:''}
  </div>
</div>
<script>window.print();window.onafterprint=()=>window.close();<\/script>
</body></html>`;
  const w=window.open("","_blank","width=780,height=950");
  w.document.write(html);
  w.document.close();
}

// ── QUOTATIONS PAGE ───────────────────────────────
function QuotationsPage({data, update}){
  const {reservations, providers, settings} = data;
  const [search, setSearch] = useState("");
  const [modal, setModal] = useState(null);
  const [printModal, setPrintModal] = useState(null);
  const cur = settings.currency||'ARS';

  // Solo cotizaciones (status === cotizacion)
  const quotations = useMemo(()=>
    reservations
      .filter(r=>r.status==="cotizacion")
      .filter(r=>{
        const q=search.toLowerCase();
        return !q||(r.destination||"").toLowerCase().includes(q)||
          String(r.fileNumber).includes(q)||
          r.passengers.some(p=>p.name.toLowerCase().includes(q));
      })
      .sort((a,b)=>b.createdAt?.localeCompare(a.createdAt||"")||0),
    [reservations, search]
  );

  const saveQuotation = r => {
    update(d=>{
      const idx=d.reservations.findIndex(x=>x.id===r.id);
      const updated=idx>=0?d.reservations.map(x=>x.id===r.id?r:x):[...d.reservations,r];
      return{...d,reservations:updated,nextFile:Math.max(d.nextFile,Number(r.fileNumber.replace(/^0+/,""))+1)};
    });
    setModal(null);
  };

  const confirmQuotation = (r) => {
    if(!window.confirm(`¿Confirmar la cotización #${r.fileNumber} y pasarla a Reservas?`)) return;
    update(d=>({...d,reservations:d.reservations.map(x=>x.id===r.id?{...x,status:"confirmada"}:x)}));
  };

  const deleteQuotation = (id) => {
    if(!window.confirm("¿Eliminar esta cotización?")) return;
    update(d=>({...d,reservations:d.reservations.filter(x=>x.id!==id)}));
  };

  return(
    <div style={{padding:28,overflowY:"auto",flex:1}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20}}>
        <div>
          <h1 style={{margin:0,fontSize:22,fontWeight:800,color:"#0F172A"}}>Cotizaciones</h1>
          <p style={{margin:"4px 0 0",fontSize:13,color:"#64748B"}}>{quotations.length} cotizaciones pendientes</p>
        </div>
        <Btn onClick={()=>setModal({isNew:true,data:newRes(data.nextFile,settings.defaultCommission)})}>
          <Plus size={14}/>Nueva cotización
        </Btn>
      </div>

      <Card style={{marginBottom:16}}>
        <div style={{position:"relative"}}>
          <Search size={14} style={{position:"absolute",left:10,top:"50%",transform:"translateY(-50%)",color:"#94A3B8"}}/>
          <input value={search} onChange={e=>setSearch(e.target.value)}
            placeholder="Buscar por destino, file o pasajero…"
            style={{...S.input,paddingLeft:32}}/>
        </div>
      </Card>

      {quotations.length===0
        ? <EmptyState icon={ClipboardList} title="Sin cotizaciones"
            sub="Creá una nueva cotización para un cliente."
            action={<Btn onClick={()=>setModal({isNew:true,data:newRes(data.nextFile,settings.defaultCommission)})}><Plus size={14}/>Nueva cotización</Btn>}/>
        : <Card style={{padding:0}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
              <thead>
                <tr style={{background:"#F8FAFC",borderBottom:"1px solid #E2E8F0"}}>
                  {["File","Destino","Pasajeros","Salida","Precio","Acciones"].map(h=>(
                    <th key={h} style={{padding:"10px 12px",textAlign:"left",fontWeight:700,fontSize:11,color:"#64748B",textTransform:"uppercase"}}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {quotations.map(r=>(
                  <tr key={r.id} style={{borderBottom:"1px solid #F1F5F9"}}
                    onMouseEnter={e=>e.currentTarget.style.background="#F8FAFC"}
                    onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
                    <td style={{padding:"10px 12px",fontWeight:700,color:"#6366F1"}}>#{r.fileNumber}</td>
                    <td style={{padding:"10px 12px",fontWeight:500}}>{r.destination||"—"}</td>
                    <td style={{padding:"10px 12px",color:"#64748B",fontSize:12}}>
                      {r.passengers.map(p=>p.name).filter(Boolean).join(", ")||"—"}
                    </td>
                    <td style={{padding:"10px 12px",color:"#64748B"}}>{r.departureDate?new Date(r.departureDate+'T12:00:00').toLocaleDateString('es-AR'):"—"}</td>
                    <td style={{padding:"10px 12px",fontWeight:600}}>{fmt$(r.salePrice,cur)}</td>
                    <td style={{padding:"10px 12px"}}>
                      <div style={{display:"flex",gap:3,flexWrap:"wrap"}}>
                        <button title="Editar" onClick={()=>setModal({isNew:false,data:r})} style={S.ghostBtn}><Pencil size={13}/></button>
                        <button title="Imprimir cotización" onClick={()=>setPrintModal(r)}
                          style={{...S.ghostBtn,color:"#6366F1",borderColor:"#C7D2FE"}}><Printer size={13}/></button>
                        <button title="Confirmar → pasar a Reservas"
                          onClick={()=>confirmQuotation(r)}
                          style={{...S.ghostBtn,color:"#059669",borderColor:"#A7F3D0",fontWeight:600}}>
                          <Check size={13}/>Confirmar
                        </button>
                        <button title="Eliminar" onClick={()=>deleteQuotation(r.id)}
                          style={{...S.ghostBtn,color:"#EF4444",borderColor:"#FCA5A5"}}><Trash2 size={13}/></button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </Card>
      }

      {modal&&(
        <ReservationModal
          initial={modal.isNew?null:modal.data}
          providers={providers}
          settings={settings}
          onSave={saveQuotation}
          onClose={()=>setModal(null)}/>
      )}

      {printModal&&(
        <PrintQuotationModal
          res={printModal}
          settings={settings}
          onClose={()=>setPrintModal(null)}/>
      )}
    </div>
  );
}

// ── PRINT QUOTATION MODAL ─────────────────────────
function PrintQuotationModal({res, settings, onClose}){
  const [depositPct, setDepositPct] = useState(30);
  const [validDays, setValidDays] = useState(5);
  const [priceBasis, setPriceBasis] = useState('en base doble');
  const [serviceDesc, setServiceDesc] = useState(res.description||"");
  const [policies, setPolicies] = useState(
    settings.quotationPolicies||
    "• Precios sujetos a disponibilidad al momento de la reserva.\n• Se requiere pasaporte con vigencia mínima de 6 meses.\n• Asistencia al viajero recomendada (consultar opciones)."
  );
  const cur = settings.currency||'ARS';
  const depositAmt = Math.round((res.salePrice||0)*(depositPct/100)*100)/100;
  const balanceAmt = (res.salePrice||0) - depositAmt;

  return(
    <Modal title={`Imprimir Cotización — File #${res.fileNumber}`} onClose={onClose} width={560}
      footer={<>
        <Btn variant="secondary" onClick={onClose}>Cancelar</Btn>
        <Btn onClick={()=>{printQuotation(res,settings,depositPct,validDays,policies,serviceDesc,priceBasis);onClose();}}>
          <Printer size={14}/>Imprimir cotización
        </Btn>
      </>}>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
        <Field label="% de seña">
          <div style={{display:"flex",alignItems:"center",gap:8}}>
            <input type="number" value={depositPct} onChange={e=>setDepositPct(Math.min(100,Math.max(0,+e.target.value)))}
              style={{...S.input,width:80}} min={0} max={100}/>
            <span style={{fontSize:12,color:"#64748B"}}>= {cur==='USD'?'US$':cur==='EUR'?'€':'$'}{depositAmt.toLocaleString('es-AR',{minimumFractionDigits:2})}</span>
          </div>
        </Field>
        <Field label="Validez (días hábiles)">
          <input type="number" value={validDays} onChange={e=>setValidDays(+e.target.value)}
            style={{...S.input}} min={1}/>
        </Field>
        <Field label="Base de precio" col={2} style={{gridColumn:'span 2'}}>
          <select value={priceBasis} onChange={e=>setPriceBasis(e.target.value)} style={{...S.input,appearance:'auto'}}>
            <option value='por persona'>Por persona</option>
            <option value='en base doble'>En base doble</option>
            <option value='en base triple'>En base triple</option>
            <option value='en base cuádruple'>En base cuádruple</option>
            <option value='por paquete'>Por paquete completo</option>
          </select>
        </Field>
      </div>
      <div style={{background:"#F8FAFC",borderRadius:8,padding:12,margin:"14px 0",display:"flex",justifyContent:"space-between",fontSize:12}}>
        <span style={{color:"#64748B"}}>Saldo (45 días antes)</span>
        <span style={{fontWeight:700}}>{cur==='USD'?'US$':cur==='EUR'?'€':'$'}{balanceAmt.toLocaleString('es-AR',{minimumFractionDigits:2})}</span>
      </div>
      <Txta label="Servicios incluidos (descripción para el cliente)"
        value={serviceDesc} onChange={setServiceDesc} rows={3}
        placeholder="El paquete incluye vuelos, alojamiento y..."/>
      <Txta label="Políticas y observaciones (una por línea)"
        value={policies} onChange={setPolicies} rows={5}
        placeholder="• Precios sujetos a disponibilidad..."/>
    </Modal>
  );
}


// ── APP ROOT ──────────────────────────────────────
export default function App(){
  const [db,setDb]=useState(null);
  const [section,setSection]=useState("dashboard");
  useEffect(()=>{loadDB().then(setDb);},[]);
  const update=useCallback(fn=>{setDb(prev=>{const next=typeof fn==="function"?fn(prev):{...prev,...fn};saveDB(next);return next;});},[]);
  if(!db)return(<div style={{display:"flex",alignItems:"center",justifyContent:"center",height:"100vh",background:"#0F172A",color:"#94A3B8",fontFamily:"system-ui",gap:10,fontSize:14}}><span style={{width:20,height:20,border:"2px solid #3B82F6",borderTopColor:"transparent",borderRadius:"50%",display:"inline-block",animation:"spin 1s linear infinite"}}/> Cargando TravelManager… <style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style></div>);
  return(<div style={{display:"flex",height:"100vh",fontFamily:"'Segoe UI',system-ui,sans-serif",background:"#F1F5F9",overflow:"hidden"}}><Sidebar section={section} setSection={setSection} settings={db.settings}/><main style={{flex:1,overflow:"hidden",display:"flex",flexDirection:"column"}}>{section==="dashboard"&&<Dashboard data={db}/>}{section==="quotations"&&<QuotationsPage data={db} update={update}/>}{section==="reservations"&&<ReservationsPage data={db} update={update}/>}{section==="clients"&&<ClientsPage data={db} update={update}/>}{section==="providers"&&<ProvidersPage data={db} update={update}/>}{section==="commissions"&&<CommissionsPage data={db}/>}{section==="reports"&&<ReportsPage data={db}/>}{section==="settings"&&<SettingsPage data={db} update={update}/>}</main></div>);
}

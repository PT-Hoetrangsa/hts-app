// === HOETRANGSA v4 — BAGIAN 1 DARI 3 ===
import { useState, useEffect, useCallback, useRef, useMemo, createContext, useContext } from "react";
import { BarChart, Bar, LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, PieChart, Pie, Cell, AreaChart, Area, ReferenceLine } from "recharts";
import * as XLSX from "xlsx";

// ─── CONSTANTS ────────────────────────────────────────────────────────────────
var SIZES=["5.5 kg","12 kg","50 kg"];
var JENIS=["Isi","Tabung+Isi"];
var DENOMS=[100000,50000,20000,10000,5000,2000,1000];
// Urutan role untuk dropdown
var ROLE_ORDER=["admin","akuntan","sales_driver","sales_freelance","checker","driver_truck","owner"];
function sortEmp(emps){return(emps||[]).slice().sort((a,b)=>{var ia=ROLE_ORDER.indexOf(a.role);var ib=ROLE_ORDER.indexOf(b.role);if(ia!==ib)return ia-ib;return(a.nama||"").localeCompare(b.nama||"");});}
// Load Plus Jakarta Sans font
(function(){
if(typeof document!=="undefined"){
if(!document.getElementById("pjs-font")){var l=document.createElement("link");l.id="pjs-font";l.rel="stylesheet";l.href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap";document.head.appendChild(l);}
if(!document.getElementById("h2c-script")){var s=document.createElement("script");s.id="h2c-script";s.src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js";document.head.appendChild(s);}
}})();
var UANG_MAKAN_DEFAULT=20000;
var SPBE_LOC={"SPPBE KCR":25000,"SPPBE MGL":115000};
var KAT_K=["BBM","Gaji","Uang Makan Karyawan","Perbaikan","Administrasi","Sewa","Parkir","Service Kendaraan","Fee LPG 50kg","Uang Jalan SPBE","Uang Bongkar DO","Pancung 12kg","Pancung 5.5kg","Belanja Modal","Belanja Tabung Kosong dr Konsumen","Kasbon / Ambilan Karyawan","Listrik PLN","Internet Wi-Fi","Air Galon","Lainnya"];
var KAT_AUTO_HARGA={"Uang Bongkar DO":50000,"Pancung 12kg":200000,"Pancung 5.5kg":95000};
var SPPBE_OPTS=["SPPBE KCR","SPPBE MGL","Lainnya"];
var PLG_KAT_27=["Rumah Tangga","Restaurant","Cafe","Warung Kopi","Warung Nasi/Rumah Makan","Kedai Bakso/Mie Ayam","Kantin Sekolah/Kampus","Catering","SPPG","UMKM Kuliner","King Fried Chicken","Toko Roti","Pabrik Roti","Produksi Industri Makanan","Peternakan Ayam","Hotel","Rumah Sakit","Kantor Pemerintah","Kantor Swasta","Kampus/Universitas","Pesantren","Pangkalan","Agen LPG","Reseller LPG","Sub Agen LPG","Canvaser LPG","SPBU","Swalayan/Grosir","Mini Market/Kios","WO/Event Organizer","Usaha Umum","Laundry","Industri Rumahan","Lainnya"]
var PLG_TITIP_KAT=PLG_KAT_27.filter(k=>!["Rumah Tangga","Industri Rumahan"].includes(k));
var KAT_OPS=["BBM","Uang Bongkar DO","Uang Jalan SPBE","Fee LPG 50kg","Pancung 12kg","Pancung 5.5kg","Parkir","Service Kendaraan","Lainnya","Listrik PLN","Internet Wi-Fi","Air Galon"];
var ABSENSI_STATUS=["Hadir","Sakit","Izin","Alpha","Cuti","Libur"];
var ROLE_LBL={admin:"Manajer",akuntan:"Kasir/Akuntan",sales_driver:"Sales Driver",sales_freelance:"Sales Freelance",checker:"Checker",driver_truck:"Driver Truck",owner:"Komisaris"};
var ROLE_TABS={owner:null,admin:null,akuntan:null,sales_driver:["dashboard","penjualan","piutang","setoran","pelanggan"],sales_freelance:["dashboard","penjualan","piutang","setoran","pelanggan"],checker:["dashboard","stok","do"],driver_truck:["dashboard","do"]};
var PENJUALAN_ROLES=["admin","akuntan","sales_driver","sales_freelance","owner"];
var DEF_MODAL={"Isi":{"5.5 kg":98881,"12 kg":212133,"50 kg":1026464},"Tabung+Isi":{"5.5 kg":381645,"12 kg":730990,"50 kg":2415419}};
var DEF_HET={"Isi":{"5.5 kg":110000,"12 kg":230000,"50 kg":1200000},"Tabung+Isi":{"5.5 kg":395000,"12 kg":730000,"50 kg":2450000}};
var PIE_COLORS=["#0C4DA2","#00A651","#F39C12","#ED1C24","#8E44AD","#1ABC9C","#E67E22","#3498DB"];
var ROMAN=["I","II","III","IV","V","VI","VII","VIII","IX","X","XI","XII"];
var BULAN_ID=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"];

var DEF_EMP=[
{id:"e01",username:"Haekal",password:"Sudirman80",role:"admin",nama:"Muhammad Haekal",posisi:"Manajer",telepon:"082246980311",alamat:"Banda Aceh",gajiPokok:5000000,uangMakan:15000,uangMakanMode:"akhir_bulan",aktif:true},
{id:"e02",username:"DAYAT",password:"Sudirman80",role:"akuntan",nama:"MANARUL HIDAYAT",posisi:"Kasir/Akuntan",telepon:"082277214045",alamat:"",gajiPokok:2000000,uangMakan:15000,uangMakanMode:"akhir_bulan",aktif:true},
{id:"e03",username:"MUSLEM",password:"Sudirman80",role:"sales_driver",nama:"MUSLEM",posisi:"Sales Driver",telepon:"082259494802",alamat:"",gajiPokok:2250000,uangMakan:15000,uangMakanMode:"harian",aktif:true},
{id:"e04",username:"USNI",password:"Sudirman80",role:"sales_driver",nama:"USNI SIREGAR",posisi:"Sales Driver",telepon:"082187102114",alamat:"",gajiPokok:2250000,uangMakan:15000,uangMakanMode:"harian",aktif:true},
{id:"e05",username:"SAYUTI",password:"Sudirman80",role:"sales_driver",nama:"T. SAYUTI",posisi:"Sales Driver",telepon:"081377494144",alamat:"",gajiPokok:2250000,uangMakan:15000,uangMakanMode:"harian",aktif:true},
{id:"e06",username:"HELMI",password:"Sudirman80",role:"sales_driver",nama:"HELMI ZAKARIA",posisi:"Sales Driver",telepon:"081377190552",alamat:"Lam Lumpu - Peukan Bada, Aceh Besar",gajiPokok:2250000,uangMakan:15000,uangMakanMode:"harian",aktif:true},
{id:"e07",username:"HAMDANI",password:"Sudirman80",role:"driver_truck",nama:"HAMDANI IBRAHIM",posisi:"Driver Truck SPBE",telepon:"085276590894",alamat:"",gajiPokok:2000000,uangMakan:15000,uangMakanMode:"harian",aktif:true},
{id:"e08",username:"BUNYAMIN",password:"Sudirman80",role:"sales_freelance",nama:"BUNYAMIN",posisi:"Sales Marketing",telepon:"085212570213",alamat:"",gajiPokok:2000000,uangMakan:15000,uangMakanMode:"harian",aktif:true},
{id:"e09",username:"SAIBAN",password:"Sudirman80",role:"checker",nama:"IBNU SAIBAN",posisi:"Checker",telepon:"082298704089",alamat:"",gajiPokok:2000000,uangMakan:15000,uangMakanMode:"akhir_bulan",aktif:true},
{id:"e10",username:"SYAHRULI",password:"Sudirman80",role:"owner",nama:"SYAHRULI",posisi:"Komisaris",telepon:"081370492714",alamat:"",gajiPokok:2500000,uangMakan:15000,uangMakanMode:"akhir_bulan",aktif:true},
{id:"e11",username:"FRIDZLUN",password:"Sudirman80",role:"sales_freelance",nama:"FRIDZLUN",posisi:"Sales Freelance",telepon:"085361224830",alamat:"",gajiPokok:0,uangMakan:15000,uangMakanMode:"harian",aktif:true},
{id:"e12",username:"Bachtiar",password:"Sudirman80",role:"sales_freelance",nama:"Bachtiar",posisi:"Sales Freelance",telepon:"081264600500",alamat:"",gajiPokok:0,uangMakan:15000,uangMakanMode:"harian",aktif:true},
];

// ─── GOOGLE SHEETS SYNC ───────────────────────────────────────────────────────
var GAS_URL="https://script.google.com/macros/s/AKfycbxSLJPfpPA_CTx9CMocRmUWKeNSH1md-_J1R4B09qlQDrsPhXql1mrWPbBVnx8u2sJ2/exec";
// Logo tetap perusahaan & Pertamina — disimpan di public/assets/, tidak perlu upload manual per device
var LOGO_HTS_DARK="./assets/logo-hts-dark-bg.png";// untuk latar gelap (header app)
var LOGO_HTS_LIGHT="./assets/logo-hts-light-bg.png";// untuk latar terang (invoice/cetak)
var LOGO_PERTAMINA_DARK="./assets/logo-pertamina-dark-bg.png";// untuk latar gelap (header app)
var LOGO_PERTAMINA_LIGHT="./assets/logo-pertamina-light-bg.png";// untuk latar terang (invoice/cetak)
var GAS_SECRET="HTS2026";
var SYNC_TABLES=["penjualan","bon","pengeluaran","pelanggan","employees","stok","doList","doTrip","absensi","payrollLog","ambilan","titipList","setoranLog","tutupBuku","config","jualanLain","kasBankTF","invoiceManual","setoranBank"];

// CORS-safe: kirim via form POST ke GAS
function gasPost(payload){
  return new Promise(function(resolve,reject){
    var id="cb_"+Date.now();
    // Gunakan no-cors fetch — tidak bisa baca response, tapi data terkirim
    // Untuk read: pakai JSONP via script tag
    fetch(GAS_URL,{
      method:"POST",
      mode:"no-cors",
      headers:{"Content-Type":"text/plain"},
      body:JSON.stringify(payload)
    }).then(()=>resolve({ok:true})).catch(e=>reject(e));
  });
}

// JSONP untuk READ (bisa baca response)
function gasReadJsonp(table){
  return new Promise(function(resolve){
    var cbName="hts_cb_"+Date.now()+"_"+Math.random().toString(36).slice(2);
    window[cbName]=function(data){
      delete window[cbName];
      if(s)document.head.removeChild(s);
      resolve(data&&data.ok?data.data:[]);
    };
    var s=document.createElement("script");
    var params=new URLSearchParams({secret:GAS_SECRET,action:"read",table,callback:cbName});
    s.src=GAS_URL+"?"+params.toString();
    s.onerror=function(){delete window[cbName];resolve([]);};
    document.head.appendChild(s);
    setTimeout(function(){if(window[cbName]){delete window[cbName];resolve([]);}},15000);
  });
}

async function gasRead(table){
  try{return await gasReadJsonp(table);}catch(e){return[];}
}

async function gasWrite(table,data){
  try{
    await gasPost({secret:GAS_SECRET,action:"write",table,data});
    return true;
  }catch(e){return false;}
}

// Push semua data ke Google Sheets
async function pushAll(data,setSyncStatus){
  setSyncStatus("pushing");
  try{
    var tasks=[];
    tasks.push(gasWrite("penjualan",data.penjualan||[]));
    tasks.push(gasWrite("bon",data.bon||[]));
    tasks.push(gasWrite("pengeluaran",data.pengeluaran||[]));
    tasks.push(gasWrite("pelanggan",data.pelanggan||[]));
    tasks.push(gasWrite("employees",data.employees||[]));
    tasks.push(gasWrite("doList",data.doList||[]));
    tasks.push(gasWrite("doTrip",data.doTrip||[]));
    tasks.push(gasWrite("absensi",data.absensi||[]));
    tasks.push(gasWrite("payrollLog",data.payrollLog||[]));
    tasks.push(gasWrite("ambilan",data.ambilan||[]));
    tasks.push(gasWrite("titipList",data.titipList||[]));
    tasks.push(gasWrite("setoranLog",data.setoranLog||[]));
    tasks.push(gasWrite("tutupBuku",data.tutupBuku||[]));
    tasks.push(gasWrite("jualanLain",data.jualanLain||[]));
    tasks.push(gasWrite("kasBankTF",data.kasBankTF||[]));
    tasks.push(gasWrite("invoiceManual",data.invoiceManual||[]));
    tasks.push(gasWrite("setoranBank",data.setoranBank||[]));
    // stok & config sebagai object → wrap dalam array
    tasks.push(gasWrite("stok",[{key:"stok",val:JSON.stringify({stock:data.stock,stokKosong:data.stokKosong,totalTabung:data.totalTabung,stokHarian:data.stokHarian,modalHistory:data.modalHistory,hetPrices:data.hetPrices,counters:data.counters,theme:data.theme,stokBatchInit:data.stokBatchInit,kas:data.kas||{},saldoAwalBank:data.saldoAwalBank||{}})}]));
    tasks.push(gasWrite("stokBatch",Object.entries(data.stokBatch||{}).map(([k,v])=>({key:k,val:JSON.stringify(v)}))));
    tasks.push(gasWrite("stockLog",data.stockLog||[]));
    tasks.push(gasWrite("config",[{key:"company",val:JSON.stringify(data.company||{})}]));
    await Promise.all(tasks);
    setSyncStatus("ok");
    return true;
  }catch(e){setSyncStatus("error");return false;}
}

// Pull semua data dari Google Sheets
async function pullAll(setSyncStatus){
  setSyncStatus("pulling");
  try{
    var [penj,bon,pen,plg,emp,doL,abs,pay,amb,titip,setor,tb,stokRaw,confRaw,jualLain,kbTF,doT,invMan,setorBank,stokBatchRaw,stockLogRaw]=await Promise.all([
      gasRead("penjualan"),gasRead("bon"),gasRead("pengeluaran"),gasRead("pelanggan"),
      gasRead("employees"),gasRead("doList"),gasRead("absensi"),gasRead("payrollLog"),
      gasRead("ambilan"),gasRead("titipList"),gasRead("setoranLog"),gasRead("tutupBuku"),
      gasRead("stok"),gasRead("config"),gasRead("jualanLain"),gasRead("kasBankTF"),gasRead("doTrip"),gasRead("invoiceManual"),gasRead("setoranBank"),gasRead("stokBatch"),gasRead("stockLog")
    ]);
    var stokMeta={};
    if(stokRaw&&stokRaw.length>0){try{stokMeta=JSON.parse(stokRaw[0].val);}catch(e){}}
    // Restore stokBatch dari sheet tersendiri
    var stokBatchParsed={};
    if(stokBatchRaw&&stokBatchRaw.length>0){stokBatchRaw.forEach(function(r){try{stokBatchParsed[r.key]=JSON.parse(r.val);}catch(e){}});}
    // Restore stockLog dari sheet tersendiri
    var stockLogParsed=stockLogRaw&&stockLogRaw.length>0?stockLogRaw:[];
    var company={};
    if(confRaw&&confRaw.length>0){try{company=JSON.parse(confRaw[0].val);}catch(e){}}
    setSyncStatus("ok");
    return{penjualan:penj,bon,pengeluaran:pen,pelanggan:plg,
      employees:emp.length>0?emp:null,
      doList:doL,doTrip:doT,absensi:abs,payrollLog:pay,ambilan:amb,
      titipList:titip,setoranLog:setor,tutupBuku:tb,
      jualanLain:jualLain,kasBankTF:kbTF,invoiceManual:invMan,setoranBank:setorBank,
      company,
      stokBatch:Object.keys(stokBatchParsed).length>0?stokBatchParsed:(stokMeta.stokBatch||{}),
      stockLog:stockLogParsed.length>0?stockLogParsed:(stokMeta.stockLog||[]),
      ...stokMeta};
  }catch(e){setSyncStatus("error");return null;}
}

var INIT={
stock:{"5.5 kg":0,"12 kg":0,"50 kg":0},stokKosong:{"5.5 kg":0,"12 kg":0,"50 kg":0},totalTabung:{"5.5 kg":0,"12 kg":0,"50 kg":0},
stockLog:[],doList:[],doTrip:[],modalHistory:[],hetPrices:{},titipList:[],
penjualan:[],bon:[],pengeluaran:[],employees:DEF_EMP.slice(),
tutupBuku:[],pelanggan:[],setoranSales:[],setoranLog:[],setoranBank:[],kas:{},saldoAwalBank:{BSI:{nominal:0,tanggal:""},BCA:{nominal:0,tanggal:""}},absensi:[],ambilan:[],payrollLog:[],jualanLain:[],kasBankTF:[],stokBatch:{},stokBatchInit:false,invoiceManual:[],
counters:{inv:{},sg:{},reg:0},theme:"light",
company:{nama:"PT. HOE TRANG SA",alamat:"Jl. Jendral Sudirman No.80, Geuce Iniem, Kec. Banda Raya, Kota Banda Aceh",telepon:"0812 6900 2121",telepon2:"(0651) 21221",email:"npso.pthoetrangsa@gmail.com",website:"pt-hoetrangsa.business.site",npwp:"",slogan:"DEALER LPG PERTAMINA",bankNama:"BSI",bankAtasNama:"PT. HOE TRANG SA",bankRekening:"812 69 2121 8",logo:"",logoPertamina:"",ttdKasir:"",ttdDirektur:"",stempelLunas:"",direkturNama:"Muhammad Haekal",kasirNama:"MANARUL HIDAYAT",soldTo:"731070",shipToKCR:"862070",shipToMGL:"782092",saBulan:"Juni 2026",sa12KCR:"2845075",sa55KCR:"2845111",sa12MGL:"",sa55MGL:"",assetArmada:0,hargaTbgKosong:{"5.5 kg":0,"12 kg":0,"50 kg":0},hppFixedSPPBE:{"SPPBE KCR":{"5.5 kg":0,"12 kg":0,"50 kg":0},"SPPBE MGL":{"5.5 kg":0,"12 kg":0,"50 kg":0}},appWidth:1600}
};

// ─── HELPERS ──────────────────────────────────────────────────────────────────
function uid(){return Date.now().toString(36)+Math.random().toString(36).slice(2,5);}
function toDay(){return new Date().toISOString().split("T")[0];}
function toMonth(){return toDay().slice(0,7);}
function fD(d){return!d?"-":new Date(d+"T00:00:00").toLocaleDateString("id-ID",{day:"2-digit",month:"long",year:"numeric"});}
function fDs(d){return!d?"-":new Date(d+"T00:00:00").toLocaleDateString("id-ID",{day:"2-digit",month:"short",year:"numeric"});}
function fR(n){return"Rp "+Number(n||0).toLocaleString("id-ID");}
function dLeft(ds){if(!ds)return null;var t=new Date();t.setHours(0,0,0,0);return Math.ceil((new Date(ds+"T00:00:00")-t)/86400000);}
function iTotal(its){return its.reduce((a,it)=>a+Number(it.qty||0)*Number(it.price||0),0);}
function daysInMonth(ym){var[y,m]=ym.split("-").map(Number);return new Date(y,m,0).getDate();}
// Kategori pengeluaran yang BUKAN biaya operasional (beli aset / piutang karyawan) — dipakai di Tutup Buku & Buku Kas Harian
var NON_OPS_CATS=["belanja modal","pancung","belanja tabung","kasbon","ambilan"];
function isNonOps(kat){kat=(kat||"").toLowerCase();return NON_OPS_CATS.some(k=>kat.includes(k));}

// ── buildStokHarian: rekonstruksi tabel stok per hari dalam sebulan ──
// ─── Buku Kas Harian (Rekap Bulanan) — semua otomatis dari data sistem ────────
function buildBukuKasHarian(data,bulan){
var dim=daysInMonth(bulan);
var rows=[];
var HARI_ID=["Minggu","Senin","Selasa","Rabu","Kamis","Jumat","Sabtu"];

for(var d=1;d<=dim;d++){
var tgl=bulan+"-"+String(d).padStart(2,"0");
var dayName=HARI_ID[new Date(tgl+"T00:00:00").getDay()];

// ── DO diterima hari ini (gabungan doTrip baru + doList lama) ──
var doQty={"5.5 kg":0,"12 kg":0,"50 kg":0};var doNilai=0;
(data.doTrip||[]).forEach(trip=>{
  if(trip.tanggal!==tgl)return;
  trip.items.forEach(it=>{if(it.status==="diterima"){doQty[it.ukuran]=(doQty[it.ukuran]||0)+it.qty;doNilai+=it.totalHPP||(it.qty*it.hppUnit)||0;}});
});
(data.doList||[]).forEach(d_=>{
  if(d_.tanggal!==tgl||(d_.status||"diterima")!=="diterima")return;
  doQty[d_.ukuran]=(doQty[d_.ukuran]||0)+(d_.qty||0);doNilai+=d_.totalHPP||0;
});

// ── Penjualan hari ini ──
var penjHari=(data.penjualan||[]).filter(p=>p.tanggal===tgl);
var penjQty={"5.5 kg":0,"12 kg":0,"50 kg":0};
penjHari.forEach(p=>(p.items||[]).forEach(it=>{if(penjQty[it.ukuran]!=null)penjQty[it.ukuran]+=Number(it.qty||0);}));
var totalPenjualan=penjHari.reduce((a,p)=>a+(p.total||0),0);
var marginKotor=penjHari.reduce((a,p)=>a+(p.margin||0),0);

// ── Penerimaan Kas & TF (penjualan cash/transfer murni, split payment dipecah) ──
var kasTF=penjHari.reduce((a,p)=>{
  if(p.splitDetail)return a+Number(p.splitDetail.cash||0)+Number(p.splitDetail.tf||0);
  if(p.bayar==="cash"||p.bayar==="transfer")return a+(p.total||0);
  return a;
},0);

// ── BON baru hari ini (piutang bertambah) ──
var bonBaru=(data.bon||[]).filter(b=>b.tanggal===tgl).reduce((a,b)=>a+(b.total||0),0);

// ── Bayar BON hari ini (piutang dilunasi/dicicil) ──
var bayarBon=0;
(data.bon||[]).forEach(b=>(b.pembayaran||[]).forEach(p=>{if((p.tanggal||"").slice(0,10)===tgl)bayarBon+=Number(p.nominal||0);}));

// ── Pengeluaran Operasional hari ini (exclude non-ops, konsisten dgn Tutup Buku) ──
var pengeluaranOps=(data.pengeluaran||[]).filter(p=>p.tanggal===tgl&&!isNonOps(p.kategori)).reduce((a,p)=>a+Number(p.nominal||0),0);

// ── Laba Bersih & Estimasi PPN ──
var labaBersih=marginKotor-pengeluaranOps;
var estimasiPPN=marginKotor*0.11;

var adaData=doNilai>0||totalPenjualan>0||pengeluaranOps>0||bonBaru>0||bayarBon>0;

rows.push({tgl,dayName,doQty,doNilai,penjQty,totalPenjualan,marginKotor,kasTF,bonBaru,bayarBon,pengeluaranOps,labaBersih,estimasiPPN,adaData});
}
return rows;
}
function buildStokHarian(data,bulan){
  var dim=daysInMonth(bulan);
  var rows=[];
  // Stok inject manual (titik awal)
  var injectMap={};
  (data.stokHarian||[]).forEach(function(r){injectMap[r.tanggal]=r;});

  // State berjalan per ukuran
  var curIsi={};var curTK={};
  SIZES.forEach(function(s){curIsi[s]=null;curTK[s]=null;});

  for(var d=1;d<=dim;d++){
    var tgl=bulan+"-"+String(d).padStart(2,"0");
    var dayName=["Minggu","Senin","Selasa","Rabu","Kamis","Jumat","Sabtu"][new Date(tgl+"T00:00:00").getDay()];

    // Cek inject manual untuk hari ini
    var inject=injectMap[tgl];
    if(inject){
      SIZES.forEach(function(s){curIsi[s]=Number(inject["isi_"+s]||0);curTK[s]=Number(inject["tk_"+s]||0);});
    }

    // Stok awal = nilai berjalan (null kalau belum ada inject)
    var awalIsi={};var awalTK={};
    SIZES.forEach(function(s){awalIsi[s]=curIsi[s]!==null?curIsi[s]:0;awalTK[s]=curTK[s]!==null?curTK[s]:0;});

    // Tabung Masuk: DO diterima + Return/Pancung hari ini
    var doHari=(data.doList||[]).filter(function(d_){return d_.tanggal===tgl&&(d_.status||"diterima")==="diterima";});
    // Gabungkan doTrip items yang diterima (multi-product DO)
    var doTripHari=[];
    (data.doTrip||[]).forEach(function(trip){
      if(trip.tanggal!==tgl)return;
      (trip.items||[]).forEach(function(it){
        if((it.status||"diterima")==="diterima"){
          doTripHari.push({ukuran:it.ukuran,qty:it.qty,tanggal:tgl,status:"diterima"});
        }
      });
    });
    // Mutasi manual masuk (Return, Pancung, Beli Tabung)
    var mutasiMasuk=(data.stockLog||[]).filter(function(l){return l.tanggal===tgl&&l.sumber==="Manual"&&(l.jenis||"").includes("Masuk")||l.tanggal===tgl&&(l.jenis||"").includes("Return")||l.tanggal===tgl&&(l.jenis||"").includes("Pancung");});

    var masukIsi={};var masukTK={};
    SIZES.forEach(function(s){
      var doQty=doHari.filter(function(d_){return d_.ukuran===s;}).reduce(function(a,d_){return a+Number(d_.qty||0);},0);
      var doTripQty=doTripHari.filter(function(d_){return d_.ukuran===s;}).reduce(function(a,d_){return a+Number(d_.qty||0);},0);
      doQty+=doTripQty;
      // Return/Pancung: +isi dari kosong (kosong berkurang seperti DO)
      var retPancQty=mutasiMasuk.filter(function(l){return l.ukuran===s&&((l.jenis||"").includes("Return")||(l.jenis||"").includes("Pancung"));}).reduce(function(a,l){return a+Number(l.qty||0);},0);
      // Mutasi masuk lain (tidak termasuk Return/Pancung)
      var mutQtyLain=mutasiMasuk.filter(function(l){return l.ukuran===s&&!(l.jenis||"").includes("Return")&&!(l.jenis||"").includes("Pancung");}).reduce(function(a,l){return a+Number(l.qty||0);},0);
      masukIsi[s]=doQty+retPancQty+mutQtyLain;
      masukTK[s]=doQty+retPancQty;// DO dan Return/Pancung sama-sama kurangi kosong
    });

    // Tabung Keluar: Penjualan + Mutasi Manual rusak/hilang
    var penjHari=(data.penjualan||[]).filter(function(p){return p.tanggal===tgl;});
    var mutasiKeluar=(data.stockLog||[]).filter(function(l){return l.tanggal===tgl&&l.sumber==="Manual"&&((l.jenis||"").includes("Rusak")||(l.jenis||"").includes("Isi Hilang")||(l.jenis||"").includes("Tbg+Isi Hilang"));});
    // Tabung Kosong Hilang: -kosong -total
    var mutasiTbgKosongHilang=(data.stockLog||[]).filter(function(l){return l.tanggal===tgl&&l.sumber==="Manual"&&(l.jenis||"").includes("Tbg Kosong Hilang");});
    // Tbg+Isi Hilang: -isi -kosong -total
    var mutasiTbgIsiHilang=(data.stockLog||[]).filter(function(l){return l.tanggal===tgl&&l.sumber==="Manual"&&(l.jenis||"").includes("Tbg+Isi Hilang");});
    // Beli Tabung dari Konsumen: +kosong, +total
    var mutasiBeli=(data.stockLog||[]).filter(function(l){return l.tanggal===tgl&&l.sumber==="Manual"&&(l.jenis||"").includes("Beli Tabung");});

    var keluarIsi={};var keluarTK={};
    SIZES.forEach(function(s){
      var refillQty=0;var tbgIsiQty=0;
      penjHari.forEach(function(p){
        (p.items||[]).forEach(function(it){
          if(it.ukuran!==s)return;
          var q=Number(it.qty||0);
          if(it.jenis==="Tabung+Isi"){tbgIsiQty+=q;}
          else{refillQty+=q;}
        });
      });
      var rusakQty=mutasiKeluar.filter(function(l){return l.ukuran===s;}).reduce(function(a,l){return a+Number(l.qty||0);},0);
      var beliQty=mutasiBeli.filter(function(l){return l.ukuran===s;}).reduce(function(a,l){return a+Number(l.qty||0);},0);
      // keluarIsi: Refill + TbgIsi + Rusak (-isi)
      var tbgKosongHilangQty=mutasiTbgKosongHilang.filter(function(l){return l.ukuran===s;}).reduce(function(a,l){return a+Number(l.qty||0);},0);
      var tbgIsiHilangQty=mutasiTbgIsiHilang.filter(function(l){return l.ukuran===s;}).reduce(function(a,l){return a+Number(l.qty||0);},0);
      // keluarIsi tambah tbgIsiHilang
      keluarIsi[s]=refillQty+tbgIsiQty+rusakQty+tbgIsiHilangQty;
      // keluarTK: tbgIsi(-kosong), refill(+kosong), rusak(+kosong), beli(+kosong), tbgKosongHilang(-kosong), tbgIsiHilang(-kosong)
      keluarTK[s]=tbgIsiQty-refillQty-rusakQty-beliQty+tbgKosongHilangQty+tbgIsiHilangQty;
    });

    // Stok Akhir
    var akhirIsi={};var akhirTK={};
    SIZES.forEach(function(s){
      akhirIsi[s]=Math.max(0,awalIsi[s]+masukIsi[s]-keluarIsi[s]);
      akhirTK[s]=Math.max(0,awalTK[s]-masukTK[s]+(-keluarTK[s]));
    });

    // Titip per hari: transaksi titip/tarik yang tanggalnya <= tgl ini
    var titipSnap={};
    SIZES.forEach(function(s){
      var bal=0;
      (data.titipList||[]).filter(function(t){return(t.tanggal||"")<=tgl;}).forEach(function(t){
        var items=t.items&&t.items.length>0?t.items:(t.ukuran===s?[{ukuran:s,qty:t.qty}]:[]);
        var m=t.tipe==="titip"?1:t.tipe==="tarik"?-1:0;
        items.filter(function(i){return i.ukuran===s;}).forEach(function(i){bal+=m*Number(i.qty||0);});
      });
      titipSnap[s]=Math.max(0,bal);
    });

    // Total per ukuran
    var total={};
    SIZES.forEach(function(s){total[s]=akhirIsi[s]+akhirTK[s]+titipSnap[s];});

    // Cek ada transaksi hari ini
    var adaTransaksi=doHari.length>0||penjHari.length>0;

    rows.push({tgl,dayName,d,adaTransaksi,inject:!!inject,
      awalIsi,awalTK,masukIsi,masukTK,keluarIsi,keluarTK,
      akhirIsi,akhirTK,titipSnap,total});

    // Update state berjalan
    SIZES.forEach(function(s){curIsi[s]=akhirIsi[s];curTK[s]=akhirTK[s];});
  }
  return rows;
}
function safeFileName(s){return(s||"file").replace(/[\/\\?%*:|"<>]/g,"_").replace(/\s+/g,"_").slice(0,80);}
function getModal(data,ukuran,jenis,tgl){tgl=tgl||toDay();var h=(data.modalHistory||[]).filter(x=>x.ukuran===ukuran&&x.jenis===jenis&&x.tanggal<=tgl);if(!h.length)return(DEF_MODAL[jenis]||{})[ukuran]||0;return h.slice().sort((a,b)=>b.tanggal.localeCompare(a.tanggal))[0].harga;}
function getHET(data,ukuran,jenis){var hp=data.hetPrices;if(hp&&hp[jenis]&&hp[jenis][ukuran]!=null)return hp[jenis][ukuran];return(DEF_HET[jenis]||{})[ukuran]||0;}
function calcMargin(items,data,tgl){return items.reduce((a,it)=>{var m=getModal(data,it.ukuran,it.jenis,tgl);return a+(Number(it.price||0)-m)*Number(it.qty||0);},0);}
function getKonsumenTitipBal(titipList){var bal={};(titipList||[]).forEach(t=>{var k=t.konsumenNama;if(!k)return;if(!bal[k])bal[k]={"5.5 kg":0,"12 kg":0,"50 kg":0,telp:t.konsumenTelp||"",alamat:t.konsumenAlamat||""};var items=t.items&&t.items.length>0?t.items:(t.ukuran&&t.qty?[{ukuran:t.ukuran,qty:t.qty}]:[]);var m=t.tipe==="titip"?1:t.tipe==="tarik"?-1:0;items.forEach(it=>{if(it.ukuran&&bal[k][it.ukuran]!==undefined)bal[k][it.ukuran]=m>0?(bal[k][it.ukuran]||0)+Number(it.qty||0):Math.max(0,(bal[k][it.ukuran]||0)-Number(it.qty||0));});});return bal;}
function getTitipTotal(titipList,ukuran){var bal=getKonsumenTitipBal(titipList);return Object.values(bal).reduce((a,v)=>a+Math.max(0,v[ukuran]||0),0);}
function getKosong(data,ukuran){return Math.max(0,((data.stokKosong||{})[ukuran]||0));}

// ═══════════════════════════════════════════════════════════════════════════
// ── SISTEM FIFO STOK ISI ──────────────────────────────────────────────────
// Setiap ukuran punya antrian batch {id,tanggal,qty,qtySisa,harga,sumber}.
// Stok masuk (DO Diterima / Retur+Pancung) ditambah ke EKOR antrian.
// Stok keluar (Penjualan / Rusak / Hilang) dimakan dari KEPALA antrian (FIFO).
// ═══════════════════════════════════════════════════════════════════════════

// Inisialisasi batch awal dari stok yang sudah ada (migrasi sekali jalan)
function ensureStokBatchInit(data){
if(data.stokBatchInit)return data;
var sb={};
SIZES.forEach(s=>{
  var qty=Math.max(0,(data.stock||{})[s]||0);
  sb[s]=qty>0?[{id:uid(),tanggal:toDay(),qty,qtySisa:qty,harga:getModal(data,s,"Isi",toDay()),sumber:"Saldo Awal (migrasi FIFO)"}]:[];
});
return{...data,stokBatch:sb,stokBatchInit:true};
}

// Tambah batch baru ke ekor antrian (DO Diterima / Retur / Pancung)
function addBatch(data,ukuran,qty,harga,sumber,tanggal){
qty=Number(qty)||0;if(qty<=0)return data;
var sb={...(data.stokBatch||{})};
var arr=(sb[ukuran]||[]).slice();
arr.push({id:uid(),tanggal:tanggal||toDay(),qty,qtySisa:qty,harga:Number(harga)||0,sumber:sumber||""});
sb[ukuran]=arr;
return{...data,stokBatch:sb};
}

// Makan qty dari kepala antrian FIFO. Return {data baru, hppTotal, hppPerUnit, detail[], qtyKurang}
function consumeFIFO(data,ukuran,qty){
qty=Number(qty)||0;
var sb={...(data.stokBatch||{})};
var arr=(sb[ukuran]||[]).map(b=>({...b}));// clone
var sisaButuh=qty;var hppTotal=0;var detail=[];
for(var i=0;i<arr.length&&sisaButuh>0;i++){
  var b=arr[i];
  if(b.qtySisa<=0)continue;
  var ambil=Math.min(b.qtySisa,sisaButuh);
  hppTotal+=ambil*b.harga;
  detail.push({batchId:b.id,tanggal:b.tanggal,qty:ambil,harga:b.harga,sumber:b.sumber});
  b.qtySisa-=ambil;
  sisaButuh-=ambil;
}
// Kalau stok batch tidak cukup (sisaButuh>0), sisanya dianggap pakai harga modal saat ini (fallback)
var qtyKurang=sisaButuh;
if(sisaButuh>0){
  var hargaFallback=getModal(data,ukuran,"Isi",toDay());
  hppTotal+=sisaButuh*hargaFallback;
  detail.push({batchId:null,tanggal:toDay(),qty:sisaButuh,harga:hargaFallback,sumber:"⚠️ Stok batch kosong — fallback harga modal"});
}
// Buang batch yang sudah qtySisa<=0 hanya kalau sangat lama (biar histori tetap bisa diaudit, tidak dihapus otomatis)
sb[ukuran]=arr;
var hppPerUnit=qty>0?hppTotal/qty:0;
return{data:{...data,stokBatch:sb},hppTotal,hppPerUnit,detail,qtyKurang};
}

// Kembalikan qty ke batch FIFO (reverse consume) — dipakai saat hapus DO/item
// Strategi: tambahkan kembali ke batch paling akhir yang cocok sumber-nya, atau buat entry baru
function reverseBatch(data,ukuran,qty,sumber,harga){
qty=Number(qty)||0;if(qty<=0)return data;
var sb={...(data.stokBatch||{})};
var arr=(sb[ukuran]||[]).map(b=>({...b}));
// Cari batch yang cocok sumbernya (paling baru) dan tambah qtySisa-nya kembali
var reversed=false;
for(var i=arr.length-1;i>=0;i--){
  if(arr[i].sumber&&sumber&&arr[i].sumber===sumber){
    arr[i].qtySisa=Math.min(arr[i].qty,(arr[i].qtySisa||0)+qty);
    reversed=true;break;
  }
}
// Kalau tidak ketemu batch yang cocok, tambah entry baru
if(!reversed){
  arr.push({id:uid(),tanggal:toDay(),qty,qtySisa:qty,harga:Number(harga)||0,sumber:(sumber||"")+" (Reverse)"});
}
sb[ukuran]=arr;
return{...data,stokBatch:sb};
}

// Hitung nilai stok saat ini berdasarkan SISA batch FIFO (bukan qty x harga terakhir)
function calcNilaiStokFIFO(data){
return SIZES.reduce((a,s)=>{
  var arr=(data.stokBatch||{})[s]||[];
  var nilai=arr.reduce((aa,b)=>aa+Math.max(0,b.qtySisa||0)*b.harga,0);
  return a+nilai;
},0);
}

// Total qty sisa di semua batch suatu ukuran (untuk validasi vs data.stock)
function totalQtyBatch(data,ukuran){return((data.stokBatch||{})[ukuran]||[]).reduce((a,b)=>a+Math.max(0,b.qtySisa||0),0);}


function terbilang(n){
if(n==null||isNaN(n))return"Nol";n=Math.floor(Math.abs(n));if(n===0)return"Nol";
var s=["","Satu","Dua","Tiga","Empat","Lima","Enam","Tujuh","Delapan","Sembilan","Sepuluh","Sebelas"];
function f(n){
if(n<12)return s[n];
if(n<20)return f(n-10)+" Belas";
if(n<100)return f(Math.floor(n/10))+" Puluh"+(n%10?" "+f(n%10):"");
if(n<200)return"Seratus"+(n-100?" "+f(n-100):"");
if(n<1000)return f(Math.floor(n/100))+" Ratus"+(n%100?" "+f(n%100):"");
if(n<2000)return"Seribu"+(n-1000?" "+f(n-1000):"");
if(n<1000000)return f(Math.floor(n/1000))+" Ribu"+(n%1000?" "+f(n%1000):"");
if(n<1000000000)return f(Math.floor(n/1000000))+" Juta"+(n%1000000?" "+f(n%1000000):"");
return f(Math.floor(n/1000000000))+" Milyar"+(n%1000000000?" "+f(n%1000000000):"");}
return f(n).trim()+" Rupiah";}

function nextInvNo(data,tanggal){var d=new Date(tanggal+"T00:00:00");var m=d.getMonth();var y=d.getFullYear();var key=y+"-"+String(m+1).padStart(2,"0");var ctr=(data.counters?.inv||{})[key]||0;var n=ctr+1;return{no:"#HTS/INV/"+ROMAN[m]+"."+String(y).slice(-2)+"/"+String(n).padStart(3,"0"),key,n};}
function nextSGNo(data,bulan){var[y,m]=bulan.split("-").map(Number);var ctr=(data.counters?.sg||{})[bulan]||0;var n=ctr+1;return{no:String(n).padStart(3,"0")+"/HTS/SG/"+ROMAN[m-1]+"/"+y,key:bulan,n};}
function nextRegNo(data){var ctr=(data.counters?.reg||0)+1;return{no:"HTS/CST/"+String(ctr).padStart(3,"0"),n:ctr};}

function calcBonusSales(empId,bulan,data){var penj=(data.penjualan||[]).filter(p=>p.salesId===empId&&(p.tanggal||"").startsWith(bulan));var q55=0,q12=0;penj.forEach(p=>(p.items||[]).forEach(it=>{if(it.ukuran==="5.5 kg")q55+=Number(it.qty||0);else if(it.ukuran==="12 kg")q12+=Number(it.qty||0);}));var r55=q55<500?500:q55;var b55=q55*r55;var r12=q12<500?500:q12;var b12=q12*r12;return{q55,r55,b55,q12,r12,b12,total:b55+b12};}

function getBiayaOpsAuto(empId,bulan,data){return(data.pengeluaran||[]).filter(p=>p.karyawanId===empId&&(p.tanggal||"").startsWith(bulan)&&KAT_OPS.includes(p.kategori)).map(p=>({id:p.id,label:p.kategori+(p.ket?" - "+p.ket:""),nominal:Number(p.nominal||0),kategori:p.kategori}));}

// Satu fungsi tunggal: Total Pinjaman = semua ambilan dari kedua sumber DIKURANGI semua cicilan payroll
// Ini fungsi yang SAMA dipakai di: Rekap Ambilan, Payroll list, Slip Gaji
function getTotalAmbilanKaryawan(empId,data){
var emp=(data.employees||[]).find(e=>e.id===empId);
var empNama=(emp?.nama||"").toLowerCase().trim();
// Sumber 1: data.ambilan (kurang setor otomatis dari setoran sales) — match by id ATAU nama
var s1=(data.ambilan||[]).filter(a=>{
  if(a.karyawanId&&a.karyawanId===empId)return true;
  if(empNama&&(a.karyawanNama||"").toLowerCase().trim()===empNama)return true;
  return false;
}).reduce((a,x)=>a+Number(x.nominal||0),0);
// Sumber 2: data.pengeluaran kategori kasbon/ambilan — match by id ATAU nama
var s2=(data.pengeluaran||[]).filter(p=>{
  var k=(p.kategori||"").toLowerCase();
  if(!k.includes("kasbon")&&!k.includes("ambilan"))return false;
  if(p.karyawanId&&p.karyawanId===empId)return true;
  if(empNama&&(p.karyawanNama||"").toLowerCase().trim()===empNama)return true;
  return false;
}).reduce((a,p)=>a+Number(p.nominal||0),0);
// Dikurangi: semua cicilan yang sudah dipotong via slip gaji
var dipotong=(data.payrollLog||[]).filter(p=>p.empId===empId).reduce((a,x)=>a+Number(x.potonganPinjaman||0),0);
return Math.max(0,s1+s2-dipotong);
}
// Alias untuk backward compatibility (bulan param diabaikan — pinjaman berjalan tidak dibatasi bulan)
function getPinjamanSaldo(empId,bulan,data){return getTotalAmbilanKaryawan(empId,data);}

function calcPayrollFull(emp,bulan,data){var abs=(data.absensi||[]).filter(a=>a.karyawanId===emp.id&&(a.tanggal||"").startsWith(bulan));var hariHadir=abs.filter(a=>a.status==="H").length;var totalHariKerja=daysInMonth(bulan);var absen=abs.filter(a=>a.status==="A").length;var isSales=["sales_driver","sales_freelance"].includes(emp.role);var bonus=isSales?calcBonusSales(emp.id,bulan,data):{q55:0,r55:500,b55:0,q12:0,r12:500,b12:0,total:0};var mode=emp.uangMakanMode||"harian";var makanAuto=(data.pengeluaran||[]).filter(p=>p.karyawanId===emp.id&&["Uang Makan Karyawan","Uang Makan","Makan Karyawan"].includes(p.kategori)&&(p.tanggal||"").startsWith(bulan)).reduce((a,p)=>a+Number(p.nominal||0),0);var uangMakanRate=emp.uangMakan||UANG_MAKAN_DEFAULT;var uangMakan=mode==="harian"?makanAuto:hariHadir*uangMakanRate;var ops=getBiayaOpsAuto(emp.id,bulan,data).filter(x=>!["Uang Makan","Makan Karyawan"].some(k=>x.label.startsWith(k)));var bongkarTotal=ops.filter(x=>x.kategori==="Uang Bongkar DO").reduce((a,x)=>a+x.nominal,0);var spbbeTotal=ops.filter(x=>x.kategori==="Uang Jalan SPBE").reduce((a,x)=>a+x.nominal,0);var bongkarCount=ops.filter(x=>x.kategori==="Uang Bongkar DO").length;var spbbeCount=ops.filter(x=>x.kategori==="Uang Jalan SPBE").length;var pinjamanSaldo=getPinjamanSaldo(emp.id,bulan,data);return{hariHadir,totalHariKerja,absen,gajiPokok:emp.gajiPokok||0,bonus,uangMakanMode:mode,uangMakan,bongkarTotal,spbbeTotal,bongkarCount,spbbeCount,pinjamanSaldo};}

function calcNilaiStok(data){return SIZES.reduce((a,s)=>{var qty=(data.stock||{})[s]||0;var modal=getModal(data,s,"Isi");return a+qty*modal;},0);}
function calcPinjamanKaryawan(data){
var dariAmbilan=(data.ambilan||[]).reduce((a,x)=>a+Number(x.nominal||0),0);
var dariKasbonPengeluaran=(data.pengeluaran||[]).filter(p=>{var k=(p.kategori||"").toLowerCase();return k.includes("kasbon")||k.includes("ambilan");}).reduce((a,p)=>a+Number(p.nominal||0),0);
var dipotongPayroll=(data.payrollLog||[]).reduce((a,x)=>a+Number(x.potonganPinjaman||0),0);
return dariAmbilan+dariKasbonPengeluaran-dipotongPayroll;
}
function calcTotalPiutang(data){return(data.bon||[]).filter(b=>b.status!=="lunas").reduce((a,b)=>a+(b.sisaTagihan||0),0);}

// ─── PRINT HELPER (Simpan PDF via dialog browser) ─────────────────────────────
// Tidak butuh CDN. User pilih "Save as PDF" di dialog print browser.

// Buat nama file otomatis
function makeFileName(type,label1,label2,ext){
  var d=new Date();
  var tgl=d.getDate().toString().padStart(2,"0")+"-"+["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Ags","Sep","Okt","Nov","Des"][d.getMonth()]+"-"+d.getFullYear();
  var clean=function(s){return(s||"").toUpperCase().replace(/[^A-Z0-9]/g,"-").replace(/-+/g,"-").replace(/^-|-$/g,"").slice(0,30);};
  if(type==="inv")return clean(label1)+"_"+clean(label2)+"_"+tgl+"."+ext;
  if(type==="slip")return clean(label1)+"_Slip-Gaji_"+clean(label2)+"."+ext;
  if(type==="tb")return "Tutup-Buku_"+clean(label1)+"."+ext;
  if(type==="do")return "Laporan-DO_"+clean(label1)+"."+ext;
  return clean(label1)+"_"+tgl+"."+ext;
}

// Cetak PDF dengan saran nama file
function doPrint(id){
var el=document.getElementById(id);
if(!el){alert("Elemen tidak ditemukan: "+id);return;}
var html=el.innerHTML;
var pw=window.open("","_blank","width=900,height=700");
if(!pw){alert("Popup diblokir browser. Izinkan popup untuk mencetak.");return;}
pw.document.write(
"<!DOCTYPE html><html><head><meta charset='utf-8'>"+
"<title>Cetak Laporan</title>"+
"<style>"+
"*{box-sizing:border-box;margin:0;padding:0;}"+
"html,body{width:210mm;}"+
"body{font-family:'Plus Jakarta Sans',Arial,sans-serif;color:#111;background:white;padding:0;}"+
"table{width:100%;border-collapse:collapse;margin-bottom:12px;}"+
"th{background:#0a1f44;color:white;padding:6px 8px;font-size:10px;font-weight:700;text-align:center;border:1px solid #ccc;}"+
"td{padding:5px 8px;font-size:11px;border:1px solid #ddd;vertical-align:top;}"+
"tr{page-break-inside:avoid;}"+
"thead{display:table-header-group;}"+
"#_inv,#_slip,#_ba,#_slip_setoran_sales,#_slip_setor{max-width:210mm!important;width:210mm!important;margin:0 auto!important;border-radius:0!important;box-shadow:none!important;}"+
"@page{margin:10mm;size:A4;}"+
"@media print{body{padding:0;}}"+
"</style></head><body>"+
html+
"</body></html>"
);
pw.document.close();
pw.focus();
setTimeout(function(){pw.print();},600);
}

// Download PNG
function doDownloadPNG(id,fileName,onDone){
var el=document.getElementById(id);
if(!el){alert("Element tidak ditemukan");return;}
if(typeof html2canvas==="undefined"){alert("html2canvas belum dimuat, coba lagi sebentar");return;}
html2canvas(el,{scale:2,useCORS:true,backgroundColor:"#ffffff",logging:false}).then(function(canvas){
  var a=document.createElement("a");
  a.href=canvas.toDataURL("image/png");
  a.download=fileName||"dokumen.png";
  a.click();
  if(onDone)onDone();
}).catch(function(e){alert("Gagal buat PNG: "+e.message);});
}

// Copy PNG ke clipboard
function doCopyPNG(id,onDone){
var el=document.getElementById(id);
if(!el){alert("Element tidak ditemukan");return;}
if(typeof html2canvas==="undefined"){alert("html2canvas belum dimuat, coba lagi sebentar");return;}
html2canvas(el,{scale:2,useCORS:true,backgroundColor:"#ffffff",logging:false}).then(function(canvas){
  canvas.toBlob(function(blob){
    if(!blob){alert("Gagal membuat gambar");return;}
    try{
      navigator.clipboard.write([new ClipboardItem({"image/png":blob})]).then(function(){
        if(onDone)onDone("copy");
        alert("✅ Gambar ter-copy! Sekarang paste ke WhatsApp atau chat.");
      }).catch(function(){
        // Fallback: download saja
        var a=document.createElement("a");
        a.href=canvas.toDataURL("image/png");
        a.download="dokumen.png";
        a.click();
        alert("Copy tidak didukung browser ini, otomatis download.");
      });
    }catch(e){alert("Copy gagal: "+e.message);}
  },"image/png");
}).catch(function(e){alert("Gagal buat PNG: "+e.message);});
}

// ─── THEME ────────────────────────────────────────────────────────────────────
var THEMES={
dark:{bg:"#0A0F18",card:"#111927",bdr:"#1F2937",nav:"#161E2D",red:"#C0392B",rdk:"#7B241C",rlt:"#F87171",org:"#E67E22",olt:"#F59E0B",grn:"#1D8348",glt:"#22C55E",blu:"#1D4ED8",blt:"#3B82F6",gry:"#475569",gl2:"#64748B",gltr:"#94A3B8",wht:"#F1F5F9",mode:"dark",inHv:"#13233D",inHvE:"#2A1517"},
light:{bg:"#F1F5F9",card:"#FFFFFF",bdr:"#E2E8F0",nav:"#F8FAFC",red:"#DC2626",rdk:"#991B1B",rlt:"#EF4444",org:"#EA580C",olt:"#F97316",grn:"#15803D",glt:"#16A34A",blu:"#0C4DA2",blt:"#2563EB",gry:"#94A3B8",gl2:"#64748B",gltr:"#475569",wht:"#0F172A",mode:"light",inHv:"#DBEAFE",inHvE:"#FEE2E2"}
};
var ThemeCtx=createContext(THEMES.light);
function useTheme(){return useContext(ThemeCtx);}

// ─── HOOKS ────────────────────────────────────────────────────────────────────
function useMobile(){var[m,setM]=useState(()=>typeof window!=="undefined"&&window.innerWidth<680);useEffect(()=>{var fn=()=>setM(window.innerWidth<680);window.addEventListener("resize",fn);return()=>window.removeEventListener("resize",fn);},[]);return m;}
function useToast(){var[toasts,setToasts]=useState([]);function toast(msg,type="success"){var id=uid();setToasts(t=>[...t,{id,msg,type}]);setTimeout(()=>setToasts(t=>t.filter(x=>x.id!==id)),3000);}return{toasts,toast};}

// ─── UI PRIMITIVES ────────────────────────────────────────────────────────────
function Toast({toasts}){var C=useTheme();if(!toasts.length)return null;return <div style={{position:"fixed",top:70,right:16,zIndex:9999,display:"flex",flexDirection:"column",gap:8}}>{toasts.map(t=><div key={t.id} style={{background:t.type==="error"?C.rdk:t.type==="warning"?C.org:C.grn,color:"#FFF",padding:"10px 16px",borderRadius:8,fontSize:13,fontWeight:600,boxShadow:"0 4px 20px rgba(0,0,0,.3)",maxWidth:320}}>{t.msg}</div>)}</div>;}
function Bdg({color="gray",children}){var C=useTheme();var d=C.mode==="dark";var m=d?{red:["#4A0E0E",C.rlt],green:["#0A2E14",C.glt],orange:["#3D200A",C.olt],blue:["#0A2040",C.blt],gray:["#1E2A35",C.gl2]}:{red:["#FEE2E2","#991B1B"],green:["#DCFCE7","#166534"],orange:["#FED7AA","#9A3412"],blue:["#DBEAFE","#1E40AF"],gray:["#E2E8F0","#475569"]};var p=m[color]||m.gray;return <span style={{background:p[0],color:p[1],padding:"2px 9px",borderRadius:20,fontSize:11,fontWeight:700,whiteSpace:"nowrap"}}>{children}</span>;}
function Card({children,style={}}){var C=useTheme();return <div style={{background:C.card,borderRadius:14,border:"1px solid "+C.bdr,padding:18,marginBottom:14,boxShadow:C.mode==="light"?"0 1px 3px rgba(0,0,0,0.04)":"0 4px 16px rgba(0,0,0,.28)",...style}}>{children}</div>;}
function STitle({icon,children}){var C=useTheme();return <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:16}}><span style={{fontSize:22}}>{icon}</span><h2 style={{margin:0,fontSize:17,fontWeight:800,color:C.wht}}>{children}</h2></div>;}
function Inp({label,type="text",value,onChange,placeholder="",ro=false,style={}}){var C=useTheme();var IS={width:"100%",border:"1px solid "+C.bdr,borderRadius:8,padding:"9px 11px",color:C.wht,fontSize:13,outline:"none",boxSizing:"border-box"};return <div style={{marginBottom:10,...style}}>{label&&<label style={{display:"block",fontSize:11,color:C.gl2,marginBottom:3,fontWeight:600}}>{label}</label>}<input type={type} value={value!=null?value:""} readOnly={ro} placeholder={placeholder} onChange={e=>onChange&&onChange(e.target.value)} style={{...IS,background:ro?C.bg:C.nav}}/></div>;}
function Sel({label,value,onChange,opts=[],style={}}){var C=useTheme();var IS={width:"100%",border:"1px solid "+C.bdr,borderRadius:8,padding:"9px 11px",color:C.wht,fontSize:13,outline:"none",boxSizing:"border-box"};return <div style={{marginBottom:10,...style}}>{label&&<label style={{display:"block",fontSize:11,color:C.gl2,marginBottom:3,fontWeight:600}}>{label}</label>}<select value={value!=null?value:""} onChange={e=>onChange(e.target.value)} style={{...IS,background:C.nav}}>{opts.map(o=>{var v=o!=null&&o.v!=null?o.v:o;var l=o!=null&&o.l!=null?o.l:o;return <option key={String(v)} value={String(v)}>{String(l)}</option>;})}</select></div>;}
function Btn({children,color="blue",sm=false,dis=false,style={},onClick}){var C=useTheme();var bgs={blue:[C.blu,C.blt],red:[C.rdk,C.red],green:[C.grn,C.glt],orange:[C.mode==="dark"?"#854000":"#C2410C",C.org],gray:[C.gry,C.gl2]};var[hov,setHov]=useState(false);var pair=bgs[color]||bgs.blue;return <button onClick={onClick} disabled={dis} onMouseEnter={()=>setHov(true)} onMouseLeave={()=>setHov(false)} style={{background:dis?C.gry:hov?pair[1]:pair[0],color:"#FFF",border:"none",borderRadius:8,padding:sm?"6px 10px":"9px 16px",fontSize:sm?11:13,fontWeight:700,cursor:dis?"not-allowed":"pointer",transition:"all .15s",...style}}>{children}</button>;}
function Tbl({headers=[],rows=[],empty="Belum ada data",widths}){var C=useTheme();var fixed=!!widths;var allSized=fixed&&widths.every(w=>w);return <div style={{overflowX:"auto"}}><table style={{width:allSized?"auto":"100%",minWidth:allSized?widths.reduce((a,w)=>a+(Number(w)||120),0)+"px":undefined,maxWidth:"100%",margin:allSized?"0 auto":undefined,borderCollapse:"collapse",fontSize:13.5,tableLayout:fixed?"fixed":"auto"}}>{fixed&&<colgroup>{headers.map((h,i)=><col key={i} style={{width:(widths[i]||200)+"px"}}/>)}</colgroup>}<thead><tr>{headers.map((h,i)=><th key={i} style={{padding:"9px 11px",background:C.nav,color:C.gl2,fontWeight:700,textAlign:"left",fontSize:11.5,borderBottom:"2px solid "+C.bdr,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{h}</th>)}</tr></thead><tbody>{!rows.length?<tr><td colSpan={headers.length} style={{padding:20,textAlign:"center",color:C.gl2}}>{empty}</td></tr>:rows.map((r,i)=><tr key={i} style={{borderBottom:"1px solid "+C.bdr}}>{r.map((c,j)=><td key={j} style={{padding:"8px 11px",color:C.wht,verticalAlign:"middle",overflow:"hidden"}}>{c}</td>)}</tr>)}</tbody></table></div>;}
function MCards({headers,rows,empty="Belum ada data"}){var C=useTheme();if(!rows.length)return <div style={{padding:20,textAlign:"center",color:C.gl2}}>{empty}</div>;return <div>{rows.map((r,i)=><div key={i} style={{background:C.nav,borderRadius:10,padding:12,marginBottom:8,border:"1px solid "+C.bdr}}>{headers.map((h,j)=><div key={j} style={{display:"flex",justifyContent:"space-between",padding:"4px 0",borderBottom:j<headers.length-1?"1px solid "+C.bdr:"none"}}><span style={{fontSize:12,color:C.gl2,marginRight:8}}>{h}</span><span style={{fontSize:13,color:C.wht,textAlign:"right"}}>{r[j]}</span></div>)}</div>)}</div>;}
function RTbl({headers,rows,empty,widths}){var m=useMobile();return m?<MCards headers={headers} rows={rows} empty={empty}/>:<Tbl headers={headers} rows={rows} empty={empty} widths={widths}/>;}
function SC({label,value,icon,color,sub}){var C=useTheme();return <div style={{background:C.card,borderRadius:12,padding:"12px 14px",border:"1px solid "+C.bdr,flex:1,minWidth:110,boxShadow:C.mode==="light"?"0 1px 3px rgba(0,0,0,0.04)":"none"}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}><div><div style={{fontSize:10,color:C.gl2,fontWeight:600,marginBottom:2}}>{label}</div><div style={{fontSize:15,fontWeight:900,color:color||C.blt,lineHeight:1.1}}>{value}</div>{sub&&<div style={{fontSize:10,color:C.gl2,marginTop:2}}>{sub}</div>}</div><span style={{fontSize:18,opacity:.7}}>{icon}</span></div></div>;}
function AutoInp({label,value="",onChange,options=[],placeholder="",onSelect}){var C=useTheme();var IS={width:"100%",border:"1px solid "+C.bdr,borderRadius:8,padding:"9px 11px",color:C.wht,fontSize:13,outline:"none",boxSizing:"border-box"};var[show,setShow]=useState(false);var filtered=options.filter(o=>!value||o.toLowerCase().includes(value.toLowerCase())).slice(0,7);return <div style={{position:"relative",marginBottom:10}}>{label&&<label style={{display:"block",fontSize:11,color:C.gl2,marginBottom:3,fontWeight:600}}>{label}</label>}<input value={value} placeholder={placeholder} onChange={e=>{onChange(e.target.value);setShow(true);}} onFocus={()=>setShow(true)} onBlur={()=>setTimeout(()=>setShow(false),200)} style={{...IS,background:C.nav}}/>{show&&filtered.length>0&&<div style={{position:"absolute",top:"100%",left:0,right:0,background:C.card,border:"1px solid "+C.bdr,borderRadius:8,zIndex:200,maxHeight:160,overflowY:"auto",boxShadow:"0 4px 16px rgba(0,0,0,.3)"}}>{filtered.map(o=><div key={o} onMouseDown={()=>{onChange(o);if(onSelect)onSelect(o);setShow(false);}} style={{padding:"9px 12px",cursor:"pointer",color:C.wht,fontSize:13,borderBottom:"1px solid "+C.bdr}}>{o}</div>)}</div>}</div>;}
function ActBtns({onEdit,onDel,onPrint}){var C=useTheme();return <div style={{display:"flex",gap:4}}>{onPrint&&<button onClick={onPrint} title="Cetak Invoice" style={{background:C.inHv,border:"1px solid "+C.blt,borderRadius:6,padding:"4px 7px",color:C.blt,cursor:"pointer",fontSize:12}}>🖨️</button>}{onEdit&&<button onClick={onEdit} style={{background:C.inHv,border:"1px solid "+C.blt,borderRadius:6,padding:"4px 7px",color:C.blt,cursor:"pointer",fontSize:12}}>✏️</button>}{onDel&&<button onClick={onDel} style={{background:C.inHvE,border:"1px solid "+C.rlt,borderRadius:6,padding:"4px 7px",color:C.rlt,cursor:"pointer",fontSize:12}}>🗑️</button>}</div>;}
function ConfirmDel({msg,onConfirm,onCancel}){var C=useTheme();return <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.65)",zIndex:8500,display:"flex",alignItems:"center",justifyContent:"center"}}><div style={{background:C.card,borderRadius:14,padding:28,maxWidth:340,width:"90%",border:"1px solid "+C.rlt}}><div style={{color:C.wht,fontWeight:700,marginBottom:8,fontSize:16}}>🗑️ Hapus Data?</div><div style={{color:C.gl2,fontSize:13,marginBottom:20}}>{msg}</div><div style={{display:"flex",gap:10}}><Btn color="red" onClick={onConfirm}>Ya, Hapus</Btn><Btn color="gray" onClick={onCancel}>Batal</Btn></div></div></div>;}
function Modal({title,children,onClose,onSave,saveLabel="💾 Simpan",width=560}){var C=useTheme();return <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.6)",zIndex:8000,display:"flex",alignItems:"center",justifyContent:"center",padding:12}}><div style={{background:C.card,borderRadius:14,padding:22,maxWidth:width,width:"100%",border:"1px solid "+C.bdr,maxHeight:"92vh",overflowY:"auto"}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}><b style={{color:C.wht,fontSize:15}}>{title}</b><button onClick={onClose} style={{background:"none",border:"none",color:C.gl2,fontSize:22,cursor:"pointer"}}>✕</button></div>{children}{onSave&&<div style={{display:"flex",gap:10,marginTop:14}}><Btn color="green" onClick={onSave}>{saveLabel}</Btn><Btn color="gray" onClick={onClose}>Batal</Btn></div>}</div></div>;}
function MonthPicker({label,value,onChange}){var C=useTheme();var IS={width:"100%",border:"1px solid "+C.bdr,borderRadius:8,padding:"9px 11px",color:C.wht,fontSize:13,outline:"none",boxSizing:"border-box"};return <div style={{marginBottom:10}}>{label&&<label style={{display:"block",fontSize:11,color:C.gl2,marginBottom:3,fontWeight:600}}>{label}</label>}<input type="month" value={value} onChange={e=>onChange(e.target.value)} style={{...IS,background:C.nav}}/></div>;}

// ─── FILTER TABLE (NEW v4) ────────────────────────────────────────────────────
// Tabel dengan: filter mini di header + sort klik header, dipakai di Penjualan, Piutang, Pengeluaran, Laporan
function FilterTbl({columns,data:rows,empty="Belum ada data",maxRows=200,renderMobileCard,footerRow}){
var C=useTheme();var mob=useMobile();
var[hdrFilters,setHdrFilters]=useState({});
var[sort,setSort]=useState({key:null,dir:"asc"});
var filtered=useMemo(()=>{
return rows.filter(row=>{
return columns.every(col=>{
if(!col.filterable)return true;
var fv=(hdrFilters[col.key]||"").trim().toLowerCase();
if(!fv)return true;
var val=col.filterVal?col.filterVal(row):row[col.key];
if(val==null)return false;
return String(val).toLowerCase().includes(fv);
});
});
},[rows,columns,hdrFilters]);
var sorted=useMemo(()=>{
if(!sort.key)return filtered;
var col=columns.find(c=>c.key===sort.key);if(!col)return filtered;
var get=col.sortVal||(r=>r[sort.key]);
var arr=filtered.slice();
arr.sort((a,b)=>{var va=get(a);var vb=get(b);if(va==null)return 1;if(vb==null)return -1;if(typeof va==="number"&&typeof vb==="number")return sort.dir==="asc"?va-vb:vb-va;return sort.dir==="asc"?String(va).localeCompare(String(vb)):String(vb).localeCompare(String(va));});
return arr;
},[filtered,sort,columns]);
var display=sorted.slice(0,maxRows);
function toggleSort(k){var col=columns.find(c=>c.key===k);if(!col||col.sortable===false)return;setSort(s=>s.key===k?{key:k,dir:s.dir==="asc"?"desc":"asc"}:{key:k,dir:"asc"});}
function setFilter(k,v){setHdrFilters(p=>({...p,[k]:v}));}
if(mob&&renderMobileCard){return <div>{display.map((r,i)=><div key={i}>{renderMobileCard(r,i)}</div>)}{!display.length&&<div style={{padding:20,textAlign:"center",color:C.gl2}}>{empty}</div>}</div>;}
return <div>
<div style={{overflowX:"auto"}}>
<table style={{width:columns.every(c=>c.width)?"auto":"100%",minWidth:columns.every(c=>c.width)?columns.reduce((a,c)=>a+(Number(c.width)||120),0)+"px":undefined,maxWidth:"100%",margin:columns.every(c=>c.width)?"0 auto":undefined,borderCollapse:"collapse",fontSize:13.5,tableLayout:"fixed"}}>
<colgroup>{columns.map(c=><col key={c.key} style={{width:c.width?(typeof c.width==="number"?c.width+"px":c.width):"200px"}}/>)}</colgroup>
<thead>
<tr>{columns.map(c=><th key={c.key} onClick={()=>toggleSort(c.key)} style={{padding:"9px 11px",background:C.nav,color:C.gl2,fontWeight:700,textAlign:"left",fontSize:11.5,borderBottom:"2px solid "+C.bdr,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis",cursor:c.sortable!==false?"pointer":"default",userSelect:"none"}}>{c.label}{c.sortable!==false&&<span style={{marginLeft:4,color:sort.key===c.key?C.blt:C.gry,fontSize:9}}>{sort.key===c.key?(sort.dir==="asc"?"▲":"▼"):"⇅"}</span>}</th>)}</tr>
<tr>{columns.map(c=><th key={c.key+"_f"} style={{padding:"4px 6px",background:C.bg,borderBottom:"1px solid "+C.bdr}}>{c.filterable===false?<span/>:c.filterType==="select"?<select value={hdrFilters[c.key]||""} onChange={e=>setFilter(c.key,e.target.value)} style={{width:"100%",background:C.card,border:"1px solid "+C.bdr,borderRadius:5,padding:"4px 6px",fontSize:11.5,color:C.wht,outline:"none"}}><option value="">Semua</option>{(c.options||[]).map(o=><option key={o.v||o} value={o.v||o}>{o.l||o}</option>)}</select>:<input value={hdrFilters[c.key]||""} onChange={e=>setFilter(c.key,e.target.value)} placeholder="Filter..." style={{width:"100%",background:C.card,border:"1px solid "+C.bdr,borderRadius:5,padding:"4px 7px",fontSize:11.5,color:C.wht,outline:"none"}}/>}</th>)}</tr>
</thead>
<tbody>
{!display.length?<tr><td colSpan={columns.length} style={{padding:20,textAlign:"center",color:C.gl2}}>{empty}</td></tr>:display.map((r,i)=><tr key={i} style={{borderBottom:"1px solid "+C.bdr}}>{columns.map(c=><td key={c.key} style={{padding:"9px 11px",color:C.wht,overflow:"hidden"}}>{c.render?c.render(r):r[c.key]}</td>)}</tr>)}
</tbody>
{footerRow&&display.length>0&&<tfoot>{footerRow(sorted)}</tfoot>}
</table>
</div>
{(Object.values(hdrFilters).some(v=>v)||sort.key)&&<div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"6px 0",fontSize:11,color:C.gl2}}><span>{sorted.length} dari {rows.length} hasil</span><Btn sm color="gray" onClick={()=>{setHdrFilters({});setSort({key:null,dir:"asc"});}}>✕ Reset Filter</Btn></div>}
</div>;
}

// ─── LOGO ─────────────────────────────────────────────────────────────────────
function LPGLogo({size=72}){return <svg width={size} height={size*0.85} viewBox="0 0 70 60" fill="none"><rect x="2" y="14" width="13" height="40" fill="#ED1C24" rx="1"/><rect x="19" y="2" width="13" height="52" fill="#F2A900" rx="1"/><rect x="36" y="10" width="13" height="44" fill="#00A651" rx="1"/><text x="26" y="14" textAnchor="middle" fontSize="7" fontWeight="700" fill="#0A2C5C">1982</text></svg>;}
function CompanyLogo({h=52,variant="dark"}){var src=variant==="light"?LOGO_HTS_LIGHT:LOGO_HTS_DARK;return <img src={src} style={{height:h,objectFit:"contain"}} alt="PT. Hoe Trang Sa" onError={e=>{e.target.style.display="none";}}/>;}
function PertaminaLogo({h=32,variant="dark"}){var src=variant==="light"?LOGO_PERTAMINA_LIGHT:LOGO_PERTAMINA_DARK;return <img src={src} style={{height:h,objectFit:"contain"}} alt="Pertamina" onError={e=>{e.target.style.display="none";}}/>;}

// === AKHIR BAGIAN 1 ===
// === BAGIAN 2 DARI 3 ===

// ─── INVOICE BARU v4 (no double logo, PDF/PNG buttons) ────────────────────────
function InvoiceView({inv,company={},onClose}){
if(!inv)return null;
var isLunas=inv.bonLunas||(inv.metodeBayar!=="BON"&&(inv.metodeBayar||"").toLowerCase().indexOf("bon")<0&&!inv.isBon);
var isSplitInv=!!(inv.splitDetail);
var sdInv=inv.splitDetail||{};
var isGabunganInv=!!(inv.isGabungan);
var total=inv.total||0;
var NAVY="#0a1f44";var NAVY2="#122d5e";var BLUE="#1565c0";
var RED="#e53935";var GREEN="#6ab04c";var TEAL="#00acc1";var TEAL_LIGHT="#e0f7fa";
var WHITE="#fff";var G100="#eef2f9";var G200="#dde4f0";var G400="#8fa3c0";var G600="#4a6080";
var FONT="'Plus Jakarta Sans',-apple-system,'Segoe UI',sans-serif";
var HARI_ID=["Minggu","Senin","Selasa","Rabu","Kamis","Jumat","Sabtu"];
function fDHari(d){if(!d)return"-";var dt=new Date(d+"T00:00:00");return HARI_ID[dt.getDay()]+", "+dt.toLocaleDateString("id-ID",{day:"2-digit",month:"long",year:"numeric"});}
var salesNama=inv.salesNama||"";
return <div id="_inv_wrap" style={{position:"fixed",inset:0,background:"#cdd3db",zIndex:9500,padding:16,overflowY:"auto"}}>
<div id="_inv" style={{maxWidth:700,margin:"0 auto",background:WHITE,borderRadius:10,overflow:"hidden",boxShadow:"0 20px 60px rgba(10,31,68,.14)",color:"#0d1f3c",fontFamily:FONT}}>

{/* HEADER navy */}
<div style={{background:"linear-gradient(135deg,"+NAVY+" 0%,"+NAVY2+" 100%)",padding:"20px 26px 17px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
<div style={{display:"flex",alignItems:"center",gap:12}}>
<CompanyLogo h={48} variant="dark"/>
</div>
{/* Pertamina */}
<div style={{display:"flex",alignItems:"center",gap:7}}>
<PertaminaLogo h={32} variant="dark"/>
</div>
</div>

{/* DIVIDER 3 warna */}
<div style={{height:3,display:"flex"}}>
<div style={{flex:1,background:BLUE}}/><div style={{flex:1,background:GREEN}}/><div style={{flex:1,background:RED}}/>
</div>

{/* TITLE BLOCK — putih */}
<div style={{background:WHITE,padding:"18px 26px 16px"}}>
{/* Judul + badge status */}
<div style={{display:"flex",alignItems:"center",justifyContent:"center",position:"relative",marginBottom:4}}>
<div style={{fontSize:26,fontWeight:800,color:NAVY,letterSpacing:3,textAlign:"center"}}>I N V O I C E</div>
{isLunas
?<div style={{position:"absolute",right:0,fontSize:10,fontWeight:700,padding:"3px 11px",borderRadius:20,background:"#D1FAE5",border:"1px solid #10B981",color:"#065F46"}}>✓ Lunas</div>
:inv.bonSebagian
?<div style={{position:"absolute",right:0,fontSize:10,fontWeight:700,padding:"3px 11px",borderRadius:20,background:"#FEF3C7",border:"1px solid #F59E0B",color:"#92400E"}}>◐ Sebagian</div>
:<div style={{position:"absolute",right:0,fontSize:10,fontWeight:700,padding:"3px 11px",borderRadius:20,background:"#FEE2E2",border:"1px solid #EF4444",color:"#991B1B"}}>— Belum Lunas</div>}
</div>
{/* No invoice di bawah judul */}
<div style={{textAlign:"center",fontSize:11,fontWeight:600,color:G400,letterSpacing:.5,marginBottom:14}}>NO. <span style={{color:NAVY,fontWeight:700}}>{inv.noInv}</span></div>
{/* Kepada kiri, Tanggal+Sales kanan */}
<div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:16}}>
<div style={{borderLeft:"3px solid "+NAVY,paddingLeft:12}}>
<div style={{fontSize:8,fontWeight:700,letterSpacing:1.2,textTransform:"uppercase",color:G400,marginBottom:4}}>Kepada Yth.</div>
<div style={{fontSize:17,fontWeight:800,color:NAVY,lineHeight:1.2,marginBottom:2}}>{inv.konsumen}</div>
<div style={{fontSize:11,color:G400}}>Di — {inv.kota||"Banda Aceh"}</div>
</div>
<div style={{display:"flex",flexDirection:"column",alignItems:"flex-end",gap:5}}>
<div style={{display:"flex",flexDirection:"column",alignItems:"flex-end"}}>
<div style={{fontSize:8,fontWeight:700,letterSpacing:1,textTransform:"uppercase",color:G400,marginBottom:1}}>Tanggal</div>
<div style={{fontSize:12,fontWeight:700,color:NAVY}}>{fDHari(inv.tanggal)}</div>
</div>
{salesNama&&<div style={{display:"flex",flexDirection:"column",alignItems:"flex-end"}}>
<div style={{fontSize:8,fontWeight:700,letterSpacing:1,textTransform:"uppercase",color:G400,marginBottom:1}}>Sales</div>
<div style={{fontSize:11.5,fontWeight:700,color:BLUE}}>{salesNama}</div>
</div>}
</div>
</div>
</div>

{/* GAP abu */}
<div style={{height:12,background:"#f4f7fc"}}/>

{/* TABEL — putih */}
<div style={{margin:"0 26px",borderRadius:8,overflow:"hidden",border:"1px solid "+G200,boxShadow:"0 2px 12px rgba(10,31,68,.06)"}}>
<table style={{width:"100%",borderCollapse:"collapse",fontFamily:FONT,fontSize:11.5}}>
<thead><tr style={{background:NAVY}}>
{[{l:"Date",w:"18%",a:"left"},{l:"Detail Produk",w:"26%",a:"left"},{l:"Qty",w:"10%",a:"center"},{l:"",w:"5%",a:"center"},{l:"Unit Price",w:"19%",a:"right"},{l:"Line Total",w:"22%",a:"right"}].map((h,i)=><th key={i} style={{width:h.w,padding:"9px 12px",color:WHITE,fontSize:8,fontWeight:700,letterSpacing:1,textTransform:"uppercase",textAlign:h.a}}>{h.l}</th>)}
</tr></thead>
<tbody>
{(inv.items||[]).map((it,i)=>{
var prodLabel=it.jenis==="Tabung+Isi"?"Tbg + Isi":"Refill";
// Untuk invoice gabungan, pakai tglDO per item; biasa pakai tanggal invoice
var tglTampil=it.tglDO?it.tglDO:inv.tanggal;
var tglHari=new Date(tglTampil+"T00:00:00");
var HARI_SHORT=["Minggu","Senin","Selasa","Rabu","Kamis","Jumat","Sabtu"];
return <tr key={i} style={{background:i%2===0?WHITE:G100}}>
<td style={{padding:"10px 14px",color:G400,fontSize:10,fontWeight:600,lineHeight:1.5,borderBottom:"1px solid "+G200,whiteSpace:"nowrap"}}>{HARI_SHORT[tglHari.getDay()]}<br/>{fDs(tglTampil)}</td>
<td style={{padding:"10px 12px",borderBottom:"1px solid "+G200}}>
<span style={{display:"inline-block",background:NAVY,color:WHITE,borderRadius:20,padding:"2px 9px",fontSize:10,fontWeight:700,marginRight:6}}>{it.ukuran}</span>
<span style={{fontSize:10.5,color:G600,marginRight:3}}>{prodLabel}</span>
<span style={{fontWeight:600,color:"#0d1f3c"}}>LPG</span>
</td>
<td style={{padding:"10px 12px",textAlign:"center",fontWeight:700,color:NAVY,borderBottom:"1px solid "+G200}}>{it.qty}</td>
<td style={{padding:"10px 12px",textAlign:"center",color:G400,fontSize:11,fontWeight:700,borderBottom:"1px solid "+G200}}>×</td>
<td style={{padding:"10px 12px",textAlign:"right",borderBottom:"1px solid "+G200}}>
<span style={{fontSize:10,fontWeight:500,color:G400,marginRight:1}}>Rp </span>
<span style={{fontWeight:700,color:NAVY}}>{Number(it.price||0).toLocaleString("id-ID")}</span>
</td>
<td style={{padding:"10px 12px",textAlign:"right",borderBottom:"1px solid "+G200}}>
<span style={{fontSize:10,fontWeight:500,color:G400,marginRight:1}}>Rp </span>
<span style={{fontWeight:700,color:NAVY,fontSize:12.5}}>{(Number(it.qty||0)*Number(it.price||0)).toLocaleString("id-ID")}</span>
</td>
</tr>;})}
{Array.from({length:Math.max(0,2-(inv.items||[]).length)}).map((_,i)=><tr key={"e"+i} style={{background:((inv.items||[]).length+i)%2===0?WHITE:G100}}><td style={{padding:"8px 12px",borderBottom:"1px solid "+G200}}>&nbsp;</td><td style={{borderBottom:"1px solid "+G200}}/><td style={{borderBottom:"1px solid "+G200}}/><td style={{borderBottom:"1px solid "+G200}}/><td style={{borderBottom:"1px solid "+G200}}/><td style={{borderBottom:"1px solid "+G200}}/></tr>)}
</tbody>
</table>
{/* TERBILANG | LUNAS | GRAND TOTAL */}
<div style={{display:"flex",alignItems:"stretch",borderTop:"2px solid "+G200}}>
<div style={{flex:1,padding:"13px 14px",background:TEAL_LIGHT}}>
<div style={{fontSize:8,fontWeight:700,letterSpacing:1,textTransform:"uppercase",color:G600,marginBottom:4}}>Terbilang</div>
<div style={{fontSize:11.5,fontWeight:500,fontStyle:"italic",color:NAVY,lineHeight:1.5}}># {terbilang(inv.isBon?(inv.sisaTagihan||0):total)} #</div>
{/* Split payment detail */}
{isSplitInv&&<div style={{marginTop:8,borderTop:"1px dashed #00acc1",paddingTop:6}}>
<div style={{fontSize:8,fontWeight:700,letterSpacing:1,color:G600,marginBottom:4,textTransform:"uppercase"}}>Rincian Pembayaran</div>
{[["💵 Cash",sdInv.cash||0,"#065F46"],["🏦 Transfer "+(inv.splitBank||""),sdInv.tf||0,"#1e3a8a"],["📃 BON (Piutang)",sdInv.bon||0,"#991B1B"]].filter(x=>Number(x[1])>0).map(x=><div key={x[0]} style={{display:"flex",justifyContent:"space-between",fontSize:10,fontWeight:600,marginBottom:2}}>
<span style={{color:G600}}>{x[0]}</span><span style={{color:x[2]}}>Rp {Number(x[1]).toLocaleString("id-ID")}</span>
</div>)}
</div>}
{inv.isBon&&inv.totalBelanja&&<div style={{marginTop:8,borderTop:"1px dashed #dc2626",paddingTop:6}}>
<div style={{fontSize:8,fontWeight:700,letterSpacing:1,color:G600,marginBottom:5,textTransform:"uppercase"}}>Rincian Piutang</div>
<div style={{display:"flex",justifyContent:"space-between",fontSize:10,fontWeight:600,marginBottom:3}}>
<span style={{color:G600}}>Total Belanja</span><span style={{color:NAVY}}>Rp {Number(inv.totalBelanja).toLocaleString("id-ID")}</span>
</div>
{inv.tfDibayar>0&&<div style={{display:"flex",justifyContent:"space-between",fontSize:10,fontWeight:600,marginBottom:2}}><span style={{color:"#059669"}}>✓ Transfer {inv.splitBank||"BSI"}</span><span style={{color:"#059669"}}>Rp {Number(inv.tfDibayar).toLocaleString("id-ID")}</span></div>}{(inv.riwayatBayar||[]).map((p,i)=>{var tglBayar=p.tanggal||p.tgl;var nom=Number(p.jumlah||p.nominal||0);if(nom<=0)return null;return<div key={i} style={{display:"flex",justifyContent:"space-between",fontSize:10,fontWeight:600,marginBottom:2}}>
<span style={{color:"#059669"}}>✓ Bayar {tglBayar?new Date(tglBayar).toLocaleDateString("id-ID",{day:"numeric",month:"short"}):""} {p.metode?"("+p.metode+")":""}</span>
<span style={{color:"#059669"}}>Rp {nom.toLocaleString("id-ID")}</span>
</div>;})}
{(inv.totalDibayar>0||inv.tfDibayar>0)&&<div style={{display:"flex",justifyContent:"space-between",fontSize:10,fontWeight:700,marginTop:3,paddingTop:3,borderTop:"1px dashed "+G200}}>
<span style={{color:"#059669"}}>Total Dibayar</span><span style={{color:"#059669"}}>Rp {(Number(inv.totalDibayar||0)+Number(inv.tfDibayar||0)).toLocaleString("id-ID")}</span>
</div>}
<div style={{display:"flex",justifyContent:"space-between",fontSize:11,fontWeight:800,marginTop:4,paddingTop:4,borderTop:"2px solid #dc2626"}}>
<span style={{color:"#dc2626"}}>Sisa Tagihan</span><span style={{color:"#dc2626"}}>Rp {Number(inv.sisaTagihan||0).toLocaleString("id-ID")}</span>
</div>
</div>}
</div>
{/* Stempel LUNAS di antara terbilang & grand total */}
<div style={{display:"flex",alignItems:"center",justifyContent:"center",padding:"0 10px",background:TEAL_LIGHT}}>
{isLunas&&(company.stempelLunas
?<img src={company.stempelLunas} style={{height:52,opacity:.9,transform:"rotate(-5deg)"}} alt="lunas"/>
:<div style={{fontSize:13,fontWeight:800,color:RED,border:"2px solid "+RED,padding:"5px 11px",borderRadius:3,letterSpacing:4,opacity:.88,transform:"rotate(-5deg)",whiteSpace:"nowrap"}}>L U N A S</div>)}
</div>
<div style={{minWidth:190,padding:"13px 16px",background:WHITE,borderLeft:"2px solid "+G200,display:"flex",flexDirection:"column",justifyContent:"center",alignItems:"flex-end"}}>

<div style={{fontSize:8,fontWeight:700,letterSpacing:1,textTransform:"uppercase",color:inv.isBon?"#dc2626":G400,marginBottom:4}}>{inv.isBon?"Sisa Tagihan":"Grand Total"}</div>
<div style={{display:"flex",alignItems:"baseline",gap:4}}>
<span style={{fontSize:13,fontWeight:600,color:inv.isBon?"#dc2626":G400}}>Rp</span>
<span style={{fontSize:24,fontWeight:800,color:inv.isBon?"#dc2626":NAVY,letterSpacing:.3}}>{Number(inv.isBon?inv.sisaTagihan:total).toLocaleString("id-ID")}</span>
</div>
</div>
</div>
</div>

{/* GAP abu */}
<div style={{height:4,background:"#f4f7fc"}}/>

{/* BOTTOM — kiri rekening+WA, kanan TTD */}
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,padding:"16px 26px 18px",background:WHITE}}>
<div style={{display:"flex",flexDirection:"column",gap:10}}>

{/* REKENING — teal-light, bold, no rek besar */}
<div style={{borderRadius:8,overflow:"hidden",border:"1px solid "+G200}}>
<div style={{background:NAVY2,padding:"8px 14px"}}>
<div style={{fontSize:13,fontWeight:800,color:WHITE,letterSpacing:.5}}>BANK {company.bankNama||"BSI"}</div>
</div>
<div style={{background:TEAL_LIGHT}}>
<div style={{padding:"8px 14px",borderBottom:"1px solid rgba(0,172,193,.15)"}}>
<div style={{fontSize:8,fontWeight:700,letterSpacing:.8,textTransform:"uppercase",color:G600,marginBottom:2}}>A/C Atas Nama</div>
<div style={{fontSize:14,fontWeight:800,color:NAVY}}>{company.bankAtasNama||company.nama}</div>
</div>
<div style={{padding:"8px 14px"}}>
<div style={{fontSize:8,fontWeight:700,letterSpacing:.8,textTransform:"uppercase",color:G600,marginBottom:2}}>No. Rekening</div>
<div style={{fontFamily:"'Courier New',Courier,monospace",fontSize:20,fontWeight:800,color:NAVY,letterSpacing:1.5}}>{company.bankRekening||"812 69 2121 8"}</div>
</div>
</div>
</div>

{/* WA Konfirmasi */}
<div style={{background:TEAL_LIGHT,borderRadius:8,border:"1px solid "+G200,padding:"10px 13px"}}>
<div style={{fontSize:8,fontWeight:700,letterSpacing:.8,textTransform:"uppercase",color:G600,marginBottom:6}}>Konfirmasi Bukti Pembayaran Ke</div>
<div style={{fontSize:22,fontWeight:800,color:NAVY,lineHeight:1,letterSpacing:.5}}>{company.telepon||"0812 6900 2121"}</div>
</div>

<div style={{fontSize:10,color:G600,fontStyle:"italic",lineHeight:1.5}}>Demikian kami sampaikan, atas perhatian dan kerjasamanya kami ucapkan terima kasih.</div>
</div>

{/* TTD kanan — kotak persegi ramping */}
<div style={{display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"flex-start",paddingTop:8,gap:4,textAlign:"center"}}>
<div style={{fontSize:10,color:G400}}>Banda Aceh, {fDHari(inv.tanggal)}</div>
<div style={{fontSize:11,fontWeight:700,color:NAVY,marginTop:1,marginBottom:10}}>{company.nama}</div>
<div style={{width:140,height:70,border:"1.5px dashed "+G200,borderRadius:8,background:WHITE,display:"flex",alignItems:"center",justifyContent:"center",fontSize:8,color:G400,fontStyle:"italic",marginBottom:10,lineHeight:1.4,textAlign:"center",flexShrink:0}}>
{company.ttdKasir?<img src={company.ttdKasir} style={{height:66,objectFit:"contain"}} alt="ttd"/>:<><div style={{fontSize:8,color:"#8fa3c0"}}>TTD Elektronik</div><div style={{fontSize:8,color:"#8fa3c0"}}>upload di Pengaturan</div></>}
</div>
<div style={{fontSize:14,fontWeight:800,color:NAVY,borderTop:"2px solid "+NAVY,paddingTop:5,display:"inline-block",minWidth:140}}>{company.kasirNama||"MANARUL HIDAYAT"}</div>
<div style={{fontSize:8,fontWeight:700,letterSpacing:1.2,textTransform:"uppercase",color:G400,marginTop:3}}>Kasir</div>
</div>
</div>

{/* FOOTER navy */}
<div style={{background:NAVY,padding:"13px 22px",display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:12}}>
{[["📍","Alamat",[company.alamat]],["📞","No.HP & WA",[company.telepon,"TELP. "+company.telepon2]],["🌐","Email & Website",[company.email,company.website]]].map((x,i)=><div key={i} style={{display:"flex",alignItems:"flex-start",gap:10}}>
<div style={{width:28,height:28,background:"rgba(255,255,255,.1)",borderRadius:"50%",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0,fontSize:13}}>{x[0]}</div>
<div><div style={{fontSize:7,fontWeight:700,letterSpacing:.8,textTransform:"uppercase",color:"rgba(255,255,255,.4)",marginBottom:3}}>{x[1]}</div>{x[2].map((v,j)=><div key={j} style={{fontSize:9.5,fontWeight:600,color:WHITE,lineHeight:1.5}}>{v}</div>)}</div>
</div>)}
</div>

</div>
<div style={{maxWidth:700,margin:"12px auto",display:"flex",gap:8,justifyContent:"center",flexWrap:"wrap"}}>
<div style={{display:"flex",flexDirection:"column",gap:6,alignItems:"center"}}>
<div style={{display:"flex",gap:6,flexWrap:"wrap",justifyContent:"center"}}>
{(()=>{var fn=makeFileName("inv",inv.konsumen,inv.noInv,"");return <>
<button onClick={()=>doPrint("_inv",fn+".pdf")} style={{background:NAVY,color:WHITE,border:"none",padding:"9px 20px",borderRadius:7,fontSize:13,cursor:"pointer",fontWeight:700,fontFamily:FONT}}>🖨️ Cetak / PDF</button>
<button onClick={()=>doDownloadPNG("_inv",fn+".png")} style={{background:"#1D6A96",color:WHITE,border:"none",padding:"9px 16px",borderRadius:7,fontSize:13,cursor:"pointer",fontWeight:700,fontFamily:FONT}}>💾 Download PNG</button>
<button onClick={()=>doCopyPNG("_inv")} style={{background:"#145A32",color:WHITE,border:"none",padding:"9px 16px",borderRadius:7,fontSize:13,cursor:"pointer",fontWeight:700,fontFamily:FONT}}>📋 Copy PNG</button>
<button onClick={onClose} style={{background:"#566573",color:WHITE,border:"none",padding:"9px 16px",borderRadius:7,fontSize:13,cursor:"pointer",fontWeight:700,fontFamily:FONT}}>✕ Tutup</button>
</>;})()}
</div>
<div style={{fontSize:10,color:"#888",fontStyle:"italic"}}>💡 Nama file PDF: <b style={{color:"#aaa"}}>{makeFileName("inv",inv.konsumen,inv.noInv,"pdf")}</b></div>
</div>
</div>
</div>;
}

// ─── KWITANSI SLIP GAJI BARU v4 (no O/X column, PDF/PNG) ──────────────────────
function SlipGajiView({slip,company={},onClose}){
if(!slip)return null;
var rows=slip.rows||[];
function rR(n){return"Rp "+Number(n||0).toLocaleString("id-ID");}
var totPgh=rows.filter(r=>r.section==="penghasilan").reduce((a,r)=>a+Number(r.jumlah||0),0);
var totPot=rows.filter(r=>r.section==="potongan"&&r.kind!=="info").reduce((a,r)=>a+Number(r.jumlah||0),0);
var totYdt=rows.filter(r=>r.section==="ydt").reduce((a,r)=>a+Number(r.jumlah||0),0);
var totDiterima=totPgh-totPot-totYdt;
var sisa=Number(slip.totalPinjaman||0)-Number(slip.potonganPinjaman||0);
var blnIdx=slip.bulan?Number(slip.bulan.split("-")[1])-1:0;
var thIdx=slip.bulan?slip.bulan.split("-")[0]:"";
var fileName=safeFileName((slip.nama||"karyawan")+"_"+(slip.noSlip||"SLIP").replace(/[\/]/g,"-"));
var penghasilanRows=rows.filter(r=>r.section==="penghasilan");
var potonganRows=rows.filter(r=>r.section==="potongan");
var ydtRows=rows.filter(r=>r.section==="ydt");
var SNAVY="#0a1f44";var SNAVY2="#122d5e";var SRED="#e53935";var SGREEN="#6ab04c";var SWHITE="#fff";var SFONT="'Plus Jakarta Sans',-apple-system,'Segoe UI',sans-serif";
return <div id="_slip_wrap" style={{position:"fixed",inset:0,background:"#cdd3db",zIndex:9700,padding:16,overflowY:"auto",fontFamily:SFONT}}>
<div id="_slip" style={{maxWidth:680,margin:"0 auto",background:SWHITE,borderRadius:10,overflow:"hidden",boxShadow:"0 20px 60px rgba(10,31,68,.14)",color:"#0d1f3c",fontSize:13}}>
{/* HEADER — sama dengan Invoice */}
<div style={{background:"linear-gradient(135deg,"+SNAVY+" 0%,"+SNAVY2+" 100%)",padding:"18px 24px 15px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
<div style={{display:"flex",alignItems:"center",gap:12}}>
<CompanyLogo h={44} variant="dark"/>
</div>
{/* Pertamina diamond */}
<div style={{display:"flex",alignItems:"center",gap:6}}>
<PertaminaLogo h={28} variant="dark"/>
</div>
</div>
{/* DIVIDER 3 warna */}
<div style={{height:3,display:"flex"}}>
<div style={{flex:1,background:"#1565c0"}}/><div style={{flex:1,background:"#6ab04c"}}/><div style={{flex:1,background:"#e53935"}}/>
</div>
<div style={{display:"flex",justifyContent:"space-between",padding:"8px 18px",borderBottom:"1px solid #ccc",fontSize:12.5}}>
<div><span style={{color:"#444"}}>No.Pembayaran :</span> <b>{slip.noSlip}</b></div>
<div><span style={{color:"#444"}}>Tanggal :</span> <b>{fDs(slip.tanggal)}</b></div>
</div>
<h2 style={{textAlign:"center",fontSize:20,fontWeight:800,padding:"10px 0 6px",letterSpacing:1,margin:0}}>KWITANSI SLIP GAJI {(BULAN_ID[blnIdx]||"").toUpperCase()} {thIdx}</h2>
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",padding:"6px 18px 10px",gap:"4px 24px",fontSize:13.5}}>
<div><b style={{minWidth:64,display:"inline-block"}}>NAMA</b> : {slip.nama}</div>
<div><b style={{minWidth:64,display:"inline-block"}}>Alamat</b> : {slip.alamat||"-"}</div>
<div><b style={{minWidth:64,display:"inline-block"}}>Jabatan</b> : {slip.posisi}</div>
<div><b style={{minWidth:64,display:"inline-block"}}>Telepon</b> : {slip.telepon||"-"}</div>
</div>
<div style={{padding:"8px 18px",display:"flex",justifyContent:"space-between",borderTop:"1px solid #aaa",borderBottom:"1px solid #aaa",fontSize:13,background:"#FAF5E8"}}>
<span>Total Hari Kerja: <b>{slip.totalHariKerja}</b></span>
<span>Hadir: <b>{slip.hariHadir}</b></span>
<span style={{color:"#d63030"}}>Absen: <b>{slip.absen}</b></span>
</div>
<table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
<thead><tr>
<th style={{width:36,padding:"8px 4px",background:"#1a3a6b",color:"white",fontSize:12,textAlign:"center"}}>NO</th>
<th style={{padding:"8px 10px",background:"#1a3a6b",color:"white",fontSize:12,textAlign:"left"}}>KETERANGAN</th>
<th style={{width:200,padding:"8px 10px",background:"#1a3a6b",color:"white",fontSize:12,textAlign:"center"}}>@</th>
<th style={{width:140,padding:"8px 10px",background:"#1a3a6b",color:"white",fontSize:12,textAlign:"right"}}>JUMLAH</th>
</tr></thead>
<tbody>
<tr><td colSpan={4} style={{background:"#EDF3FB",padding:"7px 10px"}}><b style={{textDecoration:"underline"}}>Penghasilan</b></td></tr>
{penghasilanRows.map((r,i)=><tr key={"p"+i} style={{background:"#EDF3FB"}}>
<td style={{textAlign:"center",padding:"5px 4px",fontSize:11,color:"#666"}}>{i+1}</td>
<td style={{padding:"5px 10px",fontStyle:"italic"}}>{r.label}</td>
<td style={{padding:"5px 10px",fontSize:12,color:"#444",textAlign:"center"}}>{r.qty&&Number(r.qty)>1?r.qty+" × ":""}{r.ket||r.detail||""}</td>
<td style={{padding:"5px 10px",textAlign:"right",fontWeight:600}}>{Number(r.jumlah||0)>0?rR(r.jumlah):"Rp -"}</td>
</tr>)}
<tr style={{background:"#EDF3FB",borderTop:"1px solid #333"}}><td colSpan={3} style={{padding:"7px 10px",textAlign:"right",fontWeight:800}}>Total Penghasilan</td><td style={{padding:"7px 10px",textAlign:"right",fontWeight:800}}>{rR(totPgh)} <span style={{color:"#00A651",fontWeight:900}}>+</span></td></tr>
<tr><td colSpan={4} style={{height:6,background:"white"}}/></tr>
{potonganRows.map((r,i)=><tr key={"pt"+i} style={{background:"#FDF6E1",color:"#a32d2d"}}>
<td style={{textAlign:"center",padding:"5px 4px",fontSize:11,color:"#a32d2d"}}>{i+1}</td>
<td style={{padding:"5px 10px",fontStyle:"italic"}}>{r.label}</td>
<td style={{padding:"5px 10px",fontSize:12,fontStyle:"italic",textAlign:"center"}}>{r.qty&&Number(r.qty)>1?r.qty+" × ":""}{r.ket||r.detail||""}</td>
<td style={{padding:"5px 10px",textAlign:"right",fontWeight:700}}>{r.kind==="info"?rR(r.jumlah):(Number(r.jumlah||0)>0?rR(r.jumlah)+" -":"Rp -")}</td>
</tr>)}
{Number(slip.totalPinjaman||0)>0&&<tr style={{background:"#FDF6E1"}}>
<td/>
<td style={{padding:"5px 10px",fontStyle:"italic",textDecoration:"underline",color:"#a32d2d",fontWeight:700}}>Sisa Pinjaman</td>
<td style={{padding:"5px 10px",textAlign:"center"}}><span style={{background:"#FFF35B",color:"#1a1a1a",padding:"2px 10px",borderRadius:3,border:"1px solid #d4af00",fontWeight:800}}>{rR(sisa)}</span></td>
<td/>
</tr>}
<tr style={{background:"#FDF6E1"}}>
<td colSpan={2} style={{textAlign:"right",fontWeight:800,color:"#a32d2d",padding:"6px 10px"}}>TOTAL POTONGAN :</td>
<td/>
<td style={{padding:"6px 10px",textAlign:"right",color:"#a32d2d",fontWeight:800}}>{rR(totPot)}</td>
</tr>
{ydtRows.length>0&&<>
<tr><td colSpan={4} style={{background:"#F7F7F7",padding:"7px 10px"}}><b style={{textDecoration:"underline"}}>Yang sudah diterima</b></td></tr>
{ydtRows.map((r,i)=><tr key={"y"+i} style={{background:"#F7F7F7"}}>
<td style={{textAlign:"center",padding:"5px 4px",fontSize:11,color:"#666"}}>{i+1}</td>
<td style={{padding:"5px 10px",fontStyle:"italic"}}>{r.label}</td>
<td style={{padding:"5px 10px",fontSize:12,color:"#444",textAlign:"center"}}>{r.qty&&Number(r.qty)>1?r.qty+" × ":""}{r.ket||r.detail||""}</td>
<td style={{padding:"5px 10px",textAlign:"right",fontWeight:600}}>{Number(r.jumlah||0)>0?rR(r.jumlah)+" -":"Rp -"}</td>
</tr>)}
</>}
</tbody>
</table>
<div style={{padding:"14px 18px",textAlign:"right",fontSize:14,fontWeight:800,borderTop:"1px solid #aaa"}}>
<span style={{marginRight:12,letterSpacing:0.3}}>TOTAL DITERIMA :</span>
<span style={{background:"#FFF35B",padding:"5px 16px",borderRadius:4,border:"1px solid #d4af00",fontSize:16,fontWeight:900}}>{rR(totDiterima)}</span>
</div>
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",padding:"18px 36px 14px",gap:20,fontSize:13,textAlign:"center"}}>
<div><div style={{color:"#444"}}>Penerima</div><div style={{height:54,margin:"10px 0 2px",display:"flex",alignItems:"center",justifyContent:"center"}}>{slip.ttdPenerima?<img src={slip.ttdPenerima} style={{height:54}} alt="ttd"/>:<div style={{fontSize:10,color:"#bbb",border:"1px dashed #ddd",borderRadius:4,padding:"8px 12px"}}>TTD elektronik</div>}</div><div style={{fontWeight:800,borderTop:"1.5px solid #1a1a1a",paddingTop:3,display:"inline-block",minWidth:150,textDecoration:"underline"}}>{slip.nama}</div></div>
<div><div style={{fontWeight:800}}>{company.nama}</div><div style={{height:54,margin:"10px 0 2px",display:"flex",alignItems:"center",justifyContent:"center"}}>{company.ttdDirektur?<img src={company.ttdDirektur} style={{height:54}} alt="ttd"/>:<div style={{fontSize:10,color:"#bbb",border:"1px dashed #ddd",borderRadius:4,padding:"8px 12px"}}>TTD elektronik</div>}</div><div style={{fontWeight:800,borderTop:"1.5px solid #1a1a1a",paddingTop:3,display:"inline-block",minWidth:150,textDecoration:"underline"}}>{company.direkturNama||"Muhammad Haekal"}</div><div style={{fontSize:12,color:"#444",marginTop:2}}>Direktur</div></div>
</div>
{/* FOOTER — sama dengan Invoice */}
<div style={{background:SNAVY,padding:"12px 20px",display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10}}>
{[["📍","Alamat",[company.alamat]],["📞","No.HP & WA",[company.telepon,"TELP. "+company.telepon2]],["🌐","Email & Website",[company.email,company.website]]].map((x,i)=><div key={i} style={{display:"flex",alignItems:"flex-start",gap:9}}>
<div style={{width:26,height:26,background:"rgba(255,255,255,.1)",borderRadius:"50%",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0,fontSize:12}}>{x[0]}</div>
<div><div style={{fontSize:7,fontWeight:700,letterSpacing:.8,textTransform:"uppercase",color:"rgba(255,255,255,.4)",marginBottom:2}}>{x[1]}</div>{x[2].map((v,j)=><div key={j} style={{fontSize:9,fontWeight:600,color:SWHITE,lineHeight:1.5}}>{v}</div>)}</div>
</div>)}
</div>
</div>
<div style={{maxWidth:680,margin:"12px auto",display:"flex",gap:8,justifyContent:"center",flexWrap:"wrap"}}>
<div style={{display:"flex",flexDirection:"column",gap:6,alignItems:"center"}}>
<div style={{display:"flex",gap:6,flexWrap:"wrap",justifyContent:"center"}}>
{(()=>{var fn=makeFileName("slip",slip.nama,"Gaji-"+(slip.bulan||""),"");return <>
<button onClick={()=>doPrint("_slip",fn+".pdf")} style={{background:SNAVY,color:SWHITE,border:"none",padding:"9px 18px",borderRadius:7,fontSize:13,cursor:"pointer",fontWeight:700,fontFamily:SFONT}}>🖨️ Cetak / PDF</button>
<button onClick={()=>doDownloadPNG("_slip",fn+".png")} style={{background:"#1D6A96",color:SWHITE,border:"none",padding:"9px 14px",borderRadius:7,fontSize:13,cursor:"pointer",fontWeight:700,fontFamily:SFONT}}>💾 Download PNG</button>
<button onClick={()=>doCopyPNG("_slip")} style={{background:"#145A32",color:SWHITE,border:"none",padding:"9px 14px",borderRadius:7,fontSize:13,cursor:"pointer",fontWeight:700,fontFamily:SFONT}}>📋 Copy PNG</button>
<button onClick={onClose} style={{background:"#566573",color:SWHITE,border:"none",padding:"9px 14px",borderRadius:7,fontSize:13,cursor:"pointer",fontWeight:700,fontFamily:SFONT}}>✕ Tutup</button>
</>;})()}
</div>
<div style={{fontSize:10,color:"#888",fontStyle:"italic"}}>💡 Nama file PDF: <b style={{color:"#aaa"}}>{makeFileName("slip",slip.nama,"Gaji-"+(slip.bulan||""),"pdf")}</b></div>
</div>
</div>
</div>;
}

// ─── BERITA ACARA ─────────────────────────────────────────────────────────────
function BeritaAcaraView({ba,company={},onClose}){
if(!ba)return null;
var isTitip=ba.tipe==="titip";var totalQty=(ba.items||[]).reduce((a,it)=>a+Number(it.qty||0),0);
var fileName=safeFileName((ba.konsumenNama||"konsumen")+"_"+(ba.noBA||"BA"));
return <div id="_ba_wrap" style={{position:"fixed",inset:0,background:"#e5e8ec",zIndex:9600,padding:16,overflowY:"auto",fontFamily:"'Times New Roman',serif"}}>
<div id="_ba" style={{maxWidth:680,margin:"0 auto",background:"white",borderRadius:6,padding:"24px 28px",color:"#1a1a1a",boxShadow:"0 6px 30px rgba(0,0,0,.15)",border:"2px solid #1A5276"}}>
<div style={{textAlign:"center",borderBottom:"3px double #1A5276",paddingBottom:14,marginBottom:16}}>
<div style={{display:"flex",justifyContent:"center",marginBottom:6}}><CompanyLogo h={48} variant="light"/></div>
<div style={{fontFamily:"Arial",fontWeight:900,fontSize:20,color:"#0A2C5C"}}>{company.nama}</div>
<div style={{fontFamily:"Arial",fontSize:11,color:"#444"}}>{company.alamat}</div>
<div style={{fontFamily:"Arial",fontSize:11,color:"#444"}}>Telp: {company.telepon} | {company.email}</div>
</div>
<div style={{textAlign:"center",marginBottom:16}}><div style={{fontSize:16,fontWeight:700,fontFamily:"Arial",textTransform:"uppercase",color:"#1A5276"}}>BERITA ACARA {isTitip?"PENITIPAN":"PENARIKAN"} TABUNG LPG</div><div style={{fontFamily:"Arial",fontSize:12}}>No. {ba.noBA}</div></div>
<div style={{marginBottom:14,fontFamily:"Arial",fontSize:13,lineHeight:2}}>Yang bertanda tangan di bawah ini menyatakan bahwa pada <b>{fD(ba.tanggal)}</b>, telah terjadi <b style={{color:"#1A5276"}}>{isTitip?"PENITIPAN":"PENARIKAN"} TABUNG LPG</b>:</div>
<div style={{background:"#f8f9ff",border:"1px solid #BDC3C7",borderRadius:6,padding:"12px 16px",marginBottom:14,fontFamily:"Arial",fontSize:13}}>
<div style={{fontWeight:700,color:"#1A5276",marginBottom:6}}>{isTitip?"Penerima Titipan":"Penyerah Tabung"}</div>
<div style={{display:"grid",gridTemplateColumns:"130px 1fr",gap:"4px 8px",lineHeight:1.9}}>
<span>Nama</span><span>: <b>{ba.konsumenNama}</b></span>
<span>No. HP</span><span>: {ba.konsumenTelp||"-"}</span>
<span>Alamat</span><span>: {ba.konsumenAlamat||"-"}</span>
<span>{isTitip?"Diantar oleh":"Ditarik oleh"}</span><span>: <b>{ba.salesNama||"-"}</b></span>
</div>
</div>
<table style={{width:"100%",borderCollapse:"collapse",fontFamily:"Arial",fontSize:13,marginBottom:14}}>
<thead><tr>{["No.","Ukuran","Jumlah"].map((h,i)=><th key={i} style={{background:"#1A5276",color:"#fff",padding:"7px 10px",textAlign:i===2?"center":"left"}}>{h}</th>)}</tr></thead>
<tbody>{(ba.items||[]).map((it,i)=><tr key={i} style={{borderBottom:"1px solid #ddd"}}><td style={{padding:"6px 10px"}}>{i+1}</td><td style={{padding:"6px 10px"}}>Gas LPG {it.ukuran}</td><td style={{padding:"6px 10px",textAlign:"center",fontWeight:700}}>{it.qty}</td></tr>)}<tr style={{background:"#EAF2FF",fontWeight:700}}><td colSpan={2} style={{padding:"8px 10px"}}>TOTAL</td><td style={{padding:"8px 10px",textAlign:"center",fontSize:14}}>{totalQty} Tabung</td></tr></tbody>
</table>
{ba.ket&&<div style={{marginBottom:14,fontFamily:"Arial",fontSize:12,padding:"8px 12px",background:"#FEFDE7",border:"1px solid #F1C40F"}}><b>Ket:</b> {ba.ket}</div>}
<div style={{fontFamily:"Arial",fontSize:12,marginBottom:20}}>Demikian berita acara ini dibuat. <b style={{color:"#1A5276"}}>Tabung merupakan aset milik {company.nama}</b> dan wajib dikembalikan bila tidak lagi berlangganan.</div>
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:20,fontFamily:"Arial",fontSize:12}}>
{[["Yang Menyerahkan",ba.salesNama,company.ttdKasir],["Yang Menerima",ba.konsumenNama,null]].map((x,i)=><div key={i} style={{textAlign:"center"}}><div style={{fontWeight:600}}>{x[0]}</div><div style={{height:52,display:"flex",alignItems:"center",justifyContent:"center"}}>{x[2]?<img src={x[2]} style={{height:50}} alt="ttd"/>:null}</div><div style={{borderBottom:"1px solid #333",margin:"0 20px"}}/><div style={{marginTop:4,fontWeight:700}}>{x[1]||"____________"}</div></div>)}
</div>
</div>
<div style={{maxWidth:680,margin:"14px auto",display:"flex",gap:8,justifyContent:"center",flexWrap:"wrap"}}>
<button onClick={()=>doPrint("_ba")} style={{background:"#1A5276",color:"#fff",border:"none",padding:"10px 28px",borderRadius:8,fontSize:13,cursor:"pointer",fontWeight:700}}>🖨️ Cetak / Simpan PDF</button>
<div style={{fontSize:11,color:"#666",alignSelf:"center",fontStyle:"italic"}}>Pilih "Save as PDF" di dialog print</div>
<button onClick={onClose} style={{background:"#566573",color:"#fff",border:"none",padding:"10px 22px",borderRadius:8,fontSize:13,cursor:"pointer",fontWeight:700}}>✕ Tutup</button>
</div>
</div>;
}

// ─── LOGIN ────────────────────────────────────────────────────────────────────
function LoginScreen({employees,onLogin,themeToggle,theme}){
var C=useTheme();
var[u,setU]=useState("");var[p,setP]=useState("");var[err,setErr]=useState("");
function login(){var emp=(employees||DEF_EMP).find(x=>x.username===u&&x.password===p&&x.aktif);if(emp)onLogin(emp);else setErr("Username atau password salah!");}
return <div style={{minHeight:"100vh",background:C.bg,display:"flex",alignItems:"center",justifyContent:"center",padding:16,position:"relative"}}>
<div style={{position:"absolute",top:16,right:16}}><Btn sm color="gray" onClick={themeToggle}>{theme==="light"?"🌙 Gelap":"☀️ Terang"}</Btn></div>
<div style={{background:C.card,borderRadius:20,padding:"36px 32px 28px",width:"100%",maxWidth:380,border:"1px solid "+C.bdr,boxShadow:"0 20px 80px rgba(0,0,0,.15)"}}>
<div style={{textAlign:"center",marginBottom:26}}><div style={{display:"flex",justifyContent:"center",marginBottom:14}}><CompanyLogo h={56} variant={C.mode==="dark"?"dark":"light"}/></div><div style={{fontSize:13,fontWeight:700,color:C.wht}}>Sistem Manajemen LPG</div><div style={{fontSize:11,color:C.gl2,marginTop:3}}>Distributor Resmi Pertamina</div></div>
<Inp label="Username" value={u} onChange={setU} placeholder="Masukkan username"/>
<Inp label="Password" type="password" value={p} onChange={setP} placeholder="Masukkan password"/>
{err&&<div style={{color:C.rlt,fontSize:12,marginBottom:10,textAlign:"center",padding:"6px 10px",background:C.mode==="dark"?"#3B0A0A":"#FEE2E2",borderRadius:6}}>{err}</div>}
<Btn onClick={login} style={{width:"100%",padding:"12px",fontSize:14,marginBottom:12}}>🔐 Masuk</Btn>
<div style={{background:C.nav,borderRadius:8,padding:10,fontSize:11,color:C.gl2,border:"1px solid "+C.bdr}}><b style={{color:C.wht}}>Contoh:</b> Haekal | DAYAT | MUSLEM | SAIBAN<br/>Password: Sudirman80</div>
</div>
</div>;
}

// ─── DASHBOARD ────────────────────────────────────────────────────────────────
function Dashboard({data,setTab,user}){
var C=useTheme();var td=toDay();
var pHari=(data.penjualan||[]).filter(e=>e.tanggal===td);
var cashIn=pHari.filter(e=>e.bayar==="cash").reduce((a,e)=>a+(e.total||0),0);
var tfIn=pHari.filter(e=>e.bayar==="transfer").reduce((a,e)=>a+(e.total||0),0);
var bonIn=pHari.filter(e=>e.bayar==="bon").reduce((a,e)=>a+(e.total||0),0);
var totalIn=cashIn+tfIn+bonIn;
var margin=pHari.reduce((a,e)=>a+(e.margin||0),0);
var penHari=(data.pengeluaran||[]).filter(e=>e.tanggal===td);
var totalOut=penHari.reduce((a,e)=>a+Number(e.nominal||0),0);
var labaBersih=margin-totalOut;
var bonAktif=(data.bon||[]).filter(b=>b.status!=="lunas");
var piutang=bonAktif.reduce((a,b)=>a+b.sisaTagihan,0);
var alerts=bonAktif.filter(b=>b.deadline&&dLeft(b.deadline)!==null&&dLeft(b.deadline)<=3);
var mob=useMobile();
return <div>
<div style={{marginBottom:14}}><div style={{fontSize:20,fontWeight:900,color:C.wht}}>Halo, {user?.nama} 👋</div><div style={{fontSize:12,color:C.gl2,marginTop:2}}>{ROLE_LBL[user?.role]||""} — {fDs(td)}</div></div>
{alerts.length>0&&<div style={{background:C.mode==="dark"?"#3D1A05":"#FFEDD5",border:"1px solid "+C.org,borderRadius:12,padding:"12px 16px",marginBottom:14}}><b style={{color:C.olt,fontSize:13}}>⚠️ {alerts.length} Bon Jatuh Tempo!</b>{alerts.map(b=>{var d=dLeft(b.deadline);return <div key={b.id} style={{fontSize:12,color:C.gltr,marginTop:4}}><b style={{color:C.wht}}>{b.konsumen}</b> — {fR(b.sisaTagihan)} — {d<=0?<Bdg color="red">LEWAT</Bdg>:<Bdg color="orange">{d}h lagi</Bdg>}</div>;})}</div>}
<Card style={{width:"fit-content",maxWidth:"100%",minWidth:660}}>
<div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>📊 P&L Hari Ini — {fDs(td)}</div>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(150px,180px))",gap:8}}>
{[["Omzet",totalIn,C.wht,"📈"],["Laba Kotor",margin,C.blt,"💹"],["Pengeluaran",totalOut,C.rlt,"💸"],["Laba Bersih",labaBersih,labaBersih>=0?C.glt:C.rlt,"🏆"]].map(x=><div key={x[0]} style={{background:C.nav,borderRadius:8,padding:"9px 12px",border:"1px solid "+C.bdr}}><div style={{fontSize:10,color:C.gl2,marginBottom:2}}>{x[3]} {x[0]}</div><div style={{fontSize:14,fontWeight:900,color:x[2],whiteSpace:"nowrap"}}>{fR(x[1])}</div></div>)}
</div>
</Card>
<div style={{display:"flex",gap:10,flexWrap:"wrap",marginBottom:14}}>
<SC label="Piutang Aktif" value={fR(piutang)} icon="💳" color={C.olt}/>
<SC label="Bon Aktif" value={bonAktif.length+" bon"} icon="📃" color={C.gl2}/>
</div>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(190px,1fr))",gap:12,marginBottom:14}}>
{(()=>{
var rowsLast=buildStokHarian(data,toMonth()).filter(r=>r.tgl<=td);
var lastRow=rowsLast.length>0?rowsLast[rowsLast.length-1]:null;
return <>{SIZES.map(s=>{
var isi=lastRow?lastRow.akhirIsi[s]:((data.stock||{})[s]||0);
var kosong=lastRow?lastRow.akhirTK[s]:getKosong(data,s);
var titip=lastRow?lastRow.titipSnap[s]:getTitipTotal(data.titipList,s);
var totalS=isi+kosong+titip;
return <Card key={s} style={{marginBottom:0}}>
<div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}>
<div style={{fontWeight:700,color:C.gl2,fontSize:12}}>📦 LPG {s}</div>
<div style={{fontSize:14,fontWeight:900,color:C.olt}}>{totalS} tab</div>
</div>
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:4}}>
{[["Tbg+Isi",isi,C.glt],["Titip",titip,C.blt],["Kosong",kosong,C.gl2]].map(x=><div key={x[0]} style={{background:C.nav,borderRadius:6,padding:"5px 4px",textAlign:"center",border:"1px solid "+C.bdr}}><div style={{fontSize:8,color:C.gl2}}>{x[0]}</div><div style={{fontSize:15,fontWeight:900,color:x[2]}}>{x[1]}</div></div>)}
</div>
</Card>;
})}</>;
})()}
</div>
<Card style={{marginBottom:14}}>
<div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
<div style={{fontWeight:700,color:C.gl2,fontSize:13}}>📊 Laporan Stok 7 Hari Terakhir</div>
</div>
{(()=>{
  var bulanIni=toMonth();
  var rows7=buildStokHarian(data,bulanIni).filter(r=>r.tgl<=td).slice(-7);
  var uk=["5.5 kg","12 kg","50 kg"];var ukL=["5,5kg","12kg","50kg"];
  var C2=C;
  return <div style={{overflowX:"auto"}}>
  <table style={{borderCollapse:"collapse",width:"100%",fontSize:10}}>
  <thead><tr style={{background:C2.nav}}>
    <th style={{padding:"5px 8px",color:C2.gl2,textAlign:"left",borderBottom:"1px solid "+C2.bdr,whiteSpace:"nowrap"}}>Hari/Tgl</th>
    {uk.map((s,i)=><th key={s} colSpan={4} style={{padding:"5px 6px",color:"white",textAlign:"center",borderBottom:"1px solid "+C2.bdr,borderLeft:"2px solid "+C2.bdr,background:"#1E3A5F",fontSize:9}}>{ukL[i]}: isi | TK | Titip | Total</th>)}
  </tr></thead>
  <tbody>
  {rows7.map((r,i)=><tr key={r.tgl} style={{background:i%2===0?C2.nav:C2.bg,borderBottom:"1px solid "+C2.bdr}}>
    <td style={{padding:"4px 8px",whiteSpace:"nowrap"}}>
      <div style={{fontWeight:700,color:r.tgl===td?C2.blt:C2.wht,fontSize:11}}>{r.dayName}</div>
      <div style={{fontSize:9,color:C2.gl2}}>{fDs(r.tgl)}</div>
    </td>
    {uk.map(s=>[
      <td key={"i"+s} style={{padding:"4px 6px",textAlign:"center",color:C2.glt,fontWeight:700,borderLeft:"2px solid "+C2.bdr}}>{r.akhirIsi[s]}</td>,
      <td key={"k"+s} style={{padding:"4px 6px",textAlign:"center",color:C2.gl2}}>{r.akhirTK[s]}</td>,
      <td key={"t"+s} style={{padding:"4px 6px",textAlign:"center",color:C2.blt}}>{r.titipSnap[s]}</td>,
      <td key={"o"+s} style={{padding:"4px 6px",textAlign:"center",color:C2.olt,fontWeight:700}}>{r.total[s]}</td>
    ])}
  </tr>)}
  </tbody>
  </table>
  </div>;
})()}
</Card>
<Card><div style={{fontWeight:700,color:C.gl2,fontSize:12,marginBottom:10}}>⚡ Akses Cepat</div><div style={{display:"flex",gap:8,flexWrap:"wrap"}}>{[["penjualan","🧾","Penjualan"],["stok","📦","Stok"],["piutang","💳","Piutang"],["absensi","📅","Absensi"],["absensi","💼","Payroll"],["laporan","📊","Laporan"]].map(x=><Btn key={x[0]} onClick={()=>setTab(x[0])} sm>{x[1]} {x[2]}</Btn>)}</div></Card>
</div>;
}

// ─── PENJUALAN v4 (FilterTbl, cetak invoice langsung dari riwayat) ────────────
function PenjualanMod({data,setData,setInv,user,toast}){
var C=useTheme();var mob=useMobile();
var blk={ukuran:"5.5 kg",jenis:"Isi",qty:"",price:""};
var canSelf=PENJUALAN_ROLES.includes(user?.role)&&!["owner","admin","akuntan"].includes(user?.role);
var[f,setF]=useState({tanggal:toDay(),salesId:canSelf?user.id:"",konsumen:"",konsumenId:"",items:[{...blk}],bayar:"cash",bank:"BSI",deadline:"",ket:"",splitDetail:{cash:0,tf:0,bon:0},splitBank:"BSI"});
var[delId,setDelId]=useState(null);
var[editInv,setEditInv]=useState(null);// {entry, form}
var[barFilter,setBarFilter]=useState({from:"",to:"",salesId:"",konsumen:"",bayar:""});
var[tglLap,setTglLap]=useState(toDay());
var salesEmp=sortEmp((data.employees||[]).filter(e=>e.aktif&&PENJUALAN_ROLES.includes(e.role)));
var valid=f.items.filter(it=>Number(it.qty)>0&&Number(it.price)>0);
var total=iTotal(valid);var marginEstimasi=calcMargin(valid,data,f.tanggal);
var kNames=[...new Set([...(data.pelanggan||[]).map(p=>p.nama),...(data.penjualan||[]).map(e=>e.konsumen)].filter(Boolean))];
function onKons(nama){var p=(data.pelanggan||[]).find(x=>x.nama===nama);if(p){setF(pv=>{var newItems=pv.items.map(it=>{var h=(Array.isArray(p.hargaKhusus)?p.hargaKhusus:[]).find(x=>x.ukuran===it.ukuran&&x.jenis===it.jenis);if(h)return{...it,price:String(h.harga)};var het=getHET(data,it.ukuran,it.jenis);return{...it,price:het?String(het):it.price};});return{...pv,konsumen:nama,konsumenId:p.id,items:newItems};});}else setF(pv=>({...pv,konsumen:nama,konsumenId:""}));}
function setProduct(i,ukuran,jenis){setF(p=>{var it=p.items.slice();var newIt={...it[i],ukuran,jenis};var plg=(data.pelanggan||[]).find(x=>x.id===p.konsumenId);if(plg){var h=(plg.hargaKhusus||[]).find(x=>x.ukuran===ukuran&&x.jenis===jenis);if(h){newIt.price=String(h.harga);}else newIt.price=String(getHET(data,ukuran,jenis)||"");}else newIt.price=String(getHET(data,ukuran,jenis)||"");it[i]=newIt;return{...p,items:it};});}
function setItem(i,k,v){setF(p=>{var it=p.items.slice();it[i]={...it[i],[k]:v};return{...p,items:it};});}
function makeInvObj(entry){
var emp=(data.employees||[]).find(e=>e.id===entry.salesId);
var plg=(data.pelanggan||[]).find(x=>x.id===entry.konsumenId);
var sd=entry.splitDetail||{};
var metodeBayar=entry.bayar==="bon"?"BON":entry.bayar==="transfer"?"Transfer "+(entry.bank||""):entry.bayar==="split"?"Split":"Cash";
var splitLabel=entry.bayar==="split"?[Number(sd.cash)>0?"Cash":null,Number(sd.tf)>0?"TF":null,Number(sd.bon)>0?"BON":null].filter(Boolean).join("+"):"";
return{noInv:entry.noInv,tanggal:entry.tanggal,konsumen:entry.konsumen,kota:plg?.alamat?.split(",").pop()?.trim()||"Banda Aceh",salesNama:emp?.nama||"",items:(entry.items||[]).map(it=>({ukuran:it.ukuran,jenis:it.jenis,qty:Number(it.qty),price:Number(it.price)})),total:entry.total,metodeBayar:entry.bayar==="split"?splitLabel:metodeBayar,isBon:entry.bayar==="bon"||(entry.bayar==="split"&&Number(sd.bon)>0&&Number(sd.cash)===0&&Number(sd.tf)===0),splitDetail:entry.splitDetail,splitBank:entry.splitBank||"",catatan:entry.ket||""};
}
function doSave(withPrint){
if(!valid.length||!f.konsumen)return;
var ns={...data.stock};
var nk={...(data.stokKosong||{})};
var na={...(data.totalTabung||{})};
var stokLogs=[];
valid.forEach(it=>{
  var q=Number(it.qty||0);var uk=it.ukuran;
  ns[uk]=Math.max(0,(ns[uk]||0)-q);
  if(it.jenis==="Tabung+Isi"){
    nk[uk]=Math.max(0,(nk[uk]||0)-q);
    na[uk]=Math.max(0,(na[uk]||0)-q);// totalTabung berkurang karena tabung terjual
    stokLogs.push({id:uid(),tanggal:f.tanggal,ukuran:uk,jenis:"Tbg+Isi Keluar",qty:q,ket:"Inv - "+f.konsumen,sumber:"Penjualan"});
  } else {
    nk[uk]=(nk[uk]||0)+q;
    stokLogs.push({id:uid(),tanggal:f.tanggal,ukuran:uk,jenis:"Isi Keluar",qty:q,ket:"Inv - "+f.konsumen,sumber:"Penjualan"});
  }
});
var invInfo=nextInvNo(data,f.tanggal);
var newCounters={...(data.counters||{inv:{},sg:{},reg:0})};if(!newCounters.inv)newCounters.inv={};newCounters.inv[invInfo.key]=invInfo.n;
var isSplit=f.bayar==="split";
var sd=f.splitDetail||{cash:0,tf:0,bon:0};
setData(d=>{
  // ── Konsumsi FIFO per item, hitung margin riil berdasarkan batch yang dimakan ──
  var dCur=d;var marginFIFO=0;var fifoDetailAll=[];
  var itemsFinal=valid.map(it=>{
    var q=Number(it.qty||0);
    var res=consumeFIFO(dCur,it.ukuran,q);
    dCur=res.data;
    var marginItem=(Number(it.price||0)*q)-res.hppTotal;
    marginFIFO+=marginItem;
    fifoDetailAll.push({ukuran:it.ukuran,jenis:it.jenis,qty:q,hppTotal:res.hppTotal,hppPerUnit:res.hppPerUnit,detail:res.detail,qtyKurang:res.qtyKurang});
    return{...it,qty:q,price:Number(it.price),hppFIFO:res.hppPerUnit};
  });
  var entry={id:uid(),noInv:invInfo.no,tanggal:f.tanggal,waktu:new Date().toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit"}),salesId:f.salesId,konsumen:f.konsumen,konsumenId:f.konsumenId,items:itemsFinal,total,margin:marginFIFO,fifoDetail:fifoDetailAll,bayar:f.bayar,bank:isSplit?(Number(sd.tf)>0?f.splitBank:""):f.bank,deadline:f.deadline,ket:f.ket,splitDetail:isSplit?sd:null,splitBank:isSplit?f.splitBank:""};
  var nb=(d.bon||[]).slice();
  if(f.bayar==="bon")nb.unshift({id:uid(),noInv:invInfo.no,tanggal:f.tanggal,konsumen:f.konsumen,konsumenId:f.konsumenId,salesId:f.salesId,items:valid,total,sisaTagihan:total,deadline:f.deadline,status:"belum",pembayaran:[],ket:f.ket,bank:f.bank});
  if(isSplit&&Number(sd.bon)>0)nb.unshift({id:uid(),noInv:invInfo.no+"(BON)",tanggal:f.tanggal,konsumen:f.konsumen,konsumenId:f.konsumenId,salesId:f.salesId,items:valid,total:Number(sd.bon),sisaTagihan:Number(sd.bon),deadline:f.deadline,status:"belum",pembayaran:[],ket:"Split payment — BON portion. "+f.ket,bank:""});
  if(withPrint)setTimeout(()=>setInv(makeInvObj(entry)),0);
  return{...dCur,penjualan:[entry,...(d.penjualan||[])],stock:ns,stokKosong:nk,totalTabung:na,bon:nb,counters:newCounters,stockLog:[...stokLogs,...(d.stockLog||[])].slice(0,500)};
});
setF(p=>({...p,konsumen:"",konsumenId:"",items:[{...blk}],ket:"",deadline:"",splitDetail:{cash:0,tf:0,bon:0}}));
toast("✓ Tersimpan! No: "+invInfo.no);
}
// Riwayat: flatten transactions, apply bar filter
var rows=useMemo(()=>{
return(data.penjualan||[]).filter(p=>{
if(barFilter.from&&p.tanggal<barFilter.from)return false;
if(barFilter.to&&p.tanggal>barFilter.to)return false;
if(barFilter.salesId&&p.salesId!==barFilter.salesId)return false;
if(barFilter.konsumen&&!p.konsumen.toLowerCase().includes(barFilter.konsumen.toLowerCase()))return false;
if(barFilter.bayar&&p.bayar!==barFilter.bayar)return false;
return true;
}).map(p=>{var emp=(data.employees||[]).find(e=>e.id===p.salesId);var detailStr=(p.items||[]).map(it=>it.qty+"×"+it.ukuran).join(", ");return{...p,salesNama:emp?.nama||"-",detailStr};});
},[data.penjualan,data.employees,barFilter]);
var SZ_CLR={"5.5 kg":[C.grn,"#0A2E14"],"12 kg":[C.blt,"#0A2040"],"50 kg":[C.olt,"#3D200A"]};
var cols=[
{key:"tanggal",label:"Tgl",width:84,render:r=><div><div style={{fontSize:12.5,fontWeight:600,color:C.wht,whiteSpace:"nowrap"}}>{fDs(r.tanggal)}</div><div style={{fontSize:10.5,color:C.gl2,marginTop:2}}>{r.waktu||""}</div></div>,sortVal:r=>r.tanggal,filterable:true},
{key:"noInv",label:"No.Invoice",width:115,render:r=><span style={{fontSize:12,color:C.blt,fontWeight:700,whiteSpace:"nowrap"}}>{r.noInv||"-"}</span>,filterable:true},
{key:"konsumen",label:"Konsumen",render:r=><b style={{color:C.wht}}>{r.konsumen}</b>,filterable:true,width:160},
{key:"salesNama",label:"Sales",render:r=><span style={{fontSize:12}}>{r.salesNama}</span>,filterable:true,width:90},
{key:"detailStr",label:"Produk & Qty",width:150,render:r=>{var items=r.items||[];var show=items.slice(0,3);var more=items.length-3;return <div style={{display:"flex",flexDirection:"column",gap:3}}>{show.map((it,i)=>{var clr=SZ_CLR[it.ukuran]||[C.gl2,C.bg];return <div key={i} style={{display:"flex",alignItems:"center",gap:4,paddingBottom:i<show.length-1?2:0,borderBottom:i<show.length-1?"1px dashed "+C.bdr:"none"}}><span style={{background:clr[1],border:"1px solid "+clr[0],borderRadius:3,padding:"1px 5px",fontSize:9.5,fontWeight:700,color:clr[0],whiteSpace:"nowrap",flexShrink:0,minWidth:42,textAlign:"center"}}>{it.ukuran}</span><span style={{fontSize:11,color:C.gl2,whiteSpace:"nowrap",flexShrink:0}}>{it.jenis==="Tabung+Isi"?"Tbg+Isi":"Refill"}</span><b style={{fontSize:12.5,fontWeight:800,color:C.wht,whiteSpace:"nowrap",marginLeft:"auto"}}>{it.qty}</b></div>;})}{ more>0&&<div style={{fontSize:10,color:C.gry,fontStyle:"italic"}}>+{more} item lagi</div>}</div>;},filterable:true,sortable:false},
{key:"total",label:"Total",render:r=><b style={{color:C.wht,whiteSpace:"nowrap",fontSize:13.5}}>{fR(r.total)}</b>,filterable:false,sortVal:r=>r.total,width:125},
{key:"bayar",label:"Bayar",render:r=>{var sd=r.splitDetail||{};if(r.bayar==="split"){var lbl=[Number(sd.cash)>0?"Cash":null,Number(sd.tf)>0?"TF":null,Number(sd.bon)>0?"BON":null].filter(Boolean).join("+");return <Bdg color="orange">{lbl}</Bdg>;}return r.bayar==="bon"?<Bdg color="red">BON</Bdg>:r.bayar==="transfer"?<Bdg color="blue">TF</Bdg>:<Bdg color="green">Cash</Bdg>;},filterable:true,filterType:"select",options:[{v:"cash",l:"Cash"},{v:"transfer",l:"Transfer"},{v:"bon",l:"BON"},{v:"split",l:"Split"}],width:78},
{key:"_aksi",label:"Aksi",sortable:false,filterable:false,width:115,render:r=><div style={{display:"flex",gap:5}}><button onClick={()=>setInv(makeInvObj(r))} title="Cetak Invoice" style={{background:C.inHv,border:"1px solid "+C.blt,borderRadius:7,padding:"6px 9px",color:C.blt,cursor:"pointer",fontSize:13}}>🖨️</button><button onClick={()=>{var ef={...r,items:(r.items||[]).map(it=>({...it,qty:String(it.qty),price:String(it.price)}))};setEditInv({entry:r,form:ef});}} title="Edit Invoice" style={{background:"#78350F",border:"1px solid #F59E0B",borderRadius:7,padding:"6px 9px",color:"#FCD34D",cursor:"pointer",fontSize:13}}>✏️</button><button onClick={()=>setDelId(r)} style={{background:C.inHvE,border:"1px solid "+C.rlt,borderRadius:7,padding:"6px 9px",color:C.rlt,cursor:"pointer",fontSize:13}}>🗑️</button></div>},
];
return <div>
<STitle icon="🧾" children="Input Penjualan"/>
<Card>
<div style={{display:"grid",gridTemplateColumns:mob?"1fr 1fr":"repeat(auto-fill,minmax(155px,210px))",gap:10,width:mob?"100%":"fit-content",maxWidth:"100%"}}>
<Inp label="Tanggal" type="date" value={f.tanggal} onChange={v=>setF(p=>({...p,tanggal:v}))}/>
<Sel label="Penjual" value={f.salesId} onChange={v=>setF(p=>({...p,salesId:v}))} opts={[{v:"",l:"-- Pilih --"},...salesEmp.map(e=>({v:e.id,l:e.nama+" ("+e.posisi+")"}))]} style={{gridColumn:mob?"1/-1":"auto"}}/>
<div style={{gridColumn:mob?"1/-1":"auto"}}><AutoInp label="Nama Konsumen" value={f.konsumen} onChange={v=>setF(p=>({...p,konsumen:v}))} options={kNames} placeholder="Ketik nama..." onSelect={onKons}/></div>
</div>
{f.konsumenId&&(()=>{var plg=(data.pelanggan||[]).find(x=>x.id===f.konsumenId);return plg&&(Array.isArray(plg.hargaKhusus)?plg.hargaKhusus:[]).length>0?<div style={{padding:"6px 10px",background:C.mode==="dark"?"#0A1A2A":"#DBEAFE",borderRadius:6,fontSize:11,color:C.blt,marginBottom:10}}>💲 Pelanggan memiliki <b>{plg.hargaKhusus.length}</b> harga khusus — otomatis terpasang</div>:null;})()}
<div style={{border:"1px solid "+C.bdr,borderRadius:8,overflow:"hidden",marginBottom:10,width:mob?"100%":"fit-content",maxWidth:"100%"}}>
<div style={{display:"grid",gridTemplateColumns:mob?"85px 100px 55px 95px 28px":"95px 110px 60px 110px 90px 28px",background:C.nav,padding:"6px 10px",fontSize:11,color:C.gl2,fontWeight:700,gap:5}}><span>Ukuran</span><span>Jenis</span><span>Qty</span><span>Harga</span>{!mob&&<span>Subtotal</span>}<span/></div>
{f.items.map((it,i)=><div key={i} style={{display:"grid",gridTemplateColumns:mob?"85px 100px 55px 95px 28px":"95px 110px 60px 110px 90px 28px",padding:"5px 10px",borderTop:"1px solid "+C.bdr,alignItems:"center",gap:5}}>
<select value={it.ukuran} onChange={e=>setProduct(i,e.target.value,it.jenis)} style={{background:C.nav,border:"1px solid "+C.bdr,borderRadius:6,padding:"5px 4px",color:C.wht,fontSize:12,outline:"none"}}>{SIZES.map(s=><option key={s}>{s}</option>)}</select>
<select value={it.jenis} onChange={e=>setProduct(i,it.ukuran,e.target.value)} style={{background:C.nav,border:"1px solid "+(it.jenis==="Tabung+Isi"?C.glt:C.bdr),borderRadius:6,padding:"5px 4px",color:C.wht,fontSize:12,outline:"none"}}>{JENIS.map(j=><option key={j}>{j}</option>)}</select>
<input type="number" value={it.qty} placeholder="0" onChange={e=>setItem(i,"qty",e.target.value)} style={{background:C.nav,border:"1px solid "+C.bdr,borderRadius:6,padding:"5px 4px",color:C.wht,fontSize:12,outline:"none",width:"100%"}}/>
<input type="number" value={it.price} step="1000" onChange={e=>setItem(i,"price",e.target.value)} style={{background:C.nav,border:"1px solid "+C.bdr,borderRadius:6,padding:"5px 4px",color:C.wht,fontSize:12,outline:"none",width:"100%"}}/>
{!mob&&<span style={{color:C.glt,fontWeight:700,fontSize:12}}>{it.qty&&it.price?fR(Number(it.qty)*Number(it.price)):"-"}</span>}
<button onClick={()=>setF(p=>({...p,items:p.items.filter((_,j)=>j!==i)}))} disabled={f.items.length<=1} style={{background:C.inHvE,border:"none",borderRadius:5,color:C.rlt,cursor:"pointer",fontSize:13,padding:"2px 5px",opacity:f.items.length<=1?0.3:1}}>✕</button>
</div>)}
<div style={{padding:"7px 10px",background:C.nav,borderTop:"1px solid "+C.bdr,display:"flex",justifyContent:"space-between",flexWrap:"wrap",gap:8}}>
<Btn sm color="blue" onClick={()=>setF(p=>({...p,items:[...p.items,{...blk}]}))}>+ Item</Btn>
<span style={{fontSize:13,color:C.gl2}}>Margin (estimasi): <b style={{color:C.glt}}>{fR(marginEstimasi)}</b> | Total: <b style={{color:C.wht,fontSize:14}}>{fR(total)}</b></span>
</div>
</div>
<div style={{display:"flex",gap:8,marginBottom:10,flexWrap:"wrap"}}>
{[["cash","💵 Cash",C.grn],["transfer","🏦 Transfer",C.blu],["bon","📃 BON",C.rdk],["split","✂️ Split",C.olt]].map(x=><button key={x[0]} onClick={()=>setF(p=>({...p,bayar:x[0]}))} style={{background:f.bayar===x[0]?x[2]:C.nav,color:f.bayar===x[0]?"white":C.wht,border:"1px solid "+(f.bayar===x[0]?x[2]:C.bdr),borderRadius:8,padding:"7px 14px",fontWeight:700,fontSize:13,cursor:"pointer"}}>{x[1]}</button>)}
</div>
{f.bayar==="transfer"&&<div style={{display:"flex",gap:8,marginBottom:10}}>{["BSI","BCA"].map(b=><button key={b} onClick={()=>setF(p=>({...p,bank:b}))} style={{background:f.bank===b?C.blu:C.nav,color:f.bank===b?"white":C.wht,border:"2px solid "+(f.bank===b?C.blt:C.bdr),borderRadius:8,padding:"6px 16px",fontWeight:700,cursor:"pointer"}}>{b}</button>)}</div>}
{f.bayar==="bon"&&<Inp label="Deadline" type="date" value={f.deadline} onChange={v=>setF(p=>({...p,deadline:v}))} style={{maxWidth:220}}/>}
{f.bayar==="split"&&(()=>{
var sd=f.splitDetail||{cash:0,tf:0,bon:0};
var totalSplit=(Number(sd.cash)||0)+(Number(sd.tf)||0)+(Number(sd.bon)||0);
var selisih=total-totalSplit;
return <div style={{background:C.nav,border:"2px solid "+C.olt,borderRadius:10,padding:12,marginBottom:10}}>
<div style={{fontWeight:700,color:C.olt,marginBottom:8,fontSize:12}}>✂️ Rincian Split Payment</div>
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:8}}>
{[["cash","💵 Cash",C.glt],["tf","🏦 Transfer",C.blt],["bon","📃 BON",C.rlt]].map(x=><div key={x[0]}>
<div style={{fontSize:10,color:C.gl2,marginBottom:3,fontWeight:600}}>{x[1]}</div>
<input type="number" value={sd[x[0]]||""} placeholder="0" step="1000"
onChange={e=>setF(p=>({...p,splitDetail:{...(p.splitDetail||{}),  [x[0]]:Number(e.target.value)||0}}))}
style={{background:C.bg,border:"1px solid "+(Number(sd[x[0]])>0?x[2]:C.bdr),borderRadius:6,padding:"6px 8px",color:C.wht,fontSize:12,outline:"none",width:"100%"}}/>
</div>)}
</div>
{sd.tf>0&&<div style={{display:"flex",gap:6,marginBottom:8}}>
<span style={{fontSize:11,color:C.gl2,alignSelf:"center"}}>TF ke bank:</span>
{["BSI","BCA"].map(b=><button key={b} onClick={()=>setF(p=>({...p,splitBank:b}))} style={{background:f.splitBank===b?C.blu:C.nav,color:f.splitBank===b?"white":C.wht,border:"2px solid "+(f.splitBank===b?C.blt:C.bdr),borderRadius:6,padding:"4px 12px",fontWeight:700,cursor:"pointer",fontSize:12}}>{b}</button>)}
</div>}
{sd.bon>0&&<Inp label="Deadline BON" type="date" value={f.deadline} onChange={v=>setF(p=>({...p,deadline:v}))} style={{maxWidth:180,marginBottom:8}}/>}
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:6,marginTop:4}}>
{[["Total Split",fR(totalSplit),C.wht],["Total Invoice",fR(total),C.gl2],["Selisih",fR(selisih),Math.abs(selisih)<1?C.glt:C.rlt]].map(x=><div key={x[0]} style={{background:C.bg,borderRadius:6,padding:"5px 8px",border:"1px solid "+C.bdr}}><div style={{fontSize:9,color:C.gl2}}>{x[0]}</div><div style={{fontSize:12,fontWeight:700,color:x[2]}}>{x[1]}</div></div>)}
</div>
{Math.abs(selisih)<1&&totalSplit>0&&<div style={{marginTop:6,fontSize:11,color:C.glt,fontWeight:700}}>✅ Split balance!</div>}
{Math.abs(selisih)>=1&&totalSplit>0&&<div style={{marginTop:6,fontSize:11,color:C.rlt}}>⚠️ Selisih Rp {fR(Math.abs(selisih))} — sesuaikan nominal</div>}
</div>;
})()}
<Inp label="Keterangan" value={f.ket} onChange={v=>setF(p=>({...p,ket:v}))} placeholder="Catatan opsional"/>
<div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
<Btn onClick={()=>doSave(false)} color="green" dis={!valid.length||!f.konsumen}>💾 Simpan</Btn>
<Btn onClick={()=>doSave(true)} color="blue" dis={!valid.length||!f.konsumen}>🖨️ Simpan & Cetak Invoice</Btn>
</div>
</Card>
<Card>
<div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>📊 Laporan Harian Penjualan</div>
{(()=>{
var penjLap=(data.penjualan||[]).filter(p=>p.tanggal===tglLap);
var totalOmzetLap=penjLap.reduce((a,p)=>a+(p.total||0),0);
var totalMarginLap=penjLap.reduce((a,p)=>a+(p.margin||0),0);
// Kelompok per sales
var sgMap={};
penjLap.forEach(p=>{
var empP=(data.employees||[]).find(e=>e.id===p.salesId);
var sNama=empP?.nama||p.salesNama||p.sales||"GUDANG / KASIR";
if(!sgMap[sNama])sgMap[sNama]={nama:sNama,items:[],omzet:0,margin:0,cash:0,tf:0,bon:0};
sgMap[sNama].items.push(p);
sgMap[sNama].omzet+=(p.total||0);
sgMap[sNama].margin+=(p.margin||0);
var byr=(p.bayar||"").toLowerCase();
if(byr==="split"){var sd2=p.splitDetail||{};sgMap[sNama].cash+=Number(sd2.cash||0);sgMap[sNama].tf+=Number(sd2.tf||0);sgMap[sNama].bon+=Number(sd2.bon||0);}
else if(byr==="cash")sgMap[sNama].cash+=(p.total||0);
else if(byr==="transfer"||byr==="tf")sgMap[sNama].tf+=(p.total||0);
else if(byr==="bon")sgMap[sNama].bon+=(p.total||0);
});
var sgList=Object.values(sgMap);
var totCash=sgList.reduce((a,s)=>a+s.cash,0);
var totTF=sgList.reduce((a,s)=>a+s.tf,0);
var totBon=sgList.reduce((a,s)=>a+s.bon,0);
// Total qty & nominal per ukuran keseluruhan
var totQ55=0,totN55=0,totQ12=0,totN12=0,totQ50=0,totN50=0;
penjLap.forEach(p=>(p.items||[]).forEach(it=>{var q=Number(it.qty||0);var h=Number(it.price||0);if(it.ukuran==="5.5 kg"){totQ55+=q;totN55+=q*h;}else if(it.ukuran==="12 kg"){totQ12+=q;totN12+=q*h;}else if(it.ukuran==="50 kg"){totQ50+=q;totN50+=q*h;}}));
return <>
<div style={{display:"flex",alignItems:"center",gap:8,marginBottom:10,flexWrap:"wrap"}}>
<Inp label="Tanggal" type="date" value={tglLap} onChange={setTglLap} style={{maxWidth:170,marginBottom:0}}/>
<div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
{[["Total Omzet",fR(totalOmzetLap),"#0a1f44"],["Margin",fR(totalMarginLap),"#15803D"],["Invoice",penjLap.length+" trx","#1D4ED8"]].map(x=><div key={x[0]} style={{background:"#F8FAFC",borderRadius:6,padding:"4px 10px",border:"1px solid #E2E8F0"}}><span style={{fontSize:10,color:"#64748B"}}>{x[0]}: </span><span style={{fontSize:12,fontWeight:700,color:x[2]}}>{x[1]}</span></div>)}
</div>
</div>
{penjLap.length===0?<div style={{color:"#94A3B8",fontSize:12,padding:"10px 0",fontStyle:"italic"}}>Tidak ada penjualan pada tanggal ini.</div>:
<div style={{background:"white",borderRadius:8,border:"1px solid #E2E8F0",color:"#111",fontFamily:"Arial,sans-serif",overflowX:"auto"}}>
<table style={{width:"auto",minWidth:"895px",maxWidth:"100%",margin:"0 auto",borderCollapse:"collapse",fontSize:13,tableLayout:"fixed"}}>
<colgroup>
<col style={{width:"130px"}}/><col style={{width:"150px"}}/>
<col style={{width:"55px"}}/><col style={{width:"90px"}}/>
<col style={{width:"55px"}}/><col style={{width:"90px"}}/>
<col style={{width:"55px"}}/><col style={{width:"90px"}}/>
<col style={{width:"100px"}}/><col style={{width:"80px"}}/>
</colgroup>
<thead>
<tr style={{background:"#0a1f44"}}>
<th style={{padding:"7px 8px",color:"white",fontWeight:700,textAlign:"left",border:"1px solid #1e3a5f",fontSize:11.5}} rowSpan={2}>Nama Sales</th>
<th style={{padding:"7px 8px",color:"white",fontWeight:700,textAlign:"left",border:"1px solid #1e3a5f",fontSize:11.5}} rowSpan={2}>Konsumen</th>
<th style={{padding:"7px 8px",color:"white",fontWeight:700,textAlign:"center",border:"1px solid #1e3a5f",fontSize:11.5}} colSpan={2}>5,5 kg</th>
<th style={{padding:"7px 8px",color:"white",fontWeight:700,textAlign:"center",border:"1px solid #1e3a5f",fontSize:11.5}} colSpan={2}>12 kg</th>
<th style={{padding:"7px 8px",color:"white",fontWeight:700,textAlign:"center",border:"1px solid #1e3a5f",fontSize:11.5}} colSpan={2}>50 kg</th>
<th style={{padding:"7px 8px",color:"white",fontWeight:700,textAlign:"right",border:"1px solid #1e3a5f",fontSize:11.5}} rowSpan={2}>Total</th>
<th style={{padding:"7px 8px",color:"white",fontWeight:700,textAlign:"center",border:"1px solid #1e3a5f",fontSize:11.5}} rowSpan={2}>Transaksi</th>
</tr>
<tr style={{background:"#1e3a5f"}}>
{["Qty","Harga","Qty","Harga","Qty","Harga"].map((h,i)=><th key={i} style={{padding:"5px 6px",color:"#93C5FD",fontWeight:700,textAlign:"center",border:"1px solid #2d4a6f",fontSize:10.5}}>{h}</th>)}
</tr>
</thead>
<tbody>
{sgList.map((sg,gi)=>{
var sgRows=sg.items.map((p,i)=>{
var it55=(p.items||[]).filter(it=>it.ukuran==="5.5 kg");var q55p=it55.reduce((a,it)=>a+Number(it.qty||0),0);var h55p=it55.length>0?Number(it55[0].price||0):0;
var it12=(p.items||[]).filter(it=>it.ukuran==="12 kg");var q12p=it12.reduce((a,it)=>a+Number(it.qty||0),0);var h12p=it12.length>0?Number(it12[0].price||0):0;
var it50=(p.items||[]).filter(it=>it.ukuran==="50 kg");var q50p=it50.reduce((a,it)=>a+Number(it.qty||0),0);var h50p=it50.length>0?Number(it50[0].price||0):0;
var byrl=(p.bayar||"").toLowerCase();
return <tr key={p.id} style={{background:i%2===0?"white":"#f9f9f9",borderBottom:"1px solid #e5e7eb"}}>
{i===0&&<td style={{padding:"6px 8px",fontWeight:700,color:"#0a1f44",border:"1px solid #e5e7eb",verticalAlign:"top",background:"#EFF6FF",fontSize:12.5}} rowSpan={sg.items.length+1}>{sg.nama}</td>}
<td style={{padding:"6px 8px",color:"#111",border:"1px solid #e5e7eb",fontSize:12.5}}>{p.konsumen}</td>
<td style={{padding:"6px 8px",textAlign:"center",color:q55p>0?"#15803D":"#94A3B8",border:"1px solid #e5e7eb",fontWeight:q55p>0?700:400,fontSize:12.5}}>{q55p||"—"}</td>
<td style={{padding:"6px 8px",textAlign:"right",color:q55p>0?"#374151":"#94A3B8",border:"1px solid #e5e7eb",fontSize:12,whiteSpace:"nowrap"}}>{q55p>0?fR(h55p):"—"}</td>
<td style={{padding:"6px 8px",textAlign:"center",color:q12p>0?"#1D4ED8":"#94A3B8",border:"1px solid #e5e7eb",fontWeight:q12p>0?700:400,fontSize:12.5}}>{q12p||"—"}</td>
<td style={{padding:"6px 8px",textAlign:"right",color:q12p>0?"#374151":"#94A3B8",border:"1px solid #e5e7eb",fontSize:12,whiteSpace:"nowrap"}}>{q12p>0?fR(h12p):"—"}</td>
<td style={{padding:"6px 8px",textAlign:"center",color:q50p>0?"#B45309":"#94A3B8",border:"1px solid #e5e7eb",fontWeight:q50p>0?700:400,fontSize:12.5}}>{q50p||"—"}</td>
<td style={{padding:"6px 8px",textAlign:"right",color:q50p>0?"#374151":"#94A3B8",border:"1px solid #e5e7eb",fontSize:12,whiteSpace:"nowrap"}}>{q50p>0?fR(h50p):"—"}</td>
<td style={{padding:"6px 8px",textAlign:"right",fontWeight:700,color:"#111",border:"1px solid #e5e7eb",fontSize:12.5,whiteSpace:"nowrap"}}>{fR(p.total)}</td>
<td style={{padding:"6px 8px",textAlign:"center",border:"1px solid #e5e7eb",color:byrl==="cash"?"#15803D":byrl==="bon"?"#DC2626":"#1D4ED8",fontWeight:700,fontSize:11.5}}>{p.bayar}</td>
</tr>;});
// Total laku per sales
var tq55sg=sg.items.reduce((a,p)=>a+(p.items||[]).filter(it=>it.ukuran==="5.5 kg").reduce((b,it)=>b+Number(it.qty||0),0),0);
var tq12sg=sg.items.reduce((a,p)=>a+(p.items||[]).filter(it=>it.ukuran==="12 kg").reduce((b,it)=>b+Number(it.qty||0),0),0);
var tq50sg=sg.items.reduce((a,p)=>a+(p.items||[]).filter(it=>it.ukuran==="50 kg").reduce((b,it)=>b+Number(it.qty||0),0),0);
var totalLakuRow=<tr key={"tl"+gi} style={{background:"#EFF6FF",fontWeight:700,borderBottom:"2px solid #BFDBFE"}}>
<td style={{padding:"7px 8px",color:"#1D4ED8",fontStyle:"italic",border:"1px solid #BFDBFE",fontSize:12}}>Total Laku</td>
<td style={{padding:"7px 8px",textAlign:"center",color:"#15803D",border:"1px solid #BFDBFE",fontSize:12.5}}>{tq55sg||"—"}</td>
<td style={{padding:"7px 8px",border:"1px solid #BFDBFE"}}></td>
<td style={{padding:"7px 8px",textAlign:"center",color:"#1D4ED8",border:"1px solid #BFDBFE",fontSize:12.5}}>{tq12sg||"—"}</td>
<td style={{padding:"7px 8px",border:"1px solid #BFDBFE"}}></td>
<td style={{padding:"7px 8px",textAlign:"center",color:"#B45309",border:"1px solid #BFDBFE",fontSize:12.5}}>{tq50sg||"—"}</td>
<td style={{padding:"7px 8px",border:"1px solid #BFDBFE"}}></td>
<td style={{padding:"7px 8px",textAlign:"right",color:"#0a1f44",border:"1px solid #BFDBFE",fontSize:12.5,whiteSpace:"nowrap"}}>{fR(sg.omzet)}</td>
<td style={{padding:"7px 8px",border:"1px solid #BFDBFE"}}></td>
</tr>;
return [...sgRows, totalLakuRow];
})}
{/* Grand total */}
<tr style={{background:"#0a1f44",fontWeight:700}}>
<td colSpan={2} style={{padding:"8px 9px",color:"white",border:"1px solid #1e3a5f",fontSize:12.5}}>TOTAL KESELURUHAN</td>
<td style={{padding:"8px 9px",textAlign:"center",color:"#86EFAC",border:"1px solid #1e3a5f",fontSize:12.5}}>{totQ55||"—"}</td>
<td style={{padding:"8px 9px",border:"1px solid #1e3a5f"}}></td>
<td style={{padding:"8px 9px",textAlign:"center",color:"#93C5FD",border:"1px solid #1e3a5f",fontSize:12.5}}>{totQ12||"—"}</td>
<td style={{padding:"8px 9px",border:"1px solid #1e3a5f"}}></td>
<td style={{padding:"8px 9px",textAlign:"center",color:"#FCD34D",border:"1px solid #1e3a5f",fontSize:12.5}}>{totQ50||"—"}</td>
<td style={{padding:"8px 9px",border:"1px solid #1e3a5f"}}></td>
<td style={{padding:"6px 8px",textAlign:"right",color:"white",fontSize:13,border:"1px solid #1e3a5f"}}>{fR(totalOmzetLap)}</td>
<td style={{padding:"6px 8px",border:"1px solid #1e3a5f"}}></td>
</tr>
<tr style={{background:"#0f2744"}}>
<td colSpan={2} style={{padding:"7px 8px",color:"#93C5FD",fontSize:11.5,border:"1px solid #1e3a5f"}}>Cash: {fR(totCash)} &nbsp;|&nbsp; TF: {fR(totTF)} &nbsp;|&nbsp; BON: {fR(totBon)}</td>
<td colSpan={8} style={{padding:"7px 8px",color:"#9CA3AF",fontSize:11.5,border:"1px solid #1e3a5f"}}>* BON = piutang baru, belum masuk kas</td>
</tr>
</tbody>
</table>
</div>}
</>;
})()}

<div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
<div style={{fontWeight:700,color:C.gl2,fontSize:13}}>📋 Riwayat Penjualan</div>
<button onClick={()=>{
var wb=XLSX.utils.book_new();
var detailRows=[];
rows.forEach(p=>{
  var items=p.items||[];
  var totalQtyInv=items.reduce((a,it)=>a+Number(it.qty||0),0);
  items.forEach(it=>{
    var q=Number(it.qty||0);
    var marginItem;
    if(it.hppFIFO!=null)marginItem=(Number(it.price||0)-Number(it.hppFIFO||0))*q;
    else marginItem=totalQtyInv>0?(p.margin||0)*(q/totalQtyInv):0;
    detailRows.push([p.tanggal,p.noInv||"",p.salesNama,p.konsumen,it.ukuran+" — "+(it.jenis==="Tabung+Isi"?"Tabung+Isi":"Refill/Isi"),q,Number(it.price||0),q*Number(it.price||0),Math.round(marginItem)]);
  });
});
var totalQtyAll=detailRows.reduce((a,r)=>a+r[5],0);
var totalMarginAll=detailRows.reduce((a,r)=>a+r[8],0);
var aoa=[["Tanggal","No.Invoice","Sales","Konsumen","Produk","Qty","Harga Satuan","Subtotal","Margin"],...detailRows,[],["","","","","TOTAL",totalQtyAll,"","",totalMarginAll]];
var ws=XLSX.utils.aoa_to_sheet(aoa);
ws["!cols"]=[{wch:12},{wch:20},{wch:14},{wch:18},{wch:24},{wch:8},{wch:13},{wch:14},{wch:13}];
XLSX.utils.book_append_sheet(wb,ws,"Detail Produk");
var fname="Penjualan_Detail"+(barFilter.from?"_"+barFilter.from:"")+(barFilter.to?"_s_"+barFilter.to:"")+".xlsx";
XLSX.writeFile(wb,fname);
toast("✓ Excel detail produk didownload!");
}} style={{background:"#15803D",color:"white",border:"none",padding:"7px 14px",borderRadius:7,fontSize:12,cursor:"pointer",fontWeight:700}}>📥 Export Excel (Detail Produk)</button>
</div>
<div style={{background:C.nav,borderRadius:8,padding:10,marginBottom:10,border:"1px solid "+C.bdr}}>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(140px,1fr))",gap:8}}>
<Inp label="Dari" type="date" value={barFilter.from} onChange={v=>setBarFilter(p=>({...p,from:v}))} style={{marginBottom:0}}/>
<Inp label="Sampai" type="date" value={barFilter.to} onChange={v=>setBarFilter(p=>({...p,to:v}))} style={{marginBottom:0}}/>
<Sel label="Sales" value={barFilter.salesId} onChange={v=>setBarFilter(p=>({...p,salesId:v}))} opts={[{v:"",l:"Semua"},...salesEmp.map(e=>({v:e.id,l:e.nama}))]} style={{marginBottom:0}}/>
<Inp label="Konsumen" value={barFilter.konsumen} onChange={v=>setBarFilter(p=>({...p,konsumen:v}))} placeholder="Cari..." style={{marginBottom:0}}/>
<Sel label="Bayar" value={barFilter.bayar} onChange={v=>setBarFilter(p=>({...p,bayar:v}))} opts={[{v:"",l:"Semua"},{v:"cash",l:"Cash"},{v:"transfer",l:"Transfer"},{v:"bon",l:"BON"}]} style={{marginBottom:0}}/>
</div>
{(barFilter.from||barFilter.to||barFilter.salesId||barFilter.konsumen||barFilter.bayar)&&<div style={{marginTop:8}}><Btn sm color="gray" onClick={()=>setBarFilter({from:"",to:"",salesId:"",konsumen:"",bayar:""})}>✕ Reset Bar Filter</Btn></div>}
</div>
<FilterTbl columns={cols} data={rows} empty="Belum ada penjualan" maxRows={1000}/>
</Card>

{editInv&&(()=>{
var ef=editInv.form;
var setEf=newF=>setEditInv(prev=>({...prev,form:typeof newF==="function"?newF(prev.form):newF}));
var setItem2=(i,k,v)=>setEf(p=>{var its=[...p.items];its[i]={...its[i],[k]:v};return{...p,items:its};});
var valid2=(ef.items||[]).filter(it=>Number(it.qty)>0&&Number(it.price)>0);
var total2=iTotal(valid2);
var salesEmpE=sortEmp((data.employees||[]).filter(e=>e.aktif&&PENJUALAN_ROLES.includes(e.role)));
var kNamesE=[...new Set([...(data.pelanggan||[]).map(p=>p.nama),...(data.penjualan||[]).map(e=>e.konsumen)].filter(Boolean))];
function saveEdit(){
// Reverse stok lama (qty fisik) lalu apply stok baru
var ns={...(data.stock||{})};var nk={...(data.stokKosong||{})};var na={...(data.totalTabung||{})};
(editInv.entry.items||[]).forEach(it=>{var q=Number(it.qty||0);var s=it.ukuran;if(it.jenis==="Tabung+Isi"){ns[s]=(ns[s]||0)+q;nk[s]=(nk[s]||0)+q;na[s]=(na[s]||0)+q;}else{ns[s]=(ns[s]||0)+q;nk[s]=Math.max(0,(nk[s]||0)-q);}});
valid2.forEach(it=>{var q=Number(it.qty||0);var s=it.ukuran;if(it.jenis==="Tabung+Isi"){ns[s]=Math.max(0,(ns[s]||0)-q);nk[s]=Math.max(0,(nk[s]||0)-q);na[s]=Math.max(0,(na[s]||0)-q);}else{ns[s]=Math.max(0,(ns[s]||0)-q);nk[s]=(nk[s]||0)+q;}});
var newBon=(data.bon||[]).map(b=>{if(b.noInv===editInv.entry.noInv&&b.konsumen===editInv.entry.konsumen){return{...b,konsumen:ef.konsumen,konsumenId:ef.konsumenId||b.konsumenId,total:total2,sisaTagihan:total2,items:valid2};}return b;});
setData(d=>{
  // ── Kembalikan qty ke batch FIFO asal (reverse) berdasarkan fifoDetail invoice lama ──
  var sb={...(d.stokBatch||{})};
  (editInv.entry.fifoDetail||[]).forEach(fd=>{
    (fd.detail||[]).forEach(det=>{
      if(!det.batchId)return;// batch fallback (sudah kosong saat itu), tidak bisa dikembalikan ke batch spesifik
      var arr=(sb[fd.ukuran]||[]).map(b=>b.id===det.batchId?{...b,qtySisa:b.qtySisa+det.qty}:b);
      sb[fd.ukuran]=arr;
    });
  });
  var dRev={...d,stokBatch:sb};
  // ── Konsumsi ulang FIFO dengan item yang sudah diedit ──
  var dCur=dRev;var marginFIFO=0;var fifoDetailAll=[];
  var itemsFinal=valid2.map(it=>{
    var q=Number(it.qty||0);
    var res=consumeFIFO(dCur,it.ukuran,q);
    dCur=res.data;
    var marginItem=(Number(it.price||0)*q)-res.hppTotal;
    marginFIFO+=marginItem;
    fifoDetailAll.push({ukuran:it.ukuran,jenis:it.jenis,qty:q,hppTotal:res.hppTotal,hppPerUnit:res.hppPerUnit,detail:res.detail,qtyKurang:res.qtyKurang});
    return{...it,qty:q,price:Number(it.price),hppFIFO:res.hppPerUnit};
  });
  var updatedEntry={...editInv.entry,...ef,items:itemsFinal,total:total2,margin:marginFIFO,fifoDetail:fifoDetailAll,editLog:[...(editInv.entry.editLog||[]),{by:user?.nama||"",at:new Date().toISOString(),before:{konsumen:editInv.entry.konsumen,bayar:editInv.entry.bayar,total:editInv.entry.total,items:editInv.entry.items},note:"Diedit"}]};
  return{...dCur,penjualan:(d.penjualan||[]).map(x=>x.id===editInv.entry.id?updatedEntry:x),stock:ns,stokKosong:nk,totalTabung:na,bon:newBon};
});
setEditInv(null);toast("✓ Invoice diperbarui! Stok & margin FIFO disesuaikan.");
}
return <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.8)",zIndex:999,display:"flex",alignItems:"flex-start",justifyContent:"center",padding:16,overflowY:"auto"}}>
<div style={{background:C.card,borderRadius:12,width:"100%",maxWidth:680,border:"1px solid "+C.bdr,marginTop:20}}>
<div style={{padding:"14px 18px",borderBottom:"1px solid "+C.bdr,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
<div><div style={{fontWeight:700,color:C.wht,fontSize:14}}>✏️ Edit Invoice — {editInv.entry.noInv}</div>
{(editInv.entry.editLog||[]).length>0&&<div style={{fontSize:10,color:C.gl2,marginTop:2}}>Terakhir diedit: {(editInv.entry.editLog||[]).slice(-1)[0]?.by} · {new Date((editInv.entry.editLog||[]).slice(-1)[0]?.at).toLocaleString("id-ID")}</div>}
</div>
<button onClick={()=>setEditInv(null)} style={{background:"transparent",border:"none",color:C.gl2,cursor:"pointer",fontSize:20}}>✕</button>
</div>
<div style={{padding:16}}>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(180px,260px))",gap:10,marginBottom:10}}>
<Inp label="Tanggal" type="date" value={ef.tanggal||""} onChange={v=>setEf(p=>({...p,tanggal:v}))}/>
<Sel label="Sales" value={ef.salesId||""} onChange={v=>setEf(p=>({...p,salesId:v}))} opts={[{v:"",l:"-- Pilih --"},...salesEmpE.map(e=>({v:e.id,l:e.nama}))]}/>
</div>
<div style={{marginBottom:10}}>
<label style={{fontSize:11,color:C.gl2,display:"block",marginBottom:4}}>Konsumen</label>
<input list="kl2" value={ef.konsumen||""} onChange={e=>{var v=e.target.value;var p=(data.pelanggan||[]).find(x=>x.nama===v);setEf(f=>({...f,konsumen:v,konsumenId:p?.id||f.konsumenId}));}} style={{background:C.bg,border:"1px solid "+C.bdr,borderRadius:8,padding:"8px 12px",color:C.wht,fontSize:13,outline:"none",width:"100%"}}/>
<datalist id="kl2">{kNamesE.map(n=><option key={n} value={n}/>)}</datalist>
</div>
{/* Items */}
<div style={{marginBottom:10}}>
<div style={{fontSize:11,fontWeight:700,color:C.gl2,marginBottom:6}}>Items</div>
{(ef.items||[]).map((it,i)=><div key={i} style={{display:"grid",gridTemplateColumns:"1fr 1fr 80px 90px 32px",gap:6,marginBottom:6,alignItems:"center"}}>
<Sel label="" value={it.ukuran||"5.5 kg"} onChange={v=>setItem2(i,"ukuran",v)} opts={SIZES.map(s=>({v:s,l:s}))}/>
<Sel label="" value={it.jenis||"Isi"} onChange={v=>setItem2(i,"jenis",v)} opts={[{v:"Isi",l:"Isi (Refill)"},{v:"Tabung+Isi",l:"Tabung+Isi"}]}/>
<input type="number" value={it.qty} placeholder="Qty" onChange={e=>setItem2(i,"qty",e.target.value)} style={{background:C.bg,border:"1px solid "+C.bdr,borderRadius:6,padding:"6px 8px",color:C.wht,fontSize:12,outline:"none",width:"100%"}}/>
<input type="number" value={it.price} step="1000" placeholder="Harga" onChange={e=>setItem2(i,"price",e.target.value)} style={{background:C.bg,border:"1px solid "+C.bdr,borderRadius:6,padding:"6px 8px",color:C.wht,fontSize:12,outline:"none",width:"100%"}}/>
<button onClick={()=>setEf(p=>({...p,items:p.items.filter((_,j)=>j!==i)}))} style={{background:"transparent",border:"none",color:C.rlt,cursor:"pointer",fontSize:16}}>−</button>
</div>)}
<button onClick={()=>setEf(p=>({...p,items:[...p.items,{ukuran:"5.5 kg",jenis:"Isi",qty:"",price:""}]}))} style={{background:C.nav,border:"1px solid "+C.bdr,borderRadius:6,padding:"5px 12px",color:C.gl2,cursor:"pointer",fontSize:12}}>+ Item</button>
<div style={{textAlign:"right",marginTop:6,fontSize:13,fontWeight:700,color:C.wht}}>Total: {fR(total2)}</div>
</div>
{/* Metode bayar */}
<div style={{marginBottom:10}}>
<div style={{fontSize:11,fontWeight:700,color:C.gl2,marginBottom:6}}>Metode Bayar</div>
<div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
{[["cash","💵 Cash",C.grn],["transfer","🏦 Transfer",C.blu],["bon","📃 BON",C.rdk],["split","✂️ Split",C.olt]].map(x=><button key={x[0]} onClick={()=>setEf(p=>({...p,bayar:x[0]}))} style={{background:ef.bayar===x[0]?x[2]:C.nav,color:ef.bayar===x[0]?"white":C.wht,border:"1px solid "+(ef.bayar===x[0]?x[2]:C.bdr),borderRadius:8,padding:"6px 12px",fontWeight:700,fontSize:12,cursor:"pointer"}}>{x[1]}</button>)}
</div>
{ef.bayar==="transfer"&&<div style={{display:"flex",gap:8,marginTop:8}}>{["BSI","BCA"].map(b=><button key={b} onClick={()=>setEf(p=>({...p,bank:b}))} style={{background:ef.bank===b?C.blu:C.nav,color:ef.bank===b?"white":C.wht,border:"2px solid "+(ef.bank===b?C.blt:C.bdr),borderRadius:8,padding:"5px 14px",fontWeight:700,cursor:"pointer"}}>{b}</button>)}</div>}
</div>
<Inp label="Keterangan (opsional)" value={ef.ket||""} onChange={v=>setEf(p=>({...p,ket:v}))} placeholder="Catatan..."/>
{/* Log perubahan */}
{(editInv.entry.editLog||[]).length>0&&<div style={{marginBottom:10,background:C.bg,borderRadius:8,padding:10,border:"1px solid "+C.bdr}}>
<div style={{fontSize:11,fontWeight:700,color:C.gl2,marginBottom:6}}>📋 Log Perubahan</div>
{(editInv.entry.editLog||[]).slice().reverse().map((lg,i)=><div key={i} style={{fontSize:10,color:C.gl2,marginBottom:4,paddingBottom:4,borderBottom:"1px solid "+C.bdr}}>
<b style={{color:C.wht}}>{lg.by}</b> · {new Date(lg.at).toLocaleString("id-ID")} · {lg.note}<br/>
<span style={{color:"#9CA3AF"}}>Sebelum: {lg.before?.konsumen} · {lg.before?.bayar} · {fR(lg.before?.total||0)}</span>
</div>)}
</div>}
<div style={{display:"flex",gap:8,justifyContent:"flex-end",marginTop:10}}>
<button onClick={()=>setEditInv(null)} style={{background:C.nav,border:"1px solid "+C.bdr,borderRadius:8,padding:"9px 18px",color:C.gl2,cursor:"pointer",fontWeight:700}}>Batal</button>
<button onClick={saveEdit} style={{background:C.glt,border:"none",borderRadius:8,padding:"9px 18px",color:"white",cursor:"pointer",fontWeight:700,fontSize:13}}>💾 Simpan Perubahan</button>
</div>
</div>
</div>
</div>;
})()}

{delId&&<ConfirmDel msg={"Hapus invoice "+delId.noInv+" ("+delId.konsumen+")? Stok akan dikembalikan."} onCancel={()=>setDelId(null)} onConfirm={()=>{
var ns={...(data.stock||{})};var nk={...(data.stokKosong||{})};var na={...(data.totalTabung||{})};
(delId.items||[]).forEach(it=>{
  var q=Number(it.qty||0);var uk=it.ukuran;
  ns[uk]=(ns[uk]||0)+q;
  if(it.jenis==="Tabung+Isi"){nk[uk]=(nk[uk]||0)+q;na[uk]=(na[uk]||0)+q;}
  else{nk[uk]=Math.max(0,(nk[uk]||0)-q);}
});
var log={id:uid(),tanggal:delId.tanggal,ukuran:"Multiple",jenis:"Reverse Hapus Inv "+delId.noInv,qty:0,ket:"Stok dikembalikan karena invoice dihapus",sumber:"Hapus",user:user?.nama||""};
setData(d=>{
  // Kembalikan qty ke batch FIFO asal
  var sb={...(d.stokBatch||{})};
  (delId.fifoDetail||[]).forEach(fd=>{
    (fd.detail||[]).forEach(det=>{
      if(!det.batchId)return;
      var arr=(sb[fd.ukuran]||[]).map(b=>b.id===det.batchId?{...b,qtySisa:b.qtySisa+det.qty}:b);
      sb[fd.ukuran]=arr;
    });
  });
  return{...d,stokBatch:sb,penjualan:(d.penjualan||[]).filter(x=>x.id!==delId.id),stock:ns,stokKosong:nk,totalTabung:na,stockLog:[log,...(d.stockLog||[])].slice(0,500)};
});
setDelId(null);toast("✓ Invoice dihapus & stok dikembalikan!");
}}/>}
</div>;
}

// === AKHIR BAGIAN 2 ===
// === BAGIAN 3 DARI 3 ===

// ─── STOK ─────────────────────────────────────────────────────────────────────
function StokMod({data,setData,user,toast}){
var C=useTheme();var[tab,setTab]=useState("rekap");var[ba,setBa]=useState(null);
var[stokBln,setStokBln]=useState(toMonth());
var[showInject,setShowInject]=useState(false);
var[injectTgl,setInjectTgl]=useState(toDay());
var[injectF,setInjectF]=useState({});
function RekapTab(){
var td2=toDay();
var rowsRekap=buildStokHarian(data,toMonth()).filter(r=>r.tgl<=td2);
var lastRow=rowsRekap.length>0?rowsRekap[rowsRekap.length-1]:null;
return <div>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(240px,1fr))",gap:12,marginBottom:14}}>
{SIZES.map(s=>{
var isi=lastRow?lastRow.akhirIsi[s]:((data.stock||{})[s]||0);
var kosong=lastRow?lastRow.akhirTK[s]:getKosong(data,s);
var titip=lastRow?lastRow.titipSnap[s]:getTitipTotal(data.titipList,s);
var totalS=isi+kosong+titip;
return <Card key={s} style={{marginBottom:0}}>
<div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
<div style={{fontWeight:800,color:C.wht,fontSize:14}}>📦 LPG {s}</div>
<div style={{textAlign:"right"}}><div style={{fontSize:9,color:C.gl2}}>Total Tabung</div><div style={{fontSize:18,fontWeight:900,color:C.olt}}>{totalS}</div></div>
</div>
{totalS>0&&<div style={{height:6,borderRadius:3,background:C.bdr,display:"flex",overflow:"hidden",marginBottom:8}}>
<div style={{width:(isi/totalS*100)+"%",background:C.glt}}/><div style={{width:(titip/totalS*100)+"%",background:C.blt}}/><div style={{width:(kosong/totalS*100)+"%",background:C.gl2}}/>
</div>}
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:5}}>
{[["Tbg+Isi",isi,C.glt],["Titip",titip,C.blt],["Kosong",kosong,C.gl2]].map(x=><div key={x[0]} style={{background:C.nav,borderRadius:6,padding:"5px 4px",textAlign:"center",border:"1px solid "+C.bdr}}><div style={{fontSize:8,color:C.gl2}}>{x[0]}</div><div style={{fontSize:18,fontWeight:900,color:x[2]}}>{x[1]}</div></div>)}
</div>
</Card>;
})}
</div>
<Card>
<div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
<div style={{fontWeight:700,color:C.gl2,fontSize:13}}>📋 Laporan Stok Harian</div>
<div style={{display:"flex",gap:8,alignItems:"center"}}>
<Inp label="" type="month" value={stokBln} onChange={setStokBln} style={{marginBottom:0,maxWidth:160}}/>
<Btn sm color="orange" onClick={()=>setShowInject(!showInject)}>{showInject?"✕ Tutup":"★ Inject Stok Awal"}</Btn>
</div>
</div>
{showInject&&<div style={{background:C.nav,borderRadius:8,padding:12,border:"1px solid #F59E0B",marginBottom:12}}>
<div style={{fontWeight:700,color:"#F59E0B",marginBottom:8,fontSize:12}}>★ Set Stok Awal Manual</div>
<div style={{fontSize:11,color:C.gl2,marginBottom:8}}>Tentukan titik awal perhitungan. Sistem hitung otomatis hari berikutnya.</div>
<div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"flex-end",marginBottom:8}}>
<Inp label="Tanggal" type="date" value={injectTgl} onChange={setInjectTgl} style={{maxWidth:160,marginBottom:0}}/>
{SIZES.map(s=>[
<Inp key={"ii"+s} label={"Isi "+s} type="number" value={injectF["isi_"+s]||""} onChange={v=>setInjectF(p=>({...p,["isi_"+s]:v}))} style={{maxWidth:90,marginBottom:0}}/>,
<Inp key={"ik"+s} label={"TK "+s} type="number" value={injectF["tk_"+s]||""} onChange={v=>setInjectF(p=>({...p,["tk_"+s]:v}))} style={{maxWidth:90,marginBottom:0}}/>
])}
</div>
<div style={{display:"flex",gap:8}}>
<Btn color="orange" onClick={()=>{
var rec={tanggal:injectTgl};
SIZES.forEach(s=>{rec["isi_"+s]=Number(injectF["isi_"+s]||0);rec["tk_"+s]=Number(injectF["tk_"+s]||0);});
var existing=(data.stokHarian||[]).filter(r=>r.tanggal!==injectTgl);
setData(d=>({...d,stokHarian:[rec,...existing]}));
setShowInject(false);setInjectF({});
toast("✓ Stok awal "+fDs(injectTgl)+" disimpan!");
}}>💾 Simpan</Btn>
<Btn sm color="gray" onClick={()=>{setShowInject(false);setInjectF({});}}>Batal</Btn>
</div>
{(data.stokHarian||[]).length>0&&<div style={{marginTop:8}}>
{(data.stokHarian||[]).sort((a,b)=>b.tanggal.localeCompare(a.tanggal)).map(r=><div key={r.tanggal} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"5px 8px",background:C.bg,borderRadius:6,marginBottom:3,fontSize:10}}>
<span style={{color:"#F59E0B",fontWeight:700}}>★ {fDs(r.tanggal)} — {SIZES.map(s=>"isi "+s+":"+r["isi_"+s]).join(", ")}</span>
<button onClick={()=>setData(d=>({...d,stokHarian:(d.stokHarian||[]).filter(x=>x.tanggal!==r.tanggal)}))} style={{background:C.rdk,border:"none",borderRadius:4,padding:"2px 7px",color:"white",cursor:"pointer",fontSize:10}}>✕</button>
</div>)}
</div>}
</div>}
<TabelStokBulanan data={data} bulan={stokBln}/>
</Card>
</div>;
}
function MutasiTab(){
var[f,setF]=useState({ukuran:"12 kg",jenis:"return_isi",qty:"",harga:"",ket:"",tanggal:toDay()});var[delId,setDelId]=useState(null);
var needHarga=f.jenis==="return_isi";// jenis yang MENAMBAH stok isi butuh harga modal referensi
var isKeluarIsi=["rusak","isi_hilang","tbgisi_hilang"].includes(f.jenis);// jenis yang MENGURANGI stok isi (konsumsi FIFO)
function save(){
if(!f.qty||Number(f.qty)<=0)return;
if(needHarga&&(!f.harga||Number(f.harga)<=0)){toast("⚠️ Harga modal referensi wajib diisi");return;}
var qty=Number(f.qty);var s=f.ukuran;
var jDesc="";
if(f.jenis==="return_isi")jDesc="↩️ Return (+Isi, -Kosong)";
else if(f.jenis==="beli_tbg")jDesc="🛒 Beli Tabung (+Kosong, +Total)";
else if(f.jenis==="rusak")jDesc="💥 Rusak/Bocor (-Isi, +Kosong)";
else if(f.jenis==="tbg_kosong_hilang")jDesc="🕳️ Tbg Kosong Hilang (-Kosong, -Total)";
else if(f.jenis==="isi_hilang")jDesc="👻 Isi Hilang (-Isi, +Kosong)";
else if(f.jenis==="tbgisi_hilang")jDesc="💀 Tbg+Isi Hilang (-Isi, -Kosong, -Total)";
setData(d=>{
  var ns={...(d.stock||{})};var nk={...(d.stokKosong||{})};var na={...(d.totalTabung||{})};
  var dCur=d;var kerugianFIFO=0;
  if(f.jenis==="return_isi"){
    ns[s]=(ns[s]||0)+qty;nk[s]=Math.max(0,(nk[s]||0)-qty);
    dCur=addBatch(dCur,s,qty,Number(f.harga),"Return Isi dari Konsumen",f.tanggal);
  }else if(f.jenis==="beli_tbg"){
    nk[s]=(nk[s]||0)+qty;na[s]=(na[s]||0)+qty;
  }else if(isKeluarIsi){
    ns[s]=Math.max(0,(ns[s]||0)-qty);
    var res=consumeFIFO(dCur,s,qty);dCur=res.data;kerugianFIFO=res.hppTotal;
    if(f.jenis==="rusak")nk[s]=(nk[s]||0)+qty;
    else if(f.jenis==="isi_hilang")nk[s]=(nk[s]||0)+qty;
    else if(f.jenis==="tbgisi_hilang"){nk[s]=Math.max(0,(nk[s]||0)-qty);na[s]=Math.max(0,(na[s]||0)-qty);}
  }else if(f.jenis==="tbg_kosong_hilang"){
    nk[s]=Math.max(0,(nk[s]||0)-qty);na[s]=Math.max(0,(na[s]||0)-qty);
  }
  var log={id:uid(),tanggal:f.tanggal,ukuran:s,jenis:jDesc+(kerugianFIFO>0?" — Kerugian HPP: "+fR(kerugianFIFO):""),qty,ket:f.ket,user:user?.nama||"",sumber:"Manual"};
  return{...dCur,stock:ns,stokKosong:nk,totalTabung:na,stockLog:[log,...(d.stockLog||[])].slice(0,500)};
});
setF(p=>({...p,qty:"",harga:"",ket:""}));
toast("✓ Mutasi dicatat! "+jDesc);
}
return <div><Card style={{width:"fit-content",maxWidth:"100%",minWidth:660}}><div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(155px,210px))",gap:10}}><Inp label="Tanggal" type="date" value={f.tanggal} onChange={v=>setF(p=>({...p,tanggal:v}))}/><Sel label="Ukuran" value={f.ukuran} onChange={v=>setF(p=>({...p,ukuran:v}))} opts={SIZES}/><Sel label="Jenis" value={f.jenis} onChange={v=>setF(p=>({...p,jenis:v}))} opts={[{v:"return_isi",l:"↩️ Return (+Isi)"},{v:"beli_tbg",l:"🛒 Beli Tabung dr Konsumen (+Tbg)"},{v:"rusak",l:"💥 Rusak/Bocor (-Isi)"},{v:"tbg_kosong_hilang",l:"🕳️ Tbg Kosong Hilang (-Kosong)"},{v:"isi_hilang",l:"👻 Isi Hilang (-Isi)"},{v:"tbgisi_hilang",l:"💀 Tbg+Isi Hilang (-Isi,-Kosong)"}]}/><Inp label="Qty" type="number" value={f.qty} onChange={v=>setF(p=>({...p,qty:v}))}/>{needHarga&&<Inp label="Harga Modal/Unit" type="number" step="1000" value={f.harga} onChange={v=>setF(p=>({...p,harga:v}))} placeholder="HPP referensi"/>}<Inp label="Ket" value={f.ket} onChange={v=>setF(p=>({...p,ket:v}))}/></div>
{isKeluarIsi&&<div style={{fontSize:10,color:C.olt,marginBottom:8}}>ℹ️ Stok isi akan dikurangi mengikuti FIFO (batch tertua dulu); kerugian HPP otomatis tercatat di log.</div>}
<Btn color="green" onClick={save} dis={!f.qty}>💾 Simpan Mutasi</Btn></Card><Card><div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>Log Mutasi</div><RTbl headers={["Tgl","Ukuran","Jenis","Qty","Ket","Aksi"]} rows={(data.stockLog||[]).slice(0,50).map(l=>[fDs(l.tanggal),l.ukuran,l.jenis,l.qty,l.ket||"-",<ActBtns onDel={()=>setDelId(l)}/>])}/></Card>{delId&&<ConfirmDel msg="Hapus log?" onCancel={()=>setDelId(null)} onConfirm={()=>{setData(d=>({...d,stockLog:(d.stockLog||[]).filter(x=>x.id!==delId.id)}));setDelId(null);}}/>}</div>;
}
function OpnameTab(){
var C=useTheme();
var[opDate,setOpDate]=useState(toDay());
var[opF,setOpF]=useState(()=>{var o={};SIZES.forEach(s=>{o["isi_"+s]="";o["kos_"+s]="";});return o;});
var[opResult,setOpResult]=useState(null);

function hitungSelisih(){
  var results=SIZES.map(s=>{
    var sysIsi=(data.stock||{})[s]||0;
    var sysKos=getKosong(data,s);
    var fisIsi=opF["isi_"+s]!==""?Number(opF["isi_"+s]):null;
    var fisKos=opF["kos_"+s]!==""?Number(opF["kos_"+s]):null;
    var selIsi=fisIsi!==null?fisIsi-sysIsi:null;
    var selKos=fisKos!==null?fisKos-sysKos:null;
    return{ukuran:s,sysIsi,sysKos,fisIsi,fisKos,selIsi,selKos};
  });
  setOpResult({tanggal:opDate,results});
}

function doAdjust(s,type,sel){
  var ns={...(data.stock||{})};var nk={...(data.stokKosong||{})};
  var jDesc="";
  if(type==="isi"){ns[s]=Math.max(0,(ns[s]||0)+sel);jDesc="Adj Opname Isi ("+fDs(opDate)+")";}
  else{nk[s]=Math.max(0,(nk[s]||0)+sel);jDesc="Adj Opname Kosong ("+fDs(opDate)+")";}
  var log={id:uid(),tanggal:opDate,ukuran:s,jenis:jDesc,qty:Math.abs(sel),ket:"Selisih Opname Checker",user:user?.nama||"",sumber:"Opname"};
  setData(d=>({...d,stock:ns,stokKosong:nk,stockLog:[log,...(d.stockLog||[])].slice(0,500)}));
  // Update hasil opname display
  setOpResult(prev=>prev?{...prev,results:prev.results.map(r=>{
    if(r.ukuran!==s)return r;
    var newSysIsi=type==="isi"?Math.max(0,r.sysIsi+sel):r.sysIsi;
    var newSysKos=type==="kos"?Math.max(0,r.sysKos+sel):r.sysKos;
    return{...r,sysIsi:newSysIsi,sysKos:newSysKos,selIsi:r.fisIsi!==null?r.fisIsi-newSysIsi:null,selKos:r.fisKos!==null?r.fisKos-newSysKos:null};
  })}:null);
  toast("✓ Stok disesuaikan!");
}

return <div>
<Card>
<div style={{fontWeight:700,color:C.gl2,marginBottom:12,fontSize:13}}>🔍 Opname Harian — Input Stok Fisik</div>
<Inp label="Tanggal Opname" type="date" value={opDate} onChange={setOpDate} style={{maxWidth:200,marginBottom:12}}/>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(160px,1fr))",gap:10,marginBottom:12}}>
{SIZES.map(s=>[
  <Inp key={"oi"+s} label={"Fisik Isi "+s+" (Sistem: "+(((data.stock||{})[s]||0))+")"} type="number" value={opF["isi_"+s]} onChange={v=>setOpF(p=>({...p,["isi_"+s]:v}))} placeholder="kosongkan jika tidak opname"/>,
  <Inp key={"ok"+s} label={"Fisik Kosong "+s+" (Sistem: "+getKosong(data,s)+")"} type="number" value={opF["kos_"+s]} onChange={v=>setOpF(p=>({...p,["kos_"+s]:v}))} placeholder="kosongkan jika tidak opname"/>
])}
</div>
<Btn color="blue" onClick={hitungSelisih}>🔍 Hitung Selisih</Btn>
</Card>

{opResult&&<Card style={{border:"1px solid #2980B9"}}>
<div style={{fontWeight:700,color:C.blt,marginBottom:12,fontSize:13}}>📊 Hasil Opname — {fDs(opResult.tanggal)}</div>
{opResult.results.map(r=>{
  var adaSelisih=(r.selIsi!==null&&r.selIsi!==0)||(r.selKos!==null&&r.selKos!==0);
  return <div key={r.ukuran} style={{marginBottom:10,padding:12,background:C.nav,borderRadius:8,border:"1px solid "+(adaSelisih?C.rlt:C.glt)}}>
  <div style={{fontWeight:800,color:C.wht,marginBottom:8}}>📦 {r.ukuran}</div>
  <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:6,marginBottom:adaSelisih?8:4}}>
  {[["Isi Sistem",r.sysIsi,C.glt],["Isi Fisik",r.fisIsi!==null?r.fisIsi:"—",C.glt],["Kosong Sistem",r.sysKos,C.gl2],["Kosong Fisik",r.fisKos!==null?r.fisKos:"—",C.gl2]].map(x=><div key={x[0]} style={{background:C.bg,borderRadius:6,padding:"6px 8px",textAlign:"center"}}><div style={{fontSize:9,color:C.gl2}}>{x[0]}</div><div style={{fontSize:16,fontWeight:800,color:x[2]}}>{x[1]}</div></div>)}
  </div>
  {adaSelisih?<div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
    {r.selIsi!==null&&r.selIsi!==0&&<div style={{flex:1,minWidth:160,padding:"8px 12px",background:C.rdk,borderRadius:8,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
    <span style={{color:"white",fontSize:12,fontWeight:700}}>💨 Isi: {r.selIsi>0?"+":""}{r.selIsi} tab</span>
    <button onClick={()=>doAdjust(r.ukuran,"isi",r.selIsi)} style={{background:"white",border:"none",borderRadius:6,padding:"4px 10px",fontSize:11,fontWeight:700,color:"#DC2626",cursor:"pointer"}}>✓ Adjust</button>
    </div>}
    {r.selKos!==null&&r.selKos!==0&&<div style={{flex:1,minWidth:160,padding:"8px 12px",background:C.rdk,borderRadius:8,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
    <span style={{color:"white",fontSize:12,fontWeight:700}}>📦 Kosong: {r.selKos>0?"+":""}{r.selKos} tab</span>
    <button onClick={()=>doAdjust(r.ukuran,"kos",r.selKos)} style={{background:"white",border:"none",borderRadius:6,padding:"4px 10px",fontSize:11,fontWeight:700,color:"#DC2626",cursor:"pointer"}}>✓ Adjust</button>
    </div>}
  </div>:<div style={{color:C.glt,fontSize:12,fontWeight:700}}>✅ Stok sesuai — tidak ada selisih</div>}
  </div>;
})}
</Card>}

<Card>
<div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>📋 Log Opname</div>
<RTbl headers={["Tgl","Ukuran","Keterangan","Qty","User"]} rows={(data.stockLog||[]).filter(l=>l.sumber==="Opname").slice(0,30).map(l=>[fDs(l.tanggal),l.ukuran,l.jenis,l.qty,l.user||"-"])}/>
</Card>
</div>;
}
function TitipTab(){
var blkI={ukuran:"12 kg",qty:""};
var[f,setF]=useState({tanggal:toDay(),tipe:"titip",konsumenNama:"",konsumenTelp:"",konsumenAlamat:"",salesId:"",items:[{...blkI}],ket:""});
var[delId,setDelId]=useState(null);
var[tf,setTf]=useState({from:"",to:"",tipe:"",ukuran:""});
var kNames=[...new Set([...(data.pelanggan||[]).filter(p=>PLG_TITIP_KAT.includes(p.kategori)).map(p=>p.nama),...(data.titipList||[]).map(t=>t.konsumenNama)].filter(Boolean))];
function onKons(nama){var p=(data.pelanggan||[]).find(x=>x.nama===nama);if(p)setF(pv=>({...pv,konsumenNama:nama,konsumenTelp:p.telepon||"",konsumenAlamat:p.alamat||""}));}
function setItem(i,k,v){setF(p=>{var it=p.items.slice();it[i]={...it[i],[k]:v};return{...p,items:it};});}
var validItems=f.items.filter(it=>Number(it.qty)>0);
function save(cetak){if(!f.konsumenNama||!validItems.length)return;var emp=(data.employees||[]).find(e=>e.id===f.salesId);var noBA="BA-"+Date.now().toString(36).toUpperCase().slice(-6);var rec={id:uid(),noBA,tanggal:f.tanggal,tipe:f.tipe,konsumenNama:f.konsumenNama,konsumenTelp:f.konsumenTelp,konsumenAlamat:f.konsumenAlamat,salesId:f.salesId,salesNama:emp?.nama||"",items:validItems.map(it=>({ukuran:it.ukuran,qty:Number(it.qty)})),ket:f.ket};setData(d=>({...d,titipList:[rec,...(d.titipList||[])]}));if(cetak)setBa(rec);else toast("✓ Dicatat!");setF(p=>({...p,konsumenNama:"",konsumenTelp:"",konsumenAlamat:"",items:[{...blkI}],ket:""}));}
var balMap=getKonsumenTitipBal(data.titipList);var aktifK=Object.keys(balMap).filter(k=>SIZES.some(s=>(balMap[k][s]||0)>0));

// Rekap saldo per konsumen — diurutkan per sales pengantar
var rekapTitip=[];
aktifK.forEach(function(k){
  var b=balMap[k];
  // Cari sales pengantar dari transaksi terakhir
  var lastTitip=(data.titipList||[]).filter(t=>t.konsumenNama===k&&t.tipe==="titip").slice(-1)[0];
  var salesNama=lastTitip?.salesNama||lastTitip?.salesPengantar||"(Tanpa Sales)";
  var total=SIZES.reduce(function(a,s){return a+(b[s]||0);},0);
  rekapTitip.push({konsumen:k,salesNama,s55:b["5.5 kg"]||0,s12:b["12 kg"]||0,s50:b["50 kg"]||0,total,alamat:b.alamat||""});
});
rekapTitip.sort(function(a,b){return a.salesNama.localeCompare(b.salesNama)||a.konsumen.localeCompare(b.konsumen);});

return <div>
<div style={{display:"flex",gap:10,flexWrap:"wrap",marginBottom:14}}><SC label="Konsumen Titip" value={aktifK.length} icon="🏪" color={C.blt}/>{SIZES.map(s=><SC key={s} label={"Titip "+s} value={getTitipTotal(data.titipList,s)+" tab"} icon="📦" color={C.blt}/>)}</div>

{/* Rekap Saldo + Cetak */}
<Card style={{marginBottom:12}}>
<div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
<div style={{fontWeight:700,color:C.gl2,fontSize:13}}>📊 Rekap Saldo Titip per Konsumen</div>
<button onClick={()=>{
var el=document.getElementById("_lap_titip");if(el)el.remove();
var printDiv=document.createElement("div");printDiv.id="_lap_titip";
var tot12=rekapTitip.reduce(function(a,r){return a+r.s12;},0);
var tot55=rekapTitip.reduce(function(a,r){return a+r.s55;},0);
var tot50=rekapTitip.reduce(function(a,r){return a+r.s50;},0);
var totAll=tot12+tot55+tot50;
var html='<div style="font-family:Arial,sans-serif;padding:20px;color:#111;max-width:900px;margin:0 auto">'+
'<div style="text-align:center;border-bottom:3px solid #0a1f44;padding-bottom:10px;margin-bottom:14px">'+
'<div style="font-size:18px;font-weight:900;color:#0a1f44">LAPORAN TABUNG TITIP DI KONSUMEN</div>'+
'<div style="font-size:13px;font-weight:700;color:#0a1f44">'+(data.company?.nama||"PT. HOE TRANG SA").toUpperCase()+'</div>'+
'<div style="font-size:11px;color:#555;margin-top:4px">Hanya saldo aktif (>0) | Dicetak: '+new Date().toLocaleString("id-ID")+'</div>'+
'</div>'+
'<table style="width:100%;border-collapse:collapse;font-size:11px">'+
'<thead><tr style="background:#0a1f44"><th style="color:white;padding:6px 8px;text-align:center;border:1px solid #ccc;width:30px">No</th><th style="color:white;padding:6px 8px;text-align:left;border:1px solid #ccc">Konsumen</th><th style="color:white;padding:6px 8px;text-align:left;border:1px solid #ccc">Sales</th><th style="color:white;padding:6px 8px;text-align:center;border:1px solid #ccc">5,5 kg</th><th style="color:white;padding:6px 8px;text-align:center;border:1px solid #ccc">12 kg</th><th style="color:white;padding:6px 8px;text-align:center;border:1px solid #ccc">50 kg</th><th style="color:white;padding:6px 8px;text-align:center;border:1px solid #ccc;font-weight:900">Total</th><th style="color:white;padding:6px 8px;text-align:left;border:1px solid #ccc">Alamat</th></tr></thead>'+
'<tbody>';
rekapTitip.forEach(function(r,i){
html+='<tr style="background:'+(i%2===0?'white':'#f9f9f9')+'">'+
'<td style="padding:5px 8px;border:1px solid #ddd;text-align:center">'+(i+1)+'</td>'+
'<td style="padding:5px 8px;border:1px solid #ddd;font-weight:600">'+r.konsumen+'</td>'+
'<td style="padding:5px 8px;border:1px solid #ddd">'+r.salesNama+'</td>'+
'<td style="padding:5px 8px;border:1px solid #ddd;text-align:center;font-weight:'+(r.s55>0?'700':'400')+';color:'+(r.s55>0?'#15803D':'#aaa')+'">'+(r.s55||'—')+'</td>'+
'<td style="padding:5px 8px;border:1px solid #ddd;text-align:center;font-weight:'+(r.s12>0?'700':'400')+';color:'+(r.s12>0?'#1D4ED8':'#aaa')+'">'+(r.s12||'—')+'</td>'+
'<td style="padding:5px 8px;border:1px solid #ddd;text-align:center;font-weight:'+(r.s50>0?'700':'400')+';color:'+(r.s50>0?'#D97706':'#aaa')+'">'+(r.s50||'—')+'</td>'+
'<td style="padding:5px 8px;border:1px solid #ddd;text-align:center;font-weight:900;color:#0a1f44">'+r.total+'</td>'+
'<td style="padding:5px 8px;border:1px solid #ddd;font-size:10px;color:#666">'+r.alamat+'</td>'+
'</tr>';});
html+='<tr style="background:#0a1f44;color:white;font-weight:900">'+
'<td colspan="3" style="padding:6px 8px;border:1px solid #ccc">TOTAL ('+rekapTitip.length+' konsumen)</td>'+
'<td style="padding:6px 8px;border:1px solid #ccc;text-align:center">'+tot55+'</td>'+
'<td style="padding:6px 8px;border:1px solid #ccc;text-align:center">'+tot12+'</td>'+
'<td style="padding:6px 8px;border:1px solid #ccc;text-align:center">'+tot50+'</td>'+
'<td style="padding:6px 8px;border:1px solid #ccc;text-align:center">'+totAll+'</td>'+
'<td style="border:1px solid #ccc"></td></tr>'+
'</tbody></table>'+
'<div style="margin-top:16px;font-size:10px;color:#555;text-align:right">'+data.company?.nama+' — '+data.company?.telepon+'</div>'+
'</div>';
printDiv.innerHTML=html;document.body.appendChild(printDiv);
doPrint("_lap_titip");
setTimeout(function(){var e=document.getElementById("_lap_titip");if(e)e.remove();},3000);
}} style={{background:"#0a1f44",color:"white",border:"none",padding:"7px 14px",borderRadius:7,fontSize:12,cursor:"pointer",fontWeight:700}}>🖨️ Cetak Laporan</button>
</div>
<div style={{overflowX:"auto"}}>
<table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
<thead><tr style={{background:C.nav}}>
{["No","Konsumen","Sales","5,5 kg","12 kg","50 kg","Total","Alamat"].map(h=><th key={h} style={{padding:"6px 8px",color:C.gl2,fontWeight:700,textAlign:["12 kg","5,5 kg","50 kg","Total"].includes(h)?"center":"left",borderBottom:"2px solid "+C.bdr,fontSize:11}}>{h}</th>)}
</tr></thead>
<tbody>
{rekapTitip.map((r,i)=><tr key={r.konsumen} style={{borderBottom:"1px solid "+C.bdr,background:i%2===0?C.nav:C.bg}}>
<td style={{padding:"5px 8px",color:C.gl2,fontSize:11}}>{i+1}</td>
<td style={{padding:"5px 8px",fontWeight:700,color:C.wht}}>{r.konsumen}</td>
<td style={{padding:"5px 8px",color:C.gl2,fontSize:11}}>{r.salesNama}</td>
<td style={{padding:"5px 8px",textAlign:"center",fontWeight:r.s55>0?700:400,color:r.s55>0?C.glt:C.gl2}}>{r.s55||"—"}</td>
<td style={{padding:"5px 8px",textAlign:"center",fontWeight:r.s12>0?700:400,color:r.s12>0?C.blt:C.gl2}}>{r.s12||"—"}</td>
<td style={{padding:"5px 8px",textAlign:"center",fontWeight:r.s50>0?700:400,color:r.s50>0?C.olt:C.gl2}}>{r.s50||"—"}</td>
<td style={{padding:"5px 8px",textAlign:"center",fontWeight:900,color:C.wht,fontSize:13}}>{r.total}</td>
<td style={{padding:"5px 8px",color:C.gl2,fontSize:10}}>{r.alamat||"—"}</td>
</tr>)}
<tr style={{background:C.nav,borderTop:"2px solid "+C.bdr}}>
<td colSpan={3} style={{padding:"6px 8px",fontWeight:700,color:C.wht}}>TOTAL ({rekapTitip.length} konsumen)</td>
<td style={{padding:"6px 8px",textAlign:"center",fontWeight:900,color:C.glt}}>{rekapTitip.reduce((a,r)=>a+r.s55,0)}</td>
<td style={{padding:"6px 8px",textAlign:"center",fontWeight:900,color:C.blt}}>{rekapTitip.reduce((a,r)=>a+r.s12,0)}</td>
<td style={{padding:"6px 8px",textAlign:"center",fontWeight:900,color:C.olt}}>{rekapTitip.reduce((a,r)=>a+r.s50,0)}</td>
<td style={{padding:"6px 8px",textAlign:"center",fontWeight:900,color:C.wht,fontSize:14}}>{rekapTitip.reduce((a,r)=>a+r.total,0)}</td>
<td></td>
</tr>
</tbody>
</table>
</div>
</Card>
<Card>
<div style={{display:"flex",gap:6,marginBottom:12}}>{[["titip","📦 Titip ke Konsumen",C.grn],["tarik","↩️ Tarik dari Konsumen",C.rdk],["titip_luar","🏭 Titipan Pihak Lain Masuk",C.blt],["tarik_luar","📤 Titipan Pihak Lain Keluar","#6B7280"]].map(x=><button key={x[0]} onClick={()=>setF(p=>({...p,tipe:x[0]}))} style={{background:f.tipe===x[0]?x[2]:C.nav,color:f.tipe===x[0]?"white":C.wht,border:"1px solid "+(f.tipe===x[0]?x[2]:C.bdr),borderRadius:8,padding:"8px 16px",fontWeight:700,fontSize:13,cursor:"pointer",flex:1}}>{x[1]}</button>)}</div>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(160px,1fr))",gap:10}}>
<Inp label="Tanggal" type="date" value={f.tanggal} onChange={v=>setF(p=>({...p,tanggal:v}))}/>
<Sel label="Sales" value={f.salesId} onChange={v=>setF(p=>({...p,salesId:v}))} opts={[{v:"",l:"-- Pilih --"},...(data.employees||[]).filter(e=>e.aktif).map(e=>({v:e.id,l:e.nama}))]}/>
<div style={{gridColumn:"1/-1"}}><AutoInp label="Konsumen" value={f.konsumenNama} onChange={v=>setF(p=>({...p,konsumenNama:v}))} options={kNames} placeholder="Ketik..." onSelect={onKons}/></div>
<Inp label="No. HP" value={f.konsumenTelp} onChange={v=>setF(p=>({...p,konsumenTelp:v}))}/>
<Inp label="Alamat" value={f.konsumenAlamat} onChange={v=>setF(p=>({...p,konsumenAlamat:v}))}/>
</div>
<div style={{border:"1px solid "+C.bdr,borderRadius:8,overflow:"hidden",marginBottom:10}}>
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr 36px",background:C.nav,padding:"6px 10px",fontSize:11,color:C.gl2,fontWeight:700,gap:6}}><span>Ukuran</span><span>Qty</span><span/></div>
{f.items.map((it,i)=><div key={i} style={{display:"grid",gridTemplateColumns:"1fr 1fr 36px",padding:"5px 10px",borderTop:"1px solid "+C.bdr,gap:6}}>
<select value={it.ukuran} onChange={e=>setItem(i,"ukuran",e.target.value)} style={{background:C.nav,border:"1px solid "+C.bdr,borderRadius:6,padding:"5px 6px",color:C.wht,fontSize:12,outline:"none"}}>{SIZES.map(s=><option key={s}>{s}</option>)}</select>
<input type="number" value={it.qty} placeholder="0" onChange={e=>setItem(i,"qty",e.target.value)} style={{background:C.nav,border:"1px solid "+C.bdr,borderRadius:6,padding:"5px 6px",color:C.wht,fontSize:12,outline:"none"}}/>
<button onClick={()=>setF(p=>({...p,items:p.items.filter((_,j)=>j!==i)}))} disabled={f.items.length<=1} style={{background:C.inHvE,border:"none",borderRadius:5,color:C.rlt,cursor:"pointer",fontSize:13}}>✕</button>
</div>)}
<div style={{padding:"6px 10px",background:C.nav,borderTop:"1px solid "+C.bdr}}><Btn sm color="blue" onClick={()=>setF(p=>({...p,items:[...p.items,{...blkI}]}))}>+ Item</Btn></div>
</div>
<Inp label="Keterangan" value={f.ket} onChange={v=>setF(p=>({...p,ket:v}))}/>
<div style={{display:"flex",gap:8,flexWrap:"wrap"}}><Btn color="green" onClick={()=>save(false)} dis={!f.konsumenNama||!validItems.length}>💾 Simpan</Btn><Btn color="blue" onClick={()=>save(true)} dis={!f.konsumenNama||!validItems.length}>🖨️ Simpan & Cetak BA</Btn></div>
</Card>
<Card>
<div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
<div style={{fontWeight:700,color:C.gl2,fontSize:13}}>📋 Riwayat Titip/Tarik</div>
</div>
{(()=>{
var titipFilt=(data.titipList||[]).filter(t=>{
if(tf.from&&(t.tanggal||"")<tf.from)return false;
if(tf.to&&(t.tanggal||"")>tf.to)return false;
if(tf.tipe&&t.tipe!==tf.tipe)return false;
if(tf.ukuran){var items=t.items&&t.items.length>0?t.items:[{ukuran:t.ukuran}];if(!items.some(i=>i.ukuran===tf.ukuran))return false;}
return true;
});
return <>
<div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:10,background:C.nav,padding:8,borderRadius:8,border:"1px solid "+C.bdr}}>
<Inp label="Dari" type="date" value={tf.from} onChange={v=>setTf(p=>({...p,from:v}))} style={{marginBottom:0,minWidth:130}}/>
<Inp label="Sampai" type="date" value={tf.to} onChange={v=>setTf(p=>({...p,to:v}))} style={{marginBottom:0,minWidth:130}}/>
<Sel label="Tipe" value={tf.tipe} onChange={v=>setTf(p=>({...p,tipe:v}))} opts={[{v:"",l:"Semua"},{v:"titip",l:"Titip"},{v:"tarik",l:"Tarik"}]}/>
<Sel label="Ukuran" value={tf.ukuran} onChange={v=>setTf(p=>({...p,ukuran:v}))} opts={[{v:"",l:"Semua Ukuran"},{v:"5.5 kg",l:"5,5 kg"},{v:"12 kg",l:"12 kg"},{v:"50 kg",l:"50 kg"}]}/>
{(tf.from||tf.to||tf.tipe||tf.ukuran)&&<Btn sm color="gray" onClick={()=>setTf({from:"",to:"",tipe:"",ukuran:""})}>✕ Reset</Btn>}
</div>
<RTbl headers={["Tgl","Tipe","Konsumen","Sales","Ukuran","Qty","Aksi"]} rows={titipFilt.slice(0,100).map(t=>{
var items=t.items&&t.items.length>0?t.items:[{ukuran:t.ukuran,qty:t.qty}];
var validItems=items.filter(i=>i.ukuran&&i.qty);
var m=t.tipe==="titip"?1:-1;
return[
fDs(t.tanggal),
<Bdg color={t.tipe==="titip"?"green":t.tipe==="tarik"?"red":"blue"}>{t.tipe}</Bdg>,
<b style={{color:C.wht}}>{t.konsumenNama}</b>,
t.salesPengantar||t.salesNama||"-",
<div style={{display:"flex",flexDirection:"column",gap:2}}>{validItems.map((it,i)=><Bdg key={i} color={it.ukuran==="12 kg"?"blue":it.ukuran==="5.5 kg"?"green":"orange"}>{it.ukuran}</Bdg>)}</div>,
<div style={{display:"flex",flexDirection:"column",gap:2}}>{validItems.map((it,i)=><b key={i} style={{color:m>0?C.glt:C.rlt}}>{m>0?"+":"-"}{it.qty}</b>)}</div>,
<div style={{display:"flex",gap:4}}><button onClick={()=>setBa(t)} style={{background:C.inHv,border:"1px solid "+C.blt,borderRadius:6,padding:"4px 7px",color:C.blt,cursor:"pointer",fontSize:12}}>🖨️</button><ActBtns onDel={()=>setDelId(t)}/></div>
];})}/>
</>;
})()}
</Card>
{delId&&<ConfirmDel msg="Hapus?" onCancel={()=>setDelId(null)} onConfirm={()=>{setData(d=>({...d,titipList:(d.titipList||[]).filter(x=>x.id!==delId.id)}));setDelId(null);}}/>}
</div>;
}
return <div>
<STitle icon="📦" children="Stok & Mutasi"/>
<div style={{display:"flex",gap:6,marginBottom:14,flexWrap:"wrap"}}>{[["rekap","📊 Rekap"],["mutasi","✏️ Mutasi Manual"],["opname","🔍 Opname"],["laporan","📋 Lap. Harian"],["titip","🏪 Titip/Tarik"]].map(x=><button key={x[0]} onClick={()=>setTab(x[0])} style={{background:tab===x[0]?C.blu:C.nav,color:tab===x[0]?"white":C.wht,border:"1px solid "+(tab===x[0]?C.blt:C.bdr),borderRadius:8,padding:"7px 14px",fontWeight:700,fontSize:12,cursor:"pointer"}}>{x[1]}</button>)}</div>
{tab==="rekap"&&<RekapTab/>}
{tab==="mutasi"&&<MutasiTab/>}
{tab==="opname"&&<OpnameTab/>}
{tab==="laporan"&&<div>
<Card>
<div style={{fontWeight:700,color:C.gl2,marginBottom:12,fontSize:13}}>📋 Laporan Stok Harian Bulanan</div>
<div style={{display:"flex",gap:10,alignItems:"flex-end",flexWrap:"wrap",marginBottom:14}}>
<Inp label="Pilih Bulan" type="month" value={stokBln} onChange={setStokBln} style={{maxWidth:180,marginBottom:0}}/>
<Btn sm color="orange" onClick={()=>setShowInject(!showInject)}>{showInject?"✕ Tutup":"★ Inject Stok Awal"}</Btn>
</div>
{showInject&&<div style={{background:C.nav,borderRadius:8,padding:12,border:"1px solid #F59E0B",marginBottom:14}}>
<div style={{fontWeight:700,color:"#F59E0B",marginBottom:10,fontSize:12}}>★ Set Stok Awal Manual — Titik Awal Perhitungan</div>
<div style={{fontSize:11,color:C.gl2,marginBottom:10}}>Input stok pada tanggal tertentu sebagai titik awal. Sistem akan hitung otomatis hari berikutnya.</div>
<div style={{display:"flex",gap:8,alignItems:"flex-end",flexWrap:"wrap",marginBottom:10}}>
<Inp label="Tanggal Awal" type="date" value={injectTgl} onChange={setInjectTgl} style={{maxWidth:160,marginBottom:0}}/>
</div>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(130px,1fr))",gap:8,marginBottom:10}}>
{SIZES.map(s=>[
<Inp key={"ii"+s} label={"Isi "+s} type="number" value={injectF["isi_"+s]||""} onChange={v=>setInjectF(p=>({...p,["isi_"+s]:v}))} placeholder="0"/>,
<Inp key={"ik"+s} label={"TK "+s+" (Kosong)"} type="number" value={injectF["tk_"+s]||""} onChange={v=>setInjectF(p=>({...p,["tk_"+s]:v}))} placeholder="0"/>
])}
</div>
<div style={{display:"flex",gap:8}}>
<Btn color="orange" onClick={()=>{
var rec={tanggal:injectTgl};
SIZES.forEach(s=>{rec["isi_"+s]=Number(injectF["isi_"+s]||0);rec["tk_"+s]=Number(injectF["tk_"+s]||0);});
var existing=(data.stokHarian||[]).filter(r=>r.tanggal!==injectTgl);
setData(d=>({...d,stokHarian:[rec,...existing]}));
setShowInject(false);setInjectF({});
toast("✓ Stok awal "+fDs(injectTgl)+" berhasil disimpan! Tabel akan dihitung ulang dari tanggal ini.");
}}>💾 Simpan Stok Awal</Btn>
<Btn sm color="gray" onClick={()=>{setShowInject(false);setInjectF({});}}>Batal</Btn>
</div>
{(data.stokHarian||[]).length>0&&<div style={{marginTop:10}}>
<div style={{fontSize:11,color:C.gl2,marginBottom:6,fontWeight:700}}>Inject yang tersimpan:</div>
{(data.stokHarian||[]).slice().sort((a,b)=>b.tanggal.localeCompare(a.tanggal)).map(r=><div key={r.tanggal} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"6px 10px",background:C.bg,borderRadius:6,marginBottom:4,fontSize:11}}>
<div><b style={{color:"#F59E0B"}}>★ {fDs(r.tanggal)}</b><span style={{color:C.gl2,marginLeft:8}}>{SIZES.map(s=>"isi "+s+": "+r["isi_"+s]+", TK: "+r["tk_"+s]).join(" | ")}</span></div>
<button onClick={()=>{var existing=(data.stokHarian||[]).filter(x=>x.tanggal!==r.tanggal);setData(d=>({...d,stokHarian:existing}));}} style={{background:C.rdk,border:"none",borderRadius:5,padding:"2px 8px",color:"white",cursor:"pointer",fontSize:11}}>✕</button>
</div>)}
</div>}
</div>}
<TabelStokBulanan data={data} bulan={stokBln}/>
</Card>
</div>}
{tab==="titip"&&(()=>{try{return <TitipTab/>;}catch(e){return <Card><div style={{color:C.rlt,padding:12}}>Error: {e.message}</div></Card>;}})()}
{ba&&<BeritaAcaraView ba={ba} company={data.company} onClose={()=>setBa(null)}/>}
</div>;
}

// ─── DO v4 (HPP/unit + Total HPP, push ke modalHistory) ───────────────────────
function DOMod({data,setData,user,toast}){
var C=useTheme();
var[f,setF]=useState(()=>{var hpp=getModal(data,"12 kg","Isi",toDay());return{tanggal:toDay(),trip:"Trip 1",sppbe:"SPPBE KCR",ukuran:"12 kg",qty:"",hppUnit:String(hpp||""),ket:"",dibuatOleh:user?.id||""};});
var[delId,setDelId]=useState(null);
var[editId,setEditId]=useState(null);// id DO yang sedang diedit (hanya status "gantung")
var[bulanRiwayat,setBulanRiwayat]=useState(toMonth());
var formRef=useRef(null);
function onUkuran(v){var hpp=getModal(data,v,"Isi",f.tanggal);setF(p=>({...p,ukuran:v,hppUnit:String(hpp||"")}));}
function onTanggal(v){var hpp=getModal(data,f.ukuran,"Isi",v);setF(p=>({...p,tanggal:v,hppUnit:String(hpp||"")}));}
var totalHPP=Number(f.qty||0)*Number(f.hppUnit||0);
function resetForm(){var hpp=getModal(data,"12 kg","Isi",toDay());setF({tanggal:toDay(),trip:"Trip 1",sppbe:"SPPBE KCR",ukuran:"12 kg",qty:"",hppUnit:String(hpp||""),ket:"",dibuatOleh:user?.id||""});setEditId(null);}
function mulaiEdit(d_rec){
if(d_rec.status!=="gantung"){toast("⚠️ DO yang sudah Diterima/Sangkut tidak bisa diedit — stok sudah bergerak. Hapus lalu buat ulang jika perlu.","error");return;}
var empDO=(data.employees||[]).find(e=>e.nama===d_rec.dibuatOleh);
setF({tanggal:d_rec.tanggal,trip:d_rec.trip||d_rec.noDO||"Trip 1",sppbe:d_rec.sppbe,ukuran:d_rec.ukuran,qty:String(d_rec.qty||""),hppUnit:String(d_rec.hppUnit||""),ket:d_rec.ket||"",dibuatOleh:empDO?.id||""});
setEditId(d_rec.id);
setTimeout(()=>{if(formRef.current)formRef.current.scrollIntoView({behavior:"smooth",block:"start"});},100);
toast("✏️ Mode edit DO "+(d_rec.trip||"")+" — ubah data lalu klik Simpan Perubahan");
}
function save(){
if(!f.qty||!f.trip)return;
var qty=Number(f.qty);var hpp=Number(f.hppUnit||0);
var empDO=(data.employees||[]).find(e=>e.id===f.dibuatOleh);
if(editId){
  // Edit DO berstatus gantung — replace record, tidak menyentuh stok/batch (karena belum diterima)
  var updated={id:editId,tanggal:f.tanggal,trip:f.trip,noDO:f.trip,sppbe:f.sppbe,ukuran:f.ukuran,qty,hppUnit:hpp,totalHPP:qty*hpp,ket:f.ket,dibuatOleh:empDO?.nama||user?.nama||"",status:"gantung"};
  setData(d=>({...d,doList:(d.doList||[]).map(x=>x.id===editId?updated:x)}));
  resetForm();
  toast("✓ DO diperbarui! Status tetap Gantung.");
  return;
}
var doRec={id:uid(),tanggal:f.tanggal,trip:f.trip,noDO:f.trip,sppbe:f.sppbe,ukuran:f.ukuran,qty,hppUnit:hpp,totalHPP:qty*hpp,ket:f.ket,dibuatOleh:empDO?.nama||user?.nama||"",status:"gantung"};
var newModalHistory=hpp>0?[{id:uid(),tanggal:f.tanggal,ukuran:f.ukuran,jenis:"Isi",harga:hpp,sumber:"DO "+f.trip},...(data.modalHistory||[])]:(data.modalHistory||[]);
// Stok TIDAK langsung masuk — harus klik Diterima
setData(d=>({...d,doList:[doRec,...(d.doList||[])],modalHistory:newModalHistory}));
setF(p=>({...p,qty:"",ket:""}));
toast("✓ DO dibuat! Status: Gantung. Klik '✅ Diterima' setelah barang tiba di gudang.");
}

// Terima DO → stok masuk (+ push batch FIFO dengan harga HPP DO ini)
function terimaDO(d_rec){
var qty=Number(d_rec.qty||0);var uk=d_rec.ukuran;
var ns={...(data.stock||{})};ns[uk]=(ns[uk]||0)+qty;
var log={id:uid(),tanggal:d_rec.tanggal,ukuran:uk,jenis:"Isi Ulang SPPBE",qty,ket:"DO "+d_rec.trip+" - "+d_rec.sppbe+" (Diterima)",sumber:"DO",user:user?.nama||""};
setData(d=>{
  var withBatch=addBatch(d,uk,qty,d_rec.hppUnit||0,"DO "+d_rec.trip+" ("+d_rec.sppbe+")",d_rec.tanggal);
  return{...withBatch,
    doList:(d.doList||[]).map(x=>x.id===d_rec.id?{...x,status:"diterima",tglDiterima:toDay()}:x),
    stock:ns,
    stockLog:[log,...(d.stockLog||[])].slice(0,500)
  };
});
toast("✅ DO diterima! Stok "+uk+" +"+qty+" tabung masuk gudang @ Rp "+fR(d_rec.hppUnit||0));
}

// Tandai Sangkut
function sangkutDO(d_rec){
setData(d=>({...d,doList:(d.doList||[]).map(x=>x.id===d_rec.id?{...x,status:"sangkut"}:x)}));
toast("⚠️ DO ditandai Sangkut.");
}

// Release Sangkut → stok masuk (+ push batch FIFO)
function releaseDO(d_rec){
var qty=Number(d_rec.qty||0);var uk=d_rec.ukuran;
var ns={...(data.stock||{})};ns[uk]=(ns[uk]||0)+qty;
var log={id:uid(),tanggal:toDay(),ukuran:uk,jenis:"Isi Ulang SPPBE",qty,ket:"DO "+d_rec.trip+" Release Sangkut",sumber:"DO",user:user?.nama||""};
setData(d=>{
  var withBatch=addBatch(d,uk,qty,d_rec.hppUnit||0,"DO "+d_rec.trip+" (Release Sangkut)",toDay());
  return{...withBatch,
    doList:(d.doList||[]).map(x=>x.id===d_rec.id?{...x,status:"diterima",tglDiterima:toDay()}:x),
    stock:ns,
    stockLog:[log,...(d.stockLog||[])].slice(0,500)
  };
});
toast("✅ DO di-release! Stok "+uk+" +"+qty+" tabung masuk.");
}

// ═══ DO TRIP (FORMAT BARU — 1 trip, multi produk) ═══════════════════════════
var blankItem=()=>({ukuran:"5.5 kg",qty:"",hppUnit:""});
var[ft,setFt]=useState(()=>({tanggal:toDay(),trip:"Trip 1",sppbe:"SPPBE KCR",dibuatOleh:user?.id||"",ket:"",
  items:{"5.5 kg":{qty:"",hppUnit:String(data.company?.hppFixedSPPBE?.["SPPBE KCR"]?.["5.5 kg"]||"")},
         "12 kg":{qty:"",hppUnit:String(data.company?.hppFixedSPPBE?.["SPPBE KCR"]?.["12 kg"]||"")},
         "50 kg":{qty:"",hppUnit:String(data.company?.hppFixedSPPBE?.["SPPBE KCR"]?.["50 kg"]||"")}}}));
var[editTripId,setEditTripId]=useState(null);
var[editItemDO,setEditItemDO]=useState(null);// {trip, item, newQty}
var[editItemQty,setEditItemQty]=useState("");
function mulaiEditItemDO(trip,item){
  setEditItemDO({trip,item});
  setEditItemQty(String(item.qty));
  toast("✏️ Edit qty "+item.ukuran+" — masukkan qty baru lalu klik Simpan");
}
function simpanEditItemDO(){
  if(!editItemDO)return;
  var{trip,item}=editItemDO;
  var qtyBaru=Number(editItemQty)||0;
  if(qtyBaru<=0){toast("⚠️ Qty harus > 0","error");return;}
  if(qtyBaru===item.qty){setEditItemDO(null);setEditItemQty("");return;}
  var selisih=qtyBaru-item.qty;// positif = tambah, negatif = kurang
  var sumberBatch="DO "+trip.trip+" ("+trip.sppbe+")";
  setData(d=>{
    var ns={...(d.stock||{})};var nk={...(d.stokKosong||{})};
    var dCur=d;
    if(selisih>0){
      // Tambah qty: +stok, +addBatch
      ns[item.ukuran]=(ns[item.ukuran]||0)+selisih;
      nk[item.ukuran]=Math.max(0,(nk[item.ukuran]||0)-selisih);
      dCur=addBatch(dCur,item.ukuran,selisih,item.hppUnit||0,sumberBatch,trip.tanggal);
    } else {
      // Kurangi qty: -stok, consumeFIFO dari kepala antrian
      var qKurang=Math.abs(selisih);
      ns[item.ukuran]=Math.max(0,(ns[item.ukuran]||0)-qKurang);
      nk[item.ukuran]=(nk[item.ukuran]||0)+qKurang;
      var res=consumeFIFO(dCur,item.ukuran,qKurang);
      dCur=res.data;
    }
    var log={id:uid(),tanggal:trip.tanggal,ukuran:item.ukuran,jenis:"Edit Qty DO "+trip.trip,qty:Math.abs(selisih),ket:"Qty diubah "+item.qty+"→"+qtyBaru+(selisih>0?" (+tambah)":" (-kurang)"),sumber:"Edit DO",user:user?.nama||""};
    var updatedTrip={...trip,items:trip.items.map(it=>it.id===item.id?{...it,qty:qtyBaru,totalHPP:qtyBaru*(it.hppUnit||0)}:it)};
    return{...dCur,doTrip:(d.doTrip||[]).map(x=>x.id===trip.id?updatedTrip:x),stock:ns,stokKosong:nk,stockLog:[log,...(d.stockLog||[])].slice(0,500)};
  });
  setEditItemDO(null);setEditItemQty("");
  toast("✓ Qty "+item.ukuran+" diperbarui: "+item.qty+" → "+qtyBaru);
}
var isFixedSPPBE=ft.sppbe==="SPPBE KCR"||ft.sppbe==="SPPBE MGL";
function onSppbeT(v){
var fixed=data.company?.hppFixedSPPBE?.[v]||null;
setFt(p=>({...p,sppbe:v,items:Object.fromEntries(SIZES.map(s=>[s,{qty:p.items[s].qty,hppUnit:fixed?String(fixed[s]||0):p.items[s].hppUnit}]))}));
}
function setItemT(uk,field,val){setFt(p=>({...p,items:{...p.items,[uk]:{...p.items[uk],[field]:val}}}));}
var totalNilaiTrip=SIZES.reduce((a,s)=>a+(Number(ft.items[s].qty)||0)*(Number(ft.items[s].hppUnit)||0),0);
function resetFormT(){
var fixed=data.company?.hppFixedSPPBE?.["SPPBE KCR"]||null;
setFt({tanggal:toDay(),trip:"Trip 1",sppbe:"SPPBE KCR",dibuatOleh:user?.id||"",ket:"",
  items:Object.fromEntries(SIZES.map(s=>[s,{qty:"",hppUnit:fixed?String(fixed[s]||0):""}]))});
setEditTripId(null);
}
function mulaiEditTrip(trip){
if(!trip.items.every(it=>it.status==="gantung")){toast("⚠️ Trip yang sudah ada item Diterima/Sangkut tidak bisa diedit penuh — hapus item tersebut satu per satu jika perlu.","error");return;}
var empDO=(data.employees||[]).find(e=>e.id===trip.dibuatOlehId)||(data.employees||[]).find(e=>e.nama===trip.dibuatOleh);
var itemsMap=Object.fromEntries(SIZES.map(s=>{var it=trip.items.find(x=>x.ukuran===s);return[s,{qty:it?String(it.qty):"",hppUnit:it?String(it.hppUnit):""}];}));
setFt({tanggal:trip.tanggal,trip:trip.trip,sppbe:trip.sppbe,dibuatOleh:empDO?.id||"",ket:trip.ket||"",items:itemsMap});
setEditTripId(trip.id);
setTimeout(()=>{if(formRef.current)formRef.current.scrollIntoView({behavior:"smooth",block:"start"});},100);
toast("✏️ Mode edit Trip "+(trip.trip||"")+" — ubah data lalu klik Simpan Perubahan");
}
function saveTrip(){
var activeSizes=SIZES.filter(s=>Number(ft.items[s].qty)>0);
if(!ft.trip||activeSizes.length===0){toast("⚠️ Isi minimal 1 qty produk","error");return;}
var empDO=(data.employees||[]).find(e=>e.id===ft.dibuatOleh);
var itemsArr=activeSizes.map(s=>({ukuran:s,qty:Number(ft.items[s].qty),hppUnit:Number(ft.items[s].hppUnit)||0,totalHPP:Number(ft.items[s].qty)*(Number(ft.items[s].hppUnit)||0),status:"gantung"}));
if(editTripId){
  setData(d=>({...d,doTrip:(d.doTrip||[]).map(x=>x.id===editTripId?{...x,tanggal:ft.tanggal,trip:ft.trip,sppbe:ft.sppbe,dibuatOleh:empDO?.nama||user?.nama||"",dibuatOlehId:ft.dibuatOleh,ket:ft.ket,items:itemsArr}:x)}));
  resetFormT();
  toast("✓ Trip DO diperbarui!");
  return;
}
var tripRec={id:uid(),tanggal:ft.tanggal,trip:ft.trip,sppbe:ft.sppbe,dibuatOleh:empDO?.nama||user?.nama||"",dibuatOlehId:ft.dibuatOleh,ket:ft.ket,items:itemsArr.map(it=>({...it,id:uid()}))};
// HPP fixed (KCR/MGL) ikut tercatat ke modalHistory sebagai referensi; HPP manual (Lainnya) juga dicatat
var newModalHistory=[...itemsArr.filter(it=>it.hppUnit>0).map(it=>({id:uid(),tanggal:ft.tanggal,ukuran:it.ukuran,jenis:"Isi",harga:it.hppUnit,sumber:"DO "+ft.trip+" ("+ft.sppbe+")"})),...(data.modalHistory||[])];
setData(d=>({...d,doTrip:[tripRec,...(d.doTrip||[])],modalHistory:newModalHistory}));
resetFormT();
toast("✓ Trip DO dibuat! "+itemsArr.length+" produk, status Gantung. Klik Diterima per produk setelah tiba.");
}
function terimaTripItem(trip,item){
var qty=Number(item.qty||0);var uk=item.ukuran;
var ns={...(data.stock||{})};ns[uk]=(ns[uk]||0)+qty;
var log={id:uid(),tanggal:trip.tanggal,ukuran:uk,jenis:"Isi Ulang SPPBE",qty,ket:"DO "+trip.trip+" - "+trip.sppbe+" (Diterima)",sumber:"DO Trip",user:user?.nama||""};
setData(d=>{
  var withBatch=addBatch(d,uk,qty,item.hppUnit||0,"DO "+trip.trip+" ("+trip.sppbe+")",trip.tanggal);
  return{...withBatch,
    doTrip:(d.doTrip||[]).map(x=>x.id!==trip.id?x:{...x,items:x.items.map(it=>it.id===item.id?{...it,status:"diterima",tglDiterima:toDay()}:it)}),
    stock:ns,
    stockLog:[log,...(d.stockLog||[])].slice(0,500)
  };
});
toast("✅ "+uk+" diterima! Stok +"+qty+" tabung @ "+fR(item.hppUnit||0));
}
function sangkutTripItem(trip,item){
setData(d=>({...d,doTrip:(d.doTrip||[]).map(x=>x.id!==trip.id?x:{...x,items:x.items.map(it=>it.id===item.id?{...it,status:"sangkut"}:it)})}));
toast("⚠️ "+item.ukuran+" ditandai Sangkut.");
}
function releaseTripItem(trip,item){
var qty=Number(item.qty||0);var uk=item.ukuran;
var ns={...(data.stock||{})};ns[uk]=(ns[uk]||0)+qty;
var log={id:uid(),tanggal:toDay(),ukuran:uk,jenis:"Isi Ulang SPPBE",qty,ket:"DO "+trip.trip+" Release Sangkut",sumber:"DO Trip",user:user?.nama||""};
setData(d=>{
  var withBatch=addBatch(d,uk,qty,item.hppUnit||0,"DO "+trip.trip+" (Release Sangkut)",toDay());
  return{...withBatch,
    doTrip:(d.doTrip||[]).map(x=>x.id!==trip.id?x:{...x,items:x.items.map(it=>it.id===item.id?{...it,status:"diterima",tglDiterima:toDay()}:it)}),
    stock:ns,
    stockLog:[log,...(d.stockLog||[])].slice(0,500)
  };
});
toast("✅ "+uk+" di-release! Stok +"+qty+" tabung.");
}
function hapusTripItem(trip,item){
var ns={...(data.stock||{})};var nk={...(data.stokKosong||{})};
if(item.status==="diterima"){
  ns[item.ukuran]=Math.max(0,(ns[item.ukuran]||0)-item.qty);
  nk[item.ukuran]=(nk[item.ukuran]||0)+item.qty;
}
var sisaItems=trip.items.filter(it=>it.id!==item.id);
var sumberBatch="DO "+trip.trip+" ("+trip.sppbe+")";
setData(d=>{
  var logs=item.status==="diterima"?[{id:uid(),tanggal:trip.tanggal,ukuran:item.ukuran,jenis:"Reverse Hapus Item DO "+(trip.trip||""),qty:item.qty,ket:"Stok dikembalikan karena item DO dihapus",sumber:"Hapus DO Trip",user:user?.nama||""},...(d.stockLog||[])].slice(0,500):(d.stockLog||[]);
  // Reverse FIFO batch jika item sudah diterima
  var dRev=item.status==="diterima"?reverseBatch(d,item.ukuran,item.qty,sumberBatch,item.hppUnit||0):d;
  if(sisaItems.length===0){
    return{...dRev,doTrip:(d.doTrip||[]).filter(x=>x.id!==trip.id),stock:ns,stokKosong:nk,stockLog:logs};
  }
  return{...dRev,doTrip:(d.doTrip||[]).map(x=>x.id===trip.id?{...x,items:sisaItems}:x),stock:ns,stokKosong:nk,stockLog:logs};
});
toast("✓ Item "+item.ukuran+" dihapus dari trip.");
}

// ── Gabungkan doTrip (baru) + doList (lama) jadi 1 list ternormalisasi, urut tanggal ──
var allDORows=useMemo(()=>{
var fromTrip=(data.doTrip||[]).map(trip=>({
  _src:"trip",_id:trip.id,_raw:trip,tanggal:trip.tanggal,trip:trip.trip,driver:trip.dibuatOleh,sppbe:trip.sppbe,
  per:{"5.5 kg":trip.items.find(it=>it.ukuran==="5.5 kg"),"12 kg":trip.items.find(it=>it.ukuran==="12 kg"),"50 kg":trip.items.find(it=>it.ukuran==="50 kg")},
  totalNilai:trip.items.reduce((a,it)=>a+(it.totalHPP||0),0)
}));
var fromLegacy=(data.doList||[]).map(d=>({
  _src:"legacy",_id:d.id,_raw:d,tanggal:d.tanggal,trip:(d.trip||d.noDO||"-")+" (lama)",driver:d.dibuatOleh,sppbe:d.sppbe,
  per:{"5.5 kg":d.ukuran==="5.5 kg"?{qty:d.qty,hppUnit:d.hppUnit,status:d.status||"diterima"}:null,
       "12 kg":d.ukuran==="12 kg"?{qty:d.qty,hppUnit:d.hppUnit,status:d.status||"diterima"}:null,
       "50 kg":d.ukuran==="50 kg"?{qty:d.qty,hppUnit:d.hppUnit,status:d.status||"diterima"}:null},
  totalNilai:d.totalHPP||0
}));
return[...fromTrip,...fromLegacy].sort((a,b)=>{
  if(a.tanggal!==b.tanggal)return(b.tanggal||"").localeCompare(a.tanggal||"");// terbaru di atas
  return(b._id||"").localeCompare(a._id||"");// tie-break stabil
});
},[data.doTrip,data.doList]);

var rowsBulanIni=allDORows.filter(r=>(r.tanggal||"").slice(0,7)===bulanRiwayat);
var bulanTersedia=useMemo(()=>{
var set=new Set(allDORows.map(r=>(r.tanggal||"").slice(0,7)).filter(Boolean));
set.add(bulanRiwayat);
return Array.from(set).sort().reverse();
},[allDORows,bulanRiwayat]);

function exportExcelDO(){
var wb=XLSX.utils.book_new();
var header=["Tanggal","Trip","Driver","SPPBE","Qty 5,5kg","Harga 5,5kg","Qty 12kg","Harga 12kg","Qty 50kg","Harga 50kg","Total Nilai DO","Status 5,5kg","Status 12kg","Status 50kg"];
var aoa=[header,...rowsBulanIni.map(r=>{
  var st=uk=>r.per[uk]?(r.per[uk].status==="gantung"?"Gantung":r.per[uk].status==="sangkut"?"Sangkut":"Diterima"):"-";
  return[fDs(r.tanggal),r.trip,r.driver||"-",r.sppbe,
    r.per["5.5 kg"]?r.per["5.5 kg"].qty:"",r.per["5.5 kg"]?r.per["5.5 kg"].hppUnit:"",
    r.per["12 kg"]?r.per["12 kg"].qty:"",r.per["12 kg"]?r.per["12 kg"].hppUnit:"",
    r.per["50 kg"]?r.per["50 kg"].qty:"",r.per["50 kg"]?r.per["50 kg"].hppUnit:"",
    r.totalNilai,st("5.5 kg"),st("12 kg"),st("50 kg")];
})];
var qtyTot={"5.5 kg":0,"12 kg":0,"50 kg":0};var hppTot={"5.5 kg":0,"12 kg":0,"50 kg":0};var grandTotal=0;
rowsBulanIni.forEach(r=>{SIZES.forEach(uk=>{if(r.per[uk]){qtyTot[uk]+=r.per[uk].qty||0;hppTot[uk]+=(r.per[uk].qty||0)*(r.per[uk].hppUnit||0);}});grandTotal+=r.totalNilai||0;});
aoa.push([]);
aoa.push(["","","","TOTAL",qtyTot["5.5 kg"],fR(hppTot["5.5 kg"]),qtyTot["12 kg"],fR(hppTot["12 kg"]),qtyTot["50 kg"],fR(hppTot["50 kg"]),grandTotal,"","",""]);
var ws=XLSX.utils.aoa_to_sheet(aoa);
ws["!cols"]=[{wch:12},{wch:16},{wch:14},{wch:12},{wch:9},{wch:11},{wch:9},{wch:11},{wch:9},{wch:11},{wch:14},{wch:10},{wch:10},{wch:10}];
XLSX.utils.book_append_sheet(wb,ws,"Riwayat DO");
XLSX.writeFile(wb,"Riwayat_DO_"+bulanRiwayat+".xlsx");
toast("✓ Excel Riwayat DO didownload!");
}
return <div ref={formRef}>
<div style={{marginBottom:12}}>
<div style={{display:"flex",alignItems:"center",gap:8,marginBottom:10}}>
<span style={{fontSize:20}}>🚚</span>
<span style={{fontSize:18,fontWeight:800,color:C.wht}}>Delivery Order</span>
</div>
{/* Bar kode Pertamina — memanjang tipis di bawah judul */}
<div style={{background:C.card,border:"1px solid "+C.bdr,borderRadius:8,padding:"8px 14px",display:"flex",alignItems:"center",gap:0,flexWrap:"wrap",overflow:"hidden"}}>
{/* Bulan SA */}
{data.company?.saBulan&&<div style={{display:"flex",alignItems:"center",gap:6,paddingRight:14,borderRight:"1px solid "+C.bdr,marginRight:14}}>
<span style={{fontSize:8,fontWeight:700,letterSpacing:1,textTransform:"uppercase",color:C.gl2}}>SA Aktif</span>
<span style={{fontSize:11,fontWeight:700,color:C.olt}}>{data.company.saBulan}</span>
</div>}
{/* Kode penebusan */}
{[["Sold To",data.company?.soldTo],["Ship To KCR",data.company?.shipToKCR],["Ship To MGL",data.company?.shipToMGL]].map(([l,v])=>v?<div key={l} style={{display:"flex",alignItems:"center",gap:5,paddingRight:14,borderRight:"1px solid "+C.bdr,marginRight:14}}>
<span style={{fontSize:8,color:C.gl2,fontWeight:600,whiteSpace:"nowrap"}}>{l}:</span>
<span style={{fontSize:11,fontWeight:800,color:C.blt,fontFamily:"'Courier New',monospace"}}>{v}</span>
</div>:null)}
{/* SA per SPPBE */}
{[["KCR",[["12kg",data.company?.sa12KCR],["5,5kg",data.company?.sa55KCR]]],["MGL",[["12kg",data.company?.sa12MGL],["5,5kg",data.company?.sa55MGL]]]].map(([loc,items])=><div key={loc} style={{display:"flex",alignItems:"center",gap:6,paddingRight:14,borderRight:"1px solid "+C.bdr,marginRight:14,flexWrap:"wrap"}}>
<span style={{fontSize:8,fontWeight:700,color:C.olt,letterSpacing:.5}}>{loc}:</span>
{items.map(([l,v])=><div key={l} style={{display:"flex",alignItems:"center",gap:3}}>
<span style={{fontSize:8,color:C.gl2}}>{l}</span>
{v?<span style={{fontSize:11,fontWeight:800,color:C.glt,fontFamily:"'Courier New',monospace"}}>{v}</span>:<span style={{fontSize:9,color:C.gry,fontStyle:"italic"}}>-</span>}
</div>)}
</div>)}
</div>
</div>
{editTripId&&<div style={{background:"linear-gradient(90deg,#92400e,#b45309)",border:"1px solid #f59e0b",borderRadius:10,padding:"10px 16px",marginBottom:12,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
<div>
<div style={{color:"#fef3c7",fontWeight:800,fontSize:13}}>✏️ Mode Edit Trip DO</div>
<div style={{color:"#fde68a",fontSize:11,marginTop:2}}>Data trip dimuat ke form di bawah. Ubah data lalu klik <b>Simpan Perubahan</b>.</div>
</div>
<button onClick={resetFormT} style={{background:"rgba(0,0,0,.25)",border:"1px solid #f59e0b",borderRadius:7,padding:"6px 12px",color:"#fef3c7",cursor:"pointer",fontWeight:700,fontSize:11,flexShrink:0}}>✕ Batal Edit</button>
</div>}
<Card style={{border:editTripId?"1px solid #f59e0b":undefined,width:"fit-content",maxWidth:"100%",minWidth:660}}>
<div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>🚚 Input DO — 1 Trip, Multi Produk</div>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(150px,210px))",gap:10}}>
<Inp label="Tanggal" type="date" value={ft.tanggal} onChange={v=>setFt(p=>({...p,tanggal:v}))}/>
<div style={{marginBottom:10}}>
<label style={{display:"block",fontSize:11,color:C.gl2,marginBottom:3,fontWeight:600}}>Trip</label>
<select value={ft.trip} onChange={e=>setFt(p=>({...p,trip:e.target.value}))} style={{width:"100%",border:"1px solid "+C.bdr,borderRadius:8,padding:"9px 11px",color:C.wht,fontSize:13,outline:"none",background:C.nav,boxSizing:"border-box"}}>
{["Trip 1","Trip 2","Trip 3","Trip 4","Trip 5"].map(t=>{var used=(data.doTrip||[]).filter(d=>d.tanggal===ft.tanggal&&d.trip===t);return <option key={t} value={t}>{t}{used.length>0?" (sudah dipakai)":""}</option>;})}
</select>
</div>
<Sel label="SPPBE" value={ft.sppbe} onChange={onSppbeT} opts={SPPBE_OPTS}/>
<Sel label="Nama Driver" value={ft.dibuatOleh} onChange={v=>setFt(p=>({...p,dibuatOleh:v}))} opts={[{v:"",l:"-- Pilih --"},...(data.employees||[]).filter(e=>e.aktif).map(e=>({v:e.id,l:e.nama+" ("+e.posisi+")"}))]}/>
<Inp label="Keterangan" value={ft.ket} onChange={v=>setFt(p=>({...p,ket:v}))}/>
</div>
<div style={{fontSize:10,color:isFixedSPPBE?C.olt:C.gl2,marginBottom:10,fontStyle:"italic"}}>{isFixedSPPBE?"🔒 HPP terkunci sesuai Pengaturan untuk "+ft.sppbe+" (ubah di Settings jika perlu).":"✏️ HPP bisa diisi manual karena SPPBE \"Lainnya\"."}</div>
<div style={{overflowX:"auto",marginBottom:10}}>
<table style={{width:"100%",borderCollapse:"collapse",fontSize:12,minWidth:520}}>
<thead><tr style={{background:C.nav}}>
{["Ukuran","Qty Tabung","HPP/Unit","Total HPP"].map(h=><th key={h} style={{padding:"7px 8px",color:C.gl2,textAlign:h==="Ukuran"?"left":"right",border:"1px solid "+C.bdr,fontSize:10,fontWeight:700}}>{h}</th>)}
</tr></thead>
<tbody>
{SIZES.map((s,i)=>{var it=ft.items[s];var tot=(Number(it.qty)||0)*(Number(it.hppUnit)||0);return <tr key={s} style={{background:i%2===0?C.bg:C.nav}}>
<td style={{padding:"6px 8px",border:"1px solid "+C.bdr}}><Bdg color={s==="12 kg"?"blue":s==="5.5 kg"?"green":"orange"}>{s}</Bdg></td>
<td style={{padding:"4px 8px",border:"1px solid "+C.bdr}}><input type="number" value={it.qty} onChange={e=>setItemT(s,"qty",e.target.value)} placeholder="0" style={{width:"100%",background:"transparent",border:"none",color:C.wht,fontSize:13,outline:"none",textAlign:"right"}}/></td>
<td style={{padding:"4px 8px",border:"1px solid "+C.bdr,background:isFixedSPPBE?(C.mode==="dark"?"#0A1A2A":"#EFF6FF"):"transparent"}}><input type="number" value={it.hppUnit} onChange={e=>setItemT(s,"hppUnit",e.target.value)} placeholder="0" readOnly={isFixedSPPBE} style={{width:"100%",background:"transparent",border:"none",color:isFixedSPPBE?C.blt:C.wht,fontSize:13,outline:"none",textAlign:"right",cursor:isFixedSPPBE?"not-allowed":"text"}}/></td>
<td style={{padding:"6px 8px",border:"1px solid "+C.bdr,textAlign:"right",color:tot>0?C.olt:C.gry,fontWeight:700}}>{tot>0?fR(tot):"-"}</td>
</tr>;})}
<tr style={{background:C.olt}}>
<td colSpan={3} style={{padding:"8px 8px",border:"1px solid "+C.bdr,color:"white",fontWeight:800,fontSize:12}}>TOTAL NILAI DO</td>
<td style={{padding:"8px 8px",border:"1px solid "+C.bdr,textAlign:"right",color:"white",fontWeight:900,fontSize:13}}>{fR(totalNilaiTrip)}</td>
</tr>
</tbody>
</table>
</div>
<div style={{display:"flex",gap:8}}>
<Btn color={editTripId?"orange":"green"} onClick={saveTrip} dis={!ft.trip}>{editTripId?"💾 Simpan Perubahan":"💾 Simpan DO"}</Btn>
{editTripId&&<Btn color="gray" onClick={resetFormT}>✕ Batal</Btn>}
</div>
</Card>
<Card>
<div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-end",flexWrap:"wrap",gap:10,marginBottom:10}}>
<div style={{fontWeight:700,color:C.gl2,fontSize:13}}>Riwayat DO</div>
<div style={{display:"flex",gap:8,alignItems:"flex-end"}}>
<div style={{minWidth:150}}>
<label style={{display:"block",fontSize:11,color:C.gl2,marginBottom:3,fontWeight:600}}>Bulan</label>
<select value={bulanRiwayat} onChange={e=>setBulanRiwayat(e.target.value)} style={{width:"100%",border:"1px solid "+C.bdr,borderRadius:8,padding:"7px 10px",color:C.wht,fontSize:12,outline:"none",background:C.nav,boxSizing:"border-box"}}>
{bulanTersedia.map(b=>{var[y,m]=b.split("-");var nm=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"][Number(m)-1];return <option key={b} value={b}>{nm} {y}</option>;})}
</select>
</div>
<button onClick={exportExcelDO} disabled={!rowsBulanIni.length} style={{background:rowsBulanIni.length?"#15803D":C.gry,color:"white",border:"none",padding:"8px 14px",borderRadius:8,fontSize:12,cursor:rowsBulanIni.length?"pointer":"not-allowed",fontWeight:700,whiteSpace:"nowrap"}}>📥 Export Excel</button>
</div>
</div>
<table style={{width:"auto",minWidth:"1032px",maxWidth:"100%",margin:"0 auto",borderCollapse:"collapse",fontSize:12.5,tableLayout:"fixed"}}>
<colgroup>
<col style={{width:"82px"}}/><col style={{width:"75px"}}/><col style={{width:"115px"}}/><col style={{width:"85px"}}/>
<col style={{width:"55px"}}/><col style={{width:"85px"}}/><col style={{width:"55px"}}/><col style={{width:"85px"}}/><col style={{width:"55px"}}/><col style={{width:"85px"}}/>
<col style={{width:"110px"}}/><col style={{width:"55px"}}/><col style={{width:"90px"}}/>
</colgroup>
<thead><tr style={{background:C.nav}}>
<th style={{padding:"8px 6px",color:C.gl2,textAlign:"left",border:"1px solid "+C.bdr,fontSize:11.5,fontWeight:700}}>Tgl</th>
<th style={{padding:"8px 6px",color:C.gl2,textAlign:"left",border:"1px solid "+C.bdr,fontSize:11.5,fontWeight:700}}>Trip</th>
<th style={{padding:"8px 6px",color:C.gl2,textAlign:"left",border:"1px solid "+C.bdr,fontSize:11.5,fontWeight:700}}>Driver</th>
<th style={{padding:"8px 6px",color:C.gl2,textAlign:"left",border:"1px solid "+C.bdr,fontSize:11.5,fontWeight:700}}>SPPBE</th>
<th style={{padding:"8px 4px",color:C.gl2,textAlign:"right",border:"1px solid "+C.bdr,fontSize:11.5,fontWeight:700}}>Qty 5,5</th>
<th style={{padding:"8px 4px",color:C.gl2,textAlign:"right",border:"1px solid "+C.bdr,fontSize:11.5,fontWeight:700}}>Hrg 5,5</th>
<th style={{padding:"8px 4px",color:C.gl2,textAlign:"right",border:"1px solid "+C.bdr,fontSize:11.5,fontWeight:700}}>Qty 12</th>
<th style={{padding:"8px 4px",color:C.gl2,textAlign:"right",border:"1px solid "+C.bdr,fontSize:11.5,fontWeight:700}}>Hrg 12</th>
<th style={{padding:"8px 4px",color:C.gl2,textAlign:"right",border:"1px solid "+C.bdr,fontSize:11.5,fontWeight:700}}>Qty 50</th>
<th style={{padding:"8px 4px",color:C.gl2,textAlign:"right",border:"1px solid "+C.bdr,fontSize:11.5,fontWeight:700}}>Hrg 50</th>
<th style={{padding:"8px 6px",color:C.gl2,textAlign:"right",border:"1px solid "+C.bdr,fontSize:11.5,fontWeight:700}}>Total DO</th>
<th style={{padding:"8px 4px",color:C.gl2,border:"1px solid "+C.bdr,fontSize:11.5,fontWeight:700}}>St.</th>
<th style={{padding:"8px 4px",color:C.gl2,border:"1px solid "+C.bdr,fontSize:11.5,fontWeight:700}}>Aksi</th>
</tr></thead>
<tbody>
{rowsBulanIni.length===0&&<tr><td colSpan={13} style={{padding:20,textAlign:"center",color:C.gl2,border:"1px solid "+C.bdr,fontSize:11}}>Belum ada DO di bulan ini</td></tr>}
{rowsBulanIni.map((r,ri)=>{
if(r._src==="trip"){
var trip=r._raw;
var allGantung=trip.items.every(it=>it.status==="gantung");
return <tr key={r._id} style={{background:ri%2===0?C.bg:C.nav}}>
<td style={{padding:"7px 6px",color:C.wht,border:"1px solid "+C.bdr,fontSize:12,overflow:"hidden",textOverflow:"ellipsis"}}>{fDs(r.tanggal)}</td>
<td style={{padding:"7px 6px",color:C.wht,border:"1px solid "+C.bdr,fontWeight:700,fontSize:12,overflow:"hidden",textOverflow:"ellipsis"}}>{r.trip}</td>
<td style={{padding:"7px 6px",color:C.wht,border:"1px solid "+C.bdr,fontSize:12,overflow:"hidden",textOverflow:"ellipsis"}}>{r.driver||"-"}</td>
<td style={{padding:"7px 6px",color:C.gl2,border:"1px solid "+C.bdr,fontSize:12,overflow:"hidden",textOverflow:"ellipsis"}}>{r.sppbe}</td>
{["5.5 kg","12 kg","50 kg"].flatMap(uk=>{var it=r.per[uk];return[
<td key={uk+"q"} style={{padding:"7px 4px",textAlign:"right",color:it?C.glt:C.gry,fontWeight:it?700:400,border:"1px solid "+C.bdr,fontSize:12}}>{it?it.qty:"-"}</td>,
<td key={uk+"h"} style={{padding:"7px 4px",textAlign:"right",color:it?C.wht:C.gry,border:"1px solid "+C.bdr,fontSize:11.5}}>{it?fR(it.hppUnit).replace("Rp ",""):"-"}</td>
];})}
<td style={{padding:"7px 6px",textAlign:"right",color:C.olt,fontWeight:800,border:"1px solid "+C.bdr,fontSize:12.5}}>{fR(r.totalNilai).replace("Rp ","")}</td>
<td style={{padding:"6px 4px",border:"1px solid "+C.bdr}}>
<div style={{display:"flex",flexDirection:"column",gap:2,alignItems:"center"}}>
{trip.items.map(it=><span key={it.id} title={it.ukuran} style={{fontSize:11}}>{it.status==="gantung"?"⏳":it.status==="sangkut"?"⚠️":"✅"}</span>)}
</div>
</td>
<td style={{padding:"6px 4px",border:"1px solid "+C.bdr}}>
<div style={{display:"flex",flexDirection:"column",gap:3,alignItems:"center"}}>
{allGantung&&<button onClick={()=>mulaiEditTrip(trip)} title="Edit Trip" style={{background:editTripId===trip.id?"#B45309":"#1D4ED8",border:"none",borderRadius:5,padding:"4px 6px",color:"white",fontSize:11,fontWeight:700,cursor:"pointer",width:"100%"}}>✏️</button>}
{trip.items.map(it=><div key={it.id} style={{display:"flex",gap:3,justifyContent:"center"}}>
{it.status==="gantung"&&<><button onClick={()=>terimaTripItem(trip,it)} title={it.ukuran+" Terima"} style={{background:"#15803D",border:"none",borderRadius:4,padding:"3px 5px",color:"white",fontSize:11,cursor:"pointer"}}>✅</button><button onClick={()=>sangkutTripItem(trip,it)} title={it.ukuran+" Sangkut"} style={{background:"#B45309",border:"none",borderRadius:4,padding:"3px 5px",color:"white",fontSize:11,cursor:"pointer"}}>⚠️</button></>}
{it.status==="sangkut"&&<button onClick={()=>releaseTripItem(trip,it)} title={it.ukuran+" Release"} style={{background:"#1D4ED8",border:"none",borderRadius:4,padding:"3px 5px",color:"white",fontSize:11,cursor:"pointer"}}>🔓</button>}
{it.status==="diterima"&&<button onClick={()=>mulaiEditItemDO(trip,it)} title={it.ukuran+" Edit Qty"} style={{background:editItemDO?.item?.id===it.id?"#B45309":"#1D4ED8",border:"none",borderRadius:4,padding:"3px 5px",color:"white",fontSize:11,cursor:"pointer"}}>✏️</button>}
<button onClick={()=>hapusTripItem(trip,it)} title={it.ukuran+" Hapus"} style={{background:"#991B1B",border:"none",borderRadius:4,padding:"3px 5px",color:"white",fontSize:11,cursor:"pointer"}}>🗑️</button>
</div>)}
</div>
</td>
</tr>;
}
// legacy
var d=r._raw;var st=d.status||"diterima";
var stIcon=st==="gantung"?"⏳":st==="sangkut"?"⚠️":"✅";
var aksiBtn=st==="gantung"?<div style={{display:"flex",flexDirection:"column",gap:3,alignItems:"center"}}>
<button onClick={()=>terimaDO(d)} title="Terima" style={{background:"#15803D",border:"none",borderRadius:4,padding:"3px 6px",color:"white",fontSize:11,cursor:"pointer",width:"100%"}}>✅</button>
<button onClick={()=>mulaiEdit(d)} title="Edit" style={{background:editId===d.id?"#B45309":"#1D4ED8",border:"none",borderRadius:4,padding:"3px 6px",color:"white",fontSize:11,cursor:"pointer",width:"100%"}}>✏️</button>
<button onClick={()=>sangkutDO(d)} title="Sangkut" style={{background:"#B45309",border:"none",borderRadius:4,padding:"3px 6px",color:"white",fontSize:11,cursor:"pointer",width:"100%"}}>⚠️</button>
<ActBtns onDel={()=>setDelId(d)}/>
</div>:st==="sangkut"?<div style={{display:"flex",flexDirection:"column",gap:3,alignItems:"center"}}>
<button onClick={()=>releaseDO(d)} title="Release" style={{background:"#1D4ED8",border:"none",borderRadius:4,padding:"3px 6px",color:"white",fontSize:11,cursor:"pointer",width:"100%"}}>🔓</button>
<ActBtns onDel={()=>setDelId(d)}/>
</div>:<ActBtns onDel={()=>setDelId(d)}/>;
return <tr key={r._id} style={{background:ri%2===0?C.bg:C.nav,opacity:.92}}>
<td style={{padding:"7px 6px",color:C.wht,border:"1px solid "+C.bdr,fontSize:12,overflow:"hidden",textOverflow:"ellipsis"}}>{fDs(r.tanggal)}</td>
<td style={{padding:"7px 6px",color:C.wht,border:"1px solid "+C.bdr,fontWeight:700,fontSize:11,overflow:"hidden",textOverflow:"ellipsis"}}>{r.trip}</td>
<td style={{padding:"7px 6px",color:C.wht,border:"1px solid "+C.bdr,fontSize:12,overflow:"hidden",textOverflow:"ellipsis"}}>{r.driver||"-"}</td>
<td style={{padding:"7px 6px",color:C.gl2,border:"1px solid "+C.bdr,fontSize:12,overflow:"hidden",textOverflow:"ellipsis"}}>{r.sppbe}</td>
{["5.5 kg","12 kg","50 kg"].flatMap(uk=>{var match=d.ukuran===uk;return[
<td key={uk+"q"} style={{padding:"7px 4px",textAlign:"right",color:match?C.glt:C.gry,fontWeight:match?700:400,border:"1px solid "+C.bdr,fontSize:12}}>{match?d.qty:"-"}</td>,
<td key={uk+"h"} style={{padding:"7px 4px",textAlign:"right",color:match?C.wht:C.gry,border:"1px solid "+C.bdr,fontSize:11.5}}>{match?fR(d.hppUnit||0).replace("Rp ",""):"-"}</td>
];})}
<td style={{padding:"7px 6px",textAlign:"right",color:C.olt,fontWeight:800,border:"1px solid "+C.bdr,fontSize:12.5}}>{fR(d.totalHPP||0).replace("Rp ","")}</td>
<td style={{padding:"6px 4px",border:"1px solid "+C.bdr,textAlign:"center",fontSize:11}}>{stIcon}</td>
<td style={{padding:"6px 4px",border:"1px solid "+C.bdr}}>{aksiBtn}</td>
</tr>;
})}
</tbody>
<tfoot>
{(()=>{
var qtyTot={"5.5 kg":0,"12 kg":0,"50 kg":0};var hppTot={"5.5 kg":0,"12 kg":0,"50 kg":0};var grandTotal=0;
rowsBulanIni.forEach(r=>{SIZES.forEach(uk=>{if(r.per[uk]){qtyTot[uk]+=r.per[uk].qty||0;hppTot[uk]+=(r.per[uk].qty||0)*(r.per[uk].hppUnit||0);}});grandTotal+=r.totalNilai||0;});
return <tr style={{background:C.olt}}>
<td colSpan={4} style={{padding:"9px 6px",border:"1px solid "+C.bdr,color:"white",fontWeight:800,fontSize:12.5}}>TOTAL BULAN INI</td>
{["5.5 kg","12 kg","50 kg"].flatMap(uk=>[
<td key={uk+"q"} style={{padding:"9px 4px",border:"1px solid "+C.bdr,textAlign:"right",color:"white",fontWeight:900,fontSize:12}}>{qtyTot[uk]}</td>,
<td key={uk+"h"} style={{padding:"9px 4px",border:"1px solid "+C.bdr,textAlign:"right",color:"white",fontWeight:700,fontSize:11}}>{fR(hppTot[uk]).replace("Rp ","")}</td>
])}
<td style={{padding:"9px 6px",border:"1px solid "+C.bdr,textAlign:"right",color:"white",fontWeight:900,fontSize:12.5}}>{fR(grandTotal).replace("Rp ","")}</td>
<td colSpan={2} style={{border:"1px solid "+C.bdr}}/>
</tr>;
})()}
</tfoot>
</table>
</Card>
{editItemDO&&<div style={{position:"fixed",top:0,left:0,right:0,bottom:0,background:"rgba(0,0,0,.6)",zIndex:2000,display:"flex",alignItems:"center",justifyContent:"center"}} onClick={()=>setEditItemDO(null)}>
<div style={{background:"#1e293b",border:"1px solid #f59e0b",borderRadius:14,padding:"24px 28px",minWidth:320,maxWidth:420}} onClick={e=>e.stopPropagation()}>
<div style={{fontWeight:800,fontSize:15,color:"#fcd34d",marginBottom:4}}>✏️ Edit Qty Item DO</div>
<div style={{fontSize:12,color:"#94a3b8",marginBottom:16}}>Trip: {editItemDO.trip.trip} — {editItemDO.item.ukuran} — {editItemDO.trip.sppbe}</div>
<div style={{fontSize:11,color:"#94a3b8",marginBottom:6}}>Qty saat ini: <b style={{color:"#fff"}}>{editItemDO.item.qty}</b></div>
<input type="number" value={editItemQty} onChange={e=>setEditItemQty(e.target.value)} placeholder="Qty baru"
  style={{width:"100%",border:"1px solid #f59e0b",borderRadius:8,padding:"10px 12px",color:"#fff",fontSize:14,outline:"none",background:"#0f172a",boxSizing:"border-box",marginBottom:8}}
  autoFocus/>
<div style={{fontSize:10,color:"#94a3b8",marginBottom:16}}>
{Number(editItemQty)>0&&Number(editItemQty)!==editItemDO.item.qty&&<span style={{color:Number(editItemQty)>editItemDO.item.qty?"#34d399":"#f87171"}}>
  {Number(editItemQty)>editItemDO.item.qty?"+ Tambah":"- Kurang"} {Math.abs(Number(editItemQty)-editItemDO.item.qty)} tabung → stok & FIFO batch akan disesuaikan otomatis
</span>}
</div>
<div style={{display:"flex",gap:8}}>
<button onClick={simpanEditItemDO} style={{flex:1,background:"#b45309",border:"1px solid #f59e0b",borderRadius:8,padding:"10px",color:"#fcd34d",fontWeight:800,fontSize:13,cursor:"pointer"}}>💾 Simpan</button>
<button onClick={()=>setEditItemDO(null)} style={{background:"#334155",border:"none",borderRadius:8,padding:"10px 16px",color:"#94a3b8",cursor:"pointer",fontSize:13}}>✕ Batal</button>
</div>
</div>
</div>}
{delId&&<ConfirmDel msg={"Hapus DO "+(delId.trip||"")+"? Stok akan dikembalikan otomatis."} onCancel={()=>setDelId(null)} onConfirm={()=>{
var st=delId.status||"diterima";
var qty=Number(delId.qty||0);var uk=delId.ukuran;
var ns={...(data.stock||{})};var nk={...(data.stokKosong||{})};
if(st==="diterima"){
  ns[uk]=Math.max(0,(ns[uk]||0)-qty);
  nk[uk]=(nk[uk]||0)+qty;
}
var logs=st==="diterima"?[{id:uid(),tanggal:delId.tanggal,ukuran:uk,jenis:"Reverse Hapus DO "+(delId.trip||""),qty,ket:"Stok dikembalikan karena DO dihapus",sumber:"Hapus DO",user:user?.nama||""},...(data.stockLog||[])].slice(0,500):(data.stockLog||[]);
setData(d=>{
  var dRev=st==="diterima"?reverseBatch(d,uk,qty,"DO "+(delId.trip||"")+" ("+(delId.sppbe||"")+")",delId.hppUnit||0):d;
  return{...dRev,doList:(d.doList||[]).filter(x=>x.id!==delId.id),stock:ns,stokKosong:nk,stockLog:logs};
});
if(editId===delId.id)resetForm();
setDelId(null);
toast(st==="diterima"?"✓ DO dihapus & stok dikembalikan!":"✓ DO dihapus.");
}}/>}
</div>;
}

// ─── PIUTANG v4 (filter sama dgn Penjualan, cetak invoice BON) ────────────────
function PiutangMod({data,setData,setInv,toast}){
var C=useTheme();
var[openId,setOpenId]=useState(null);var[delId,setDelId]=useState(null);
var[bF,setBF]=useState({nominal:"",metode:"cash",bank:"BSI",salesPenerimaId:""});
var[barFilter,setBarFilter]=useState({from:"",to:"",salesId:"",konsumen:"",status:""});
var[showGabung,setShowGabung]=useState(false);
var[editPayBon,setEditPayBon]=useState(null);// {bon, payIdx} untuk edit/cancel pembayaran
var[gabungPilih,setGabungPilih]=useState([]);// array bon id yang dipilih
var[gabungKons,setGabungKons]=useState("");// filter konsumen untuk gabung
var[tabBon,setTabBon]=useState("aktif");// aktif | lunas
var salesList=sortEmp((data.employees||[]).filter(e=>e.aktif));

function makeBonInvObj(b){var plg=(data.pelanggan||[]).find(x=>x.id===b.konsumenId);var emp=(data.employees||[]).find(x=>x.id===b.salesId);var totalDibayar=(b.pembayaran||[]).reduce((a,p)=>a+Number(p.jumlah||p.nominal||0),0);var sisaTagihan=b.status==="lunas"?0:(b.sisaTagihan!=null?b.sisaTagihan:b.total-totalDibayar);var noInvAsli=(b.noInv||"").replace(/\(BON\)$/,"");var penjAsli=(data.penjualan||[]).find(x=>x.noInv===noInvAsli&&x.konsumenId===b.konsumenId);var totalBelanja=penjAsli?penjAsli.total:b.total;var tfDibayar=penjAsli&&penjAsli.splitDetail?Number(penjAsli.splitDetail.tf||0):0;var splitBankAsli=penjAsli?penjAsli.splitBank||"":"";return{noInv:b.noInv||"#HTS/INV/-/-",tanggal:b.tanggal,konsumen:b.konsumen,kota:plg?.alamat?.split(",").pop()?.trim()||"Banda Aceh",salesNama:emp?.nama||"",items:(b.items||[]).map(it=>({ukuran:it.ukuran,jenis:it.jenis,qty:Number(it.qty),price:Number(it.price),tglDO:it.tglDO||b.tanggal})),total:sisaTagihan,totalBelanja,tfDibayar,splitBank:splitBankAsli,metodeBayar:b.status==="lunas"?"BON (LUNAS)":"BON",isBon:b.status!=="lunas",isGabungan:b.isGabungan||false,catatan:b.ket||"",bonLunas:b.status==="lunas",bonSebagian:b.status==="sebagian",totalDibayar,sisaTagihan,riwayatBayar:(b.pembayaran||[])};}

// ── Invoice Gabungan Object ──
function makeGabungInvObj(bons,noInvBaru){
var plg=(data.pelanggan||[]).find(x=>x.id===bons[0]?.konsumenId);
var emp=(data.employees||[]).find(x=>x.id===bons[0]?.salesId);
// items: satu row per BON, dengan tanggal DO
var items=bons.flatMap(b=>(b.items||[]).map(it=>({...it,qty:Number(it.qty),price:Number(it.price),tglDO:b.tanggal})));
var total=bons.reduce((a,b)=>a+(b.sisaTagihan||b.total||0),0);
var allLunas=bons.every(b=>b.status==="lunas");
var anySebagian=bons.some(b=>b.status==="sebagian"||(b.pembayaran||[]).length>0)&&!allLunas;
var totalSisa=bons.reduce((a,b)=>a+(b.status==="lunas"?0:(b.sisaTagihan||b.total||0)),0);
var totalAll=bons.reduce((a,b)=>a+(b.total||0),0);
var totalDibayarAll=bons.reduce((a,b)=>a+(b.pembayaran||[]).reduce((aa,p)=>aa+Number(p.nominal||0),0),0);
return{noInv:noInvBaru,tanggal:toDay(),konsumen:bons[0]?.konsumen||"",kota:plg?.alamat?.split(",").pop()?.trim()||"Banda Aceh",salesNama:emp?.nama||"",items,total:totalAll,sisaTagihan:totalSisa,metodeBayar:allLunas?"BON (LUNAS)":"BON (GABUNGAN)",isBon:!allLunas,isGabungan:true,bonLunas:allLunas,bonSebagian:anySebagian,totalDibayar:totalDibayarAll,bonAsal:bons.map(b=>b.id),catatan:"Invoice gabungan dari "+bons.length+" BON"};
}

function bayar(b){
if(!bF.nominal)return;var nom=Number(bF.nominal);
var newSisa=Math.max(0,(b.sisaTagihan||0)-nom);var st=newSisa===0?"lunas":"sebagian";
var payRec={id:uid(),tanggal:toDay(),jumlah:nom,metode:bF.metode,bank:bF.metode==="transfer"?bF.bank:"",salesPenerimaId:bF.salesPenerimaId,salesPenerimaNama:salesList.find(e=>e.id===bF.salesPenerimaId)?.nama||""};
// Kalau BON gabungan dan lunas → lunasi semua BON asal sekaligus
var newBon=(data.bon||[]).map(x=>{
  if(x.id===b.id)return{...x,sisaTagihan:newSisa,status:st,pembayaran:[...(x.pembayaran||[]),payRec]};
  // Kalau ini BON asal dari gabungan yang baru lunas
  if(st==="lunas"&&b.isGabungan&&(b.bonAsal||[]).includes(x.id))return{...x,status:"digabung",sisaTagihan:0};
  return x;
});
var newSet=data.setoranSales||[];
if(bF.metode==="cash"&&bF.salesPenerimaId)newSet=[{id:uid(),tanggal:toDay(),salesId:bF.salesPenerimaId,sumber:"piutang",refId:b.id,nominal:nom,disetor:false,konsumen:b.konsumen},...newSet];
setData(d=>({...d,bon:newBon,setoranSales:newSet}));
setBF({nominal:"",metode:"cash",bank:"BSI",salesPenerimaId:""});setOpenId(null);toast("✓ Pembayaran tercatat!");
}

// ── Gabung Invoice ──
function doGabung(){
if(gabungPilih.length<2){toast("Pilih minimal 2 BON untuk digabung");return;}
var bons=(data.bon||[]).filter(b=>gabungPilih.includes(b.id));
var invInfo=nextInvNo(data,toDay());
var newCounters={...(data.counters||{inv:{},sg:{},reg:0})};if(!newCounters.inv)newCounters.inv={};newCounters.inv[invInfo.key]=invInfo.n;
var noInvBaru=invInfo.no;
// Tandai BON asal sebagai "digabung"
var newBon=(data.bon||[]).map(b=>{
  if(!gabungPilih.includes(b.id))return b;
  // BON lunas → tetap lunas (jangan ubah status)
  if(b.status==="lunas")return b;
  // BON belum/sebagian → tandai digabung
  return{...b,status:"digabung",sisaTagihan:b.sisaTagihan||b.total||0};
});
// Buat BON gabungan baru
// Total yang masih perlu dibayar (BON lunas tidak dihitung)
var totalGabung=bons.reduce((a,b)=>a+(b.status==="lunas"?0:(b.sisaTagihan||b.total||0)),0);
// Total keseluruhan untuk info
var totalGabungAll=bons.reduce((a,b)=>a+(b.total||0),0);
var semuaLunas=bons.every(b=>b.status==="lunas");
var bonGabung={id:uid(),noInv:noInvBaru,tanggal:toDay(),konsumen:bons[0].konsumen,konsumenId:bons[0].konsumenId,salesId:bons[0].salesId,items:bons.flatMap(b=>(b.items||[]).map(it=>({...it,qty:Number(it.qty),price:Number(it.price),tglDO:b.tanggal,statusBon:b.status}))),total:totalGabungAll,sisaTagihan:totalGabung,status:semuaLunas?"lunas":"belum",pembayaran:[],isGabungan:true,bonAsal:gabungPilih,ket:"Gabungan dari: "+bons.map(b=>b.noInv).join(", ")};
newBon=[bonGabung,...newBon];
setData(d=>({...d,bon:newBon,counters:newCounters}));
toast("✓ Invoice gabungan dibuat: "+noInvBaru);
setShowGabung(false);setGabungPilih([]);setGabungKons("");
// Tampilkan invoice gabungan
setInv(makeGabungInvObj(bons,noInvBaru));
}

var rows=useMemo(()=>{
return(data.bon||[]).filter(b=>{
if(barFilter.from&&b.tanggal<barFilter.from)return false;
if(barFilter.to&&b.tanggal>barFilter.to)return false;
if(barFilter.salesId&&b.salesId!==barFilter.salesId&&!(b.salesNama||b.sales||'').toLowerCase().includes((salesList.find(e=>e.id===barFilter.salesId)?.nama||'').toLowerCase()))return false;
if(barFilter.konsumen&&!b.konsumen.toLowerCase().includes(barFilter.konsumen.toLowerCase()))return false;
if(barFilter.status&&b.status!==barFilter.status)return false;
return true;
}).map(b=>{var emp=(data.employees||[]).find(e=>e.id===b.salesId);return{...b,salesNama:emp?.nama||b.salesNama||b.sales||"-",dl:dLeft(b.deadline)};});
},[data.bon,data.employees,barFilter]);

// Pisahkan berdasarkan tab: Aktif (belum/sebagian/digabung) vs Riwayat Lunas
var rowsAktif=rows.filter(r=>r.status!=="lunas");
var rowsLunas=rows.filter(r=>r.status==="lunas");
var rowsTab=tabBon==="lunas"?rowsLunas:rowsAktif;

// BON aktif untuk modal gabung
var bonAktifGabung=(data.bon||[]).filter(b=>b.status!=="digabung").filter(b=>!gabungKons||b.konsumen.toLowerCase().includes(gabungKons.toLowerCase()));

var cols=[
{key:"tanggal",label:"Tgl",width:88,render:r=>fDs(r.tanggal),sortVal:r=>r.tanggal,filterable:true},
{key:"konsumen",label:"Konsumen",render:r=><b style={{color:C.wht}}>{r.konsumen}</b>,filterable:true,width:170},
{key:"salesNama",label:"Sales",width:120,filterable:true},
{key:"deadline",label:"Jatuh Tempo",width:140,render:r=>r.deadline?<span style={{whiteSpace:"nowrap",color:r.dl<0&&r.status!=="lunas"?C.rlt:r.dl<=3&&r.status!=="lunas"?C.olt:C.gl2}}>{fDs(r.deadline)}{r.status!=="lunas"&&r.dl!=null?" ("+(r.dl<0?Math.abs(r.dl)+"h LEWAT":r.dl+"h)"):""}</span>:"-",filterable:false},
{key:"total",label:"Total",width:130,render:r=><span style={{whiteSpace:"nowrap"}}>{fR(r.total)}</span>,filterable:false},
{key:"sisaTagihan",label:"Sisa",width:130,render:r=><b style={{color:r.status==="lunas"?C.glt:r.status==="digabung"?C.gl2:C.rlt,whiteSpace:"nowrap"}}>{fR(r.sisaTagihan)}</b>,filterable:false},
{key:"status",label:"Status",width:110,render:r=>r.status==="lunas"?<Bdg color="green">Lunas</Bdg>:r.status==="sebagian"?<Bdg color="orange">Sebagian</Bdg>:r.status==="digabung"?<Bdg color="gray">Digabung</Bdg>:r.isGabungan?<Bdg color="blue">Gabungan</Bdg>:<Bdg color="red">Belum</Bdg>,filterable:true,filterType:"select",options:[{v:"lunas",l:"Lunas"},{v:"sebagian",l:"Sebagian"},{v:"belum",l:"Belum"},{v:"digabung",l:"Digabung"}]},
{key:"_aksi",label:"Aksi",width:150,sortable:false,filterable:false,render:r=><div style={{display:"flex",gap:5}}><button onClick={()=>{if(r.isGabungan){var bonsAsal=(data.bon||[]).filter(b=>(r.bonAsal||[]).includes(b.id));var invObj=bonsAsal.length>0?makeGabungInvObj(bonsAsal,r.noInv):makeBonInvObj(r);
// Status lunas/sebagian & nominal dibayar ditentukan dari record gabungan ITU SENDIRI (r), bukan hasil hitung ulang dari BON asal yang sudah berubah status jadi "digabung" setelah penggabungan
var lunasReal=r.status==="lunas";
var totalDibayarReal=(r.pembayaran||[]).reduce((a,p)=>a+Number(p.nominal||0),0);
invObj={...invObj,bonLunas:lunasReal,bonSebagian:r.status==="sebagian",totalDibayar:totalDibayarReal,isBon:!lunasReal,metodeBayar:lunasReal?"BON (LUNAS)":"BON (GABUNGAN)",sisaTagihan:lunasReal?0:r.sisaTagihan};
setInv(invObj);}else{setInv(makeBonInvObj(r));}}} title="Cetak Invoice BON" style={{background:C.inHv,border:"1px solid "+C.blt,borderRadius:7,padding:"6px 9px",color:C.blt,cursor:"pointer",fontSize:13}}>🖨️</button>{(r.status==="belum"||r.status==="sebagian")&&<button onClick={()=>{setOpenId(r.id);setBF({nominal:"",metode:"cash",bank:"BSI",salesPenerimaId:r.salesId||""});}} style={{background:C.grn,border:"none",borderRadius:7,padding:"6px 9px",color:"white",cursor:"pointer",fontSize:13}}>💳</button>}{(r.pembayaran||[]).length>0&&<button onClick={()=>setEditPayBon({bon:r,payIdx:null})} title="Edit/Cancel Pembayaran" style={{background:"#78350F",border:"1px solid #F59E0B",borderRadius:7,padding:"6px 9px",color:"#FCD34D",cursor:"pointer",fontSize:13}}>✏️</button>}<button onClick={()=>setDelId(r)} style={{background:C.inHvE,border:"1px solid "+C.rlt,borderRadius:7,padding:"6px 9px",color:C.rlt,cursor:"pointer",fontSize:13}}>🗑️</button></div>},
];
var bonActive=(data.bon||[]).filter(b=>b.status!=="lunas"&&b.status!=="digabung");
var totPiutang=bonActive.reduce((a,b)=>a+b.sisaTagihan,0);
return <div>
<STitle icon="💳" children="Piutang / BON"/>
<div style={{display:"flex",gap:10,flexWrap:"wrap",marginBottom:14}}>
<SC label="Total Bon Aktif" value={bonActive.length} icon="📃" color={C.olt}/>
<SC label="Total Piutang" value={fR(totPiutang)} icon="💰" color={C.rlt}/>
<SC label="Bon Lunas" value={(data.bon||[]).filter(b=>b.status==="lunas").length} icon="✅" color={C.glt}/>
</div>

{/* ── MODAL GABUNG INVOICE ── */}
{showGabung&&<div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.75)",zIndex:999,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
<div style={{background:C.card,borderRadius:12,width:"100%",maxWidth:600,maxHeight:"85vh",overflow:"hidden",display:"flex",flexDirection:"column",border:"1px solid "+C.bdr}}>
<div style={{padding:"14px 18px",borderBottom:"1px solid "+C.bdr,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
<div style={{fontWeight:700,color:C.wht,fontSize:14}}>🔗 Gabung Invoice BON</div>
<button onClick={()=>{setShowGabung(false);setGabungPilih([]);setGabungKons("");}} style={{background:"transparent",border:"none",color:C.gl2,cursor:"pointer",fontSize:18}}>✕</button>
</div>
<div style={{padding:"12px 18px",borderBottom:"1px solid "+C.bdr}}>
<Inp label="Filter Konsumen" value={gabungKons} onChange={setGabungKons} placeholder="Ketik nama konsumen..." style={{marginBottom:0}}/>

</div>
<div style={{flex:1,overflowY:"auto",padding:"12px 18px"}}>
{bonAktifGabung.length===0?<div style={{color:C.gl2,fontStyle:"italic",fontSize:12}}>Tidak ada BON aktif{gabungKons?" untuk konsumen ini":""}</div>:
bonAktifGabung.map(b=>{
var isPilih=gabungPilih.includes(b.id);
return <div key={b.id} onClick={()=>setGabungPilih(prev=>prev.includes(b.id)?prev.filter(x=>x!==b.id):[...prev,b.id])} style={{display:"flex",alignItems:"center",gap:10,padding:"10px 12px",background:isPilih?C.nav:C.bg,borderRadius:8,marginBottom:6,border:"2px solid "+(isPilih?C.blt:C.bdr),cursor:"pointer"}}>
<input type="checkbox" checked={isPilih} onChange={()=>{}} style={{width:16,height:16,cursor:"pointer"}}/>
<div style={{flex:1}}>
<div style={{fontWeight:700,color:C.wht,fontSize:12}}>{b.konsumen}</div>
<div style={{display:"flex",gap:6,alignItems:"center",flexWrap:"wrap"}}>
<span style={{fontSize:10,color:C.gl2}}>{b.noInv} · {fDs(b.tanggal)}</span>
{b.status==="lunas"?<span style={{fontSize:9,fontWeight:700,color:"#065F46",background:"#D1FAE5",borderRadius:10,padding:"1px 6px"}}>✓ Lunas</span>:b.status==="sebagian"?<span style={{fontSize:9,fontWeight:700,color:"#92400E",background:"#FEF3C7",borderRadius:10,padding:"1px 6px"}}>Sebagian</span>:<span style={{fontSize:9,fontWeight:700,color:"#991B1B",background:"#FEE2E2",borderRadius:10,padding:"1px 6px"}}>Belum Bayar</span>}
</div>
<div style={{fontSize:10,color:C.gl2}}>{(b.items||[]).map(it=>it.qty+"×"+it.ukuran).join(", ")}</div>
</div>
<div style={{textAlign:"right"}}>
<div style={{fontWeight:700,color:C.rlt,fontSize:13}}>{fR(b.sisaTagihan)}</div>
<div style={{fontSize:9,color:C.gl2}}>sisa tagihan</div>
</div>
</div>;
})}
</div>
{gabungPilih.length>0&&<div style={{padding:"10px 18px",borderTop:"1px solid "+C.bdr,background:C.nav}}>
<div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}>
<span style={{fontSize:12,color:C.gl2}}>{gabungPilih.length} BON dipilih</span>
<div style={{textAlign:"right"}}>
{(()=>{
var pilihBons=(data.bon||[]).filter(b=>gabungPilih.includes(b.id));
var ttlAll=pilihBons.reduce((a,b)=>a+(b.total||0),0);
var ttlSisa=pilihBons.reduce((a,b)=>a+(b.status==="lunas"?0:(b.sisaTagihan||0)),0);
var adaLunas=pilihBons.some(b=>b.status==="lunas");
return <><div style={{fontSize:14,fontWeight:900,color:C.wht}}>Total: {fR(ttlAll)}</div>
{adaLunas&&ttlSisa>0&&<div style={{fontSize:11,color:C.rlt}}>Sisa tagihan: {fR(ttlSisa)}</div>}
{adaLunas&&ttlSisa===0&&<div style={{fontSize:11,color:C.glt}}>✓ Semua sudah lunas</div>}
</>;
})()}
</div>
</div>
<Btn color="blue" onClick={doGabung} dis={gabungPilih.length<2}>🔗 Gabung & Buat Invoice Baru ({gabungPilih.length} BON)</Btn>
</div>}
</div>
</div>}

<Card>
<div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
<div style={{fontWeight:700,color:C.gl2,fontSize:13}}>📋 Daftar Bon</div>
<Btn sm color="blue" onClick={()=>setShowGabung(true)}>🔗 Gabung Invoice</Btn>
</div>
<div style={{background:C.nav,borderRadius:8,padding:10,marginBottom:10,border:"1px solid "+C.bdr}}>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(140px,1fr))",gap:8}}>
<Inp label="Dari" type="date" value={barFilter.from} onChange={v=>setBarFilter(p=>({...p,from:v}))} style={{marginBottom:0}}/>
<Inp label="Sampai" type="date" value={barFilter.to} onChange={v=>setBarFilter(p=>({...p,to:v}))} style={{marginBottom:0}}/>
<Sel label="Sales" value={barFilter.salesId} onChange={v=>setBarFilter(p=>({...p,salesId:v}))} opts={[{v:"",l:"Semua"},...salesList.map(e=>({v:e.id,l:e.nama}))]} style={{marginBottom:0}}/>
<Inp label="Konsumen" value={barFilter.konsumen} onChange={v=>setBarFilter(p=>({...p,konsumen:v}))} placeholder="Cari..." style={{marginBottom:0}}/>
<Sel label="Status" value={barFilter.status} onChange={v=>setBarFilter(p=>({...p,status:v}))} opts={tabBon==="lunas"?[{v:"",l:"Semua"}]:[{v:"",l:"Semua"},{v:"belum",l:"Belum"},{v:"sebagian",l:"Sebagian"},{v:"digabung",l:"Digabung"}]} style={{marginBottom:0}}/>
</div>
{(barFilter.from||barFilter.to||barFilter.salesId||barFilter.konsumen||barFilter.status)&&<div style={{marginTop:8}}><Btn sm color="gray" onClick={()=>setBarFilter({from:"",to:"",salesId:"",konsumen:"",status:""})}>✕ Reset Filter</Btn></div>}
</div>
{/* Tombol Cetak Laporan BON */}
<div style={{display:"flex",gap:8,marginBottom:14}}>
{[["aktif","📋 Daftar Bon ("+rowsAktif.length+")"],["lunas","✅ Riwayat BON Lunas ("+rowsLunas.length+")"]].map(x=><button key={x[0]} onClick={()=>setTabBon(x[0])} style={{background:tabBon===x[0]?C.blu:C.nav,color:tabBon===x[0]?"white":C.wht,border:"1px solid "+(tabBon===x[0]?C.blt:C.bdr),borderRadius:8,padding:"7px 14px",fontWeight:700,fontSize:12,cursor:"pointer"}}>{x[1]}</button>)}
</div>
<div style={{marginBottom:10,display:"flex",gap:8,alignItems:"center"}}>
<button onClick={()=>{
var el=document.getElementById("_lap_bon");if(el)el.remove();
// Build print content
var salesGroups={};
rows.filter(b=>b.status!=="lunas").forEach(b=>{var sn=b.salesNama||"(Tanpa Sales)";if(!salesGroups[sn])salesGroups[sn]=[];salesGroups[sn].push(b);});
var totalSisa=rows.filter(b=>b.status!=="lunas").reduce((a,b)=>a+(b.sisaTagihan||0),0);
var filterInfo=[(barFilter.from||barFilter.to)?("Periode: "+(barFilter.from?fDs(barFilter.from):"")+(barFilter.to?" s/d "+fDs(barFilter.to):"")):"",barFilter.salesId?("Sales: "+(salesList.find(e=>e.id===barFilter.salesId)?.nama||"")):"",barFilter.konsumen?("Konsumen: "+barFilter.konsumen):"",barFilter.status?("Status: "+barFilter.status):""].filter(Boolean).join(" | ")||"Semua Data";
var printDiv=document.createElement("div");printDiv.id="_lap_bon";
var html='<div style="font-family:Arial,sans-serif;padding:20px;color:#111;max-width:900px;margin:0 auto">'+
'<div style="text-align:center;border-bottom:3px solid #0a1f44;padding-bottom:10px;margin-bottom:14px">'+
'<div style="font-size:18px;font-weight:900;color:#0a1f44">LAPORAN PIUTANG / BON AKTIF</div>'+
'<div style="font-size:13px;font-weight:700;color:#0a1f44">'+(data.company?.nama||"PT. HOE TRANG SA").toUpperCase()+'</div>'+
'<div style="font-size:11px;color:#555;margin-top:4px">'+filterInfo+' | Dicetak: '+new Date().toLocaleString("id-ID")+'</div>'+
'</div>';
Object.entries(salesGroups).forEach(function([sn,bons]){
var subTotal=bons.reduce(function(a,b){return a+(b.sisaTagihan||0);},0);
html+='<div style="margin-bottom:14px">';
html+='<div style="font-weight:700;font-size:12px;color:#0a1f44;background:#EFF6FF;padding:5px 8px;border-radius:4px;margin-bottom:6px">Sales: '+sn+'</div>';
html+='<table style="width:100%;border-collapse:collapse;font-size:11px"><thead><tr style="background:#0a1f44">';
['No','Konsumen','No.Invoice','Tanggal','Jatuh Tempo','Sisa Tagihan','Status'].forEach(function(h){html+='<th style="color:white;padding:5px 7px;text-align:'+(h==='Sisa Tagihan'?'right':'left');if(h==='No')html+=';width:30px';html+='">'+h+'</th>';});
html+='</tr></thead><tbody>';
bons.forEach(function(b,i){
var dl=dLeft(b.deadline);
var dlTxt=b.deadline?(fDs(b.deadline)+(b.status!=="lunas"&&dl!=null?" ("+(dl<0?Math.abs(dl)+"h LEWAT":dl+"h")+")":'')):'—';
html+='<tr style="background:'+(i%2===0?'white':'#f9f9f9')+'">';
html+='<td style="padding:4px 7px;border:1px solid #ddd">'+(i+1)+'</td>';
html+='<td style="padding:4px 7px;border:1px solid #ddd;font-weight:600">'+b.konsumen+'</td>';
html+='<td style="padding:4px 7px;border:1px solid #ddd">'+( b.noInv||'-')+'</td>';
html+='<td style="padding:4px 7px;border:1px solid #ddd">'+fDs(b.tanggal)+'</td>';
html+='<td style="padding:4px 7px;border:1px solid #ddd;color:'+(dl!==null&&dl<0?'#DC2626':dl!==null&&dl<=3?'#D97706':'#111')+'">'+dlTxt+'</td>';
html+='<td style="padding:4px 7px;border:1px solid #ddd;text-align:right;font-weight:700;color:#DC2626">'+fR(b.sisaTagihan||0)+'</td>';
html+='<td style="padding:4px 7px;border:1px solid #ddd">'+b.status+'</td>';
html+='</tr>';});
html+='<tr style="background:#DBEAFE;font-weight:700"><td colspan="5" style="padding:5px 7px;border:1px solid #ddd">Sub-total '+sn+'</td><td style="padding:5px 7px;border:1px solid #ddd;text-align:right;color:#DC2626">'+fR(subTotal)+'</td><td style="border:1px solid #ddd"></td></tr>';
html+='</tbody></table></div>';});
html+='<div style="background:#0a1f44;color:white;padding:8px 12px;border-radius:6px;display:flex;justify-content:space-between;font-weight:700;font-size:13px"><span>TOTAL PIUTANG ('+rows.filter(b=>b.status!=="lunas").length+' bon aktif)</span><span>'+fR(totalSisa)+'</span></div>';
html+='</div>';
printDiv.innerHTML=html;
document.body.appendChild(printDiv);
doPrint("_lap_bon");
setTimeout(function(){var e=document.getElementById("_lap_bon");if(e)e.remove();},3000);
}} style={{background:"#0a1f44",color:"white",border:"none",padding:"8px 16px",borderRadius:7,fontSize:12,cursor:"pointer",fontWeight:700}}>🖨️ Cetak Laporan BON</button>
<span style={{fontSize:10,color:C.gl2,fontStyle:"italic"}}>Cetak sesuai filter aktif — hanya bon belum lunas</span>
</div>
<FilterTbl columns={cols} data={rowsTab} empty={tabBon==="lunas"?"Belum ada riwayat BON lunas":"Belum ada bon aktif"} maxRows={150}
footerRow={sorted=>{
if(tabBon==="lunas"){
  var totalLunas=sorted.reduce((a,r)=>a+(r.total||0),0);
  if(!totalLunas)return null;
  var nCols2=cols.length;
  return <tr style={{background:C.nav,borderTop:"2px solid "+C.bdr}}>
  <td colSpan={5} style={{padding:"8px 10px",color:C.gl2,fontSize:11,fontWeight:600}}>{sorted.length} bon lunas</td>
  <td style={{padding:"8px 10px",textAlign:"left",color:C.glt,fontWeight:900,fontSize:14}}>{fR(totalLunas)}</td>
  <td colSpan={nCols2-6}/>
  </tr>;
}
var aktif=sorted.filter(r=>r.status!=="lunas"&&r.status!=="digabung");
var totalSisa=aktif.reduce((a,r)=>a+(r.sisaTagihan||0),0);
if(!totalSisa)return null;
// cari index kolom sisa (index 5)
var nCols=cols.length;
return <tr style={{background:C.nav,borderTop:"2px solid "+C.bdr}}>
<td colSpan={5} style={{padding:"8px 10px",color:C.gl2,fontSize:11,fontWeight:600}}>{aktif.length} bon aktif</td>
<td style={{padding:"8px 10px",textAlign:"left",color:C.rlt,fontWeight:900,fontSize:14}}>{fR(totalSisa)}</td>
<td colSpan={nCols-6}/>
</tr>;
}}/>
</Card>
{openId&&(()=>{var b=(data.bon||[]).find(x=>x.id===openId);if(!b)return null;var paid=(b.pembayaran||[]).reduce((a,p)=>a+p.jumlah,0);return <Card style={{border:"1px solid "+C.blt}}>
<div style={{display:"flex",justifyContent:"space-between",flexWrap:"wrap",gap:8,marginBottom:8}}>
<div><div style={{fontWeight:800,color:C.wht,fontSize:15}}>💳 Bayar: {b.konsumen}</div><div style={{fontSize:11,color:C.gl2}}>{fDs(b.tanggal)} · {b.noInv}</div></div>
<Btn sm color="gray" onClick={()=>setOpenId(null)}>Batal</Btn>
</div>
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:10}}>{[["Total",b.total,C.wht],["Dibayar",paid,C.glt],["Sisa",b.sisaTagihan,C.rlt]].map(x=><div key={x[0]} style={{background:C.nav,borderRadius:8,padding:"6px 10px",textAlign:"center",border:"1px solid "+C.bdr}}><div style={{fontSize:10,color:C.gl2}}>{x[0]}</div><div style={{fontSize:13,fontWeight:900,color:x[2]}}>{fR(x[1])}</div></div>)}</div>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(180px,260px))",gap:10,marginBottom:10}}>
<div><Inp label="Nominal" type="number" value={bF.nominal} onChange={v=>setBF(p=>({...p,nominal:v}))} placeholder={String(b.sisaTagihan)} style={{marginBottom:4}}/><button onClick={()=>setBF(p=>({...p,nominal:String(b.sisaTagihan)}))} style={{background:C.inHv,border:"1px solid "+C.blt,borderRadius:5,padding:"3px 8px",color:C.blt,cursor:"pointer",fontSize:11}}>Isi Semua</button></div>
<Sel label="Sales Penerima" value={bF.salesPenerimaId} onChange={v=>setBF(p=>({...p,salesPenerimaId:v}))} opts={[{v:"",l:"Admin/Kantor"},...salesList.map(e=>({v:e.id,l:e.nama}))]}/>
</div>
<div style={{display:"flex",gap:6,marginBottom:8}}>{["cash","transfer"].map(m=><button key={m} onClick={()=>setBF(p=>({...p,metode:m}))} style={{background:bF.metode===m?C.blu:C.nav,color:bF.metode===m?"white":C.wht,border:"1px solid "+(bF.metode===m?C.blt:C.bdr),borderRadius:6,padding:"7px 12px",fontWeight:700,fontSize:12,cursor:"pointer",flex:1}}>{m==="cash"?"💵 Cash":"🏦 Transfer"}</button>)}</div>
<Btn color="green" onClick={()=>bayar(b)} dis={!bF.nominal}>💾 Catat Pembayaran</Btn>
</Card>;})()}
{/* Total sisa mengikuti filter */}
{(()=>{
// rows sudah terfilter barFilter dari useMemo
// hanya exclude lunas & digabung untuk rekap aktif
var rowsFiltered=rows.filter(r=>r.status!=="lunas"&&r.status!=="digabung");
var totalSisaFilter=rowsFiltered.reduce((a,r)=>a+(r.sisaTagihan||0),0);
var totalFilter=rowsFiltered.reduce((a,r)=>a+(r.total||0),0);
return <Card style={{border:"1px solid "+C.rlt}}>
<div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
<div style={{fontWeight:700,color:C.gl2,fontSize:13}}>📊 Rekap Piutang Aktif{barFilter.konsumen||barFilter.salesId||barFilter.from||barFilter.status?" — filter aktif":""}</div>
<div style={{fontSize:10,color:C.gl2,fontStyle:"italic"}}>* mengikuti filter atas</div>
</div>
<div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:8}}>
{[["Jumlah BON",rowsFiltered.length+" bon",C.wht],["Total Invoice",fR(totalFilter),C.olt],["Total Sisa Tagihan",fR(totalSisaFilter),C.rlt]].map(x=><div key={x[0]} style={{background:C.nav,borderRadius:8,padding:"10px 14px",border:"1px solid "+C.bdr}}>
<div style={{fontSize:10,color:C.gl2,marginBottom:3}}>{x[0]}</div>
<div style={{fontSize:14,fontWeight:900,color:x[2]}}>{x[1]}</div>
</div>)}
</div>
</Card>;
})()}
{delId&&<ConfirmDel msg="Hapus bon?" onCancel={()=>setDelId(null)} onConfirm={()=>{setData(d=>({...d,bon:(d.bon||[]).filter(x=>x.id!==delId.id)}));setDelId(null);}}/>}

{/* ── MODAL EDIT/CANCEL PEMBAYARAN BON ── */}
{editPayBon&&<div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.8)",zIndex:999,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
<div style={{background:C.card,borderRadius:12,width:"100%",maxWidth:520,border:"1px solid "+C.bdr,maxHeight:"85vh",overflowY:"auto"}}>
<div style={{padding:"14px 18px",borderBottom:"1px solid "+C.bdr,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
<div style={{fontWeight:700,color:C.wht,fontSize:14}}>✏️ Riwayat Pembayaran — {editPayBon.bon.konsumen}</div>
<button onClick={()=>setEditPayBon(null)} style={{background:"transparent",border:"none",color:C.gl2,cursor:"pointer",fontSize:20}}>✕</button>
</div>
<div style={{padding:16}}>
<div style={{background:C.nav,borderRadius:8,padding:"8px 12px",marginBottom:12,display:"flex",justifyContent:"space-between"}}>
<span style={{fontSize:12,color:C.gl2}}>Invoice: <b style={{color:C.wht}}>{editPayBon.bon.noInv}</b></span>
<span style={{fontSize:12,color:C.gl2}}>Total: <b style={{color:C.rlt}}>{fR(editPayBon.bon.total)}</b></span>
</div>
{(editPayBon.bon.pembayaran||[]).length===0
?<div style={{color:C.gl2,fontSize:12,fontStyle:"italic"}}>Belum ada pembayaran</div>
:(editPayBon.bon.pembayaran||[]).map((pay,pi)=>{
var isEdit=editPayBon.payIdx===pi;
var bonFresh=(data.bon||[]).find(b=>b.id===editPayBon.bon.id)||editPayBon.bon;
var payFresh=(bonFresh.pembayaran||[])[pi]||pay;
return <div key={pi} style={{background:isEdit?C.nav:C.bg,borderRadius:8,padding:"10px 12px",marginBottom:8,border:"2px solid "+(isEdit?C.olt:C.bdr)}}>
<div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:isEdit?8:0}}>
<div>
<div style={{fontSize:13,fontWeight:700,color:C.wht}}>{fR(payFresh.jumlah||payFresh.nominal||0)}</div>
<div style={{fontSize:10,color:C.gl2}}>{fDs(payFresh.tanggal)} · <span style={{color:(payFresh.metode||"").toLowerCase()==="cash"?C.glt:C.blt,fontWeight:700}}>{payFresh.metode||"Cash"}</span>{payFresh.bank?" · "+payFresh.bank:""}</div>
{payFresh.salesPenerimaNama&&<div style={{fontSize:10,color:C.gl2}}>Diterima: {payFresh.salesPenerimaNama}</div>}
</div>
<div style={{display:"flex",gap:5}}>
<button onClick={()=>setEditPayBon(p=>({...p,payIdx:isEdit?null:pi,editForm:{...payFresh,jumlah:payFresh.jumlah||payFresh.nominal||0}}))} style={{background:isEdit?C.olt:C.nav,border:"1px solid "+(isEdit?C.olt:C.bdr),borderRadius:6,padding:"4px 9px",color:isEdit?"white":C.gl2,cursor:"pointer",fontSize:11,fontWeight:700}}>{isEdit?"▲":"✏️ Edit"}</button>
<button onClick={()=>{
var b2=(data.bon||[]).find(b=>b.id===editPayBon.bon.id);
if(!b2){toast("BON tidak ditemukan");return;}
var newPays=(b2.pembayaran||[]).filter((_,idx)=>idx!==pi);
var newSisa=Math.max(0,b2.total-newPays.reduce((a,p)=>a+(Number(p.jumlah||p.nominal||0)),0));
var newStatus=newPays.length===0?"belum":newSisa<=0?"lunas":"sebagian";
var log={id:uid(),type:"cancel_pay",by:"Admin",at:new Date().toISOString(),before:{jumlah:payFresh.jumlah||payFresh.nominal||0,metode:payFresh.metode,tanggal:payFresh.tanggal},note:"Pembayaran dicancel"};
var updBon={...b2,pembayaran:newPays,sisaTagihan:newSisa,status:newStatus,editLog:[...(b2.editLog||[]),log]};
setData(d=>({...d,bon:(d.bon||[]).map(b=>b.id===updBon.id?updBon:b)}));
setEditPayBon({bon:updBon,payIdx:null,editForm:null});
toast("✓ Dicancel. Status: "+newStatus);
}} style={{background:C.rdk,border:"1px solid "+C.rlt,borderRadius:6,padding:"4px 9px",color:"white",cursor:"pointer",fontSize:11,fontWeight:700}}>🗑️ Cancel</button>
</div>
</div>
{isEdit&&editPayBon.editForm&&<div>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(180px,260px))",gap:8,marginBottom:8}}>
<div><div style={{fontSize:10,color:C.gl2,marginBottom:3}}>Nominal</div>
<input type="number" value={editPayBon.editForm.jumlah||""} onChange={e=>setEditPayBon(p=>({...p,editForm:{...p.editForm,jumlah:Number(e.target.value)||0}}))} style={{background:C.bg,border:"1px solid "+C.bdr,borderRadius:6,padding:"7px 9px",color:C.wht,fontSize:12,outline:"none",width:"100%"}}/></div>
<div><div style={{fontSize:10,color:C.gl2,marginBottom:3}}>Tanggal</div>
<input type="date" value={editPayBon.editForm.tanggal||""} onChange={e=>setEditPayBon(p=>({...p,editForm:{...p.editForm,tanggal:e.target.value}}))} style={{background:C.bg,border:"1px solid "+C.bdr,borderRadius:6,padding:"7px 9px",color:C.wht,fontSize:12,outline:"none",width:"100%"}}/></div>
</div>
<div style={{marginBottom:8}}>
<div style={{fontSize:10,color:C.gl2,marginBottom:4}}>Metode</div>
<div style={{display:"flex",gap:6}}>
{[["cash","💵 Cash",C.grn],["transfer","🏦 Transfer",C.blu]].map(x=><button key={x[0]} onClick={()=>setEditPayBon(p=>({...p,editForm:{...p.editForm,metode:x[0]}}))} style={{background:(editPayBon.editForm.metode||"cash")===x[0]?x[2]:C.nav,color:(editPayBon.editForm.metode||"cash")===x[0]?"white":C.wht,border:"1px solid "+((editPayBon.editForm.metode||"cash")===x[0]?x[2]:C.bdr),borderRadius:7,padding:"5px 14px",fontWeight:700,fontSize:12,cursor:"pointer"}}>{x[1]}</button>)}
</div>
{(editPayBon.editForm.metode||"cash")==="transfer"&&<div style={{display:"flex",gap:6,marginTop:6}}>{["BSI","BCA"].map(bk=><button key={bk} onClick={()=>setEditPayBon(p=>({...p,editForm:{...p.editForm,bank:bk}}))} style={{background:(editPayBon.editForm.bank||"BSI")===bk?C.blu:C.nav,color:(editPayBon.editForm.bank||"BSI")===bk?"white":C.wht,border:"2px solid "+((editPayBon.editForm.bank||"BSI")===bk?C.blt:C.bdr),borderRadius:7,padding:"4px 14px",fontWeight:700,cursor:"pointer",fontSize:12}}>{bk}</button>)}</div>}
</div>
<button onClick={()=>{
var b2=(data.bon||[]).find(b=>b.id===editPayBon.bon.id);
if(!b2){toast("BON tidak ditemukan");return;}
var pi2=editPayBon.payIdx;
var ef2=editPayBon.editForm;
if(pi2===null||pi2===undefined||!ef2){toast("Pilih pembayaran yang akan diedit");return;}
var jumlahBaru=Number(ef2.jumlah||0);
if(jumlahBaru<=0){toast("Nominal harus lebih dari 0");return;}
var log={id:uid(),type:"edit_pay",by:"Admin",at:new Date().toISOString(),before:{...(b2.pembayaran||[])[pi2]},note:"Pembayaran diedit: "+fR((b2.pembayaran||[])[pi2]?.jumlah||0)+" → "+fR(jumlahBaru)};
var newPays=(b2.pembayaran||[]).map((p,idx)=>idx===pi2?{...p,jumlah:jumlahBaru,nominal:jumlahBaru,metode:ef2.metode||p.metode,bank:ef2.bank||p.bank,tanggal:ef2.tanggal||p.tanggal}:p);
var newSisa=Math.max(0,b2.total-newPays.reduce((a,p)=>a+(Number(p.jumlah||p.nominal||0)),0));
var newStatus=newSisa<=0?"lunas":newPays.length===0?"belum":"sebagian";
var updBon={...b2,pembayaran:newPays,sisaTagihan:newSisa,status:newStatus,editLog:[...(b2.editLog||[]),log]};
setData(d=>({...d,bon:(d.bon||[]).map(b=>b.id===updBon.id?updBon:b)}));
setEditPayBon({bon:updBon,payIdx:null,editForm:null});
toast("✓ Disimpan! Status: "+newStatus);
}} style={{background:C.glt,border:"none",borderRadius:8,padding:"7px 18px",color:"white",cursor:"pointer",fontWeight:700,fontSize:12,width:"100%"}}>💾 Simpan Perubahan</button>
</div>}
</div>;})}
{/* Log perubahan */}
{((data.bon||[]).find(b=>b.id===editPayBon.bon.id)?.editLog||[]).length>0&&<div style={{marginTop:10,background:C.bg,borderRadius:8,padding:10,border:"1px solid "+C.bdr}}>
<div style={{fontSize:11,fontWeight:700,color:C.gl2,marginBottom:6}}>📋 Log Perubahan</div>
{((data.bon||[]).find(b=>b.id===editPayBon.bon.id)?.editLog||[]).slice().reverse().map((lg,i)=><div key={i} style={{fontSize:10,color:C.gl2,marginBottom:4,paddingBottom:4,borderBottom:"1px solid "+C.bdr}}>
<b style={{color:C.wht}}>{lg.by}</b> · {new Date(lg.at).toLocaleString("id-ID")} · <span style={{color:lg.type==="cancel_pay"?C.rlt:C.olt}}>{lg.note}</span>
{lg.before&&<><br/><span style={{color:"#9CA3AF"}}>Sebelum: {fR(lg.before.jumlah||0)} · {lg.before.metode} · {fDs(lg.before.tanggal)}</span></>}
</div>)}
</div>}
</div>
</div>
</div>}
</div>;
}

// ─── PAYROLL v4 (NO O/X toggle in editor, role-aware default rows) ────────────
function PayrollMod({data,setData,toast}){
var C=useTheme();
var[bulan,setBulan]=useState(toMonth());
var[viewSlip,setViewSlip]=useState(null);
var[editSlip,setEditSlip]=useState(null);
var emps=sortEmp((data.employees||[]).filter(e=>e.aktif));

function buildSlipRows(emp){
var r=calcPayrollFull(emp,bulan,data);
var uangMakanRate=emp.uangMakan||UANG_MAKAN_DEFAULT;
var rows=[];
// ── Penghasilan ──
rows.push({id:uid(),section:"penghasilan",label:"Gaji Pokok",qty:1,ket:"",jumlah:r.gajiPokok});
rows.push({id:uid(),section:"penghasilan",label:"Intensif 12kg",qty:r.bonus.q12,ket:fR(r.bonus.r12)+"/tabung",jumlah:r.bonus.b12});
rows.push({id:uid(),section:"penghasilan",label:"Intensif 5,5kg",qty:r.bonus.q55,ket:fR(r.bonus.r55)+"/tabung",jumlah:r.bonus.b55});
rows.push({id:uid(),section:"penghasilan",label:"Bonus",qty:1,ket:"",jumlah:0});
rows.push({id:uid(),section:"penghasilan",label:"Uang Makan",qty:r.hariHadir,ket:fR(uangMakanRate)+"/hari (dari absen)",jumlah:r.uangMakan});
rows.push({id:uid(),section:"penghasilan",label:"Uang Bongkar DO",qty:r.bongkarCount,ket:"Rp 50.000/DO",jumlah:r.bongkarTotal});
rows.push({id:uid(),section:"penghasilan",label:"Uang Do SPBBE",qty:r.spbbeCount,ket:"per trip",jumlah:r.spbbeTotal});
// ── Potongan ──
var potAbsen=r.absen>0?Math.round(r.gajiPokok/r.totalHariKerja)*r.absen:0;
rows.push({id:uid(),section:"potongan",label:"Potongan Absensi",qty:r.absen,ket:fR(r.absen>0?Math.round(r.gajiPokok/r.totalHariKerja):0)+"/hari",jumlah:potAbsen});
rows.push({id:uid(),section:"potongan",label:"Total Pinjaman Berjalan",qty:1,ket:"saldo s/d bulan ini",jumlah:r.pinjamanSaldo,kind:"info"});
rows.push({id:uid(),section:"potongan",label:"Potongan Pinjaman",qty:1,ket:"cicilan bulan ini",jumlah:0});
// ── Yang sudah diterima ──
rows.push({id:uid(),section:"ydt",label:"Uang Makan",qty:1,ket:"diambil tengah bulan",jumlah:0});
return{empId:emp.id,nama:emp.nama,posisi:emp.posisi,alamat:emp.alamat||"",telepon:emp.telepon||"",bulan,tanggal:toDay(),hariHadir:r.hariHadir,totalHariKerja:r.totalHariKerja,absen:r.absen,totalPinjaman:r.pinjamanSaldo,potonganPinjaman:0,rows};
}

function SlipEditor({slip,onClose,onSave}){
var[s,setS]=useState(JSON.parse(JSON.stringify(slip)));
function rR(n){return"Rp "+Number(n||0).toLocaleString("id-ID");}
var totPgh=s.rows.filter(r=>r.section==="penghasilan").reduce((a,r)=>a+Number(r.jumlah||0),0);
var totPot=s.rows.filter(r=>r.section==="potongan"&&r.kind!=="info").reduce((a,r)=>a+Number(r.jumlah||0),0);
var totYdt=s.rows.filter(r=>r.section==="ydt").reduce((a,r)=>a+Number(r.jumlah||0),0);
var sisaPinjaman=Math.max(0,Number(s.totalPinjaman||0)-Number(s.potonganPinjaman||0));
var totDiterima=totPgh-totPot-totYdt;
function updRow(id,k,v){setS(p=>{
  var newRows=p.rows.map(r=>r.id===id?{...r,[k]:v}:r);
  var editedRow=p.rows.find(r=>r.id===id);
  // Jika yang diedit adalah baris "Potongan Pinjaman" kolom jumlah, sync balik ke s.potonganPinjaman
  if(editedRow&&editedRow.label==="Potongan Pinjaman"&&k==="jumlah"){
    return{...p,rows:newRows,potonganPinjaman:Number(v)||0};
  }
  return{...p,rows:newRows};
});}
function addRow(section){setS(p=>({...p,rows:[...p.rows,{id:uid(),section,label:"Item baru",qty:1,ket:"",jumlah:0}]}));}
function delRow(id){setS(p=>({...p,rows:p.rows.filter(r=>r.id!==id)}));}
function updPotPinjaman(v){var n=Number(v)||0;setS(p=>({...p,potonganPinjaman:n,rows:p.rows.map(r=>r.label==="Potongan Pinjaman"?{...r,jumlah:n}:r)}));}
var iStyle={background:"transparent",border:"1px solid",borderRadius:6,padding:"4px 7px",color:"inherit",fontSize:11,outline:"none",width:"100%",boxSizing:"border-box"};
return <Modal title={"✏️ Edit Kwitansi — "+s.nama} onClose={onClose} width={860} saveLabel="💾 Simpan & Cetak" onSave={()=>onSave({...s})}>
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10,marginBottom:12}}>
<Inp label="Total Hari Kerja" type="number" value={s.totalHariKerja} onChange={v=>setS(p=>({...p,totalHariKerja:Number(v)}))}/>
<Inp label="Hadir (dari Absensi)" type="number" value={s.hariHadir} onChange={v=>setS(p=>({...p,hariHadir:Number(v)}))}/>
<Inp label="Absen/Alpha" type="number" value={s.absen} onChange={v=>setS(p=>({...p,absen:Number(v)}))}/>
</div>
{[["penghasilan","💚 Penghasilan","#22c55e"],["potongan","🔴 Potongan","#ef4444"],["ydt","🟠 Yang Sudah Diterima","#f59e0b"]].map(([sec,title,col])=><div key={sec} style={{marginBottom:12,border:"1px solid #334155",borderRadius:8,overflow:"hidden"}}>
<div style={{padding:"7px 12px",background:"#1e293b",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
<b style={{color:col,fontSize:12}}>{title}</b>
<Btn sm color="blue" onClick={()=>addRow(sec)}>+ Baris</Btn>
</div>
{/* header kolom */}
<div style={{display:"grid",gridTemplateColumns:"2fr 70px 1.5fr 110px 28px",gap:4,padding:"4px 10px",background:"#0f172a"}}>
{["Label","Qty","Keterangan","Jumlah (Rp)",""].map((h,i)=><span key={i} style={{fontSize:9,color:"#64748b",fontWeight:700,textAlign:i===3?"right":"left"}}>{h}</span>)}
</div>
{s.rows.filter(r=>r.section===sec).length===0&&<div style={{padding:10,color:"#475569",fontSize:11,textAlign:"center"}}>Belum ada baris</div>}
{s.rows.filter(r=>r.section===sec).map(r=>{
var isPinjamanRow=r.label==="Potongan Pinjaman";
var lockField=r.kind==="info"||isPinjamanRow;
return <div key={r.id} style={{display:"grid",gridTemplateColumns:"2fr 70px 1.5fr 110px 28px",gap:4,padding:"5px 10px",borderTop:"1px solid #1e293b",alignItems:"center",background:r.kind==="info"?"#1e1a09":isPinjamanRow?"#1a1f2e":"transparent"}}>
<input value={r.label} onChange={e=>updRow(r.id,"label",e.target.value)} style={{...iStyle,borderColor:"#334155",color:isPinjamanRow?"#94a3b8":"#e2e8f0"}} readOnly={isPinjamanRow}/>
<input type="number" value={r.qty??1} onChange={e=>updRow(r.id,"qty",Number(e.target.value))} style={{...iStyle,borderColor:"#334155",color:"#94a3b8",textAlign:"center"}} disabled={lockField}/>
<input value={isPinjamanRow?"⬇️ isi di kotak bawah":(r.ket||"")} onChange={e=>updRow(r.id,"ket",e.target.value)} placeholder="keterangan" style={{...iStyle,borderColor:"#334155",color:"#94a3b8",fontStyle:isPinjamanRow?"italic":"normal"}} disabled={lockField}/>
<input type="number" value={r.jumlah||0} onChange={e=>updRow(r.id,"jumlah",Number(e.target.value))} title={isPinjamanRow?"Isi lewat kotak \"Cicilan/Potongan Bulan Ini\" di bawah":""} style={{...iStyle,borderColor:"transparent",color:r.kind==="info"?"#f59e0b":isPinjamanRow?"#94a3b8":col,textAlign:"right",fontWeight:700,background:"transparent",cursor:lockField?"not-allowed":"text"}} readOnly={lockField}/>
<button onClick={()=>delRow(r.id)} style={{background:"transparent",border:"none",color:"#ef4444",cursor:"pointer",fontSize:13,padding:0}} title="Hapus">✕</button>
</div>;})}
</div>)}
{/* Sistem Pinjaman Berjalan */}
<div style={{background:"#1c1a05",border:"1px solid #f59e0b",borderRadius:8,padding:"12px 14px",marginBottom:12}}>
<div style={{fontWeight:700,color:"#f59e0b",marginBottom:10,fontSize:12}}>💰 Sistem Pinjaman Berjalan</div>
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10}}>
<div>
<label style={{display:"block",fontSize:10,color:"#94a3b8",marginBottom:3,fontWeight:600}}>Total Pinjaman (saldo berjalan)</label>
<div style={{background:"#334155",padding:"9px 11px",borderRadius:8,fontWeight:800,fontSize:13,color:"#f59e0b"}}>{rR(s.totalPinjaman)}</div>
<div style={{fontSize:9,color:"#64748b",marginTop:3}}>Semua ambilan s/d bulan ini − cicilan sebelumnya</div>
</div>
<div>
<Inp label="Cicilan/Potongan Bulan Ini" type="number" value={s.potonganPinjaman} onChange={updPotPinjaman}/>
<div style={{fontSize:9,color:"#64748b",marginTop:-6}}>Otomatis update baris "Potongan Pinjaman" di tabel</div>
</div>
<div>
<label style={{display:"block",fontSize:10,color:"#94a3b8",marginBottom:3,fontWeight:600}}>Sisa Pinjaman Setelah Potong</label>
<div style={{background:"#fef08a",color:"#1a1a1a",padding:"9px 11px",borderRadius:8,fontWeight:800,fontSize:14}}>{rR(sisaPinjaman)}</div>
</div>
</div>
</div>
{/* Ringkasan total */}
<div style={{background:"#0f1f0f",borderRadius:10,padding:"14px 16px",border:"1px solid #22c55e"}}>
<div style={{display:"flex",justifyContent:"space-between",marginBottom:5}}><span style={{color:"#64748b",fontSize:12}}>Total Penghasilan (+)</span><b style={{color:"#22c55e"}}>{rR(totPgh)}</b></div>
<div style={{display:"flex",justifyContent:"space-between",marginBottom:5}}><span style={{color:"#64748b",fontSize:12}}>Total Potongan (−)</span><b style={{color:"#ef4444"}}>{rR(totPot)}</b></div>
<div style={{display:"flex",justifyContent:"space-between",marginBottom:8}}><span style={{color:"#64748b",fontSize:12}}>Yang sudah diterima (−)</span><b style={{color:"#f59e0b"}}>{rR(totYdt)}</b></div>
<div style={{display:"flex",justifyContent:"space-between",borderTop:"1px solid #22c55e",paddingTop:8}}><b style={{color:"#e2e8f0",fontSize:14}}>TOTAL DITERIMA</b><b style={{color:"#22c55e",fontSize:18}}>{rR(totDiterima)}</b></div>
</div>
</Modal>;
}

return <div>
<STitle icon="💼" children="Payroll & Kwitansi Slip Gaji"/>
<Card><MonthPicker label="Pilih Bulan Payroll" value={bulan} onChange={setBulan}/></Card>
<Card>
<div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>Karyawan — {bulan}</div>
<RTbl headers={["Nama","Posisi","Hadir","Pinjaman","Aksi"]} widths={[200,160,110,150,150]} rows={emps.map(e=>{var r=calcPayrollFull(e,bulan,data);return[<b style={{color:C.wht}}>{e.nama}</b>,e.posisi,<Bdg color="blue">{r.hariHadir}/{r.totalHariKerja}</Bdg>,<b style={{color:r.pinjamanSaldo>0?C.olt:C.gl2,whiteSpace:"nowrap"}}>{fR(r.pinjamanSaldo)}</b>,<Btn sm color="blue" onClick={()=>setEditSlip(buildSlipRows(e))}>📝 Buat Slip</Btn>];})}/>
</Card>
{(data.payrollLog||[]).length>0&&<Card><div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}><div style={{fontWeight:700,color:C.gl2,fontSize:13}}>📜 Riwayat Slip</div><div style={{display:"flex",gap:6}}><Bdg color="green">Aktif: {(data.payrollLog||[]).filter(p=>!p.arsip).length}</Bdg><Bdg color="gray">Arsip: {(data.payrollLog||[]).filter(p=>p.arsip).length}</Bdg></div></div><RTbl headers={["Bulan","Nama","No.Slip","Total","Status","Aksi"]} widths={[90,160,140,150,100,220]} rows={(data.payrollLog||[]).map(p=>[
p.bulan,p.nama,p.noSlip,
<b style={{color:C.glt,whiteSpace:"nowrap"}}>{fR(p.totalDiterima)}</b>,
p.arsip?<Bdg color="gray">Diarsip</Bdg>:<Bdg color="green">Aktif</Bdg>,
<div style={{display:"flex",gap:5}}>
<Btn sm color="blue" onClick={()=>setViewSlip(p)}>🖨️ Cetak</Btn>
{!p.arsip
?<Btn sm color="gray" onClick={()=>setData(d=>({...d,payrollLog:(d.payrollLog||[]).map(x=>x.id===p.id?{...x,arsip:true}:x)}))}>📦 Arsip</Btn>
:<Btn sm color="green" onClick={()=>setData(d=>({...d,payrollLog:(d.payrollLog||[]).map(x=>x.id===p.id?{...x,arsip:false}:x)}))}>↩️ Restore</Btn>}
</div>
])}/></Card>}
{editSlip&&<SlipEditor slip={editSlip} onClose={()=>setEditSlip(null)} onSave={s=>{
var sgInfo=nextSGNo(data,s.bulan);
var newCounters={...(data.counters||{inv:{},sg:{},reg:0})};if(!newCounters.sg)newCounters.sg={};newCounters.sg[sgInfo.key]=sgInfo.n;
var totPgh=s.rows.filter(r=>r.section==="penghasilan").reduce((a,r)=>a+Number(r.jumlah||0),0);
var totPot=s.rows.filter(r=>r.section==="potongan"&&r.kind!=="info").reduce((a,r)=>a+Number(r.jumlah||0),0);
var totYdt=s.rows.filter(r=>r.section==="ydt").reduce((a,r)=>a+Number(r.jumlah||0),0);
var totalDiterima=totPgh-totPot-totYdt;
var rec={id:uid(),noSlip:sgInfo.no,empId:s.empId,nama:s.nama,posisi:s.posisi,alamat:s.alamat,telepon:s.telepon,bulan:s.bulan,tanggal:s.tanggal,hariHadir:s.hariHadir,totalHariKerja:s.totalHariKerja,absen:s.absen,totalPinjaman:s.totalPinjaman,potonganPinjaman:s.potonganPinjaman,rows:s.rows,totalDiterima};
setData(d=>({...d,payrollLog:[rec,...(d.payrollLog||[])],counters:newCounters}));
setEditSlip(null);setViewSlip(rec);toast("✓ Slip "+sgInfo.no+" tersimpan!");
}}/>}
{viewSlip&&<SlipGajiView slip={viewSlip} company={data.company} onClose={()=>setViewSlip(null)}/>}
</div>;
}

// ─── PELANGGAN ────────────────────────────────────────────────────────────────
function PelangganMod({data,setData,toast}){
var C=useTheme();var mob=useMobile();
var blk={nama:"",kategori:"Rumah Tangga",kategoriCustom:"",telepon:"",alamat:""};
var[f,setF]=useState({...blk});var[edit,setEdit]=useState(null);var[delId,setDelId]=useState(null);
var[search,setSearch]=useState("");var[filterKat,setFilterKat]=useState("");
var[hargaModal,setHargaModal]=useState(null);var[hf,setHf]=useState({ukuran:"12 kg",jenis:"Isi",harga:""});
var[regEdit,setRegEdit]=useState(null);
var isCustom=f.kategori==="Lainnya";
var allKats=[...new Set([...PLG_KAT_27,"Lainnya",...(data.pelanggan||[]).map(p=>p.kategori).filter(k=>!PLG_KAT_27.includes(k))])];
var filtered=(data.pelanggan||[]).filter(p=>{if(filterKat&&p.kategori!==filterKat)return false;if(search&&!((p.nama||"")+" "+(p.telepon||"")+" "+(p.regNo||"")).toLowerCase().includes(search.toLowerCase()))return false;return true;});
function getKat(){return isCustom?(f.kategoriCustom||"Lainnya"):f.kategori;}
function save(){if(!f.nama)return;if(edit){setData(d=>({...d,pelanggan:(d.pelanggan||[]).map(p=>p.id===edit.id?{...p,...f,kategori:getKat()}:p)}));setEdit(null);}else{var regInfo=nextRegNo(data);var newCounters={...(data.counters||{inv:{},sg:{},reg:0}),reg:regInfo.n};setData(d=>({...d,pelanggan:[{id:uid(),...f,kategori:getKat(),regNo:regInfo.no,hargaKhusus:[]},...(d.pelanggan||[])],counters:newCounters}));}setF({...blk});toast("✓ Pelanggan disimpan!");}
function saveHarga(plg){if(!hf.harga)return;var baseHK=Array.isArray(plg.hargaKhusus)?plg.hargaKhusus:[];var newHK=[...baseHK.filter(x=>!(x.ukuran===hf.ukuran&&x.jenis===hf.jenis)),{ukuran:hf.ukuran,jenis:hf.jenis,harga:Number(hf.harga)}];setData(d=>({...d,pelanggan:(d.pelanggan||[]).map(p=>p.id!==plg.id?p:{...p,hargaKhusus:newHK})}));setHargaModal(prev=>prev?{...prev,hargaKhusus:newHK}:null);toast("✓ Harga khusus disimpan!");}
function delHarga(plg,ukuran,jenis){var newHK=(plg.hargaKhusus||[]).filter(x=>!(x.ukuran===ukuran&&x.jenis===jenis));setData(d=>({...d,pelanggan:(d.pelanggan||[]).map(p=>p.id!==plg.id?p:{...p,hargaKhusus:newHK})}));setHargaModal(prev=>prev?{...prev,hargaKhusus:newHK}:null);}
return <div>
<STitle icon="👥" children="Kelola Pelanggan"/>
<Card style={{width:"fit-content",maxWidth:"100%",minWidth:660}}>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(155px,210px))",gap:10}}>
<Inp label="Nama Pelanggan" value={f.nama} onChange={v=>setF(p=>({...p,nama:v}))} style={{gridColumn:mob?"1/-1":"auto"}}/>
<div style={{marginBottom:10,gridColumn:mob?"1/-1":"auto"}}>
<label style={{display:"block",fontSize:11,color:C.gl2,marginBottom:3,fontWeight:600}}>Kategori</label>
<select value={f.kategori} onChange={e=>setF(p=>({...p,kategori:e.target.value}))} style={{width:"100%",border:"1px solid "+C.bdr,borderRadius:8,padding:"9px 11px",color:C.wht,fontSize:13,outline:"none",background:C.nav,boxSizing:"border-box"}}>{[...PLG_KAT_27,"Lainnya"].map(k=><option key={k}>{k}</option>)}</select>
{isCustom&&<input value={f.kategoriCustom} placeholder="Ketik kategori..." onChange={e=>setF(p=>({...p,kategoriCustom:e.target.value}))} style={{width:"100%",border:"1px solid "+C.bdr,borderRadius:8,padding:"9px 11px",color:C.wht,fontSize:13,outline:"none",background:C.nav,marginTop:4,boxSizing:"border-box"}}/>}
</div>
<Inp label="No. Telepon" value={f.telepon} onChange={v=>setF(p=>({...p,telepon:v}))}/>
<Inp label="Alamat" value={f.alamat} onChange={v=>setF(p=>({...p,alamat:v}))}/>
</div>
<Btn color="green" onClick={save} dis={!f.nama}>➕ {edit?"Update":"Tambah"} Pelanggan</Btn>
{!edit&&<div style={{fontSize:11,color:C.gl2,marginTop:6}}>No. Registrasi otomatis: <b style={{color:C.blt}}>HTS/CST/{String((data.counters?.reg||0)+1).padStart(3,"0")}</b></div>}
</Card>
<Card>
<div style={{display:"grid",gridTemplateColumns:"1fr auto",gap:10,marginBottom:10}}><Inp label="" value={search} onChange={setSearch} placeholder="🔍 Cari nama / HP / No.Reg..." style={{marginBottom:0}}/><Sel label="" value={filterKat} onChange={setFilterKat} opts={[{v:"",l:"Semua"},...allKats.map(k=>({v:k,l:k}))]} style={{marginBottom:0}}/></div>
<div style={{fontSize:12,color:C.gl2,marginBottom:8}}>{filtered.length} pelanggan</div>
<RTbl headers={["No.Reg","Nama","Kategori","Telepon","Harga Khusus","Aksi"]} widths={[140,190,150,140,140,120]} rows={filtered.map(p=>{var hk=p.hargaKhusus||[];return[<button onClick={()=>setRegEdit(p)} style={{background:"none",border:"none",color:C.blt,fontSize:12,fontWeight:700,cursor:"pointer",padding:0}}>{p.regNo||"-"}</button>,<b style={{color:C.wht}}>{p.nama}</b>,<Bdg color={PLG_TITIP_KAT.includes(p.kategori)?"blue":"gray"}>{p.kategori||"-"}</Bdg>,p.telepon||"-",<button onClick={()=>{var cur=(data.pelanggan||[]).find(x=>x.id===p.id)||p;setHargaModal({...cur,hargaKhusus:Array.isArray(cur.hargaKhusus)?cur.hargaKhusus:[]});}} style={{background:hk.length>0?C.inHv:C.nav,border:"1px solid "+(hk.length>0?C.blt:C.bdr),borderRadius:5,padding:"4px 8px",color:hk.length>0?C.blt:C.gl2,cursor:"pointer",fontSize:12}}>{hk.length>0?hk.length+" harga":"+ Tambah"}</button>,<ActBtns onEdit={()=>{setEdit(p);setF({nama:p.nama,kategori:PLG_KAT_27.includes(p.kategori)?p.kategori:"Lainnya",kategoriCustom:PLG_KAT_27.includes(p.kategori)?"":p.kategori,telepon:p.telepon||"",alamat:p.alamat||""});}} onDel={()=>setDelId(p)}/>];})}/>
</Card>
{regEdit&&<Modal title={"Edit No. Registrasi — "+regEdit.nama} onClose={()=>setRegEdit(null)} onSave={()=>{setData(d=>({...d,pelanggan:(d.pelanggan||[]).map(p=>p.id===regEdit.id?{...p,regNo:regEdit.regNo}:p)}));setRegEdit(null);toast("✓ Diperbarui!");}}><Inp label="No. Registrasi" value={regEdit.regNo} onChange={v=>setRegEdit(p=>({...p,regNo:v}))} placeholder="HTS/CST/001"/></Modal>}
{hargaModal&&<Modal title={"💲 Harga Khusus — "+hargaModal.nama} onClose={()=>setHargaModal(null)} width={500}>
<div style={{marginBottom:14}}>
<div style={{fontWeight:700,color:C.gl2,marginBottom:8,fontSize:13}}>Harga Terdaftar:</div>
{!(hargaModal.hargaKhusus||[]).length&&<div style={{color:C.gl2,fontSize:12,marginBottom:10,fontStyle:"italic"}}>Belum ada. Pakai HET default.</div>}
{(hargaModal.hargaKhusus||[]).map((h,i)=><div key={i} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"8px 12px",background:C.nav,borderRadius:8,marginBottom:6}}>
<span style={{fontSize:13,color:C.wht}}><b>{h.ukuran}</b> — {h.jenis}</span>
<div style={{display:"flex",alignItems:"center",gap:8}}><span style={{color:C.glt,fontWeight:700}}>{fR(h.harga)}</span><span style={{color:C.gl2,fontSize:10}}>(HET: {fR(getHET(data,h.ukuran,h.jenis))})</span><button onClick={()=>delHarga(hargaModal,h.ukuran,h.jenis)} style={{background:C.inHvE,border:"none",borderRadius:5,color:C.rlt,cursor:"pointer",fontSize:11,padding:"3px 7px"}}>🗑️</button></div>
</div>)}
</div>
<div style={{borderTop:"1px solid "+C.bdr,paddingTop:14}}>
<div style={{fontWeight:700,color:C.gl2,marginBottom:8,fontSize:13}}>Tambah / Update:</div>
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10}}>
<Sel label="Ukuran" value={hf.ukuran} onChange={v=>setHf(p=>({...p,ukuran:v}))} opts={SIZES}/>
<Sel label="Jenis" value={hf.jenis} onChange={v=>setHf(p=>({...p,jenis:v}))} opts={JENIS}/>
<Inp label={"Harga (HET: "+fR(getHET(data,hf.ukuran,hf.jenis))+")"} type="number" value={hf.harga} onChange={v=>setHf(p=>({...p,harga:v}))}/>
</div>
<Btn color="green" onClick={()=>{saveHarga(hargaModal);setHf({ukuran:"12 kg",jenis:"Isi",harga:""});}} dis={!hf.harga}>💾 Simpan Harga</Btn>
</div>
</Modal>}
{edit&&<Modal title={"Edit: "+edit.nama} onSave={save} onClose={()=>{setEdit(null);setF({...blk});}}><Inp label="Nama" value={f.nama} onChange={v=>setF(p=>({...p,nama:v}))}/><div style={{marginBottom:10}}><label style={{display:"block",fontSize:11,color:C.gl2,marginBottom:3,fontWeight:600}}>Kategori</label><select value={f.kategori} onChange={e=>setF(p=>({...p,kategori:e.target.value}))} style={{width:"100%",border:"1px solid "+C.bdr,borderRadius:8,padding:"9px 11px",color:C.wht,fontSize:13,outline:"none",background:C.nav,boxSizing:"border-box"}}>{[...PLG_KAT_27,"Lainnya"].map(k=><option key={k}>{k}</option>)}</select>{isCustom&&<input value={f.kategoriCustom} placeholder="Ketik..." onChange={e=>setF(p=>({...p,kategoriCustom:e.target.value}))} style={{width:"100%",border:"1px solid "+C.bdr,borderRadius:8,padding:"9px 11px",color:C.wht,fontSize:13,outline:"none",background:C.nav,marginTop:4,boxSizing:"border-box"}}/>}</div><Inp label="Telepon" value={f.telepon} onChange={v=>setF(p=>({...p,telepon:v}))}/><Inp label="Alamat" value={f.alamat} onChange={v=>setF(p=>({...p,alamat:v}))}/></Modal>}
{delId&&<ConfirmDel msg={"Hapus \""+delId.nama+"\"?"} onCancel={()=>setDelId(null)} onConfirm={()=>{setData(d=>({...d,pelanggan:(d.pelanggan||[]).filter(x=>x.id!==delId.id)}));setDelId(null);}}/>}
</div>;
}

// === AKHIR BAGIAN 3 ===
// === BAGIAN 4 DARI 4 ===

// ─── PENGELUARAN v4 (fixed + filter) ──────────────────────────────────────────
function PengeluaranMod({data,setData,user,toast}){
var C=useTheme();
var[f,setF]=useState({tanggal:toDay(),kategori:"BBM",kategoriCustom:"",keperluan:"perusahaan",ket:"",nominal:"",metode:"cash",bank:"BSI",hppReferensi:""});
var[delId,setDelId]=useState(null);
var[barFilter,setBarFilter]=useState({from:"",to:"",kategori:"",keperluan:""});
var karList=sortEmp((data.employees||[]).filter(e=>e.aktif));
var karAbsensi=karList.filter(e=>e.ikutAbsensi);
var isLainnya=f.kategori==="Lainnya";
var isPancung=f.kategori==="Pancung 12kg"||f.kategori==="Pancung 5.5kg";
var ukuranPancung=f.kategori==="Pancung 12kg"?"12 kg":"5.5 kg";
function save(){
if(!f.nominal||Number(f.nominal)<=0)return;
if(isPancung&&(!f.qty||Number(f.qty)<=0)){toast("⚠️ Qty tabung dipancung wajib diisi");return;}
if(isPancung&&(!f.hppReferensi||Number(f.hppReferensi)<=0)){toast("⚠️ Harga Modal per Unit Isi (HPP) wajib diisi");return;}
var kat=isLainnya?(f.kategoriCustom||"Lainnya"):f.kategori;
var isKar=f.keperluan!=="perusahaan";
var karId=isKar?f.keperluan:null;
var karNm=isKar?(karList.find(e=>e.id===karId)?.nama||""):"";
var rec={id:uid(),tanggal:f.tanggal,kategori:kat,keperluan:f.keperluan,karyawanId:karId,karyawanNama:karNm,ket:f.ket,nominal:Number(f.nominal),metode:f.metode,bank:f.metode==="transfer"?f.bank:"",
  qtyPancung:isPancung?Number(f.qty||0):undefined,hppReferensi:isPancung?Number(f.hppReferensi||0):undefined,ukuranPancung:isPancung?ukuranPancung:undefined};
setData(d=>{
  if(!isPancung)return{...d,pengeluaran:[rec,...(d.pengeluaran||[])]};
  // ── Pancung: otomatis +stok isi & push batch FIFO, -stok kosong ──
  var qty=Number(f.qty||0);var hpp=Number(f.hppReferensi||0);
  var ns={...(d.stock||{})};ns[ukuranPancung]=(ns[ukuranPancung]||0)+qty;
  var nk={...(d.stokKosong||{})};nk[ukuranPancung]=Math.max(0,(nk[ukuranPancung]||0)-qty);
  var withBatch=addBatch(d,ukuranPancung,qty,hpp,"Pancung — "+(karNm||"Perusahaan"),f.tanggal);
  var stockLog={id:uid(),tanggal:f.tanggal,ukuran:ukuranPancung,jenis:"✂️ Pancung (+Isi, -Kosong)",qty,ket:"Biaya jasa: "+fR(Number(f.nominal))+(f.ket?" — "+f.ket:""),sumber:"Pancung",user:user?.nama||""};
  return{...withBatch,pengeluaran:[rec,...(d.pengeluaran||[])],stock:ns,stokKosong:nk,stockLog:[stockLog,...(d.stockLog||[])].slice(0,500)};
});
setF(p=>({...p,nominal:"",ket:"",qty:"",hppReferensi:""}));
toast(isPancung?"✓ Pancung dicatat! Biaya jasa + stok isi otomatis terupdate.":"✓ Pengeluaran dicatat!");
}
var penFiltered=useMemo(()=>{
return(data.pengeluaran||[]).filter(p=>{
if(barFilter.from&&(p.tanggal||"")<barFilter.from)return false;
if(barFilter.to&&(p.tanggal||"")>barFilter.to)return false;
if(barFilter.kategori&&(p.kategori||"")!==barFilter.kategori)return false;
if(barFilter.keperluan){
if(barFilter.keperluan==="perusahaan"&&p.keperluan!=="perusahaan")return false;
if(barFilter.keperluan!=="perusahaan"&&p.karyawanId!==barFilter.keperluan)return false;
}
return true;
});
},[data.pengeluaran,barFilter]);
var totalPen=penFiltered.reduce((a,p)=>a+Number(p.nominal||0),0);
var cashPen=penFiltered.filter(p=>p.metode==="cash").reduce((a,p)=>a+Number(p.nominal||0),0);
var tfPen=penFiltered.filter(p=>p.metode==="transfer").reduce((a,p)=>a+Number(p.nominal||0),0);
var cols=[
{key:"tanggal",label:"Tgl",width:90,render:r=>fDs(r.tanggal),sortVal:r=>r.tanggal,filterable:true},
{key:"kategori",label:"Kategori",width:170,render:r=><Bdg color="gray">{r.kategori}</Bdg>,filterable:true},
{key:"keperluan",label:"Keperluan",width:150,render:r=>r.keperluan==="perusahaan"?<Bdg color="gray">Perusahaan</Bdg>:<Bdg color="blue">{r.karyawanNama||r.keperluan}</Bdg>,filterable:true},
{key:"ket",label:"Keterangan",width:280,render:r=>r.ket||"-",filterable:true},
{key:"nominal",label:"Nominal",width:140,render:r=><b style={{color:C.rlt,whiteSpace:"nowrap"}}>{fR(r.nominal)}</b>,sortVal:r=>r.nominal,filterable:false},
{key:"metode",label:"Metode",width:100,render:r=><Bdg color={r.metode==="transfer"?"blue":"green"}>{r.metode}</Bdg>,filterable:true,filterType:"select",options:[{v:"cash",l:"Cash"},{v:"transfer",l:"Transfer"}]},
{key:"_aksi",label:"Aksi",width:70,sortable:false,filterable:false,render:r=><button onClick={()=>setDelId(r)} style={{background:C.inHvE,border:"1px solid "+C.rlt,borderRadius:7,padding:"6px 9px",color:C.rlt,cursor:"pointer",fontSize:13}}>🗑️</button>},
];
return <div>
<STitle icon="💸" children="Pengeluaran"/>
<Card style={{width:"fit-content",maxWidth:"100%",minWidth:660}}>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(180px,220px))",gap:10,marginBottom:10}}>
<Inp label="Tanggal" type="date" value={f.tanggal} onChange={v=>setF(p=>({...p,tanggal:v}))} style={{marginBottom:0}}/>
<div>
<label style={{display:"block",fontSize:11,color:C.gl2,marginBottom:3,fontWeight:600}}>Kategori</label>
<select value={f.kategori} onChange={e=>setF(p=>({...p,kategori:e.target.value}))} style={{width:"100%",border:"1px solid "+C.bdr,borderRadius:8,padding:"9px 11px",color:C.wht,fontSize:13,outline:"none",background:C.nav,boxSizing:"border-box"}}>{KAT_K.map(k=><option key={k}>{k}</option>)}</select>
{isLainnya&&<input value={f.kategoriCustom} placeholder="Ketik kategori..." onChange={e=>setF(p=>({...p,kategoriCustom:e.target.value}))} style={{width:"100%",border:"1px solid "+C.bdr,borderRadius:8,padding:"9px 11px",color:C.wht,fontSize:13,outline:"none",background:C.nav,marginTop:4,boxSizing:"border-box"}}/>}
</div>
<Sel label="Keperluan / Atas Nama" value={f.keperluan} onChange={v=>setF(p=>({...p,keperluan:v}))} opts={[{v:"perusahaan",l:"🏢 Perusahaan"},...karList.map(e=>({v:e.id,l:e.nama}))]} style={{marginBottom:0}}/>
</div>
<Inp label="Keterangan" value={f.ket} onChange={v=>setF(p=>({...p,ket:v}))} placeholder="Detail pengeluaran..." style={{marginBottom:10}}/>
<div style={{display:"grid",gridTemplateColumns:"90px auto 180px auto 180px",gap:4,alignItems:"flex-end",marginBottom:isPancung?6:10,width:"fit-content",maxWidth:"100%",minWidth:660}}>
<Inp label={isPancung?"Qty Tabung":"Qty (opsional)"} type="number" value={f.qty||""} onChange={v=>{var n=Number(v)||0;var h=Number(f.hargaSatuan)||0;setF(p=>({...p,qty:v,nominal:n&&h?String(n*h):p.nominal}));}} placeholder="1" style={{marginBottom:0}}/>
<div style={{textAlign:"center",color:C.gl2,fontSize:14,paddingBottom:10}}>×</div>
<Inp label={isPancung?"Biaya Jasa/Tbg":"Harga Satuan"} type="number" step="1000" value={f.hargaSatuan||(isPancung?String(KAT_AUTO_HARGA[f.kategori]||""):"")} onChange={v=>{var h=Number(v)||0;var q=Number(f.qty)||0;setF(p=>({...p,hargaSatuan:v,nominal:q&&h?String(q*h):p.nominal}));}} placeholder="opsional" style={{marginBottom:0}}/>
<div style={{textAlign:"center",color:C.gl2,fontSize:14,paddingBottom:10}}>=</div>
<Inp label="Total (Rp)" type="number" value={f.nominal} onChange={v=>setF(p=>({...p,nominal:v}))} style={{marginBottom:0}}/>
</div>
{isPancung&&<div style={{background:C.nav,borderRadius:8,padding:"10px 12px",border:"1px solid "+C.olt,marginBottom:10}}>
<div style={{fontSize:11,fontWeight:700,color:C.olt,marginBottom:6}}>✂️ Hasil Pancung → Stok Isi {ukuranPancung}</div>
<Inp label="Harga Modal per Unit Isi (HPP Referensi)" type="number" step="1000" value={f.hppReferensi||""} onChange={v=>setF(p=>({...p,hppReferensi:v}))} placeholder="Harga modal isi yang dihasilkan dari pancung" style={{marginBottom:6}}/>
<div style={{fontSize:10,color:C.gl2}}>Stok isi {ukuranPancung} otomatis <b style={{color:C.glt}}>+{f.qty||0}</b> tabung @ <b style={{color:C.olt}}>{fR(Number(f.hppReferensi)||0)}</b>/unit • Stok kosong otomatis −{f.qty||0}</div>
</div>}
<div style={{display:"flex",gap:8,marginBottom:10}}>{["cash","transfer"].map(m=><button key={m} onClick={()=>setF(p=>({...p,metode:m}))} style={{background:f.metode===m?C.blu:C.nav,color:f.metode===m?"white":C.wht,border:"1px solid "+(f.metode===m?C.blt:C.bdr),borderRadius:8,padding:"6px 14px",fontWeight:700,fontSize:12,cursor:"pointer"}}>{m==="cash"?"💵 Cash":"🏦 Transfer"}</button>)}</div>
{f.metode==="transfer"&&<div style={{display:"flex",gap:8,marginBottom:10}}>{["BSI","BCA"].map(b=><button key={b} onClick={()=>setF(p=>({...p,bank:b}))} style={{background:f.bank===b?C.blu:C.nav,color:f.bank===b?"white":C.wht,border:"2px solid "+(f.bank===b?C.blt:C.bdr),borderRadius:8,padding:"6px 14px",fontWeight:700,cursor:"pointer"}}>{b}</button>)}</div>}
<Btn color="red" onClick={save} dis={!f.nominal||Number(f.nominal)<=0}>💾 Catat Pengeluaran</Btn>
</Card>
<div style={{display:"flex",gap:10,flexWrap:"wrap",marginBottom:12}}>
<SC label="Total Pengeluaran" value={fR(totalPen)} icon="💸" color={C.rlt}/>
<SC label="Cash" value={fR(cashPen)} icon="💵" color={C.gl2}/>
<SC label="Transfer" value={fR(tfPen)} icon="🏦" color={C.blt}/>
</div>
<Card>
<div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
<div style={{fontWeight:700,color:C.gl2,fontSize:13}}>📋 Riwayat Pengeluaran</div>
<button onClick={()=>{
var wb=XLSX.utils.book_new();
var header=["Tanggal","Kategori","Keperluan/Atas Nama","Keterangan","Qty","Harga Satuan","Nominal","Metode","Bank"];
var aoa=[header,...penFiltered.map(p=>[
  fDs(p.tanggal),p.kategori,
  p.keperluan==="perusahaan"?"Perusahaan":(p.karyawanNama||p.keperluan||""),
  p.ket||"",p.qtyPancung||p.qty||"",p.hargaSatuan||"",Number(p.nominal||0),
  p.metode||"",p.metode==="transfer"?(p.bank||""):""
])];
aoa.push([]);
aoa.push(["","","","","","TOTAL",penFiltered.reduce((a,p)=>a+Number(p.nominal||0),0),"",""]);
var ws=XLSX.utils.aoa_to_sheet(aoa);
ws["!cols"]=[{wch:12},{wch:22},{wch:18},{wch:28},{wch:8},{wch:14},{wch:14},{wch:10},{wch:8}];
XLSX.utils.book_append_sheet(wb,ws,"Pengeluaran");
var fname="Pengeluaran"+(barFilter.from?"_"+barFilter.from:"")+(barFilter.to?"_s_"+barFilter.to:"")+".xlsx";
XLSX.writeFile(wb,fname);
toast("✓ Excel pengeluaran didownload!");
}} style={{background:"#15803D",color:"white",border:"none",padding:"7px 14px",borderRadius:7,fontSize:12,cursor:"pointer",fontWeight:700}}>📥 Export Excel</button>
</div>
<div style={{background:C.nav,borderRadius:8,padding:10,marginBottom:10,border:"1px solid "+C.bdr}}>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(140px,1fr))",gap:8}}>
<Inp label="Dari" type="date" value={barFilter.from} onChange={v=>setBarFilter(p=>({...p,from:v}))} style={{marginBottom:0}}/>
<Inp label="Sampai" type="date" value={barFilter.to} onChange={v=>setBarFilter(p=>({...p,to:v}))} style={{marginBottom:0}}/>
<Sel label="Kategori" value={barFilter.kategori} onChange={v=>setBarFilter(p=>({...p,kategori:v}))} opts={[{v:"",l:"Semua"},...KAT_K.map(k=>({v:k,l:k}))]} style={{marginBottom:0}}/>
<Sel label="Keperluan" value={barFilter.keperluan} onChange={v=>setBarFilter(p=>({...p,keperluan:v}))} opts={[{v:"",l:"Semua"},{v:"perusahaan",l:"Perusahaan"},...karList.map(e=>({v:e.id,l:e.nama}))]} style={{marginBottom:0}}/>
</div>
{(barFilter.from||barFilter.to||barFilter.kategori||barFilter.keperluan)&&<div style={{marginTop:8}}><Btn sm color="gray" onClick={()=>setBarFilter({from:"",to:"",kategori:"",keperluan:""})}>✕ Reset Filter</Btn></div>}
</div>
<FilterTbl columns={cols} data={penFiltered} empty="Belum ada pengeluaran" maxRows={100}/>
</Card>
{delId&&<ConfirmDel msg={"Hapus pengeluaran "+fR(delId.nominal)+"?"} onCancel={()=>setDelId(null)} onConfirm={()=>{setData(d=>({...d,pengeluaran:(d.pengeluaran||[]).filter(x=>x.id!==delId.id)}));setDelId(null);toast("✓ Dihapus");}}/>}
</div>;
}

// ─── JUALAN LAIN (gas kaleng, aksesoris, dll — non-LPG, tanpa stok) ───────────
function JualanLainMod({data,setData,user,toast}){
var C=useTheme();
var[tgl,setTgl]=useState(toDay());
var[nama,setNama]=useState("");
var[qty,setQty]=useState("1");
var[harga,setHarga]=useState("");
var[bayar,setBayar]=useState("cash");
var[ket,setKet]=useState("");
var[editId,setEditId]=useState(null);
var[delId,setDelId]=useState(null);
var[fFrom,setFFrom]=useState("");
var[fTo,setFTo]=useState("");

var list=(data.jualanLain||[]).slice().sort((a,b)=>(b.tanggal||"").localeCompare(a.tanggal||"")||(b.id||"").localeCompare(a.id||""));
var filtered=list.filter(x=>{
if(fFrom&&x.tanggal<fFrom)return false;
if(fTo&&x.tanggal>fTo)return false;
return true;
});
var totalQty=Number(qty)||0;var totalHarga=Number(harga)||0;var subtotal=totalQty*totalHarga;

function resetForm(){setTgl(toDay());setNama("");setQty("1");setHarga("");setBayar("cash");setKet("");setEditId(null);}

function simpan(){
if(!nama.trim()){toast("⚠️ Nama barang wajib diisi");return;}
if(subtotal<=0){toast("⚠️ Qty & harga harus diisi");return;}
var rec={id:editId||uid(),tanggal:tgl,nama:nama.trim(),qty:totalQty,harga:totalHarga,total:subtotal,bayar,ket:ket.trim()};
if(editId){
  setData(d=>({...d,jualanLain:(d.jualanLain||[]).map(x=>x.id===editId?rec:x)}));
  toast("✓ Jualan lain diperbarui!");
}else{
  setData(d=>({...d,jualanLain:[rec,...(d.jualanLain||[])]}));
  toast("✓ Jualan lain tersimpan! Total: "+fR(subtotal));
}
resetForm();
}

function mulaiEdit(x){
setEditId(x.id);setTgl(x.tanggal);setNama(x.nama);setQty(String(x.qty));setHarga(String(x.harga));setBayar(x.bayar||"cash");setKet(x.ket||"");
}

var totalFiltered=filtered.reduce((a,x)=>a+(x.total||0),0);
var totalCashF=filtered.filter(x=>x.bayar==="cash").reduce((a,x)=>a+(x.total||0),0);
var totalTFF=filtered.filter(x=>x.bayar==="transfer").reduce((a,x)=>a+(x.total||0),0);

return <div>
<STitle icon="🛒" children="Jualan Lain (Non-LPG)"/>
<div style={{fontSize:11,color:C.gl2,marginBottom:12,marginTop:-8}}>Gas kaleng, korek, selang, regulator, aksesoris, dll — tidak masuk hitungan stok LPG.</div>

<Card style={{border:"1px solid "+(editId?C.olt:C.bdr),width:"fit-content",maxWidth:"100%",minWidth:660}}>
<div style={{fontWeight:700,color:editId?C.olt:C.gl2,marginBottom:10,fontSize:13}}>{editId?"✏️ Edit Jualan Lain":"➕ Tambah Jualan Lain"}</div>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(180px,260px))",gap:10}}>
<Inp label="Tanggal" type="date" value={tgl} onChange={setTgl}/>
<Inp label="Nama Barang" value={nama} onChange={setNama} placeholder="Gas Kaleng / Selang / dll"/>
<Inp label="Qty" type="number" value={qty} onChange={setQty} placeholder="1"/>
<Inp label="Harga Satuan" type="number" value={harga} onChange={setHarga} placeholder="0"/>
<Sel label="Metode Bayar" value={bayar} onChange={setBayar} opts={[{v:"cash",l:"💵 Cash"},{v:"transfer",l:"🏦 Transfer"}]}/>
<Inp label="Keterangan (opsional)" value={ket} onChange={setKet} placeholder="—"/>
</div>
<div style={{background:C.nav,borderRadius:8,padding:"10px 14px",border:"1px solid "+C.bdr,marginBottom:10,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
<span style={{fontSize:12,color:C.gl2}}>Subtotal</span>
<b style={{fontSize:18,color:C.glt}}>{fR(subtotal)}</b>
</div>
<div style={{display:"flex",gap:8}}>
<Btn color={editId?"orange":"green"} onClick={simpan}>{editId?"💾 Simpan Perubahan":"✓ Simpan"}</Btn>
{editId&&<Btn color="gray" onClick={resetForm}>✕ Batal</Btn>}
</div>
</Card>

<Card>
<div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>📋 Riwayat Jualan Lain</div>
<div style={{display:"flex",gap:8,marginBottom:10,flexWrap:"wrap"}}>
<Inp label="Dari" type="date" value={fFrom} onChange={setFFrom} style={{marginBottom:0,maxWidth:160}}/>
<Inp label="Sampai" type="date" value={fTo} onChange={setFTo} style={{marginBottom:0,maxWidth:160}}/>
{(fFrom||fTo)&&<div style={{alignSelf:"flex-end"}}><Btn sm color="gray" onClick={()=>{setFFrom("");setFTo("");}}>✕ Reset</Btn></div>}
</div>
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:12}}>
{[["Total Penjualan",totalFiltered,C.wht],["Cash",totalCashF,C.glt],["Transfer",totalTFF,C.blt]].map(x=><div key={x[0]} style={{background:C.nav,borderRadius:8,padding:"8px 10px",textAlign:"center",border:"1px solid "+C.bdr}}><div style={{fontSize:10,color:C.gl2}}>{x[0]}</div><div style={{fontSize:13,fontWeight:900,color:x[2]}}>{fR(x[1])}</div></div>)}
</div>
<RTbl headers={["Tanggal","Barang","Qty","Harga","Total","Bayar","Aksi"]} widths={[100,260,80,140,140,100,110]} rows={filtered.slice(0,100).map(x=>[
fDs(x.tanggal),x.nama+(x.ket?" — "+x.ket:""),String(x.qty),<span style={{whiteSpace:"nowrap"}}>{fR(x.harga)}</span>,<span style={{whiteSpace:"nowrap"}}>{fR(x.total)}</span>,
x.bayar==="cash"?"💵 Cash":"🏦 TF",
<div style={{display:"flex",gap:5}}>
<button onClick={()=>mulaiEdit(x)} style={{background:C.blu,border:"none",borderRadius:6,padding:"5px 9px",color:"white",cursor:"pointer",fontSize:12,fontWeight:700}}>✏️</button>
<button onClick={()=>setDelId(x)} style={{background:C.rdk,border:"none",borderRadius:6,padding:"5px 9px",color:"white",cursor:"pointer",fontSize:12,fontWeight:700}}>🗑️</button>
</div>
])} empty="Belum ada data jualan lain"/>
</Card>
{delId&&<ConfirmDel msg={"Hapus jualan \""+delId.nama+"\" senilai "+fR(delId.total)+"?"} onCancel={()=>setDelId(null)} onConfirm={()=>{setData(d=>({...d,jualanLain:(d.jualanLain||[]).filter(x=>x.id!==delId.id)}));setDelId(null);toast("✓ Dihapus");}}/>}
</div>;
}

// ─── INVOICE MANUAL — form bebas, tidak terhubung Penjualan/BON/Piutang ──────
function InvoiceManualMod({data,setData,setInv,user,toast}){
var C=useTheme();
// blankItem inherits tanggal from caller so per-item date is independent
var blankItem=(tgl)=>({id:uid(),tanggal:tgl||toDay(),produk:"",qty:"",price:""});
var suggestNo=(tgl)=>{var t=tgl||toDay();var info=nextInvNo(data,t);return info.no;};
var blankForm=()=>{var tgl=toDay();return{noInv:suggestNo(tgl),tanggal:tgl,konsumen:"",konsumenManual:"",kota:"",salesId:"",salesManual:"",status:"belum",totalDibayar:"",catatan:"",items:[blankItem(tgl)]};};
var[f,setF]=useState(blankForm());
var[editId,setEditId]=useState(null);
var[delId,setDelId]=useState(null);
var[filterKons,setFilterKons]=useState("");

var plgList=(data.pelanggan||[]).slice().sort((a,b)=>(a.nama||"").localeCompare(b.nama||""));
var salesList=(data.employees||[]).filter(e=>e.aktif);

// When header tanggal changes, update suggest no but keep item dates independent
function onTglChange(v){setF(p=>({...p,tanggal:v,noInv:suggestNo(v)}));}

function setItem(id,k,v){setF(p=>({...p,items:p.items.map(it=>it.id===id?{...it,[k]:v}:it)}));}
function addItem(){setF(p=>({...p,items:[...p.items,blankItem(p.tanggal)]}));}
function delItem(id){setF(p=>({...p,items:p.items.filter(it=>it.id!==id)}));}

// Quick-add product by size with HET auto-filled, editable
function addProduk(ukuran){
  var het=getHET(data,ukuran,"Isi");
  setF(p=>({...p,items:[...p.items,{id:uid(),tanggal:p.tanggal,produk:ukuran,qty:"",price:String(het||"")}]}));
}

var validItems=f.items.filter(it=>it.produk&&Number(it.qty)>0);
var totalForm=validItems.reduce((a,it)=>a+Number(it.qty||0)*Number(it.price||0),0);

function resetForm(){setF(blankForm());setEditId(null);}

function mulaiEdit(rec){
setF({noInv:rec.noInv,tanggal:rec.tanggal,
  konsumen:rec.konsumenId||"",konsumenManual:rec.konsumenId?"":(rec.konsumen||""),
  kota:rec.kota||"",salesId:rec.salesId||"",salesManual:rec.salesId?"":(rec.salesNama||""),
  status:rec.status||"belum",totalDibayar:rec.totalDibayar?String(rec.totalDibayar):"",catatan:rec.catatan||"",
  items:(rec.items||[]).map(it=>({id:uid(),tanggal:it.tglDO||rec.tanggal,produk:it.ukuran||"",qty:String(it.qty||""),price:String(it.price||"")}))});
setEditId(rec.id);
}

function save(){
if(!f.noInv||!validItems.length){toast("⚠️ No. Invoice & minimal 1 produk wajib diisi","error");return;}
var plg=plgList.find(p=>p.id===f.konsumen);
var konsumenNama=plg?plg.nama:(f.konsumenManual||"-");
var kotaFinal=f.kota||plg?.alamat?.split(",").pop()?.trim()||"Banda Aceh";
var emp=salesList.find(e=>e.id===f.salesId);
var salesNamaFinal=emp?emp.nama:(f.salesManual||"");
var totalDibayarNum=Number(f.totalDibayar)||0;
var rec={
  id:editId||uid(),noInv:f.noInv,tanggal:f.tanggal,
  konsumenId:plg?.id||"",konsumen:konsumenNama,kota:kotaFinal,
  salesId:emp?.id||"",salesNama:salesNamaFinal,
  items:validItems.map(it=>({ukuran:it.produk,jenis:"",qty:Number(it.qty),price:Number(it.price),tglDO:it.tanggal})),
  total:totalForm,status:f.status,
  totalDibayar:f.status==="sebagian"?totalDibayarNum:(f.status==="lunas"?totalForm:0),
  sisaTagihan:f.status==="lunas"?0:f.status==="sebagian"?Math.max(0,totalForm-totalDibayarNum):totalForm,
  catatan:f.catatan,dibuatOleh:user?.nama||"",updatedAt:new Date().toISOString()
};
setData(d=>({...d,invoiceManual:editId?(d.invoiceManual||[]).map(x=>x.id===editId?rec:x):[rec,...(d.invoiceManual||[])]}));
toast(editId?"✓ Invoice manual diperbarui!":"✓ Invoice manual disimpan!");
resetForm();
}

function bukaCetak(rec){
var inv={
  noInv:rec.noInv,tanggal:rec.tanggal,konsumen:rec.konsumen,kota:rec.kota||"Banda Aceh",
  salesNama:rec.salesNama||"",items:(rec.items||[]).map(it=>({ukuran:it.ukuran,jenis:it.jenis||"",qty:it.qty,price:it.price,tglDO:it.tglDO})),
  total:rec.total,catatan:rec.catatan||"",
  metodeBayar:rec.status==="lunas"?"LUNAS":rec.status==="sebagian"?"SEBAGIAN":"BELUM LUNAS",
  isBon:rec.status!=="lunas",bonLunas:rec.status==="lunas",bonSebagian:rec.status==="sebagian",
  totalDibayar:rec.totalDibayar||0,sisaTagihan:rec.sisaTagihan!=null?rec.sisaTagihan:rec.total,
  isGabungan:false
};
setInv(inv);
}

var rowsFiltered=(data.invoiceManual||[]).filter(r=>!filterKons||(r.konsumen||"").toLowerCase().includes(filterKons.toLowerCase())||(r.noInv||"").toLowerCase().includes(filterKons.toLowerCase()));

// Quick-pick button style
var qBtn=(active)=>({border:"1px solid "+C.bdr,borderRadius:8,padding:"7px 14px",cursor:"pointer",fontWeight:700,fontSize:12,background:active?C.blt:"transparent",color:active?"#fff":C.gl2,transition:"all .15s"});

return <div>
<STitle icon="🧾" children="Invoice Manual"/>
<div style={{fontSize:11,color:C.gl2,marginBottom:12,marginTop:-8}}>Form invoice bebas — diisi & disesuaikan sendiri (no. invoice, konsumen, sales, tanggal, produk, qty). Tidak memengaruhi data Penjualan, BON, Piutang, stok, atau laporan apapun. Murni dokumen tampilan/cetak.</div>

{editId&&<div style={{background:"linear-gradient(90deg,#92400e,#b45309)",border:"1px solid #f59e0b",borderRadius:10,padding:"10px 16px",marginBottom:12,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
<div style={{color:"#fef3c7",fontWeight:800,fontSize:13}}>✏️ Mode Edit Invoice Manual</div>
<button onClick={resetForm} style={{background:"rgba(0,0,0,.25)",border:"1px solid #f59e0b",borderRadius:7,padding:"6px 12px",color:"#fef3c7",cursor:"pointer",fontWeight:700,fontSize:11}}>✕ Batal Edit</button>
</div>}

<Card style={{border:editId?"1px solid #f59e0b":undefined,width:"fit-content",maxWidth:"100%",minWidth:660}}>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(170px,220px))",gap:10}}>

{/* No. Invoice — auto-suggest, langsung editable */}
<div style={{marginBottom:10}}>
  <label style={{display:"block",fontSize:11,color:C.gl2,marginBottom:3,fontWeight:600}}>No. Invoice</label>
  <input value={f.noInv} onChange={e=>setF(p=>({...p,noInv:e.target.value}))} placeholder="#HTS/INV/VI.26/001"
    style={{width:"100%",border:"1px solid "+C.bdr,borderRadius:8,padding:"9px 11px",color:C.wht,fontSize:12,outline:"none",background:C.nav,boxSizing:"border-box",fontFamily:"monospace"}}/>
  {!editId&&<div style={{fontSize:10,color:C.gl2,marginTop:3,cursor:"pointer"}} onClick={()=>setF(p=>({...p,noInv:suggestNo(p.tanggal)}))}>
    🔄 Suggest: <span style={{color:C.blt,fontFamily:"monospace"}}>{suggestNo(f.tanggal)}</span>
  </div>}
</div>

<Inp label="Tanggal" type="date" value={f.tanggal} onChange={onTglChange}/>
<div style={{marginBottom:10}}>
<label style={{display:"block",fontSize:11,color:C.gl2,marginBottom:3,fontWeight:600}}>Konsumen (pilih atau ketik manual)</label>
<select value={f.konsumen||""} onChange={e=>setF(p=>({...p,konsumen:e.target.value}))} style={{width:"100%",border:"1px solid "+C.bdr,borderRadius:8,padding:"9px 11px",color:C.wht,fontSize:13,outline:"none",background:C.nav,boxSizing:"border-box",marginBottom:6}}>
<option value="">— Ketik manual di bawah —</option>
{plgList.map(p=><option key={p.id} value={p.id}>{p.nama}</option>)}
</select>
{!f.konsumen&&<input value={f.konsumenManual} onChange={e=>setF(p=>({...p,konsumenManual:e.target.value}))} placeholder="Nama konsumen manual" style={{width:"100%",border:"1px solid "+C.bdr,borderRadius:8,padding:"9px 11px",color:C.wht,fontSize:13,outline:"none",background:C.nav,boxSizing:"border-box"}}/>}
</div>
<Inp label="Kota / Lokasi" value={f.kota} onChange={v=>setF(p=>({...p,kota:v}))} placeholder="Banda Aceh"/>
<div style={{marginBottom:10}}>
<label style={{display:"block",fontSize:11,color:C.gl2,marginBottom:3,fontWeight:600}}>Sales (pilih atau ketik manual)</label>
<select value={f.salesId||""} onChange={e=>setF(p=>({...p,salesId:e.target.value}))} style={{width:"100%",border:"1px solid "+C.bdr,borderRadius:8,padding:"9px 11px",color:C.wht,fontSize:13,outline:"none",background:C.nav,boxSizing:"border-box",marginBottom:6}}>
<option value="">— Ketik manual di bawah —</option>
{salesList.map(e=><option key={e.id} value={e.id}>{e.nama}</option>)}
</select>
{!f.salesId&&<input value={f.salesManual} onChange={e=>setF(p=>({...p,salesManual:e.target.value}))} placeholder="Nama sales manual (opsional)" style={{width:"100%",border:"1px solid "+C.bdr,borderRadius:8,padding:"9px 11px",color:C.wht,fontSize:13,outline:"none",background:C.nav,boxSizing:"border-box"}}/>}
</div>
<Sel label="Status Pembayaran" value={f.status} onChange={v=>setF(p=>({...p,status:v}))} opts={[{v:"belum",l:"Belum Lunas"},{v:"sebagian",l:"Sebagian"},{v:"lunas",l:"Lunas"}]}/>
{f.status==="sebagian"&&<Inp label="Telah Dibayar (Rp)" type="number" value={f.totalDibayar} onChange={v=>setF(p=>({...p,totalDibayar:v}))}/>}
</div>

{/* Quick-pick produk */}
<div style={{marginTop:4,marginBottom:12}}>
  <div style={{fontSize:11,fontWeight:700,color:C.gl2,marginBottom:6}}>📦 Tambah Produk Cepat</div>
  <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
    {SIZES.map(s=>{
      var het=getHET(data,s,"Isi");
      return <button key={s} onClick={()=>addProduk(s)} style={qBtn(false)}>
        <span style={{fontSize:13}}>➕ {s}</span>
        <span style={{display:"block",fontSize:10,color:C.gl2,fontWeight:400,marginTop:1}}>HET {fR(het)}</span>
      </button>;
    })}
    <button onClick={addItem} style={{...qBtn(false),color:C.gl2,fontSize:11,padding:"7px 12px"}}>+ Baris Bebas</button>
  </div>
</div>

{/* Header kolom */}
{f.items.length>0&&<div style={{display:"grid",gridTemplateColumns:"130px 1.4fr 80px 145px 28px",gap:6,marginBottom:4}}>
  {["Tanggal","Produk","Qty","Harga Satuan",""].map((h,i)=><div key={i} style={{fontSize:10,color:C.gl2,fontWeight:700,paddingLeft:2}}>{h}</div>)}
</div>}

{f.items.map((it)=><div key={it.id} style={{display:"grid",gridTemplateColumns:"130px 1.4fr 80px 145px 28px",gap:6,marginBottom:6,alignItems:"center"}}>
  <input type="date" value={it.tanggal} onChange={e=>setItem(it.id,"tanggal",e.target.value)}
    style={{border:"1px solid "+C.bdr,borderRadius:8,padding:"8px 8px",color:C.wht,fontSize:12,outline:"none",background:C.nav,width:"100%",boxSizing:"border-box"}}/>
  <input value={it.produk} onChange={e=>setItem(it.id,"produk",e.target.value)} placeholder="mis. 12 kg / Jasa Antar"
    style={{border:"1px solid "+C.bdr,borderRadius:8,padding:"8px 10px",color:C.wht,fontSize:12,outline:"none",background:C.nav,width:"100%",boxSizing:"border-box"}}/>
  <input type="number" value={it.qty} onChange={e=>setItem(it.id,"qty",e.target.value)} placeholder="0"
    style={{border:"1px solid "+C.bdr,borderRadius:8,padding:"8px 8px",color:C.wht,fontSize:12,outline:"none",background:C.nav,width:"100%",boxSizing:"border-box",textAlign:"right"}}/>
  <input type="number" value={it.price} onChange={e=>setItem(it.id,"price",e.target.value)} placeholder="0"
    style={{border:"1px solid "+C.bdr,borderRadius:8,padding:"8px 10px",color:C.wht,fontSize:12,outline:"none",background:C.nav,width:"100%",boxSizing:"border-box",textAlign:"right"}}/>
  <button onClick={()=>delItem(it.id)} style={{background:C.inHvE,border:"none",borderRadius:6,color:C.rlt,cursor:"pointer",fontSize:13,padding:"8px 6px"}} title="Hapus baris">✕</button>
</div>)}

<div style={{marginTop:14,marginBottom:10}}><Inp label="Catatan (opsional)" value={f.catatan} onChange={v=>setF(p=>({...p,catatan:v}))}/></div>

<div style={{display:"flex",justifyContent:"space-between",alignItems:"center",background:C.nav,borderRadius:8,padding:"10px 14px",marginBottom:12}}>
<span style={{fontSize:12,color:C.gl2,fontWeight:700}}>Total Invoice</span>
<span style={{fontSize:18,fontWeight:900,color:C.olt}}>{fR(totalForm)}</span>
</div>

<div style={{display:"flex",gap:8}}>
<Btn color={editId?"orange":"green"} onClick={save} dis={!f.noInv||!validItems.length}>{editId?"💾 Simpan Perubahan":"💾 Simpan Invoice Manual"}</Btn>
{editId&&<Btn color="gray" onClick={resetForm}>✕ Batal</Btn>}
</div>
</Card>

<Card>
<div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10,flexWrap:"wrap",gap:8}}>
<div style={{fontWeight:700,color:C.gl2,fontSize:13}}>Riwayat Invoice Manual</div>
<input value={filterKons} onChange={e=>setFilterKons(e.target.value)} placeholder="Cari konsumen / no. invoice..." style={{border:"1px solid "+C.bdr,borderRadius:8,padding:"7px 11px",color:C.wht,fontSize:12,outline:"none",background:C.nav,minWidth:200}}/>
</div>
<RTbl headers={["No. Invoice","Tgl","Konsumen","Sales","Total","Status","Aksi"]} widths={[170,100,200,150,150,110,140]} rows={rowsFiltered.slice(0,150).map(r=>{
var stBadge=r.status==="lunas"?<Bdg color="green">Lunas</Bdg>:r.status==="sebagian"?<Bdg color="orange">Sebagian</Bdg>:<Bdg color="red">Belum</Bdg>;
return[<b style={{color:C.blt,fontSize:12,whiteSpace:"nowrap"}}>{r.noInv}</b>,fDs(r.tanggal),r.konsumen,r.salesNama||"-",<b style={{color:C.olt,whiteSpace:"nowrap"}}>{fR(r.total)}</b>,stBadge,
<div style={{display:"flex",gap:4}}>
<button onClick={()=>bukaCetak(r)} title="Cetak" style={{background:C.inHv,border:"1px solid "+C.blt,borderRadius:6,padding:"5px 8px",color:C.blt,cursor:"pointer",fontSize:12}}>🖨️</button>
<button onClick={()=>mulaiEdit(r)} title="Edit" style={{background:"#78350F",border:"1px solid #F59E0B",borderRadius:6,padding:"5px 8px",color:"#FCD34D",cursor:"pointer",fontSize:12}}>✏️</button>
<button onClick={()=>setDelId(r)} title="Hapus" style={{background:C.inHvE,border:"1px solid "+C.rlt,borderRadius:6,padding:"5px 8px",color:C.rlt,cursor:"pointer",fontSize:12}}>🗑️</button>
</div>];
})} empty="Belum ada invoice manual"/>
</Card>
{delId&&<ConfirmDel msg={"Hapus invoice manual "+delId.noInv+" ("+delId.konsumen+")?"} onCancel={()=>setDelId(null)} onConfirm={()=>{setData(d=>({...d,invoiceManual:(d.invoiceManual||[]).filter(x=>x.id!==delId.id)}));setDelId(null);toast("✓ Dihapus");}}/>}
</div>;
}

// ─── FIFO DETAIL — khusus Admin (Haekal) ──────────────────────────────────────
function FIFODetailMod({data,setData,user,toast}){
var C=useTheme();
var[tabF,setTabF]=useState("batch");// batch | invoice
var[ukuranSel,setUkuranSel]=useState("12 kg");
var[fFrom,setFFrom]=useState("");
var[fTo,setFTo]=useState("");

var batchList=((data.stokBatch||{})[ukuranSel]||[]).slice().sort((a,b)=>(a.tanggal||"").localeCompare(b.tanggal||""));
var batchAktif=batchList.filter(b=>b.qtySisa>0);
var totalQtySisa=batchAktif.reduce((a,b)=>a+b.qtySisa,0);
var nilaiStokUkuran=batchAktif.reduce((a,b)=>a+b.qtySisa*b.harga,0);
var qtySistem=(data.stock||{})[ukuranSel]||0;
var selisihBatch=totalQtySisa-qtySistem;

var nilaiStokTotal=calcNilaiStokFIFO(data);
var nilaiStokLama=calcNilaiStok(data);

var[showHPP,setShowHPP]=useState(false);

var penjualanFIFO=(data.penjualan||[]).filter(p=>p.fifoDetail&&p.fifoDetail.length>0).filter(p=>{
if(fFrom&&p.tanggal<fFrom)return false;
if(fTo&&p.tanggal>fTo)return false;
return true;
}).map(p=>{
var emp=(data.employees||[]).find(e=>e.id===p.salesId);
var per={"5.5 kg":{qty:0,harga:0,hpp:0,margin:0},"12 kg":{qty:0,harga:0,hpp:0,margin:0},"50 kg":{qty:0,harga:0,hpp:0,margin:0}};
(p.items||[]).forEach(it=>{
  if(!per[it.ukuran])return;
  var q=Number(it.qty||0);var pr=Number(it.price||0);
  per[it.ukuran].qty+=q;
  per[it.ukuran].harga=pr;// harga satuan terakhir kalau ada >1 baris ukuran sama
  var hppU=it.hppFIFO!=null?Number(it.hppFIFO):0;
  per[it.ukuran].hpp=hppU;
  per[it.ukuran].margin+=(pr-hppU)*q;
});
return{...p,salesNama:emp?.nama||p.salesNama||"-",per};
}).sort((a,b)=>(b.tanggal||"").localeCompare(a.tanggal||""));
var totalMarginAll=penjualanFIFO.reduce((a,p)=>a+(p.margin||0),0);
var totalAll=penjualanFIFO.reduce((a,p)=>a+(p.total||0),0);

return <div>
<STitle icon="🔬" children="FIFO Detail (Admin Only)"/>
<div style={{fontSize:11,color:C.gl2,marginBottom:12,marginTop:-8}}>Rincian batch stok & margin riil per transaksi berdasarkan metode FIFO. Hanya terlihat oleh role Admin.</div>

<div style={{display:"flex",gap:8,marginBottom:14}}>
{[["batch","📦 Sisa Batch"],["invoice","🧾 Detail per Invoice"]].map(x=><button key={x[0]} onClick={()=>setTabF(x[0])} style={{background:tabF===x[0]?C.blu:C.nav,color:tabF===x[0]?"white":C.wht,border:"1px solid "+(tabF===x[0]?C.blt:C.bdr),borderRadius:8,padding:"7px 14px",fontWeight:700,fontSize:12,cursor:"pointer"}}>{x[1]}</button>)}
</div>

{tabF==="batch"&&<>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(180px,260px))",gap:10,marginBottom:14}}>
<div style={{background:C.card,borderRadius:10,border:"1px solid "+C.bdr,padding:14}}>
<div style={{fontSize:11,color:C.gl2,marginBottom:4}}>Nilai Stok (Metode FIFO)</div>
<div style={{fontSize:18,fontWeight:900,color:C.glt}}>{fR(nilaiStokTotal)}</div>
</div>
<div style={{background:C.card,borderRadius:10,border:"1px solid "+C.bdr,padding:14}}>
<div style={{fontSize:11,color:C.gl2,marginBottom:4}}>Nilai Stok (Metode Lama — Harga Terakhir)</div>
<div style={{fontSize:18,fontWeight:900,color:C.gl2}}>{fR(nilaiStokLama)}</div>
</div>
</div>

<div style={{display:"flex",gap:8,marginBottom:12}}>
{SIZES.map(s=><button key={s} onClick={()=>setUkuranSel(s)} style={{background:ukuranSel===s?C.olt:C.nav,color:ukuranSel===s?"white":C.wht,border:"1px solid "+(ukuranSel===s?C.olt:C.bdr),borderRadius:8,padding:"7px 14px",fontWeight:700,fontSize:12,cursor:"pointer"}}>{s}</button>)}
</div>

<div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:12}}>
<div style={{background:C.nav,borderRadius:8,padding:"8px 10px",textAlign:"center",border:"1px solid "+C.bdr}}><div style={{fontSize:10,color:C.gl2}}>Qty Sisa Batch</div><div style={{fontSize:16,fontWeight:900,color:C.glt}}>{totalQtySisa}</div></div>
<div style={{background:C.nav,borderRadius:8,padding:"8px 10px",textAlign:"center",border:"1px solid "+C.bdr}}><div style={{fontSize:10,color:C.gl2}}>Qty Sistem (data.stock)</div><div style={{fontSize:16,fontWeight:900,color:C.wht}}>{qtySistem}</div></div>
<div style={{background:C.nav,borderRadius:8,padding:"8px 10px",textAlign:"center",border:"1px solid "+(selisihBatch!==0?C.rlt:C.bdr)}}><div style={{fontSize:10,color:C.gl2}}>Selisih</div><div style={{fontSize:16,fontWeight:900,color:selisihBatch!==0?C.rlt:C.glt}}>{selisihBatch}</div></div>
</div>
{selisihBatch!==0&&<div style={{fontSize:10,color:C.rlt,marginBottom:10,fontStyle:"italic"}}>⚠️ Ada selisih antara qty batch FIFO dan qty sistem — kemungkinan ada mutasi stok yang belum tersinkron ke batch (mis. opname manual, migrasi awal, atau data lama sebelum FIFO aktif).</div>}

<Card>
<div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>📦 Antrian Batch {ukuranSel} (FIFO — kepala antrian dipakai dulu)</div>
<RTbl headers={["#","Tanggal Masuk","Qty Awal","Qty Sisa","Harga/Unit","Nilai Sisa","Sumber"]} rows={batchList.map((b,i)=>[
String(i+1),fDs(b.tanggal),String(b.qty),
<b style={{color:b.qtySisa>0?C.glt:C.gl2}}>{b.qtySisa}</b>,
fR(b.harga),
fR(b.qtySisa*b.harga),
b.sumber||"-"
])} empty="Belum ada batch untuk ukuran ini"/>
</Card>
</>}

{tabF==="invoice"&&<>
<div style={{display:"flex",gap:8,marginBottom:12,flexWrap:"wrap",alignItems:"flex-end"}}>
<Inp label="Dari" type="date" value={fFrom} onChange={setFFrom} style={{marginBottom:0,maxWidth:160}}/>
<Inp label="Sampai" type="date" value={fTo} onChange={setFTo} style={{marginBottom:0,maxWidth:160}}/>
{(fFrom||fTo)&&<Btn sm color="gray" onClick={()=>{setFFrom("");setFTo("");}}>✕ Reset</Btn>}
<button onClick={()=>setShowHPP(!showHPP)} style={{background:showHPP?C.olt:C.nav,color:showHPP?"white":C.wht,border:"1px solid "+(showHPP?C.olt:C.bdr),borderRadius:8,padding:"9px 14px",fontWeight:700,fontSize:12,cursor:"pointer"}}>{showHPP?"🙈 Sembunyikan HPP":"👁️ Tampilkan HPP"}</button>
</div>
{penjualanFIFO.length===0&&<div style={{textAlign:"center",padding:30,color:C.gl2,fontSize:12}}>Belum ada transaksi dengan rincian FIFO (FIFO mulai berlaku sejak fitur ini aktif).</div>}
{penjualanFIFO.length>0&&<Card>
<div style={{overflowX:"auto"}}>
<table style={{width:"100%",borderCollapse:"collapse",fontSize:11,minWidth:showHPP?1000:760}}>
<thead><tr style={{background:C.nav}}>
{["Tgl","No.Inv","Sales","Konsumen","Qty 5.5kg","Harga 5.5kg",...(showHPP?["HPP 5.5kg"]:[]),"Qty 12kg","Harga 12kg",...(showHPP?["HPP 12kg"]:[]),"Qty 50kg","Harga 50kg",...(showHPP?["HPP 50kg"]:[]),"Total","Margin"].map((h,i)=><th key={i} style={{padding:"6px 8px",color:C.gl2,textAlign:["Tgl","No.Inv","Sales","Konsumen"].includes(h)?"left":"right",border:"1px solid "+C.bdr,fontSize:10,fontWeight:700,whiteSpace:"nowrap"}}>{h}</th>)}
</tr></thead>
<tbody>
{penjualanFIFO.slice(0,150).map((p,ri)=><tr key={p.id} style={{background:ri%2===0?C.bg:C.nav}}>
<td style={{padding:"5px 8px",color:C.wht,border:"1px solid "+C.bdr,whiteSpace:"nowrap"}}>{fDs(p.tanggal)}</td>
<td style={{padding:"5px 8px",color:C.blt,border:"1px solid "+C.bdr,whiteSpace:"nowrap",fontSize:10}}>{p.noInv||"-"}</td>
<td style={{padding:"5px 8px",color:C.wht,border:"1px solid "+C.bdr,whiteSpace:"nowrap"}}>{p.salesNama}</td>
<td style={{padding:"5px 8px",color:C.wht,border:"1px solid "+C.bdr}}>{p.konsumen}</td>
{["5.5 kg","12 kg","50 kg"].flatMap(uk=>{
var cells=[
<td key={uk+"-qty"} style={{padding:"5px 8px",textAlign:"right",color:p.per[uk].qty>0?C.glt:C.gry,border:"1px solid "+C.bdr}}>{p.per[uk].qty||"-"}</td>,
<td key={uk+"-harga"} style={{padding:"5px 8px",textAlign:"right",color:p.per[uk].qty>0?C.wht:C.gry,border:"1px solid "+C.bdr}}>{p.per[uk].qty>0?fR(p.per[uk].harga):"-"}</td>
];
if(showHPP)cells.push(<td key={uk+"-hpp"} style={{padding:"5px 8px",textAlign:"right",color:p.per[uk].qty>0?C.olt:C.gry,border:"1px solid "+C.bdr,background:C.mode==="dark"?"#3D2E0A":"#FEF3C7"}}>{p.per[uk].qty>0?fR(p.per[uk].hpp):"-"}</td>);
return cells;
})}
<td style={{padding:"5px 8px",textAlign:"right",color:C.wht,fontWeight:700,border:"1px solid "+C.bdr}}>{fR(p.total)}</td>
<td style={{padding:"5px 8px",textAlign:"right",color:p.margin>=0?C.glt:C.rlt,fontWeight:700,border:"1px solid "+C.bdr}}>{fR(p.margin)}</td>
</tr>)}
</tbody>
<tfoot><tr style={{background:C.nav,borderTop:"2px solid "+C.bdr}}>
<td colSpan={showHPP?12:9} style={{padding:"8px 8px",color:C.gl2,fontWeight:700,border:"1px solid "+C.bdr,fontSize:11}}>TOTAL ({penjualanFIFO.length} invoice)</td>
<td style={{padding:"8px 8px",textAlign:"right",color:C.wht,fontWeight:900,border:"1px solid "+C.bdr}}>{fR(totalAll)}</td>
<td style={{padding:"8px 8px",textAlign:"right",color:C.glt,fontWeight:900,border:"1px solid "+C.bdr}}>{fR(totalMarginAll)}</td>
</tr></tfoot>
</table>
</div>
</Card>}
</>}
</div>;
}
function SetoranMod({data,setData,user,toast}){
var C=useTheme();
var isSales=user&&["sales_driver","sales_freelance"].includes(user.role);
var[salesId,setSalesId]=useState(isSales?user.id:"");
var[tgl,setTgl]=useState(toDay());
var[pecah,setPecah]=useState(()=>{var o={};DENOMS.forEach(d=>{o[d]="";});return o;});
var[jadikanPinjaman,setJadikanPinjaman]=useState(false);
var[checkedPen,setCheckedPen]=useState({});
// Mode edit riwayat inline
var[editingLogId,setEditingLogId]=useState(null);
var[showPrint,setShowPrint]=useState(null);
var[viewLog,setViewLog]=useState(null);
var[delLog,setDelLog]=useState(null);
var formRef=useRef(null);
var salesList=sortEmp((data.employees||[]).filter(e=>e.aktif));
var emp=(data.employees||[]).find(e=>e.id===salesId);

// ── DATA HARI INI ──
var penjHari=(data.penjualan||[]).filter(p=>p.salesId===salesId&&p.tanggal===tgl);
var penjCash=penjHari.filter(p=>(p.bayar||"").toLowerCase()==="cash");
var penjTF=penjHari.filter(p=>(p.bayar||"").toLowerCase()==="transfer"||(p.bayar||"").toLowerCase()==="tf");
var penjBon=penjHari.filter(p=>(p.bayar||"").toLowerCase()==="bon");
var penjSplit=penjHari.filter(p=>(p.bayar||"").toLowerCase()==="split");
// Hitung nominal per metode — termasuk split
var cashPenjualan=penjCash.reduce((a,p)=>a+(p.total||0),0)+penjSplit.reduce((a,p)=>a+Number((p.splitDetail||{}).cash||0),0);
var tfPenjualan=penjTF.reduce((a,p)=>a+(p.total||0),0)+penjSplit.reduce((a,p)=>a+Number((p.splitDetail||{}).tf||0),0);
var bonPenjualan=penjBon.reduce((a,p)=>a+(p.total||0),0)+penjSplit.reduce((a,p)=>a+Number((p.splitDetail||{}).bon||0),0);

// Bayar BON hari ini
var bonBayarList=(data.bon||[]).flatMap(b=>(b.pembayaran||[]).filter(px=>px.tanggal===tgl&&b.salesId===salesId).map(px=>({...px,konsumen:b.konsumen})));
var bonCash=bonBayarList.filter(b=>(b.metode||"cash").toLowerCase()==="cash").reduce((a,b)=>a+Number(b.jumlah||b.nominal||0),0);
var bonTF=bonBayarList.filter(b=>(b.metode||"").toLowerCase()==="transfer"||(b.metode||"").toLowerCase()==="tf").reduce((a,b)=>a+Number(b.jumlah||b.nominal||0),0);
var totalBonBayar=bonCash+bonTF;

// Pengeluaran operasional hari ini
var allPenHari=(data.pengeluaran||[]).filter(p=>p.tanggal===tgl);
var penSales=allPenHari.filter(p=>p.karyawanId===salesId||p.karyawanNama===(emp?.nama||""));
var penLain=allPenHari.filter(p=>p.karyawanId!==salesId&&p.karyawanNama!==(emp?.nama||""));

// Init centang: auto centang atas nama sales, bisa tambah manual
useEffect(()=>{
  if(!salesId)return;
  var cp={};
  allPenHari.forEach(p=>{cp[p.id]=!!(p.karyawanId===salesId||p.karyawanNama===(emp?.nama||""));});
  setCheckedPen(cp);
},[salesId,tgl]);

var totalPenChecked=allPenHari.filter(p=>checkedPen[p.id]).reduce((a,p)=>a+Number(p.nominal||0),0);
var totalCashWajibSetor=Math.max(0,cashPenjualan+bonCash-totalPenChecked);
var totalTunai=DENOMS.reduce((a,d)=>a+Number(pecah[d]||0)*d,0);
var selisih=totalTunai-totalCashWajibSetor;

// qty per ukuran helper
var getQty=(p,uk)=>(p.items||[]).filter(it=>it.ukuran===uk).reduce((a,it)=>a+Number(it.qty||0),0);
var getHarga=(p,uk)=>{var it=(p.items||[]).filter(it=>it.ukuran===uk);return it.length>0?Number(it[0].price||0):0;};

// Load data riwayat ke form (mode edit)
function loadEditLog(r){
setSalesId(r.salesId||"");
setTgl(r.tanggal||toDay());
// Isi pecahan — handle key string (dari JSON/localStorage) atau int (dari state)
var p={};
DENOMS.forEach(d=>{
  var val=(r.pecah||{})[d]||(r.pecah||{})[String(d)]||"";
  p[d]=val||"";
});
setPecah(p);
setJadikanPinjaman(false);
setEditingLogId(r.id);
// Restore checkedPen dari log jika tersimpan
if(r.checkedPenIds&&r.checkedPenIds.length>0){
  var cp={};r.checkedPenIds.forEach(id=>{cp[id]=true;});
  setCheckedPen(cp);
}
setTimeout(()=>{if(formRef.current)formRef.current.scrollIntoView({behavior:"smooth",block:"start"});},100);
toast("✏️ Mode edit setoran "+fDs(r.tanggal)+" — ubah data lalu klik Simpan Perubahan");
}

function batalEdit(){
setEditingLogId(null);
setSalesId(isSales?user.id:"");
setTgl(toDay());
setPecah(()=>{var o={};DENOMS.forEach(d=>{o[d]="";});return o;});
setJadikanPinjaman(false);
}

function konfirmasi(){
var newAmb=[...(data.ambilan||[])];
if(selisih<0&&jadikanPinjaman){
  newAmb.unshift({id:uid(),karyawanId:salesId,karyawanNama:emp?.nama||"",nominal:Math.abs(selisih),ket:"Kurang setor "+fDs(tgl),tanggal:tgl});
}
var logEntry={
  id:editingLogId||uid(),
  tanggal:tgl,salesId,salesNama:emp?.nama||"",
  cashPenjualan,tfPenjualan,bonPenjualan,bonCash,bonTF,totalBonBayar,
  totalPotongan:totalPenChecked,totalCashWajibSetor,totalTunai,selisih,
  jadikanPinjaman:selisih<0&&jadikanPinjaman,
  pecah:Object.fromEntries(DENOMS.map(d=>[String(d),Number(pecah[d])||0])),
  checkedPenIds:Object.keys(checkedPen).filter(k=>checkedPen[k]),
  _edited:!!editingLogId,
  _editedAt:editingLogId?new Date().toISOString():undefined
};
var newLog=editingLogId
  ?(data.setoranLog||[]).map(x=>x.id===editingLogId?logEntry:x)
  :[logEntry,...(data.setoranLog||[])];
setData(d=>({...d,ambilan:newAmb,setoranLog:newLog}));
setPecah(()=>{var o={};DENOMS.forEach(d=>{o[d]="";});return o;});
setJadikanPinjaman(false);
setEditingLogId(null);
toast(editingLogId?"✓ Setoran diperbarui! Wajib setor: "+fR(totalCashWajibSetor):"✓ Setoran dikonfirmasi! Wajib setor: "+fR(totalCashWajibSetor));
}

var riwayat=(data.setoranLog||[]).filter(s=>!salesId||s.salesId===salesId).slice(0,20);

return <div ref={formRef}>
<STitle icon="💰" children="Setoran Sales"/>

{/* ── BANNER MODE EDIT ── */}
{editingLogId&&<div style={{background:"linear-gradient(90deg,#92400e,#b45309)",border:"1px solid #f59e0b",borderRadius:10,padding:"10px 16px",marginBottom:12,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
<div>
<div style={{color:"#fef3c7",fontWeight:800,fontSize:13}}>✏️ Mode Edit Setoran</div>
<div style={{color:"#fde68a",fontSize:11,marginTop:2}}>Data riwayat dimuat ke form. Ubah pengeluaran, pecahan kas, lalu klik <b>Simpan Perubahan</b>.</div>
</div>
<button onClick={batalEdit} style={{background:"rgba(0,0,0,.25)",border:"1px solid #f59e0b",borderRadius:7,padding:"6px 12px",color:"#fef3c7",cursor:"pointer",fontWeight:700,fontSize:11,flexShrink:0}}>✕ Batal Edit</button>
</div>}

<Card style={{width:"fit-content",maxWidth:"100%",minWidth:660}}>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(180px,260px))",gap:10}}>
<Sel label="Sales" value={salesId} onChange={v=>{if(!editingLogId)setSalesId(v);}} opts={[{v:"",l:"-- Pilih --"},...salesList.map(e=>({v:e.id,l:e.nama}))]}/>
<Inp label="Tanggal" type="date" value={tgl} onChange={v=>{if(!editingLogId)setTgl(v);}} ro={!!editingLogId}/>
</div>
{editingLogId&&<div style={{fontSize:10,color:"#f59e0b",marginTop:4}}>⚠️ Sales & tanggal dikunci saat mode edit. Batal edit untuk ganti.</div>}
</Card>

{salesId&&<>
{/* ── 1. PENJUALAN per konsumen ── */}
<Card style={{border:"1px solid "+C.blt,overflow:"hidden"}}>
<div style={{fontWeight:700,color:C.blt,marginBottom:8,fontSize:13}}>📦 Penjualan — {fDs(tgl)}</div>
{penjHari.length===0?<div style={{color:C.gl2,fontSize:12,fontStyle:"italic"}}>Tidak ada penjualan hari ini</div>:<>
<div style={{overflowX:"auto"}}>
<table style={{width:"100%",borderCollapse:"collapse",fontSize:11,minWidth:600}}>
<thead>
<tr style={{background:C.nav}}>
<th style={{padding:"5px 8px",color:C.gl2,textAlign:"left",border:"1px solid "+C.bdr,fontSize:10}} rowSpan={2}>Konsumen</th>
<th style={{padding:"5px 8px",color:C.gl2,textAlign:"center",border:"1px solid "+C.bdr,fontSize:10}} colSpan={2}>5,5 kg</th>
<th style={{padding:"5px 8px",color:C.gl2,textAlign:"center",border:"1px solid "+C.bdr,fontSize:10}} colSpan={2}>12 kg</th>
<th style={{padding:"5px 8px",color:C.gl2,textAlign:"center",border:"1px solid "+C.bdr,fontSize:10}} colSpan={2}>50 kg</th>
<th style={{padding:"5px 8px",color:C.gl2,textAlign:"right",border:"1px solid "+C.bdr,fontSize:10}} rowSpan={2}>Total</th>
<th style={{padding:"5px 8px",color:C.gl2,textAlign:"center",border:"1px solid "+C.bdr,fontSize:10}} rowSpan={2}>Transaksi</th>
</tr>
<tr style={{background:C.nav}}>
{["Qty","Harga","Qty","Harga","Qty","Harga"].map((h,i)=><th key={i} style={{padding:"4px 6px",color:C.gl2,textAlign:"center",border:"1px solid "+C.bdr,fontSize:9}}>{h}</th>)}
</tr>
</thead>
<tbody>
{penjHari.map((p,i)=>{
var q55=getQty(p,"5.5 kg");var h55=getHarga(p,"5.5 kg");
var q12=getQty(p,"12 kg");var h12=getHarga(p,"12 kg");
var q50=getQty(p,"50 kg");var h50=getHarga(p,"50 kg");
var byr=(p.bayar||"").toLowerCase();
return <tr key={p.id} style={{background:i%2===0?C.bg:C.nav,borderBottom:"1px solid "+C.bdr}}>
<td style={{padding:"4px 8px",color:C.wht,fontWeight:600,border:"1px solid "+C.bdr}}>{p.konsumen}</td>
<td style={{padding:"4px 6px",textAlign:"center",color:q55>0?C.glt:C.gl2,border:"1px solid "+C.bdr}}>{q55||"—"}</td>
<td style={{padding:"4px 6px",textAlign:"right",color:q55>0?C.wht:C.gl2,border:"1px solid "+C.bdr,fontSize:10}}>{q55>0?fR(h55):"—"}</td>
<td style={{padding:"4px 6px",textAlign:"center",color:q12>0?C.blt:C.gl2,border:"1px solid "+C.bdr}}>{q12||"—"}</td>
<td style={{padding:"4px 6px",textAlign:"right",color:q12>0?C.wht:C.gl2,border:"1px solid "+C.bdr,fontSize:10}}>{q12>0?fR(h12):"—"}</td>
<td style={{padding:"4px 6px",textAlign:"center",color:q50>0?C.olt:C.gl2,border:"1px solid "+C.bdr}}>{q50||"—"}</td>
<td style={{padding:"4px 6px",textAlign:"right",color:q50>0?C.wht:C.gl2,border:"1px solid "+C.bdr,fontSize:10}}>{q50>0?fR(h50):"—"}</td>
<td style={{padding:"4px 8px",textAlign:"right",fontWeight:700,color:C.wht,border:"1px solid "+C.bdr}}>{fR(p.total)}</td>
<td style={{padding:"4px 6px",textAlign:"center",border:"1px solid "+C.bdr,color:byr==="cash"?C.glt:byr==="bon"?C.rlt:C.blt,fontWeight:700,fontSize:10}}>{p.bayar}</td>
</tr>;})}
{/* Total Laku */}
<tr style={{background:C.nav,borderTop:"2px solid "+C.bdr,fontWeight:700}}>
<td style={{padding:"5px 8px",color:C.blt,fontStyle:"italic",border:"1px solid "+C.bdr,fontSize:10}}>Total Laku</td>
{(()=>{
var tq55=penjHari.reduce((a,p)=>a+getQty(p,"5.5 kg"),0);
var tq12=penjHari.reduce((a,p)=>a+getQty(p,"12 kg"),0);
var tq50=penjHari.reduce((a,p)=>a+getQty(p,"50 kg"),0);
return <>
<td style={{padding:"5px 6px",textAlign:"center",color:C.glt,border:"1px solid "+C.bdr}}>{tq55||"—"}</td>
<td style={{border:"1px solid "+C.bdr}}></td>
<td style={{padding:"5px 6px",textAlign:"center",color:C.blt,border:"1px solid "+C.bdr}}>{tq12||"—"}</td>
<td style={{border:"1px solid "+C.bdr}}></td>
<td style={{padding:"5px 6px",textAlign:"center",color:C.olt,border:"1px solid "+C.bdr}}>{tq50||"—"}</td>
<td style={{border:"1px solid "+C.bdr}}></td>
<td style={{padding:"5px 8px",textAlign:"right",color:C.wht,border:"1px solid "+C.bdr}}>{fR(penjHari.reduce((a,p)=>a+(p.total||0),0))}</td>
<td style={{border:"1px solid "+C.bdr}}></td>
</>;})()} 
</tr>
</tbody>
</table>
</div>
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:6,marginTop:8}}>
{[["Cash",fR(cashPenjualan),C.glt,penjCash.length+" trx"],["Transfer",fR(tfPenjualan),C.blt,penjTF.length+" trx"],["BON (piutang)",fR(bonPenjualan),"#9CA3AF",penjBon.length+" trx"]].map(x=><div key={x[0]} style={{background:C.nav,borderRadius:6,padding:"6px 10px",border:"1px solid "+C.bdr}}><div style={{fontSize:9,color:C.gl2}}>{x[0]}</div><div style={{fontSize:12,fontWeight:700,color:x[2]}}>{x[1]}</div><div style={{fontSize:9,color:C.gl2}}>{x[3]}</div></div>)}
</div>
</>}
</Card>

{/* ── 2. BAYAR BON ── */}
<Card style={{border:"1px solid "+C.olt}}>
<div style={{fontWeight:700,color:C.olt,marginBottom:8,fontSize:13}}>💳 Bayar Bon / Piutang</div>
{bonBayarList.length===0?<div style={{color:C.gl2,fontSize:12,fontStyle:"italic"}}>Tidak ada pembayaran BON hari ini</div>:<>
<table style={{width:"100%",borderCollapse:"collapse",fontSize:11,marginBottom:8}}>
<thead><tr style={{background:C.nav}}>
{["Konsumen","Nominal","Metode"].map(h=><th key={h} style={{padding:"5px 8px",color:C.gl2,textAlign:h==="Nominal"?"right":"left",border:"1px solid "+C.bdr,fontSize:10}}>{h}</th>)}
</tr></thead>
<tbody>
{bonBayarList.map((b,i)=><tr key={i} style={{background:i%2===0?C.bg:C.nav,borderBottom:"1px solid "+C.bdr}}>
<td style={{padding:"4px 8px",color:C.wht,border:"1px solid "+C.bdr}}>{b.konsumen}</td>
<td style={{padding:"4px 8px",textAlign:"right",fontWeight:700,color:C.wht,border:"1px solid "+C.bdr}}>{fR(b.jumlah||b.nominal||0)}</td>
<td style={{padding:"4px 8px",color:(b.metode||"cash").toLowerCase()==="cash"?C.glt:C.blt,fontWeight:700,border:"1px solid "+C.bdr,fontSize:10}}>{b.metode||"Cash"}</td>
</tr>)}
</tbody>
</table>
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:6}}>
{[["Cash",fR(bonCash),C.glt],["Transfer",fR(bonTF),C.blt],["TOTAL BAYAR BON",fR(totalBonBayar),C.olt]].map(x=><div key={x[0]} style={{background:C.nav,borderRadius:6,padding:"6px 10px",border:"1px solid "+C.bdr}}><div style={{fontSize:9,color:C.gl2}}>{x[0]}</div><div style={{fontSize:12,fontWeight:700,color:x[2]}}>{x[1]}</div></div>)}
</div>
</>}
</Card>

{/* ── 3. OPERASIONAL SALES ── */}
<Card style={{border:"1px solid "+C.bdr}}>
<div style={{fontWeight:700,color:C.gl2,marginBottom:6,fontSize:13}}>💸 Operasional Sales</div>
<div style={{fontSize:11,color:C.gl2,marginBottom:8}}>✅ Auto tercentang = atas nama {emp?.nama||"sales ini"}. Klik untuk ubah.</div>
{allPenHari.length===0?<div style={{color:C.gl2,fontSize:12,fontStyle:"italic"}}>Tidak ada pengeluaran hari ini</div>:
<table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
<thead><tr style={{background:C.nav}}>
{["","Kategori","Keterangan","Atas Nama","Metode","Nominal"].map(h=><th key={h} style={{padding:"5px 7px",color:C.gl2,textAlign:h==="Nominal"?"right":"left",border:"1px solid "+C.bdr,fontSize:10}}>{h}</th>)}
</tr></thead>
<tbody>
{allPenHari.map((p,i)=>{
var isMe=p.karyawanId===salesId||p.karyawanNama===(emp?.nama||"");
var checked=!!checkedPen[p.id];
return <tr key={p.id} style={{background:checked?(i%2===0?C.nav:C.bg):(i%2===0?C.bg:"transparent"),borderBottom:"1px solid "+C.bdr,cursor:"pointer",opacity:checked?1:.6}} onClick={()=>setCheckedPen(prev=>({...prev,[p.id]:!prev[p.id]}))}>
<td style={{padding:"4px 7px",border:"1px solid "+C.bdr}}><input type="checkbox" checked={checked} onChange={()=>{}} style={{cursor:"pointer"}}/></td>
<td style={{padding:"4px 7px",color:C.wht,border:"1px solid "+C.bdr,fontWeight:600}}>{p.kategori}</td>
<td style={{padding:"4px 7px",color:C.gl2,border:"1px solid "+C.bdr,fontSize:10}}>{p.ket||"—"}</td>
<td style={{padding:"4px 7px",border:"1px solid "+C.bdr}}>
<span style={{color:isMe?C.glt:C.gl2,fontSize:10,fontWeight:isMe?700:400}}>{p.karyawanNama||"—"}{isMe?" (auto)":""}</span>
</td>
<td style={{padding:"4px 7px",color:(p.metode||"cash").toLowerCase()==="cash"?C.glt:C.blt,border:"1px solid "+C.bdr,fontWeight:700,fontSize:10}}>{p.metode||"Cash"}</td>
<td style={{padding:"4px 7px",textAlign:"right",fontWeight:700,color:checked?C.rlt:C.gl2,border:"1px solid "+C.bdr}}>{fR(p.nominal)}</td>
</tr>;})}
<tr style={{background:C.nav,borderTop:"2px solid "+C.bdr,fontWeight:700}}>
<td colSpan={5} style={{padding:"5px 7px",color:C.gl2,border:"1px solid "+C.bdr,fontSize:11}}>Total Operasional (yang dicentang)</td>
<td style={{padding:"5px 7px",textAlign:"right",color:C.rlt,border:"1px solid "+C.bdr,fontSize:13}}>{fR(totalPenChecked)}</td>
</tr>
</tbody>
</table>}
</Card>

{/* ── 4. RINGKASAN WAJIB SETOR ── */}
<Card style={{border:"2px solid "+C.blt}}>
<div style={{fontWeight:700,color:C.blt,marginBottom:10,fontSize:13}}>🧾 Ringkasan Wajib Setor Cash</div>
<table style={{width:"100%",borderCollapse:"collapse",fontSize:12,marginBottom:12}}>
<tbody>
{[
["Cash Penjualan",cashPenjualan,C.glt,false],
["Bayar BON Cash",bonCash,C.glt,false],
["Pengeluaran Cash (dipotong)","-"+fR(totalPenChecked),C.rlt,false],
["TOTAL CASH WAJIB SETOR",totalCashWajibSetor,C.wht,true],
["Total TF Penjualan (info)",tfPenjualan,"#9CA3AF",false],
["Bayar BON TF (info)",bonTF,"#9CA3AF",false],
["Total BON Piutang Baru (info)",bonPenjualan,"#9CA3AF",false],
].map((r,i)=><tr key={i} style={{borderBottom:"1px solid "+C.bdr,background:r[3]?C.nav:"transparent"}}>
<td style={{padding:"6px 10px",color:r[3]?C.wht:C.gl2,fontWeight:r[3]?700:400}}>{r[0]}</td>
<td style={{padding:"6px 10px",textAlign:"right",color:r[2],fontWeight:r[3]?800:600,fontSize:r[3]?15:12}}>{typeof r[1]==="string"?r[1]:fR(r[1])}</td>
</tr>)}
</tbody>
</table>
</Card>

{/* ── 5. PECAHAN KAS ── */}
<Card>
<div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>💵 Pecahan Kas Fisik yang Disetor</div>
<div style={{border:"1px solid "+C.bdr,borderRadius:8,overflow:"hidden",marginBottom:10}}>
{DENOMS.map(d=><div key={d} style={{display:"grid",gridTemplateColumns:"1fr 85px 105px",padding:"5px 11px",borderTop:"1px solid "+C.bdr,alignItems:"center"}}>
<span style={{color:C.wht,fontSize:13,fontWeight:600}}>{fR(d)}</span>
<input type="number" value={pecah[d]} placeholder="0" onChange={e=>setPecah(u=>({...u,[d]:e.target.value}))} style={{background:C.nav,border:"1px solid "+C.bdr,borderRadius:6,padding:"4px 7px",color:C.wht,fontSize:12,outline:"none",width:74}}/>
<span style={{color:Number(pecah[d]||0)>0?C.olt:C.gl2,fontWeight:700,fontSize:12}}>{Number(pecah[d]||0)>0?fR(Number(pecah[d])*d):"-"}</span>
</div>)}
<div style={{display:"grid",gridTemplateColumns:"1fr 85px 105px",padding:"9px 11px",background:C.nav,borderTop:"2px solid "+C.bdr}}><b style={{color:C.wht}}>Total Tunai</b><span/><b style={{color:C.glt}}>{fR(totalTunai)}</b></div>
</div>
</Card>

{/* ── 6. REKONSILIASI ── */}
<Card style={{border:"2px solid "+(Math.abs(selisih)<1000?C.glt:C.rlt)}}>
<div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>🔄 Rekonsiliasi</div>
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:12}}>
{[["Wajib Setor",totalCashWajibSetor,C.wht],["Tunai Fisik",totalTunai,C.glt],["Selisih",selisih,selisih>=0?C.glt:C.rlt]].map(x=><div key={x[0]} style={{background:C.nav,borderRadius:8,padding:"8px 10px",border:"1px solid "+C.bdr,textAlign:"center"}}><div style={{fontSize:10,color:C.gl2}}>{x[0]}</div><div style={{fontSize:13,fontWeight:900,color:x[2]}}>{fR(x[1])}</div></div>)}
</div>
<div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"10px 14px",background:Math.abs(selisih)<1000?C.grn:C.rdk,borderRadius:8,marginBottom:12}}>
<span style={{color:"white",fontWeight:700}}>SELISIH {selisih>=0?"LEBIH":"KURANG"}</span>
<b style={{color:"white",fontSize:18}}>{fR(Math.abs(selisih))}</b>
</div>
{selisih<0&&<div style={{marginBottom:8}}>
<label style={{display:"flex",alignItems:"center",gap:8,cursor:"pointer",padding:"8px 12px",background:C.rdk,borderRadius:8}}>
<input type="checkbox" checked={jadikanPinjaman} onChange={e=>setJadikanPinjaman(e.target.checked)} style={{width:16,height:16}}/>
<span style={{color:"white",fontSize:12,fontWeight:700}}>💳 Jadikan pinjaman karyawan (kurang setor {fR(Math.abs(selisih))})</span>
</label>
</div>}
<Btn color={editingLogId?"orange":"green"} onClick={konfirmasi} dis={!salesId}>{editingLogId?"💾 Simpan Perubahan":"✓ Konfirmasi Setoran"}</Btn>
</Card>
</>}

{riwayat.length>0&&<Card>
<div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>📋 Riwayat Setoran</div>
<div style={{overflowX:"auto"}}>
<table style={{width:"100%",borderCollapse:"collapse",fontSize:11,minWidth:620}}>
<thead><tr style={{background:C.nav}}>
{["Tanggal","Sales","Cash Penj","Bayar BON","Potongan","Wajib Setor","Tunai","Selisih","Aksi"].map(h=><th key={h} style={{padding:"6px 8px",color:C.gl2,textAlign:h==="Selisih"||h==="Wajib Setor"||h==="Tunai"||h==="Cash Penj"||h==="Potongan"||h==="Bayar BON"?"right":"left",border:"1px solid "+C.bdr,fontSize:10,fontWeight:700}}>{h}</th>)}
</tr></thead>
<tbody>
{riwayat.map((r,ri)=>{
var ok=Math.abs(r.selisih||0)<1000;
return <tr key={r.id||ri} style={{background:ri%2===0?C.bg:C.nav,borderBottom:"1px solid "+C.bdr}}>
<td style={{padding:"5px 8px",color:C.wht,border:"1px solid "+C.bdr,fontWeight:600,whiteSpace:"nowrap"}}>{fDs(r.tanggal)}{r._edited&&<span style={{fontSize:8,color:C.olt,marginLeft:4,fontWeight:700}}>✏️</span>}</td>
<td style={{padding:"5px 8px",color:C.wht,border:"1px solid "+C.bdr}}>{r.salesNama}</td>
<td style={{padding:"5px 8px",textAlign:"right",color:C.glt,border:"1px solid "+C.bdr}}>{fR(r.cashPenjualan||0)}</td>
<td style={{padding:"5px 8px",textAlign:"right",color:C.olt,border:"1px solid "+C.bdr}}>{fR(r.bonCash||r.bonBayarHari||0)}</td>
<td style={{padding:"5px 8px",textAlign:"right",color:C.rlt,border:"1px solid "+C.bdr}}>{fR(r.totalPotongan||0)}</td>
<td style={{padding:"5px 8px",textAlign:"right",color:C.wht,fontWeight:700,border:"1px solid "+C.bdr}}>{fR(r.totalCashWajibSetor||r.bersihSetor||0)}</td>
<td style={{padding:"5px 8px",textAlign:"right",color:C.glt,border:"1px solid "+C.bdr}}>{fR(r.totalTunai||0)}</td>
<td style={{padding:"5px 8px",textAlign:"right",border:"1px solid "+C.bdr}}><b style={{color:ok?C.glt:C.rlt}}>{fR(r.selisih||0)}</b></td>
<td style={{padding:"4px 6px",border:"1px solid "+C.bdr,whiteSpace:"nowrap"}}>
<div style={{display:"flex",gap:4}}>
<button onClick={()=>setViewLog(r)} style={{background:C.gry,border:"none",borderRadius:5,padding:"3px 8px",color:"white",cursor:"pointer",fontSize:10,fontWeight:700}}>👁️</button>
<button onClick={()=>loadEditLog(r)} style={{background:editingLogId===r.id?C.olt:C.blu,border:"none",borderRadius:5,padding:"3px 8px",color:"white",cursor:"pointer",fontSize:10,fontWeight:700}}>{editingLogId===r.id?"📝 Aktif":"✏️"}</button>
<button onClick={()=>setShowPrint(r)} style={{background:C.grn,border:"none",borderRadius:5,padding:"3px 8px",color:"white",cursor:"pointer",fontSize:10,fontWeight:700}}>🖨️</button>
<button onClick={()=>setDelLog(r)} style={{background:C.rdk,border:"none",borderRadius:5,padding:"3px 8px",color:"white",cursor:"pointer",fontSize:10,fontWeight:700}}>🗑️</button>
</div>
</td>
</tr>;
})}
</tbody>
</table>
</div>
</Card>}

{/* ── MODAL VIEW DETAIL RIWAYAT ── */}
{viewLog&&!showPrint&&(()=>{
var r=viewLog;
var totalTunaiV=r.totalTunai||DENOMS.reduce((a,d)=>a+Number((r.pecah||{})[d]||(r.pecah||{})[String(d)]||0)*d,0);
var selisihV=r.selisih!=null?r.selisih:totalTunaiV-(r.totalCashWajibSetor||0);
var ok=Math.abs(selisihV)<1000;
return <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.75)",zIndex:999,display:"flex",alignItems:"center",justifyContent:"center",padding:16,overflowY:"auto"}}>
<div style={{background:C.card,borderRadius:12,width:"100%",maxWidth:480,padding:20,border:"1px solid "+C.bdr,maxHeight:"90vh",overflowY:"auto"}}>
{/* Header */}
<div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:14}}>
<div>
<div style={{fontWeight:800,color:C.wht,fontSize:14}}>📋 Detail Setoran</div>
<div style={{color:C.gl2,fontSize:12,marginTop:2}}>{fDs(r.tanggal)} · {r.salesNama}{r._edited&&<span style={{color:C.olt,marginLeft:6,fontSize:10}}>✏️ Pernah diedit</span>}</div>
</div>
<button onClick={()=>setViewLog(null)} style={{background:C.gry,border:"none",borderRadius:7,padding:"5px 10px",color:C.gl2,cursor:"pointer",fontWeight:700,fontSize:12}}>✕</button>
</div>

{/* Ringkasan kas */}
<div style={{background:C.nav,borderRadius:8,border:"1px solid "+C.bdr,overflow:"hidden",marginBottom:12}}>
<div style={{padding:"7px 12px",background:C.bg,borderBottom:"1px solid "+C.bdr,fontSize:11,fontWeight:700,color:C.gl2}}>RINGKASAN KAS</div>
{[
["Cash Penjualan",r.cashPenjualan||0,C.glt],
["Bayar BON Cash",r.bonCash||r.bonBayarHari||0,C.olt],
["Potongan Operasional","−"+fR(r.totalPotongan||0),C.rlt],
["WAJIB SETOR",r.totalCashWajibSetor||r.bersihSetor||0,C.wht],
].map((x,i)=><div key={i} style={{display:"flex",justifyContent:"space-between",padding:"6px 12px",borderBottom:"1px solid "+C.bdr,background:i===3?C.bg:"transparent"}}>
<span style={{color:i===3?C.wht:C.gl2,fontWeight:i===3?700:400,fontSize:12}}>{x[0]}</span>
<span style={{color:x[2],fontWeight:i===3?800:600,fontSize:i===3?14:12}}>{typeof x[1]==="string"?x[1]:fR(x[1])}</span>
</div>)}
{r.tfPenjualan>0&&<div style={{display:"flex",justifyContent:"space-between",padding:"5px 12px",borderBottom:"1px solid "+C.bdr}}>
<span style={{color:C.gl2,fontSize:11}}>TF Penjualan (info)</span>
<span style={{color:"#6B7280",fontSize:11}}>{fR(r.tfPenjualan||0)}</span>
</div>}
{r.bonTF>0&&<div style={{display:"flex",justifyContent:"space-between",padding:"5px 12px"}}>
<span style={{color:C.gl2,fontSize:11}}>Bayar BON TF (info)</span>
<span style={{color:"#6B7280",fontSize:11}}>{fR(r.bonTF||0)}</span>
</div>}
</div>

{/* Pecahan kas */}
<div style={{background:C.nav,borderRadius:8,border:"1px solid "+C.bdr,overflow:"hidden",marginBottom:12}}>
<div style={{padding:"7px 12px",background:C.bg,borderBottom:"1px solid "+C.bdr,fontSize:11,fontWeight:700,color:C.gl2}}>💵 PECAHAN KAS FISIK</div>
{(()=>{
var adaPecah=DENOMS.some(d=>Number((r.pecah||{})[d]||(r.pecah||{})[String(d)]||0)>0);
if(!adaPecah)return <div style={{padding:"10px 12px",color:C.gl2,fontSize:11,fontStyle:"italic"}}>Pecahan tidak dicatat</div>;
return DENOMS.map(d=>{
var lbr=Number((r.pecah||{})[d]||(r.pecah||{})[String(d)]||0);
if(!lbr)return null;
return <div key={d} style={{display:"grid",gridTemplateColumns:"1fr 60px 1fr",padding:"5px 12px",borderBottom:"1px solid "+C.bdr,alignItems:"center"}}>
<span style={{color:C.wht,fontSize:12}}>{fR(d)}</span>
<span style={{color:C.glt,fontWeight:700,fontSize:13,textAlign:"center"}}>{lbr}</span>
<span style={{color:C.olt,fontSize:12,textAlign:"right"}}>{fR(lbr*d)}</span>
</div>;
});
})()}
<div style={{display:"grid",gridTemplateColumns:"1fr 60px 1fr",padding:"7px 12px",background:C.bg,borderTop:"2px solid "+C.bdr}}>
<b style={{color:C.wht,fontSize:12}}>Total Tunai</b><span/>
<b style={{color:C.glt,fontSize:14,textAlign:"right"}}>{fR(totalTunaiV)}</b>
</div>
</div>

{/* Rekonsiliasi */}
<div style={{padding:"10px 14px",background:ok?C.grn:C.rdk,borderRadius:8,display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
<span style={{color:"white",fontWeight:700,fontSize:12}}>SELISIH {selisihV>=0?"LEBIH":"KURANG"}</span>
<b style={{color:"white",fontSize:16}}>{fR(Math.abs(selisihV))}</b>
</div>

{/* Tombol aksi */}
<div style={{display:"flex",gap:8,justifyContent:"flex-end",flexWrap:"wrap"}}>
<button onClick={()=>{setViewLog(null);setDelLog(r);}} style={{background:C.rdk,border:"none",borderRadius:8,padding:"8px 14px",color:"white",cursor:"pointer",fontWeight:700,fontSize:12}}>🗑️ Hapus</button>
<button onClick={()=>{setViewLog(null);setShowPrint(r);}} style={{background:C.grn,border:"none",borderRadius:8,padding:"8px 14px",color:"white",cursor:"pointer",fontWeight:700,fontSize:12}}>🖨️ Cetak</button>
<button onClick={()=>{setViewLog(null);loadEditLog(r);}} style={{background:C.blu,border:"none",borderRadius:8,padding:"8px 14px",color:"white",cursor:"pointer",fontWeight:700,fontSize:12}}>✏️ Edit</button>
</div>
</div>
</div>;
})()}

{/* ── MODAL SLIP CETAK SETORAN SALES ── */}
{showPrint&&(()=>{
var r=showPrint;
var totalTunaiSlip=r.totalTunai||DENOMS.reduce((a,d)=>a+Number((r.pecah||{})[d]||0)*d,0);
var selisihSlip=r.selisih!=null?r.selisih:totalTunaiSlip-(r.totalCashWajibSetor||0);
var tglD=new Date((r.tanggal||toDay())+"T00:00:00");
var hariLabel=["Minggu","Senin","Selasa","Rabu","Kamis","Jumat","Sabtu"][tglD.getDay()];
var tglLabel=tglD.toLocaleDateString("id-ID",{day:"2-digit",month:"long",year:"numeric"});
return <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.8)",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center",padding:16,overflowY:"auto"}}>
<div style={{background:"white",borderRadius:12,width:"100%",maxWidth:430,boxShadow:"0 20px 60px rgba(0,0,0,.5)",overflow:"hidden"}}>
<div id="_slip_setoran_sales" style={{background:"white",padding:20,fontFamily:"Arial,sans-serif",color:"#111",borderRadius:12}}>
{/* Header */}
<div style={{background:"linear-gradient(135deg,#0a1f44 0%,#122d5e 100%)",padding:"14px 18px 12px",display:"flex",justifyContent:"space-between",alignItems:"center",margin:-20,marginBottom:0}}>
<div style={{display:"flex",alignItems:"center",gap:10}}>
<CompanyLogo h={40} variant="dark"/>
</div>
<div style={{display:"flex",alignItems:"center",gap:5}}>
<PertaminaLogo h={26} variant="dark"/>
</div>
</div>
{/* 3-color divider */}
<div style={{height:3,display:"flex",margin:"0 -20px 14px"}}><div style={{flex:1,background:"#1565c0"}}/><div style={{flex:1,background:"#6ab04c"}}/><div style={{flex:1,background:"#e53935"}}/></div>
{/* Judul */}
<div style={{textAlign:"center",marginBottom:12}}>
<div style={{fontSize:13,fontWeight:700,color:"#0a1f44",letterSpacing:.5}}>BUKTI SETORAN SALES</div>
<div style={{fontSize:12,color:"#555"}}>{hariLabel} &nbsp;|&nbsp; {tglLabel}</div>
</div>
{/* Info Sales */}
<div style={{width:"90%",margin:"0 auto 10px",background:"#F0F7FF",borderRadius:8,border:"1.5px solid #BFDBFE",overflow:"hidden"}}>
<div style={{background:"#0a1f44",padding:"7px 14px"}}><span style={{fontSize:11,color:"#93C5FD",fontWeight:700,letterSpacing:1}}>SALES: {r.salesNama||"—"}</span></div>
<div style={{padding:"10px 14px",display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
{[["Cash Penjualan",r.cashPenjualan||0,"#166534"],["Bayar BON Cash",r.bonCash||r.bonBayarHari||0,"#854d0e"],["Potongan Operasional",r.totalPotongan||0,"#991b1b"],["WAJIB SETOR",r.totalCashWajibSetor||r.bersihSetor||0,"#0a1f44"]].map(x=><div key={x[0]} style={{background:"white",borderRadius:6,padding:"7px 10px",border:"1px solid #DBEAFE"}}>
<div style={{fontSize:9,color:"#475569",marginBottom:2}}>{x[0]}</div>
<div style={{fontSize:12,fontWeight:700,color:x[2]}}>Rp {Number(x[1]).toLocaleString("id-ID")}</div>
</div>)}
</div>
</div>
{/* Tabel Pecahan */}
<table style={{width:"90%",margin:"0 auto",borderCollapse:"collapse",fontSize:12,marginBottom:12}}>
<thead><tr style={{background:"#0a1f44"}}>
{["PECAHAN","LBR","JUMLAH"].map(h=><th key={h} style={{padding:"7px 10px",color:"white",textAlign:h==="JUMLAH"?"right":"left",border:"1px solid #1e3a5f",fontSize:10,letterSpacing:.8}}>{h}</th>)}
</tr></thead>
<tbody>
{(()=>{
return <>
{DENOMS.map(d=>{
var lbr=Number((r.pecah||{})[d]||0);
var jml=lbr*d;
return <tr key={d} style={{borderBottom:"1px solid #E5E7EB",background:lbr>0?"white":"#FAFAFA"}}>
<td style={{padding:"6px 10px",color:"#111",border:"1px solid #E2E8F0",fontSize:12}}>Rp {Number(d).toLocaleString("id-ID")}</td>
<td style={{padding:"6px 8px",textAlign:"left",color:lbr>0?"#0a1f44":"#CBD5E1",fontWeight:lbr>0?700:400,border:"1px solid #E2E8F0",fontSize:13}}>{lbr||"—"}</td>
<td style={{padding:"6px 10px",textAlign:"right",color:jml>0?"#0a1f44":"#CBD5E1",fontWeight:jml>0?700:400,border:"1px solid #E2E8F0",fontSize:12}}>{jml>0?"Rp "+jml.toLocaleString("id-ID"):"—"}</td>
</tr>;})}
<tr style={{background:"#0a1f44",fontWeight:700}}>
<td colSpan={2} style={{padding:"8px 10px",color:"white",border:"1px solid #1e3a5f",fontSize:12}}>TOTAL TUNAI</td>
<td style={{padding:"8px 10px",textAlign:"right",color:"white",fontSize:14,border:"1px solid #1e3a5f",fontWeight:900}}>Rp {totalTunaiSlip.toLocaleString("id-ID")}</td>
</tr>
{/* Selisih */}
<tr style={{background:Math.abs(selisihSlip)<1000?"#f0fdf4":"#fef2f2"}}>
<td colSpan={2} style={{padding:"7px 10px",color:Math.abs(selisihSlip)<1000?"#166534":"#991b1b",border:"1px solid #E2E8F0",fontSize:11,fontWeight:700}}>SELISIH {selisihSlip>=0?"(LEBIH)":"(KURANG)"}</td>
<td style={{padding:"7px 10px",textAlign:"right",color:Math.abs(selisihSlip)<1000?"#166534":"#991b1b",border:"1px solid #E2E8F0",fontSize:13,fontWeight:900}}>Rp {Math.abs(selisihSlip).toLocaleString("id-ID")}</td>
</tr>
</>;
})()}
</tbody>
</table>
{/* Tanda tangan */}
<div style={{width:"90%",margin:"0 auto 10px",display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
{["Disetor Oleh","Diterima Oleh"].map(lb=><div key={lb} style={{border:"1px solid #DBEAFE",borderRadius:6,padding:"6px 10px",textAlign:"center",background:"#F8FAFF"}}>
<div style={{fontSize:9,color:"#94A3B8",marginBottom:20}}>{lb}</div>
<div style={{borderTop:"1px solid #CBD5E1",paddingTop:4,fontSize:9,color:"#475569"}}>{lb==="Disetor Oleh"?r.salesNama||"—":"Kasir"}</div>
</div>)}
</div>
<div style={{textAlign:"center",fontSize:9,color:"#94A3B8",marginBottom:4}}>
Jl. Jendral Sudirman No. 80 · Banda Aceh &nbsp;|&nbsp; 081269002121 / (0651) 21221
</div>
</div>
{/* Tombol aksi */}
<div style={{padding:"10px 16px",display:"flex",gap:8,flexWrap:"wrap",justifyContent:"center",background:"#f8faff"}}>
<button onClick={()=>doPrint("_slip_setoran_sales")} style={{background:"#0a1f44",color:"white",border:"none",padding:"8px 16px",borderRadius:7,fontSize:12,cursor:"pointer",fontWeight:700}}>🖨️ Cetak / PDF</button>
<button onClick={()=>doDownloadPNG("_slip_setoran_sales","Slip-Setoran-"+r.salesNama+"-"+r.tanggal+".png")} style={{background:"#1D6A96",color:"white",border:"none",padding:"8px 14px",borderRadius:7,fontSize:12,cursor:"pointer",fontWeight:700}}>💾 Download PNG</button>
<button onClick={()=>doCopyPNG("_slip_setoran_sales")} style={{background:"#145A32",color:"white",border:"none",padding:"8px 14px",borderRadius:7,fontSize:12,cursor:"pointer",fontWeight:700}}>📋 Copy PNG</button>
<button onClick={()=>setShowPrint(null)} style={{background:"#6B7280",color:"white",border:"none",padding:"8px 14px",borderRadius:7,fontSize:12,cursor:"pointer",fontWeight:700}}>✕ Tutup</button>
</div>
</div>
</div>;
})()}

{delLog&&<ConfirmDel msg={"Hapus riwayat setoran "+delLog.salesNama+" tanggal "+fDs(delLog.tanggal)+"? Data tidak bisa dikembalikan."} onCancel={()=>setDelLog(null)} onConfirm={()=>{
setData(d=>({...d,setoranLog:(d.setoranLog||[]).filter(x=>x.id!==delLog.id)}));
if(editingLogId===delLog.id)batalEdit();
setDelLog(null);
toast("✓ Riwayat setoran dihapus");
}}/>}

</div>;
}

// ─── TUTUP BUKU v4 (harian + bulanan, PDF/PNG) ────────────────────────────────
// ── TabelStokBulanan: tampil tabel per hari dalam sebulan ──
function TabelStokBulanan({data,bulan,compact}){
  var C=useTheme();
  var rows=buildStokHarian(data,bulan);
  var uk=["12 kg","5.5 kg","50 kg"];
  var ukLabel=["12kg","5,5kg","50kg"];
  var headerGroups=[
    {label:"STOK AWAL",cols:["isi","TK"],color:"#1E40AF"},
    {label:"TABUNG MASUK",cols:["isi","TK"],color:"#065F46"},
    {label:"TABUNG KELUAR",cols:["isi","TK"],color:"#7F1D1D"},
    {label:"STOK AKHIR",cols:["isi","TK","Titip","Total"],color:"#1E3A5F"},
  ];
  var thS=function(c,align){return{padding:"5px 6px",color:"white",fontWeight:700,fontSize:9,textAlign:align||"center",borderRight:"1px solid rgba(255,255,255,.15)",background:c,whiteSpace:"nowrap"};};
  var tdS=function(color,bold){return{padding:"4px 6px",textAlign:"center",color:color||C.wht,fontWeight:bold?700:400,fontSize:11,borderRight:"1px solid "+C.bdr,borderBottom:"1px solid "+C.bdr,whiteSpace:"nowrap"};};

  return <div style={{overflowX:"auto"}}>
  <table style={{borderCollapse:"collapse",fontSize:11,width:"100%",minWidth:900}}>
  <thead>
  <tr>
    <th rowSpan={3} style={{...thS("#0F172A"),padding:"6px 8px",fontSize:10,minWidth:28}}>No</th>
    <th rowSpan={3} style={{...thS("#0F172A"),minWidth:90,textAlign:"left",padding:"6px 8px"}}>Hari/Tanggal</th>
    {["STOK AWAL","TABUNG MASUK","TABUNG KELUAR"].map((g,gi)=><th key={g} colSpan={uk.length*2} style={{...thS(["#1E40AF","#065F46","#7F1D1D"][gi]),padding:"6px 4px",letterSpacing:.5}}>{g}</th>)}
    <th colSpan={uk.length*4} style={{...thS("#1E3A5F"),padding:"6px 4px",letterSpacing:.5}}>STOK AKHIR</th>
  </tr>
  <tr>
    {["STOK AWAL","TABUNG MASUK","TABUNG KELUAR"].map((g,gi)=>uk.map((s,si)=><th key={g+s} colSpan={2} style={{...thS(["#1E40AF","#065F46","#7F1D1D"][gi]),fontSize:9}}>{ukLabel[si]}</th>))}
    {uk.map((s,si)=><th key={"akhir"+s} colSpan={4} style={{...thS("#1E3A5F"),fontSize:9}}>{ukLabel[si]}</th>)}
  </tr>
  <tr>
    {["STOK AWAL","TABUNG MASUK","TABUNG KELUAR"].map((g,gi)=>uk.map(s=>["isi","TK"].map(c=><th key={g+s+c} style={{...thS(["#1E40AF","#065F46","#7F1D1D"][gi]),fontSize:8}}>{c}</th>)))}
    {uk.map(s=>["isi","TK","Titip","Total"].map(c=><th key={"akhir"+s+c} style={{...thS("#1E3A5F"),fontSize:8,fontWeight:c==="Total"?900:700}}>{c}</th>))}
  </tr>
  </thead>
  <tbody>
  {rows.map((r,i)=>{
    var isLibur=!r.adaTransaksi&&!r.inject;
    var bg=i%2===0?C.nav:C.bg;
    return <tr key={r.tgl} style={{background:isLibur?"rgba(100,100,100,.1)":bg}}>
    <td style={tdS(C.gl2)}>{r.d}</td>
    <td style={{...tdS(),textAlign:"left",padding:"4px 8px"}}>
      <div style={{fontWeight:700,color:isLibur?C.gl2:C.wht,fontSize:11}}>{r.dayName}</div>
      <div style={{fontSize:9,color:C.gl2}}>{fDs(r.tgl)}{r.inject?<span style={{color:"#F59E0B",marginLeft:4}}>★</span>:""}{isLibur?<span style={{color:C.gl2,marginLeft:4}}>—</span>:""}</div>
    </td>
    {/* STOK AWAL */}
    {uk.map(s=>[
      <td key={"ai"+s} style={tdS(C.blt)}>{r.awalIsi[s]}</td>,
      <td key={"ak"+s} style={tdS(C.gl2)}>{r.awalTK[s]}</td>
    ])}
    {/* TABUNG MASUK */}
    {uk.map(s=>[
      <td key={"mi"+s} style={tdS(r.masukIsi[s]>0?C.glt:C.gl2,r.masukIsi[s]>0)}>{r.masukIsi[s]||"-"}</td>,
      <td key={"mk"+s} style={tdS(r.masukTK[s]>0?"#F59E0B":C.gl2,r.masukTK[s]>0)}>{r.masukTK[s]||"-"}</td>
    ])}
    {/* TABUNG KELUAR */}
    {uk.map(s=>[
      <td key={"ki"+s} style={tdS(r.keluarIsi[s]>0?C.rlt:C.gl2,r.keluarIsi[s]>0)}>{r.keluarIsi[s]||"-"}</td>,
      <td key={"kk"+s} style={tdS(r.keluarTK[s]!==0?"#F59E0B":C.gl2,r.keluarTK[s]!==0)}>{r.keluarTK[s]!==0?Math.abs(r.keluarTK[s]):"-"}</td>
    ])}
    {/* STOK AKHIR */}
    {uk.map(s=>[
      <td key={"eisi"+s} style={tdS(C.glt,true)}>{r.akhirIsi[s]}</td>,
      <td key={"etk"+s} style={tdS(C.gl2)}>{r.akhirTK[s]}</td>,
      <td key={"etitip"+s} style={tdS(C.blt)}>{r.titipSnap[s]}</td>,
      <td key={"etot"+s} style={tdS(C.olt,true)}>{r.total[s]}</td>
    ])}
    </tr>;
  })}
  </tbody>
  </table>
  <div style={{marginTop:6,fontSize:10,color:C.gl2}}>★ = Inject manual stok awal &nbsp;|&nbsp; — = Tidak ada transaksi &nbsp;|&nbsp; TK = Tabung Kosong</div>
  </div>;
}

function TabelStokHarian({data,tgl}){
var C=useTheme();
var prevTB=(data.tutupBuku||[]).filter(r=>r.tanggal&&r.tanggal<tgl).sort((a,b)=>b.tanggal.localeCompare(a.tanggal))[0];
var rows=SIZES.map(s=>{
  var stokAwal=prevTB?.detail?.stokSnapshot?.[s]?.isi??((data.stock||{})[s]||0);
  var doMasuk=(data.doList||[]).filter(d=>d.tanggal===tgl&&d.ukuran===s&&(d.status||"diterima")==="diterima").reduce((a,d)=>a+Number(d.qty||0),0);
  var terjual=(data.penjualan||[]).filter(p=>p.tanggal===tgl).reduce((a,p)=>a+(p.items||[]).filter(it=>it.ukuran===s).reduce((b,it)=>b+Number(it.qty||0),0),0);
  var sisaAkhir=(data.stock||{})[s]||0;
  var totalHariIni=stokAwal+doMasuk;
  return{s,stokAwal,doMasuk,totalHariIni,terjual,sisaAkhir};
});
return <div style={{overflowX:"auto"}}>
<table style={{width:"100%",borderCollapse:"collapse",fontSize:12,fontFamily:"inherit"}}>
<thead><tr style={{background:C.nav}}>
<th style={{padding:"8px 12px",color:C.gl2,fontWeight:700,textAlign:"left",borderBottom:"2px solid "+C.bdr,fontSize:11}}>Keterangan</th>
{SIZES.map(s=><th key={s} style={{padding:"8px 12px",color:C.wht,fontWeight:700,textAlign:"center",borderBottom:"2px solid "+C.bdr,fontSize:11}}>{s}</th>)}
</tr></thead>
<tbody>
{[
["Stok Awal",rows.map(r=>r.stokAwal),C.gl2,false],
["DO Masuk (Diterima)",rows.map(r=>r.doMasuk),C.blt,false],
["Total Stok Harian",rows.map(r=>r.totalHariIni),C.wht,true],
["Penjualan",rows.map(r=>r.terjual),C.rlt,false],
["SISA STOK AKHIR",rows.map(r=>r.sisaAkhir),C.glt,true],
].map((row,i)=><tr key={i} style={{borderBottom:"1px solid "+C.bdr,background:row[3]?C.nav:"transparent"}}>
<td style={{padding:"7px 12px",color:row[3]?C.wht:C.gl2,fontWeight:row[3]?700:400}}>{row[0]}</td>
{row[1].map((v,j)=><td key={j} style={{padding:"7px 12px",textAlign:"center",fontWeight:row[3]?800:600,fontSize:row[3]?14:12,color:row[2]}}>{v}</td>)}
</tr>)}
</tbody>
</table>
</div>;
}
function TutupBukuMod({data,setData,toast}){
var C=useTheme();var[tab,setTab]=useState("harian");
var[tgl,setTgl]=useState(toDay());var[bln,setBln]=useState(toMonth());
var[cashLaci,setCashLaci]=useState("");var[rekBSI,setRekBSI]=useState("");var[rekBCA,setRekBCA]=useState("");
var[pecah,setPecah]=useState(()=>{var o={};DENOMS.forEach(d=>{o[d]="";});return o;});
var[jadikanPinjaman,setJadikanPinjaman]=useState(false);
var[viewTB,setViewTB]=useState(null);// modal lihat
var[editTB,setEditTB]=useState(null);// record yang sedang diedit
useEffect(()=>{if(editTB){setTgl(editTB.tanggal);setCashLaci(String(editTB.cashLaci||""));setRekBSI(String(editTB.rekBSI||""));setRekBCA(String(editTB.rekBCA||""));setPemasukanLain(String(editTB.pemasukanLain||""));var p={};DENOMS.forEach(function(d){p[d]=(editTB.pecah||{})[d]||(editTB.pecah||{})[String(d)]||"";});setPecah(p);}else{setTgl(toDay());setCashLaci("");setRekBSI("");setRekBCA("");setPecah(function(){var o={};DENOMS.forEach(function(d){o[d]="";});return o;});}},[editTB]);

// Harian P&L
var pHari=(data.penjualan||[]).filter(e=>e.tanggal===tgl);
var omzetH=pHari.reduce((a,e)=>a+(e.total||0),0);
var marginH=pHari.reduce((a,e)=>a+(e.margin||0),0);
var hppH=omzetH-marginH;
// Jualan Lain (gas kaleng, aksesoris) — TIDAK masuk omzet/laba, hanya cash/TF masuk tambahan
var jualLainH=(data.jualanLain||[]).filter(j=>j.tanggal===tgl);
var omzetJualLainH=jualLainH.reduce((a,j)=>a+Number(j.total||0),0);
// Kategori pengeluaran non-operasional & fungsi isNonOps sekarang pakai versi global (lihat dekat daysInMonth)
var penH=(data.pengeluaran||[]).filter(e=>e.tanggal===tgl);
var totalOutH=penH.filter(e=>!isNonOps(e.kategori)).reduce((a,e)=>a+Number(e.nominal||0),0);
var labaBersihH=marginH-totalOutH;
var cashInH=pHari.filter(e=>e.bayar==="cash").reduce((a,e)=>a+(e.total||0),0)+jualLainH.filter(j=>j.bayar==="cash").reduce((a,j)=>a+Number(j.total||0),0);
var tfInH=pHari.filter(e=>e.bayar==="transfer").reduce((a,e)=>a+(e.total||0),0)+jualLainH.filter(j=>j.bayar==="transfer").reduce((a,j)=>a+Number(j.total||0),0);
var bonInH=pHari.filter(e=>e.bayar==="bon").reduce((a,e)=>a+(e.total||0),0);

// Asset kalkulasi
var piutangA=calcTotalPiutang(data);
var nilaiStokA=calcNilaiStokFIFO(data);
var pinjamanA=Math.max(0,calcPinjamanKaryawan(data));
var kosong55=getKosong(data,"5.5 kg");var kosong12=getKosong(data,"12 kg");var kosong50=getKosong(data,"50 kg");
var totalPecah=DENOMS.reduce((a,d)=>a+Number(pecah[d]||0)*d,0);
var cashTotal=(Number(cashLaci)||0)+(Number(rekBSI)||0)+(Number(rekBCA)||0);

// DO Gantung = DO belum diklik Diterima/Sangkut
// DO Gantung: hanya DO yang dibuat pada/sebelum tanggal tutup buku & belum diterima
var doGantung=(data.doList||[]).filter(d=>d.status==="gantung"&&(d.tanggal||"")<=tgl);
var nilaiDOGantung=doGantung.reduce((a,d)=>a+Number(d.totalHPP||0),0);
var doSangkut=(data.doList||[]).filter(d=>d.status==="sangkut"&&(d.tanggal||"")<=tgl);
var nilaiDOSangkut=doSangkut.reduce((a,d)=>a+Number(d.totalHPP||0),0);

// Asset Tabung Milik PT
var titipLuarBal={};
(data.titipList||[]).forEach(t=>{if(t.tipe==="titip_luar"||t.tipe==="tarik_luar"){var m=t.tipe==="titip_luar"?1:-1;(t.items||[]).forEach(it=>{titipLuarBal[it.ukuran]=(titipLuarBal[it.ukuran]||0)+m*Number(it.qty||0);});}});

var hargaTbg=data.company?.hargaTbgKosong||{};
var assetTabungMilikPT=SIZES.reduce((a,s)=>{
  var isiS=(data.stock||{})[s]||0;
  var kosS=getKosong(data,s);
  var titipS=getTitipTotal(data.titipList,s);
  var totalS=isiS+kosS+titipS;
  var titipLuar=Math.max(0,titipLuarBal[s]||0);
  var milikPT=Math.max(0,totalS-titipLuar);
  return a+milikPT*(hargaTbg[s]||0);
},0);
var assetArmada=Number(data.company?.assetArmada)||0;

// PIUTANG & MODAL BERJALAN = Stok Isi + DO Gantung + BON Konsumen + Pinjaman Karyawan
var piutangModal=nilaiStokA+nilaiDOGantung+piutangA+pinjamanA;

// TOTAL CASH FLOW = Total Cash + Bank + Piutang & Modal Berjalan
var cashFlowOmset=cashTotal+piutangModal;

// Total Asset
var assetValue=cashFlowOmset+assetTabungMilikPT+assetArmada;

// Kurang setor yang dijadikan pinjaman hari ini — HANYA INFO, tidak mempengaruhi laba.
// Uang ini sudah otomatis masuk piutang karyawan (pinjamanA) di sisi aset,
// sehingga cash flow tetap balance tanpa perlu mengurangi laba.
var kurangSetorHari=(data.ambilan||[]).filter(a=>a.tanggal===tgl&&(a.ket||"").toLowerCase().includes("kurang setor")).reduce((a,x)=>a+Number(x.nominal||0),0);

// Prev tutup buku
var prevTB=(data.tutupBuku||[]).filter(r=>tab==="harian"?r.tanggal<tgl:r.bulan&&r.bulan<bln).sort((a,b)=>b.tanggal.localeCompare(a.tanggal))[0];
var cashFlowKemarin=prevTB?.cashFlowOmset||0;
var tglTBKemarin=prevTB?.tanggal||null;// tanggal TB terakhir yang dipakai
// labaBersihEfektif: HANYA laba bersih operasional. Pemasukan Lainnya (topup modal/saham)
// adalah penambah kas/aset, BUKAN laba — tidak dihitung di sini.
var labaBersihEfektif=labaBersihH;
var deltaSelisih=cashFlowOmset-cashFlowKemarin-labaBersihEfektif-(Number(pemasukanLain)||0);

// Pemasukan lainnya
var[pemasukanLain,setPemasukanLain]=useState("");

// Bulanan aggregation
var pBln=(data.penjualan||[]).filter(e=>(e.tanggal||"").startsWith(bln));
var omzetB=pBln.reduce((a,e)=>a+(e.total||0),0);
var marginB=pBln.reduce((a,e)=>a+(e.margin||0),0);
var penB=(data.pengeluaran||[]).filter(e=>(e.tanggal||"").startsWith(bln));
var totalOutB=penB.reduce((a,e)=>a+Number(e.nominal||0),0);
var labaBersihB=marginB-totalOutB;
var dim=daysInMonth(bln);
var doBln=(data.doList||[]).filter(e=>(e.tanggal||"").startsWith(bln)&&(e.status||"diterima")==="diterima");
var doTripBln=[];(data.doTrip||[]).filter(t=>(t.tanggal||"").startsWith(bln)).forEach(function(trip){(trip.items||[]).forEach(function(it){if((it.status||"diterima")==="diterima"){doTripBln.push({tanggal:trip.tanggal,totalHPP:Number(it.qty||0)*Number(it.hppUnit||0)});}});});
var hppBln=doBln.reduce((a,e)=>a+Number(e.totalHPP||0),0)+doTripBln.reduce((a,e)=>a+Number(e.totalHPP||0),0);
var cashFlowKumul=0;
var chartDataB=[];for(var d=1;d<=dim;d++){var ds=bln+"-"+String(d).padStart(2,"0");var pp=pBln.filter(x=>x.tanggal===ds);var oz=pp.reduce((a,x)=>a+(x.total||0),0);var mg=pp.reduce((a,x)=>a+(x.margin||0),0);var pn=penB.filter(x=>x.tanggal===ds).reduce((a,x)=>a+Number(x.nominal||0),0);var hpp=doBln.filter(x=>x.tanggal===ds).reduce((a,x)=>a+Number(x.totalHPP||0),0)+doTripBln.filter(x=>x.tanggal===ds).reduce((a,x)=>a+Number(x.totalHPP||0),0);var labaBersih=mg-pn;cashFlowKumul+=labaBersih;chartDataB.push({hari:String(d),omzet:oz,hpp,marginKotor:mg,pengeluaran:pn,labaBersih,cashFlow:cashFlowKumul});}
// Top 5 hari omzet
var top5=chartDataB.slice().sort((a,b)=>b.omzet-a.omzet).slice(0,5);
// Pengeluaran per kategori bulanan
var katPenMap={};penB.forEach(p=>{katPenMap[p.kategori]=(katPenMap[p.kategori]||0)+Number(p.nominal||0);});
var katPenArr=Object.entries(katPenMap).sort((a,b)=>b[1]-a[1]);

function saveHarian(){
var penjualanDetail=(data.penjualan||[]).filter(e=>e.tanggal===tgl).map(e=>({noInv:e.noInv,konsumen:e.konsumen,total:e.total,bayar:e.bayar,items:e.items}));
var pengeluaranDetail=(data.pengeluaran||[]).filter(e=>e.tanggal===tgl).map(e=>({kategori:e.kategori,nominal:e.nominal,ket:e.ket,karyawan:e.karyawanNama||""}));
var doDetail=(data.doList||[]).filter(e=>e.tanggal===tgl).map(e=>({trip:e.trip,sppbe:e.sppbe,ukuran:e.ukuran,qty:e.qty,totalHPP:e.totalHPP,status:e.status}));
var stokSnap={};SIZES.forEach(s=>{stokSnap[s]={isi:(data.stock||{})[s]||0,kosong:getKosong(data,s),titip:getTitipTotal(data.titipList,s)};});
var katPenH={};penH.forEach(p=>{katPenH[p.kategori]=(katPenH[p.kategori]||0)+Number(p.nominal||0);});
var rec={
  id:editTB?editTB.id:uid(),
  tanggal:tgl,
  omzet:omzetH,hpp:hppH,labaKotor:marginH,totalOut:totalOutH,
  pemasukanLain:Number(pemasukanLain)||0,
  labaBersih:labaBersihEfektif,
  kurangSetorHari,
  cashIn:cashInH,tfIn:tfInH,bonIn:bonInH,
  cashLaci:Number(cashLaci)||0,rekBSI:Number(rekBSI)||0,rekBCA:Number(rekBCA)||0,totalPecah,pecah:{...pecah},
  piutangA,nilaiStokA,pinjamanA,
  nilaiDOGantung,nilaiDOSangkut,
  cashFlowOmset,assetTabungMilikPT,assetArmada,assetValue,
  cashFlowKemarin,deltaSelisih,
  lastEdited:editTB?new Date().toISOString():undefined,
  detail:{penjualan:penjualanDetail,pengeluaran:pengeluaranDetail,doMasuk:doDetail,stokSnapshot:stokSnap,katPengeluaran:katPenH,jumlahTransaksi:penjualanDetail.length,jumlahDO:doDetail.length}
};
if(editTB){
  // Replace record lama
  setData(d=>({...d,tutupBuku:(d.tutupBuku||[]).map(x=>x.id===editTB.id?rec:x)}));
  setEditTB(null);
  toast("✓ Tutup buku diperbarui! Data hari berikutnya otomatis sinkron.");
}else{
  // Hapus record lama dengan tanggal yang sama (hanya 1 per tanggal)
  setData(d=>({...d,tutupBuku:[rec,...(d.tutupBuku||[]).filter(x=>x.tanggal!==tgl)]}));
  toast("✓ Tutup buku harian tersimpan!");
}
}
function saveBulanan(){
var rec={id:uid(),bulan:bln,tanggal:bln+"-28",omzet:omzetB,labaKotor:marginB,totalOut:totalOutB,labaBersih:labaBersihB,assetValue,piutangA,nilaiStokA,pinjamanA,cashLaci:Number(cashLaci)||0,rekBSI:Number(rekBSI)||0,rekBCA:Number(rekBCA)||0};
setData(d=>({...d,tutupBuku:[rec,...(d.tutupBuku||[])]}));
toast("✓ Tutup buku bulanan tersimpan!");
}

function AssetSection(){return <Card style={{border:"1px solid "+C.olt,width:"fit-content",maxWidth:"100%",minWidth:660}}>
<div style={{fontWeight:700,color:C.olt,marginBottom:12,fontSize:13}}>🏦 Komponen Asset</div>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(150px,210px))",gap:10,marginBottom:10}}>
<Inp label="Cash di Laci" type="number" value={cashLaci} onChange={setCashLaci}/>
<Inp label="Rekening BSI" type="number" value={rekBSI} onChange={setRekBSI}/>
<Inp label="Rekening BCA" type="number" value={rekBCA} onChange={setRekBCA}/>
</div>
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
{[["Piutang Aktif (auto)",fR(piutangA),C.olt],["Nilai Stok Isi (auto)",fR(nilaiStokA),C.blt],["Pinjaman Karyawan (auto)",fR(pinjamanA),C.gl2],["Tabung Kosong",kosong55+"×5.5 / "+kosong12+"×12 / "+kosong50+"×50 (qty)",C.gl2]].map(x=><div key={x[0]} style={{background:C.nav,borderRadius:8,padding:"8px 10px",border:"1px solid "+C.bdr}}><div style={{fontSize:10,color:C.gl2}}>{x[0]}</div><div style={{fontSize:12,fontWeight:700,color:x[2]}}>{x[1]}</div></div>)}
</div>
<div style={{display:"flex",justifyContent:"space-between",padding:"12px 14px",background:C.nav,borderRadius:8,border:"2px solid "+C.olt}}>
<b style={{fontSize:14,color:C.wht}}>TOTAL ASSET VALUE</b>
<div style={{textAlign:"right"}}>
<b style={{fontSize:20,color:C.olt}}>{fR(assetValue)}</b>
{deltaAsset!==null&&<div style={{fontSize:11,color:deltaAsset>=0?C.glt:C.rlt,marginTop:2}}>{deltaAsset>=0?"▲":"▼"} {fR(Math.abs(deltaAsset))} vs sebelumnya</div>}
</div>
</div>
</Card>;}

return <div>
<STitle icon="📒" children="Tutup Buku"/>
<div style={{display:"flex",gap:6,marginBottom:14}}>{[["harian","📅 Harian"],["bulanan","📆 Bulanan"]].map(x=><button key={x[0]} onClick={()=>setTab(x[0])} style={{background:tab===x[0]?C.blu:C.nav,color:tab===x[0]?"white":C.wht,border:"1px solid "+(tab===x[0]?C.blt:C.bdr),borderRadius:8,padding:"8px 18px",fontWeight:700,fontSize:13,cursor:"pointer"}}>{x[1]}</button>)}</div>

{tab==="harian"&&<div id="_tb_hari">
<Card>
<div style={{display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
<Inp label="Tanggal" type="date" value={tgl} onChange={setTgl} style={{maxWidth:220,marginBottom:0}}/>
{editTB&&<div style={{background:"#78350F",border:"1px solid #F59E0B",borderRadius:8,padding:"8px 14px",fontSize:12,color:"#FCD34D"}}>
✏️ Mode Edit — {fDs(editTB.tanggal)} · <button onClick={()=>setEditTB(null)} style={{background:"none",border:"none",color:"#FCD34D",cursor:"pointer",fontSize:12,textDecoration:"underline"}}>Batalkan</button>
</div>}
</div>
{editTB&&<div style={{marginTop:8,padding:"8px 12px",background:"#7C2D12",borderRadius:6,fontSize:11,color:"#FCA5A5"}}>⚠️ Mengedit tutup buku akan mempengaruhi perhitungan hari berikutnya. Pastikan data sudah benar sebelum simpan.</div>}
</Card>

{/* 1. CASH WAJIB SETOR KASIR */}
{(()=>{
var allPenTgl=(data.pengeluaran||[]).filter(p=>p.tanggal===tgl);
var _penjTgl=(data.penjualan||[]).filter(p=>p.tanggal===tgl);
var cashPenjTgl=_penjTgl.filter(p=>(p.bayar||"").toLowerCase()==="cash").reduce((a,p)=>a+(p.total||0),0)+_penjTgl.filter(p=>(p.bayar||"").toLowerCase()==="split").reduce((a,p)=>a+Number((p.splitDetail||{}).cash||0),0);
var tfPenjTgl=_penjTgl.filter(p=>(p.bayar||"").toLowerCase()==="transfer"||(p.bayar||"").toLowerCase()==="tf").reduce((a,p)=>a+(p.total||0),0)+_penjTgl.filter(p=>(p.bayar||"").toLowerCase()==="split").reduce((a,p)=>a+Number((p.splitDetail||{}).tf||0),0);
var bonBayarCashTgl=(data.bon||[]).reduce((a,b)=>{var px=(b.pembayaran||[]).filter(p=>p.tanggal===tgl&&(p.metode||"cash").toLowerCase()==="cash");return a+px.reduce((s,p)=>s+Number(p.jumlah||p.nominal||0),0);},0);
var bonBayarTFTgl=(data.bon||[]).reduce((a,b)=>{var px=(b.pembayaran||[]).filter(p=>p.tanggal===tgl&&((p.metode||"").toLowerCase()==="transfer"||(p.metode||"").toLowerCase()==="tf"));return a+px.reduce((s,p)=>s+Number(p.jumlah||p.nominal||0),0);},0);
var penCashTgl=allPenTgl.filter(p=>(p.metode||"cash").toLowerCase()==="cash").reduce((a,p)=>a+Number(p.nominal||0),0);
var penTFTgl=allPenTgl.filter(p=>(p.metode||"").toLowerCase()==="transfer"||(p.metode||"").toLowerCase()==="tf").reduce((a,p)=>a+Number(p.nominal||0),0);
// Jualan Lain (gas kaleng, aksesoris dll)
var jualanLainTgl=(data.jualanLain||[]).filter(j=>j.tanggal===tgl);
var jualanLainCashTgl=jualanLainTgl.filter(j=>j.bayar==="cash").reduce((a,j)=>a+Number(j.total||0),0);
var jualanLainTFTgl=jualanLainTgl.filter(j=>j.bayar==="transfer").reduce((a,j)=>a+Number(j.total||0),0);
// Tarik TF (cash laci -> bank) & Setor TF (bank -> cash laci) hari ini
var tarikTFTgl=(data.kasBankTF||[]).filter(k=>k.tanggal===tgl&&k.jenis==="tarik").reduce((a,k)=>a+Number(k.nominal||0),0);
var setorTFTgl=(data.kasBankTF||[]).filter(k=>k.tanggal===tgl&&k.jenis==="setor").reduce((a,k)=>a+Number(k.nominal||0),0);
var totalCashMasuk=cashPenjTgl+bonBayarCashTgl+jualanLainCashTgl;
// Kurang setor sales hari ini = uang yang tidak masuk laci, sudah jadi pinjaman karyawan
var kurangSetorKasir=(data.ambilan||[]).filter(a=>a.tanggal===tgl&&(a.ket||"").toLowerCase().includes("kurang setor")).reduce((a,x)=>a+Number(x.nominal||0),0);
var wajibSetorKasir=Math.max(0,totalCashMasuk-penCashTgl-kurangSetorKasir-tarikTFTgl+setorTFTgl);
return <Card style={{border:"2px solid "+C.glt}}>
<div style={{fontWeight:700,color:C.glt,marginBottom:10,fontSize:13}}>🏦 Cash Wajib Setor Kasir — {fDs(tgl)}</div>
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
<div>
<div style={{fontSize:11,fontWeight:700,color:C.gl2,marginBottom:6,textTransform:"uppercase",letterSpacing:.5}}>Cash Masuk</div>
{[["Cash Penjualan (semua sales)",cashPenjTgl,C.glt],["Bayar BON Cash (semua)",bonBayarCashTgl,C.glt],...(jualanLainCashTgl>0?[["Jualan Lain (Cash)",jualanLainCashTgl,C.glt]]:[]),...(setorTFTgl>0?[["Setor TF (Bank → Laci)",setorTFTgl,C.glt]]:[])].map(x=><div key={x[0]} style={{display:"flex",justifyContent:"space-between",padding:"5px 8px",background:C.bg,borderRadius:5,marginBottom:3,border:"1px solid "+C.bdr}}>
<span style={{fontSize:11,color:C.gl2}}>{x[0]}</span><b style={{color:x[2],fontSize:12}}>{fR(x[1])}</b>
</div>)}
<div style={{display:"flex",justifyContent:"space-between",padding:"6px 8px",background:C.nav,borderRadius:5,marginBottom:8,border:"1px solid "+C.glt}}>
<b style={{fontSize:12,color:C.wht}}>Total Cash Masuk</b><b style={{color:C.glt,fontSize:13}}>{fR(totalCashMasuk)}</b>
</div>
<div style={{fontSize:11,fontWeight:700,color:C.gl2,marginBottom:6,textTransform:"uppercase",letterSpacing:.5}}>Pengeluaran Cash</div>
{allPenTgl.filter(p=>(p.metode||"cash").toLowerCase()==="cash").map((p,i)=><div key={i} style={{display:"flex",justifyContent:"space-between",padding:"4px 8px",background:C.bg,borderRadius:5,marginBottom:2,border:"1px solid "+C.bdr}}>
<span style={{fontSize:10,color:C.gl2}}>{p.kategori}{p.ket?" — "+p.ket:""} ({p.karyawanNama||"—"})</span><b style={{color:C.rlt,fontSize:11}}>{fR(p.nominal)}</b>
</div>)}
<div style={{display:"flex",justifyContent:"space-between",padding:"6px 8px",background:C.nav,borderRadius:5,border:"1px solid "+C.rlt}}>
<b style={{fontSize:12,color:C.wht}}>Total Pengeluaran Cash</b><b style={{color:C.rlt,fontSize:13}}>{fR(penCashTgl)}</b>
</div>
</div>
<div>
<div style={{background:C.nav,borderRadius:8,padding:14,border:"2px solid "+C.glt,marginBottom:8}}>
<div style={{fontSize:11,color:C.gl2,marginBottom:4}}>Cash Masuk</div>
<div style={{fontSize:15,fontWeight:900,color:C.glt,marginBottom:8}}>{fR(totalCashMasuk)}</div>
<div style={{fontSize:11,color:C.gl2,marginBottom:4}}>Pengeluaran Cash</div>
<div style={{fontSize:15,fontWeight:900,color:C.rlt,marginBottom:8}}>- {fR(penCashTgl)}</div>
{kurangSetorKasir>0&&<><div style={{fontSize:11,color:C.gl2,marginBottom:4}}>Kurang Setor Sales</div>
<div style={{fontSize:15,fontWeight:900,color:C.rlt,marginBottom:8}}>- {fR(kurangSetorKasir)}</div></>}
{tarikTFTgl>0&&<><div style={{fontSize:11,color:C.gl2,marginBottom:4}}>Tarik TF (Laci → Bank)</div>
<div style={{fontSize:15,fontWeight:900,color:C.rlt,marginBottom:8}}>- {fR(tarikTFTgl)}</div></>}
<div style={{height:1,background:C.bdr,marginBottom:8}}/>
<div style={{fontSize:11,color:C.glt,fontWeight:700,marginBottom:4}}>WAJIB SETOR KE BANK</div>
<div style={{fontSize:22,fontWeight:900,color:C.wht,marginBottom:10}}>{fR(wajibSetorKasir)}</div>
<div style={{height:1,background:C.bdr,marginBottom:10}}/>
<div style={{fontSize:11,color:C.gl2,marginBottom:4}}>Total Tunai Fisik (pecahan)</div>
<div style={{fontSize:15,fontWeight:700,color:totalPecah>=wajibSetorKasir?C.glt:C.olt,marginBottom:6}}>{fR(totalPecah)}</div>
<div style={{fontSize:11,color:C.gl2,marginBottom:4}}>Selisih Cash</div>
<div style={{fontSize:18,fontWeight:900,color:Math.abs(totalPecah-wajibSetorKasir)<1000?C.glt:C.rlt}}>{fR(totalPecah-wajibSetorKasir)}</div>
</div>
<div style={{background:C.nav,borderRadius:8,padding:12,border:"1px solid "+C.bdr}}>
<div style={{fontSize:10,color:C.gl2,marginBottom:6,fontWeight:700}}>INFO TRANSFER (masuk rekening langsung)</div>
{[["TF Penjualan",tfPenjTgl],["Bayar BON TF",bonBayarTFTgl],["Pengeluaran TF",penTFTgl]].map(x=><div key={x[0]} style={{display:"flex",justifyContent:"space-between",marginBottom:3}}>
<span style={{fontSize:10,color:C.gl2}}>{x[0]}</span><span style={{fontSize:11,color:"#9CA3AF",fontWeight:600}}>{fR(x[1])}</span>
</div>)}
</div>
</div>
</div>
</Card>;
})()}

{/* 2. INPUT CASH FISIK (di dalam CASH FLOW) */}
<Card style={{border:"1px solid "+C.blt,width:"fit-content",maxWidth:"100%",minWidth:660}}>
<div style={{fontSize:11,fontWeight:700,color:C.gl2,marginBottom:6}}>📥 Input Cash Fisik:</div>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(150px,210px))",gap:8,marginBottom:10}}>
<Inp label="Cash di Laci (Rp)" type="number" value={cashLaci} onChange={setCashLaci}/>
<Inp label="Saldo Bank BSI (Rp)" type="number" value={rekBSI} onChange={setRekBSI}/>
<Inp label="Saldo Bank BCA (Rp)" type="number" value={rekBCA} onChange={setRekBCA}/>
</div>
<div style={{fontSize:11,fontWeight:700,color:C.gl2,marginBottom:6}}>🪙 Pecahan Kas Fisik (opsional):</div>
<div style={{border:"1px solid "+C.bdr,borderRadius:8,overflow:"hidden",marginBottom:10}}>
{DENOMS.map(d=><div key={d} style={{display:"grid",gridTemplateColumns:"1fr 85px 105px",padding:"5px 11px",borderTop:"1px solid "+C.bdr,alignItems:"center"}}><span style={{color:C.wht,fontSize:13,fontWeight:600}}>{fR(d)}</span><input type="number" value={pecah[d]} placeholder="0" onChange={e=>setPecah(u=>({...u,[d]:e.target.value}))} style={{background:C.nav,border:"1px solid "+C.bdr,borderRadius:6,padding:"4px 7px",color:C.wht,fontSize:12,outline:"none",width:74}}/><span style={{color:Number(pecah[d]||0)>0?C.olt:C.gl2,fontWeight:700,fontSize:12}}>{Number(pecah[d]||0)>0?fR(Number(pecah[d])*d):"-"}</span></div>)}
<div style={{display:"grid",gridTemplateColumns:"1fr 85px 105px",padding:"9px 11px",background:C.nav,borderTop:"2px solid "+C.bdr}}><b style={{color:C.wht}}>Total Tunai</b><span/><b style={{color:C.glt}}>{fR(totalPecah)}</b></div>
</div>
</Card>

<div style={{display:"flex",gap:14,flexWrap:"wrap",alignItems:"flex-start"}}>
<div style={{flex:"0 1 420px",minWidth:380}}>
{/* 3. P&L HARI INI */}
<Card style={{border:"1px solid "+C.glt,width:"100%"}}>
<div style={{fontWeight:700,color:C.glt,marginBottom:12,fontSize:13}}>📊 P&L Hari Ini</div>
<div style={{border:"1px solid "+C.bdr,borderRadius:8,overflow:"hidden",marginBottom:10}}>
{[["Omzet",omzetH,C.wht,false],["HPP / Modal",hppH,C.gl2,false],["Laba Kotor",marginH,C.blt,true],["Pengeluaran Operasional",-totalOutH,C.rlt,false],["LABA BERSIH EFEKTIF (Operasional)",labaBersihEfektif,labaBersihEfektif>=0?C.glt:C.rlt,true],...(omzetJualLainH>0?[["Cash/TF Jualan Lain (info kas, non-laba)",omzetJualLainH,C.olt,false]]:[]),...(Number(pemasukanLain)>0?[["Pemasukan Lainnya (topup, non-laba)",Number(pemasukanLain),C.olt,false]]:[])].map((x,i)=><div key={i} style={{display:"flex",justifyContent:"space-between",gap:10,padding:x[3]?"10px 14px":"7px 14px",background:x[3]?C.nav:"transparent",borderBottom:"1px solid "+C.bdr}}><span style={{fontSize:x[3]?13:12,color:x[3]?C.wht:C.gl2,fontWeight:x[3]?700:400}}>{x[0]}</span><span style={{fontSize:x[3]?15:13,fontWeight:x[3]?900:600,color:x[2],whiteSpace:"nowrap"}}>{fR(x[1])}</span></div>)}
</div>
<Inp label="Pemasukan Lainnya (topup saham dll)" type="number" value={pemasukanLain} onChange={setPemasukanLain} placeholder="0"/>
<div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:8,marginTop:10}}>
{[["Cash",cashInH,C.glt],["Transfer",tfInH,C.blt],["BON",bonInH,C.olt]].map(x=><div key={x[0]} style={{background:C.nav,borderRadius:8,padding:"8px 10px",textAlign:"center",border:"1px solid "+C.bdr}}><div style={{fontSize:10,color:C.gl2}}>{x[0]}</div><div style={{fontSize:13,fontWeight:900,color:x[2]}}>{fR(x[1])}</div></div>)}
</div>
</Card>
</div>
<div style={{flex:"0 1 380px",minWidth:340}}>
{/* 4. CASH FLOW / OMSET */}
<Card style={{border:"1px solid "+C.blt,width:"100%"}}>
<div style={{fontWeight:700,color:C.blt,marginBottom:12,fontSize:13}}>💰 CASH FLOW / OMSET — {fDs(tgl)}</div>
<div style={{border:"1px solid "+C.bdr,borderRadius:8,overflow:"hidden",marginBottom:10}}>
{[
["Cash di Laci",Number(cashLaci)||0,C.wht,false],
["Bank BSI",Number(rekBSI)||0,"#9CA3AF",false],
["Bank BCA",Number(rekBCA)||0,"#9CA3AF",false],
["Total Cash + Bank",cashTotal,C.wht,true],
["",null,null,false],
["Stok Isi Gudang (HPP)",nilaiStokA,C.gl2,false],
["DO Gantung",nilaiDOGantung,nilaiDOGantung>0?"#F59E0B":C.gl2,false],
["BON Konsumen",piutangA,C.gl2,false],
["Pinjaman Karyawan",pinjamanA,C.gl2,false],
["PIUTANG & MODAL BERJALAN",piutangModal,C.wht,true],
["",null,null,false],
["TOTAL CASH FLOW",cashFlowOmset,cashFlowOmset>=0?C.glt:C.rlt,true],
].map((x,i)=>x[1]===null&&x[0]===""?<div key={i} style={{height:8,background:"transparent"}}/> :<div key={i} style={{display:"flex",justifyContent:"space-between",gap:10,padding:x[3]?"10px 14px":"7px 14px",background:x[3]?C.nav:"transparent",borderBottom:"1px solid "+C.bdr}}>
<span style={{fontSize:x[3]?13:12,color:x[3]?C.wht:C.gl2,fontWeight:x[3]?700:400,whiteSpace:"nowrap"}}>{x[0]}</span>
<span style={{fontSize:x[3]?15:13,fontWeight:x[3]?900:600,color:x[2],whiteSpace:"nowrap"}}>{fR(x[1]||0)}</span>
</div>)}
</div>
</Card>
</div>
</div>

<div style={{display:"flex",gap:14,flexWrap:"wrap",alignItems:"flex-start"}}>
<div style={{flex:"1 1 360px",minWidth:300}}>
{/* 5. VERIFIKASI CASH FLOW */}
<Card style={{border:"1px solid "+(Math.abs(deltaSelisih)<1000?C.glt:C.olt)}}>
<div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>✅ Verifikasi Cash Flow</div>
<div style={{border:"1px solid "+C.bdr,borderRadius:8,overflow:"hidden"}}>
{[[(tglTBKemarin?"Cash Flow "+fDs(tglTBKemarin):"Belum ada TB sebelumnya"),cashFlowKemarin,C.gl2,false],["Laba Bersih Efektif",labaBersihEfektif,C.glt,false],...(Number(pemasukanLain)>0?[["  ↳ Pemasukan Lainnya (topup)",Number(pemasukanLain),C.olt,false]]:[]),...(kurangSetorHari>0?[["  ↳ Kurang Setor (sudah masuk Piutang Karyawan)",kurangSetorHari,C.olt,false]]:[]),["Total Cash Flow Hari Ini",cashFlowOmset,C.blt,true],["SELISIH",deltaSelisih,Math.abs(deltaSelisih)<1000?C.glt:C.rlt,true]].map((x,i)=><div key={i} style={{display:"flex",justifyContent:"space-between",padding:x[3]?"10px 14px":"7px 14px",background:x[3]?C.nav:"transparent",borderBottom:"1px solid "+C.bdr}}><span style={{fontSize:x[3]?13:12,color:x[3]?C.wht:C.gl2,fontWeight:x[3]?700:400}}>{x[0]}</span><span style={{fontSize:x[3]?15:13,fontWeight:x[3]?900:600,color:x[2]}}>{fR(x[1])}</span></div>)}
</div>
{Math.abs(deltaSelisih)<1000&&<div style={{marginTop:8,padding:"6px 12px",background:C.grn,borderRadius:6,fontSize:12,fontWeight:700,color:"white"}}>✅ Selisih = 0. Cash flow balance!</div>}
{Math.abs(deltaSelisih)>=1000&&<div style={{marginTop:8,padding:"6px 12px",background:C.rdk,borderRadius:6,fontSize:12,color:"white"}}>⚠️ Ada selisih {fR(Math.abs(deltaSelisih))}. Periksa input cash atau ada transaksi yang terlewat.</div>}
</Card>
</div>
<div style={{flex:"1 1 360px",minWidth:300}}>
{/* 6. TOTAL ASSET */}
<Card style={{border:"1px solid "+C.olt}}>
<div style={{fontWeight:700,color:C.olt,marginBottom:12,fontSize:13}}>🏦 TOTAL ASSET</div>
<div style={{border:"1px solid "+C.bdr,borderRadius:8,overflow:"hidden",marginBottom:10}}>
{[["Total Cash Flow",cashFlowOmset,C.blt,false],["Asset Tabung Milik PT",assetTabungMilikPT,C.gl2,false],["Asset Armada",assetArmada,C.gl2,false],["TOTAL ASSET",assetValue,C.olt,true]].map((x,i)=><div key={i} style={{display:"flex",justifyContent:"space-between",padding:x[3]?"12px 14px":"8px 14px",background:x[3]?C.nav:"transparent",borderBottom:"1px solid "+C.bdr}}><span style={{fontSize:x[3]?14:12,color:x[3]?C.wht:C.gl2,fontWeight:x[3]?800:400}}>{x[0]}</span><span style={{fontSize:x[3]?18:13,fontWeight:x[3]?900:600,color:x[2]}}>{fR(x[1])}</span></div>)}
</div>
</Card>
</div>
</div>

{/* 7. LAPORAN STOK HARIAN */}
<Card>
<div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>📊 Laporan Stok Harian</div>
{(()=>{
var bulanTgl=tgl.slice(0,7);
var rowsTB=buildStokHarian(data,bulanTgl).filter(r=>r.tgl<=tgl).slice(-3);
var uk=["5.5 kg","12 kg","50 kg"];var ukL=["5,5kg","12kg","50kg"];
return <div style={{overflowX:"auto"}}>
<table style={{borderCollapse:"collapse",width:"100%",fontSize:10}}>
<thead><tr style={{background:C.nav}}>
<th style={{padding:"5px 8px",color:C.gl2,textAlign:"left",borderBottom:"1px solid "+C.bdr}}>Hari/Tgl</th>
{uk.map((s,i)=><th key={s} colSpan={4} style={{padding:"5px 6px",color:"white",textAlign:"center",borderBottom:"1px solid "+C.bdr,borderLeft:"2px solid "+C.bdr,background:"#1E3A5F",fontSize:9}}>{ukL[i]}: isi | TK | Titip | Total</th>)}
</tr></thead>
<tbody>
{rowsTB.map((r,i)=><tr key={r.tgl} style={{background:r.tgl===tgl?C.nav:C.bg,borderBottom:"1px solid "+C.bdr}}>
<td style={{padding:"4px 8px"}}><div style={{fontWeight:700,color:r.tgl===tgl?C.blt:C.gl2,fontSize:11}}>{r.dayName}</div><div style={{fontSize:9,color:C.gl2}}>{fDs(r.tgl)}</div></td>
{uk.map(s=>[<td key={"i"+s} style={{padding:"4px 6px",textAlign:"center",color:C.glt,fontWeight:700,borderLeft:"2px solid "+C.bdr}}>{r.akhirIsi[s]}</td>,<td key={"k"+s} style={{padding:"4px 6px",textAlign:"center",color:C.gl2}}>{r.akhirTK[s]}</td>,<td key={"t"+s} style={{padding:"4px 6px",textAlign:"center",color:C.blt}}>{r.titipSnap[s]}</td>,<td key={"o"+s} style={{padding:"4px 6px",textAlign:"center",color:C.olt,fontWeight:700}}>{r.total[s]}</td>])}
</tr>)}
</tbody>
</table>
</div>;
})()}
</Card>

{/* 8. REKAP TABUNG */}
<Card>
<div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>📦 Rekap Tabung</div>
{(()=>{
var rowsTBRekap=buildStokHarian(data,tgl.slice(0,7)).filter(r=>r.tgl<=tgl);
var tbRow=rowsTBRekap.length>0?rowsTBRekap[rowsTBRekap.length-1]:null;
var titipLuarBalTB={};
(data.titipList||[]).filter(t=>t.tipe==="titip_luar"||t.tipe==="tarik_luar").forEach(t=>{var m=t.tipe==="titip_luar"?1:-1;(t.items||[]).forEach(it=>{titipLuarBalTB[it.ukuran]=(titipLuarBalTB[it.ukuran]||0)+m*Number(it.qty||0);});});
var rows=[
["Di Gudang (Isi)",SIZES.map(s=>tbRow?tbRow.akhirIsi[s]:((data.stock||{})[s]||0)),C.glt,false],
["Kosong di Gudang",SIZES.map(s=>tbRow?tbRow.akhirTK[s]:getKosong(data,s)),C.gl2,false],
["Titip ke Konsumen",SIZES.map(s=>tbRow?tbRow.titipSnap[s]:getTitipTotal(data.titipList,s)),C.blt,false],
["Total Keseluruhan",SIZES.map(s=>tbRow?tbRow.total[s]:0),C.olt,true],
["Titipan Pihak Lain di PT",SIZES.map(s=>Math.max(0,titipLuarBalTB[s]||0)),"#6B7280",false],
["MILIK PT HOE TRANGSA",SIZES.map(s=>Math.max(0,(tbRow?tbRow.total[s]:0)-(titipLuarBalTB[s]||0))),C.wht,true],
];
return <div style={{overflowX:"auto"}}>
<table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
<thead><tr style={{background:C.nav}}>
{["Keterangan","5,5 kg","12 kg","50 kg"].map(h=><th key={h} style={{padding:"8px 10px",color:C.gl2,fontWeight:700,textAlign:h==="Keterangan"?"left":"center",borderBottom:"2px solid "+C.bdr}}>{h}</th>)}
</tr></thead>
<tbody>
{rows.map((r,i)=><tr key={i} style={{borderBottom:"1px solid "+C.bdr,background:r[3]?C.nav:"transparent"}}>
<td style={{padding:"7px 10px",color:r[3]?C.wht:C.gl2,fontWeight:r[3]?700:400}}>{r[0]}</td>
{r[1].map((v,j)=><td key={j} style={{padding:"7px 10px",textAlign:"center",color:r[2],fontWeight:r[3]?800:600,fontSize:r[3]?14:12}}>{v}</td>)}
</tr>)}
</tbody>
</table>
</div>;
})()}
</Card>

<div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
<Btn color={editTB?"orange":"green"} onClick={saveHarian}>{editTB?"💾 Simpan Perubahan":"💾 Simpan Tutup Buku"}</Btn>
<button onClick={()=>doPrint("_tb_hari")} style={{background:C.blu,color:"#fff",border:"none",padding:"9px 16px",borderRadius:8,fontSize:13,cursor:"pointer",fontWeight:700}}>🖨️ Cetak / PDF</button>
<button onClick={()=>doDownloadPNG("_tb_hari",makeFileName("tb",tgl,"","png"))} style={{background:"#1D6A96",color:"#fff",border:"none",padding:"9px 14px",borderRadius:8,fontSize:13,cursor:"pointer",fontWeight:700}}>💾 PNG</button>
<button onClick={()=>doCopyPNG("_tb_hari")} style={{background:"#145A32",color:"#fff",border:"none",padding:"9px 14px",borderRadius:8,fontSize:13,cursor:"pointer",fontWeight:700}}>📋 Copy</button>
<span style={{fontSize:10,color:"#888",fontStyle:"italic",alignSelf:"center"}}>💡 PDF: <b>{makeFileName("tb",tgl,"","pdf")}</b></span>
</div>
</div>}

{tab==="bulanan"&&<div id="_tb_bln">
<Card><MonthPicker label="Pilih Bulan" value={bln} onChange={setBln}/></Card>
<Card style={{border:"1px solid "+C.blt}}>
<div style={{fontWeight:700,color:C.blt,marginBottom:12,fontSize:13}}>📊 P&L Bulanan — {BULAN_ID[Number(bln.split("-")[1])-1]} {bln.split("-")[0]}</div>
{/* Summary cards */}
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(140px,1fr))",gap:8,marginBottom:14}}>
{[["Total Omzet",omzetB,C.wht],["HPP/Modal (DO)",hppBln,C.olt],["Margin Kotor",marginB,C.blt],["Pengeluaran Ops",totalOutB,C.rlt],["Laba Bersih",labaBersihB,labaBersihB>=0?C.glt:C.rlt]].map(x=><div key={x[0]} style={{background:C.nav,borderRadius:8,padding:"8px 10px",border:"1px solid "+C.bdr}}><div style={{fontSize:10,color:C.gl2}}>{x[0]}</div><div style={{fontSize:13,fontWeight:900,color:x[2]}}>{fR(x[1])}</div></div>)}
</div>
{/* Grafik 1: Omzet & HPP per hari */}
{chartDataB.some(d=>d.omzet>0)&&<>
<div style={{fontSize:11,fontWeight:700,color:C.gl2,marginBottom:6}}>📈 Omzet & Modal HPP per Hari</div>
<ResponsiveContainer width="100%" height={180}><AreaChart data={chartDataB} margin={{top:4,right:8,bottom:0,left:8}}>
<CartesianGrid strokeDasharray="3 3" stroke={C.bdr}/>
<XAxis dataKey="hari" stroke={C.gl2} fontSize={9}/>
<YAxis stroke={C.gl2} fontSize={9} tickFormatter={v=>(v/1e6).toFixed(1)+"jt"}/>
<Tooltip contentStyle={{background:C.card,border:"1px solid "+C.bdr,color:C.wht,fontSize:11}} formatter={v=>fR(v)}/>
<Area type="monotone" dataKey="omzet" stroke={C.blt} fill={C.blt} fillOpacity={0.25} name="Omzet"/>
<Area type="monotone" dataKey="hpp" stroke={C.olt} fill={C.olt} fillOpacity={0.2} name="HPP/Modal"/>
</AreaChart></ResponsiveContainer>
{/* Grafik 2: Margin & Pengeluaran */}
<div style={{fontSize:11,fontWeight:700,color:C.gl2,marginBottom:6,marginTop:14}}>📊 Margin Kotor & Pengeluaran per Hari</div>
<ResponsiveContainer width="100%" height={160}><BarChart data={chartDataB} margin={{top:4,right:8,bottom:0,left:8}}>
<CartesianGrid strokeDasharray="3 3" stroke={C.bdr}/>
<XAxis dataKey="hari" stroke={C.gl2} fontSize={9}/>
<YAxis stroke={C.gl2} fontSize={9} tickFormatter={v=>(v/1e6).toFixed(1)+"jt"}/>
<Tooltip contentStyle={{background:C.card,border:"1px solid "+C.bdr,color:C.wht,fontSize:11}} formatter={v=>fR(v)}/>
<Bar dataKey="marginKotor" fill={C.blt} name="Margin Kotor" opacity={0.8}/>
<Bar dataKey="pengeluaran" fill={C.rlt} name="Pengeluaran" opacity={0.8}/>
</BarChart></ResponsiveContainer>
{/* Grafik 3: Cash Flow Kumulatif */}
<div style={{fontSize:11,fontWeight:700,color:C.gl2,marginBottom:6,marginTop:14}}>💰 Cash Flow Kumulatif (Laba Bersih per Hari)</div>
<ResponsiveContainer width="100%" height={160}><AreaChart data={chartDataB} margin={{top:4,right:8,bottom:0,left:8}}>
<CartesianGrid strokeDasharray="3 3" stroke={C.bdr}/>
<XAxis dataKey="hari" stroke={C.gl2} fontSize={9}/>
<YAxis stroke={C.gl2} fontSize={9} tickFormatter={v=>(v/1e6).toFixed(1)+"jt"}/>
<Tooltip contentStyle={{background:C.card,border:"1px solid "+C.bdr,color:C.wht,fontSize:11}} formatter={v=>fR(v)}/>
<Area type="monotone" dataKey="cashFlow" stroke={C.glt} fill={C.glt} fillOpacity={0.3} name="Cash Flow Kumulatif"/>
<Area type="monotone" dataKey="labaBersih" stroke="#F59E0B" fill="#F59E0B" fillOpacity={0.15} name="Laba Bersih Harian"/>
</AreaChart></ResponsiveContainer>
</>}
</Card>
{top5.some(x=>x.omzet>0)&&<Card><div style={{fontWeight:700,color:C.gl2,marginBottom:8,fontSize:13}}>🏆 Top 5 Hari (Omzet)</div><RTbl headers={["Hari","Omzet","Laba Bersih"]} rows={top5.map(x=>[bln+"-"+x.hari,<b style={{color:C.blt}}>{fR(x.omzet)}</b>,<b style={{color:x.labaBersih>=0?C.glt:C.rlt}}>{fR(x.labaBersih)}</b>])}/></Card>}
{katPenArr.length>0&&<Card><div style={{fontWeight:700,color:C.gl2,marginBottom:8,fontSize:13}}>💸 Pengeluaran per Kategori</div><RTbl headers={["Kategori","Total","% dari Pengeluaran"]} rows={katPenArr.map(([k,v])=>[k,<b style={{color:C.rlt}}>{fR(v)}</b>,(totalOutB>0?(v/totalOutB*100).toFixed(1):0)+"%"])}/></Card>}
<Card style={{width:"fit-content",maxWidth:"100%",minWidth:660}}>
<div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>💰 Input Cash untuk Simpan Bulanan</div>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(150px,210px))",gap:8,marginBottom:10}}>
<Inp label="Cash di Laci (Rp)" type="number" value={cashLaci} onChange={setCashLaci}/>
<Inp label="Saldo Bank BSI (Rp)" type="number" value={rekBSI} onChange={setRekBSI}/>
<Inp label="Saldo Bank BCA (Rp)" type="number" value={rekBCA} onChange={setRekBCA}/>
</div>
</Card>
<div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
<Btn color="green" onClick={saveBulanan}>💾 Simpan</Btn>
<button onClick={()=>doPrint("_tb_bln")} style={{background:C.blu,color:"#fff",border:"none",padding:"9px 16px",borderRadius:8,fontSize:13,cursor:"pointer",fontWeight:700}}>🖨️ Cetak</button>
<span style={{fontSize:11,color:"#888",fontStyle:"italic",alignSelf:"center"}}>Pilih "Save as PDF" di dialog cetak</span>
<button onClick={()=>{var wb=XLSX.utils.book_new();var s=[["Bulan",bln],["Omzet",omzetB],["Laba Kotor",marginB],["Total Pengeluaran",totalOutB],["Laba Bersih",labaBersihB],["Asset Value",assetValue]];XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(s),"Tutup Buku");XLSX.writeFile(wb,"TutupBuku_"+bln+".xlsx");toast("✓ Excel didownload!");}} style={{background:"#15803D",color:"#fff",border:"none",padding:"9px 16px",borderRadius:8,fontSize:13,cursor:"pointer",fontWeight:700}}>📥 Excel</button>
</div>
</div>}

{(data.tutupBuku||[]).length>0&&<Card style={{marginTop:16}}>
<div style={{fontWeight:700,color:C.gl2,marginBottom:8,fontSize:13}}>📜 Riwayat Tutup Buku</div>
<RTbl headers={["Tanggal/Bulan","Transaksi","Omzet","Laba Bersih","Asset","Aksi"]} widths={[150,150,150,150,150,260]} rows={(data.tutupBuku||[]).slice(0,30).map(r=>[
  <div><div style={{fontWeight:700,color:C.wht}}>{r.bulan?BULAN_ID[Number(r.bulan.split("-")[1])-1]+" "+r.bulan.split("-")[0]:fDs(r.tanggal)}</div><div style={{fontSize:11,color:C.gl2}}>{r.bulan?"Bulanan":"Harian"}</div></div>,
  <span style={{color:C.blt,fontWeight:700}}>{r.detail?.jumlahTransaksi||"-"} inv{r.detail?.jumlahDO?", "+r.detail.jumlahDO+" DO":""}</span>,
  <span style={{whiteSpace:"nowrap"}}>{fR(r.omzet||0)}</span>,
  <b style={{color:(r.labaBersih||0)>=0?C.glt:C.rlt,whiteSpace:"nowrap"}}>{fR(r.labaBersih||0)}</b>,
  <span style={{whiteSpace:"nowrap"}}>{fR(r.assetValue||0)}</span>,
  <div style={{display:"flex",gap:5}}>
<button onClick={()=>setViewTB(r)} style={{background:C.inHv,border:"1px solid "+C.blt,borderRadius:7,padding:"6px 9px",color:C.blt,cursor:"pointer",fontSize:12,fontWeight:700}}>👁️ Lihat</button>
<button onClick={()=>{setEditTB(r);setTab("harian");window.scrollTo(0,0);}} style={{background:"#78350F",border:"1px solid #F59E0B",borderRadius:7,padding:"6px 9px",color:"#FCD34D",cursor:"pointer",fontSize:12,fontWeight:700}}>✏️ Edit</button>
<button onClick={()=>{setViewTB(r);setTimeout(()=>doPrint("_tb_view"),300);}} style={{background:C.nav,border:"1px solid "+C.bdr,borderRadius:7,padding:"6px 9px",color:C.gl2,cursor:"pointer",fontSize:12,fontWeight:700}}>🖨️ Cetak</button>
</div>
])}/>
</Card>}

{/* Modal Detail Tutup Buku */}
{viewTB&&<div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.7)",zIndex:9000,overflowY:"auto",padding:16}}>
<div id="_tb_view" style={{maxWidth:680,margin:"0 auto",background:C.card,borderRadius:12,border:"1px solid "+C.bdr,padding:20}}>
<div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
<div><div style={{fontSize:16,fontWeight:800,color:C.wht}}>📒 Laporan Tutup Buku</div><div style={{fontSize:12,color:C.gl2}}>{viewTB.bulan?BULAN_ID[Number(viewTB.bulan.split("-")[1])-1]+" "+viewTB.bulan.split("-")[0]:fDs(viewTB.tanggal)}</div></div>
<button onClick={()=>setViewTB(null)} style={{background:C.nav,border:"1px solid "+C.bdr,borderRadius:8,padding:"6px 14px",color:C.wht,cursor:"pointer",fontWeight:700}}>✕ Tutup</button>
</div>

{/* Ringkasan P&L */}

{/* SECTION 1: CASH FLOW */}
<div style={{background:C.nav,borderRadius:8,border:"1px solid "+C.bdr,overflow:"hidden",marginBottom:12}}>
<div style={{padding:"8px 14px",borderBottom:"1px solid "+C.bdr,fontWeight:700,color:C.blt,fontSize:12}}>💰 CASH FLOW / OMSET</div>
{[["Cash di Laci",viewTB.cashLaci,C.wht],["Bank BSI",viewTB.rekBSI,C.gl2],["Bank BCA",viewTB.rekBCA,C.gl2],["Total Cash + Bank",(viewTB.cashLaci||0)+(viewTB.rekBSI||0)+(viewTB.rekBCA||0),C.wht,true],["Stok Isi Gudang (HPP)",viewTB.nilaiStokA,C.gl2],["DO Gantung",viewTB.nilaiDOGantung||0,"#F59E0B"],["BON Konsumen",viewTB.piutangA,C.gl2],["Pinjaman Karyawan",viewTB.pinjamanA,C.gl2],["PIUTANG & MODAL BERJALAN",(viewTB.nilaiStokA||0)+(viewTB.nilaiDOGantung||0)+(viewTB.piutangA||0)+(viewTB.pinjamanA||0),C.wht,true],["TOTAL CASH FLOW",viewTB.cashFlowOmset,C.glt,true]].map((x,i)=><div key={i} style={{display:"flex",justifyContent:"space-between",padding:x[3]?"9px 14px":"6px 14px",background:x[3]?C.card:"transparent",borderBottom:"1px solid "+C.bdr}}><span style={{color:x[3]?C.wht:C.gl2,fontWeight:x[3]?700:400,fontSize:12}}>{x[0]}</span><span style={{color:x[2],fontWeight:x[3]?800:600,fontSize:x[3]?14:12}}>{fR(x[1]||0)}</span></div>)}
</div>

{/* SECTION 2: P&L */}
<div style={{background:C.nav,borderRadius:8,border:"1px solid "+C.bdr,overflow:"hidden",marginBottom:12}}>
<div style={{padding:"8px 14px",borderBottom:"1px solid "+C.bdr,fontWeight:700,color:C.glt,fontSize:12}}>📊 P&L — {fDs(viewTB.tanggal)}</div>
{[["Omzet",viewTB.omzet,C.wht],["HPP / Modal",viewTB.hpp,C.gl2],["Laba Kotor",viewTB.labaKotor,C.blt,true],["Pengeluaran Ops",viewTB.totalOut,C.rlt],["Pemasukan Lainnya",viewTB.pemasukanLain||0,"#F59E0B"],["LABA BERSIH",viewTB.labaBersih,viewTB.labaBersih>=0?C.glt:C.rlt,true]].map((x,i)=><div key={i} style={{display:"flex",justifyContent:"space-between",padding:x[3]?"9px 14px":"6px 14px",background:x[3]?C.card:"transparent",borderBottom:"1px solid "+C.bdr}}><span style={{color:x[3]?C.wht:C.gl2,fontWeight:x[3]?700:400,fontSize:12}}>{x[0]}</span><span style={{color:x[2],fontWeight:x[3]?800:600,fontSize:x[3]?14:12}}>{fR(x[1]||0)}</span></div>)}
</div>

{/* SECTION 3: VERIFIKASI */}
<div style={{background:C.nav,borderRadius:8,border:"1px solid "+C.bdr,overflow:"hidden",marginBottom:12}}>
<div style={{padding:"8px 14px",borderBottom:"1px solid "+C.bdr,fontWeight:700,color:C.blt,fontSize:12}}>✅ Verifikasi Cash Flow</div>
{[["Cash Flow Kemarin",viewTB.cashFlowKemarin||0,C.gl2],["Laba Hari Ini",viewTB.labaBersih,C.glt],["Cash Flow Hari Ini",viewTB.cashFlowOmset,C.blt,true],["Selisih",viewTB.deltaSelisih||0,Math.abs(viewTB.deltaSelisih||0)<1000?C.glt:C.rlt,true]].map((x,i)=><div key={i} style={{display:"flex",justifyContent:"space-between",padding:x[3]?"9px 14px":"6px 14px",background:x[3]?C.card:"transparent",borderBottom:"1px solid "+C.bdr}}><span style={{color:x[3]?C.wht:C.gl2,fontWeight:x[3]?700:400,fontSize:12}}>{x[0]}</span><span style={{color:x[2],fontWeight:x[3]?800:600,fontSize:x[3]?14:12}}>{fR(x[1]||0)}</span></div>)}
</div>

{/* SECTION 4: TOTAL ASSET */}
<div style={{background:C.nav,borderRadius:8,border:"1px solid "+C.bdr,overflow:"hidden",marginBottom:12}}>
<div style={{padding:"8px 14px",borderBottom:"1px solid "+C.bdr,fontWeight:700,color:C.olt,fontSize:12}}>🏦 TOTAL ASSET</div>
{[["Total Cash Flow",viewTB.cashFlowOmset,C.blt],["Asset Tabung Milik PT",viewTB.assetTabungMilikPT||0,C.gl2],["Asset Armada",viewTB.assetArmada||0,C.gl2],["TOTAL ASSET",viewTB.assetValue,C.olt,true]].map((x,i)=><div key={i} style={{display:"flex",justifyContent:"space-between",padding:x[3]?"9px 14px":"6px 14px",background:x[3]?C.card:"transparent",borderBottom:"1px solid "+C.bdr}}><span style={{color:x[3]?C.wht:C.gl2,fontWeight:x[3]?700:400,fontSize:12}}>{x[0]}</span><span style={{color:x[2],fontWeight:x[3]?800:600,fontSize:x[3]?14:12}}>{fR(x[1]||0)}</span></div>)}
</div>

{/* SECTION 5: REKAP TABUNG */}
{viewTB.detail?.stokSnapshot&&<div style={{background:C.nav,borderRadius:8,border:"1px solid "+C.bdr,overflow:"hidden",marginBottom:12}}>
<div style={{padding:"8px 14px",borderBottom:"1px solid "+C.bdr,fontWeight:700,color:C.gl2,fontSize:12}}>📦 Rekap Tabung</div>
<table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
<thead><tr style={{background:C.card}}>{["Keterangan","5,5 kg","12 kg","50 kg"].map(h=><th key={h} style={{padding:"6px 10px",color:C.gl2,fontWeight:700,textAlign:h==="Keterangan"?"left":"center",borderBottom:"1px solid "+C.bdr}}>{h}</th>)}</tr></thead>
<tbody>
{[["(Tbg+Isi) Di Gudang","isi",C.glt],["Tbg Kosong Di Gudang","kosong",C.gl2],["Titip di Konsumen","titip",C.blt]].map((row,i)=><tr key={i} style={{borderBottom:"1px solid "+C.bdr}}>
<td style={{padding:"6px 10px",color:C.gl2}}>{row[0]}</td>
{SIZES.map((s,j)=><td key={j} style={{padding:"6px 10px",textAlign:"center",color:row[2],fontWeight:600}}>{(viewTB.detail.stokSnapshot[s]||{})[row[1]]||0}</td>)}
</tr>)}
<tr style={{background:C.card,borderBottom:"1px solid "+C.bdr}}>
<td style={{padding:"7px 10px",color:C.wht,fontWeight:700}}>Total Keseluruhan</td>
{SIZES.map((s,j)=>{var sn=viewTB.detail.stokSnapshot[s]||{};var tot=(sn.isi||0)+(sn.kosong||0)+(sn.titip||0);return <td key={j} style={{padding:"7px 10px",textAlign:"center",color:C.olt,fontWeight:800,fontSize:14}}>{tot}</td>;})}
</tr>
</tbody>
</table>
</div>}

{/* Komposisi Bayar */}
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:12}}>
{[["Cash",viewTB.cashIn,C.glt],["Transfer",viewTB.tfIn,C.blt],["BON",viewTB.bonIn,C.olt]].map(x=><div key={x[0]} style={{background:C.nav,borderRadius:8,padding:"8px 10px",textAlign:"center",border:"1px solid "+C.bdr}}><div style={{fontSize:10,color:C.gl2}}>{x[0]}</div><div style={{fontSize:14,fontWeight:800,color:x[2]}}>{fR(x[1]||0)}</div></div>)}
</div>

{/* DO Masuk */}
{viewTB.detail?.doMasuk?.length>0&&<div style={{marginBottom:12}}>
<div style={{fontWeight:700,color:C.gl2,fontSize:12,marginBottom:6}}>🚚 DO Masuk ({viewTB.detail.doMasuk.length} DO)</div>
<div style={{background:C.nav,borderRadius:8,border:"1px solid "+C.bdr,overflow:"hidden"}}>
{viewTB.detail.doMasuk.map((d,i)=><div key={i} style={{display:"flex",justifyContent:"space-between",padding:"6px 12px",borderBottom:"1px solid "+C.bdr,fontSize:12}}><span style={{color:C.gl2}}>{d.trip} — {d.sppbe}</span><span style={{color:C.wht,fontWeight:700}}>{d.ukuran} × {d.qty} tab · HPP {fR(d.totalHPP||0)}</span></div>)}
</div>
</div>}

{/* Pengeluaran per kategori */}
{viewTB.detail?.katPengeluaran&&Object.keys(viewTB.detail.katPengeluaran).length>0&&<div style={{marginBottom:12}}>
<div style={{fontWeight:700,color:C.gl2,fontSize:12,marginBottom:6}}>💸 Pengeluaran per Kategori</div>
<div style={{background:C.nav,borderRadius:8,border:"1px solid "+C.bdr,overflow:"hidden"}}>
{Object.entries(viewTB.detail.katPengeluaran).sort((a,b)=>b[1]-a[1]).map(([k,v],i)=><div key={i} style={{display:"flex",justifyContent:"space-between",padding:"6px 12px",borderBottom:"1px solid "+C.bdr,fontSize:12}}><span style={{color:C.gl2}}>{k}</span><span style={{color:C.rlt,fontWeight:700}}>{fR(v)}</span></div>)}
</div>
</div>}

{/* Stok Snapshot */}
{viewTB.detail?.stokSnapshot&&<div style={{marginBottom:12}}>
<div style={{fontWeight:700,color:C.gl2,fontSize:12,marginBottom:6}}>📦 Snapshot Stok Saat Tutup Buku</div>
<div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:6}}>
{SIZES.map(s=>{var sn=viewTB.detail.stokSnapshot[s]||{};return <div key={s} style={{background:C.nav,borderRadius:8,padding:"8px 10px",border:"1px solid "+C.bdr}}>
<div style={{fontSize:11,fontWeight:700,color:C.wht,marginBottom:4}}>{s}</div>
{[["Isi",sn.isi,C.glt],["Titip",sn.titip,C.blt],["Kosong",sn.kosong,C.gl2]].map(x=><div key={x[0]} style={{display:"flex",justifyContent:"space-between",fontSize:11}}><span style={{color:C.gl2}}>{x[0]}</span><span style={{color:x[2],fontWeight:700}}>{x[1]||0}</span></div>)}
</div>;})}
</div>
</div>}

{/* Daftar Penjualan */}
{viewTB.detail?.penjualan?.length>0&&<div>
<div style={{fontWeight:700,color:C.gl2,fontSize:12,marginBottom:6}}>🧾 Transaksi Penjualan ({viewTB.detail.penjualan.length} invoice)</div>
<div style={{background:C.nav,borderRadius:8,border:"1px solid "+C.bdr,overflow:"hidden",maxHeight:200,overflowY:"auto"}}>
{viewTB.detail.penjualan.map((p,i)=><div key={i} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"6px 12px",borderBottom:"1px solid "+C.bdr,fontSize:12}}>
<div><div style={{color:C.blt,fontWeight:700}}>{p.noInv}</div><div style={{color:C.gl2,fontSize:11}}>{p.konsumen}</div></div>
<div style={{textAlign:"right"}}><div style={{color:C.wht,fontWeight:700}}>{fR(p.total)}</div><div style={{fontSize:10}}><Bdg color={p.bayar==="bon"?"red":p.bayar==="transfer"?"blue":"green"}>{p.bayar}</Bdg></div></div>
</div>)}
</div>
</div>}

<div style={{marginTop:14,display:"flex",gap:8,flexWrap:"wrap"}}>
{(()=>{var fn=makeFileName("tb",viewTB.tanggal,"","");return <>
<button onClick={()=>doPrint("_tb_view",fn+".pdf")} style={{background:C.blu,border:"none",borderRadius:8,padding:"8px 16px",color:"#fff",cursor:"pointer",fontWeight:700,fontSize:12}}>🖨️ Cetak / PDF</button>
<button onClick={()=>doDownloadPNG("_tb_view",fn+".png")} style={{background:"#1D6A96",border:"none",borderRadius:8,padding:"8px 14px",color:"#fff",cursor:"pointer",fontWeight:700,fontSize:12}}>💾 Download PNG</button>
<button onClick={()=>doCopyPNG("_tb_view")} style={{background:"#145A32",border:"none",borderRadius:8,padding:"8px 14px",color:"#fff",cursor:"pointer",fontWeight:700,fontSize:12}}>📋 Copy PNG</button>
<button onClick={()=>setViewTB(null)} style={{background:C.nav,border:"1px solid "+C.bdr,borderRadius:8,padding:"8px 14px",color:C.wht,cursor:"pointer",fontWeight:700,fontSize:12}}>✕ Tutup</button>
</>;})()}
</div>
</div>
</div>}
</div>;
}

// ─── LAPORAN v4 (FilterTbl di semua tab, 2 tab baru) ──────────────────────────
function LaporanMod({data,toast}){
var C=useTheme();
var[mode,setMode]=useState("bulanan");var[bln,setBln]=useState(toMonth());var[tgl,setTgl]=useState(toDay());var[tab,setTab]=useState("harian");
var[tglHarian,setTglHarian]=useState(toDay());
var penjAll=data.penjualan||[];
var penjFilt=mode==="bulanan"?penjAll.filter(p=>(p.tanggal||"").startsWith(bln)):penjAll.filter(p=>p.tanggal===tgl);
var penFilt=mode==="bulanan"?(data.pengeluaran||[]).filter(p=>(p.tanggal||"").startsWith(bln)):(data.pengeluaran||[]).filter(p=>p.tanggal===tgl);
var omzet=penjFilt.reduce((a,p)=>a+(p.total||0),0);
var margin=penjFilt.reduce((a,p)=>a+(p.margin||0),0);
var pengeluaran=penFilt.reduce((a,p)=>a+Number(p.nominal||0),0);
var labaBersih=margin-pengeluaran;
var cash=penjFilt.filter(p=>p.bayar==="cash").reduce((a,p)=>a+p.total,0);
var tf=penjFilt.filter(p=>p.bayar==="transfer").reduce((a,p)=>a+p.total,0);
var bon=penjFilt.filter(p=>p.bayar==="bon").reduce((a,p)=>a+p.total,0);

// Per sales
var salesMap={};penjFilt.forEach(p=>{var emp=(data.employees||[]).find(e=>e.id===p.salesId);var key=p.salesId||"?";if(!salesMap[key])salesMap[key]={id:key,nama:emp?.nama||"Tanpa Sales",omzet:0,margin:0,trx:0,q55:0,q12:0,q50:0};salesMap[key].omzet+=p.total||0;salesMap[key].margin+=p.margin||0;salesMap[key].trx++;(p.items||[]).forEach(it=>{if(it.ukuran==="5.5 kg")salesMap[key].q55+=Number(it.qty||0);else if(it.ukuran==="12 kg")salesMap[key].q12+=Number(it.qty||0);else if(it.ukuran==="50 kg")salesMap[key].q50+=Number(it.qty||0);});});
var salesArr=Object.values(salesMap);

// Per kategori
var katMap={};penjFilt.forEach(p=>{var plg=(data.pelanggan||[]).find(x=>x.id===p.konsumenId);var k=plg?.kategori||"Tanpa Kategori";if(!katMap[k])katMap[k]={kategori:k,omzet:0,trx:0,unik:new Set()};katMap[k].omzet+=p.total||0;katMap[k].trx++;katMap[k].unik.add(p.konsumen);});
var katArr=Object.values(katMap).map(x=>({...x,unik:x.unik.size}));

// Per produk
var prodArr=SIZES.map(s=>{var qty=0,omz=0;penjFilt.forEach(p=>(p.items||[]).forEach(it=>{if(it.ukuran===s){qty+=Number(it.qty||0);omz+=Number(it.qty||0)*Number(it.price||0);}}));return{ukuran:s,qty,omzet:omz};});

// Per pelanggan
var plgMap={};penjFilt.forEach(p=>{var k=p.konsumenId||p.konsumen;if(!plgMap[k])plgMap[k]={id:k,nama:p.konsumen,regNo:(data.pelanggan||[]).find(x=>x.id===p.konsumenId)?.regNo||"-",kategori:(data.pelanggan||[]).find(x=>x.id===p.konsumenId)?.kategori||"-",omzet:0,trx:0};plgMap[k].omzet+=p.total||0;plgMap[k].trx++;});
var plgArr=Object.values(plgMap);

// Per sales-kategori matrix (compact)
var skMap={};penjFilt.forEach(p=>{var emp=(data.employees||[]).find(e=>e.id===p.salesId);var sn=emp?.nama||"?";var plg=(data.pelanggan||[]).find(x=>x.id===p.konsumenId);var kat=plg?.kategori||"Tanpa Kategori";var key=sn+"||"+kat;if(!skMap[key])skMap[key]={sales:sn,kategori:kat,omzet:0,trx:0};skMap[key].omzet+=p.total||0;skMap[key].trx++;});
var skArr=Object.values(skMap);

// Detail flat
var detailRows=penjFilt.map(p=>{var emp=(data.employees||[]).find(e=>e.id===p.salesId);return{...p,salesNama:emp?.nama||"-",detail:(p.items||[]).map(it=>it.qty+"×"+it.ukuran).join(", ")};});

// Chart (hanya bulanan)
// Rekap pengeluaran per kategori bulan ini
var penBlnAll=(data.pengeluaran||[]).filter(e=>(e.tanggal||"").startsWith(bln));
var katMap={};
penBlnAll.forEach(p=>{var k=p.kategori||"Lainnya";katMap[k]=(katMap[k]||0)+Number(p.nominal||0);});
var katArr=Object.entries(katMap).sort((a,b)=>b[1]-a[1]).map(([k,v])=>({kategori:k,total:v}));
var totalPenBln=penBlnAll.reduce((a,p)=>a+Number(p.nominal||0),0);

// Tren pengeluaran per bulan (6 bulan terakhir)
var bulanList=[];
var nowM=new Date();for(var mi=5;mi>=0;mi--){var d2=new Date(nowM.getFullYear(),nowM.getMonth()-mi,1);var ym=d2.getFullYear()+"-"+String(d2.getMonth()+1).padStart(2,"0");bulanList.push(ym);}
var trenPen=bulanList.map(ym=>{var total=(data.pengeluaran||[]).filter(e=>(e.tanggal||"").startsWith(ym)).reduce((a,p)=>a+Number(p.nominal||0),0);return{bulan:ym.slice(5)+"/"+ym.slice(2,4),total};});

var chartData=[];
if(mode==="bulanan"){
  var dim2=daysInMonth(bln);
  var cfKumul=0;
  var doBlnLap=(data.doList||[]).filter(e=>(e.tanggal||"").startsWith(bln)&&(e.status||"diterima")==="diterima");
  // Gabungkan doTrip items yang diterima untuk HPP
  var doTripBlnLap=[];
  (data.doTrip||[]).filter(t=>(t.tanggal||"").startsWith(bln)).forEach(function(trip){
    (trip.items||[]).forEach(function(it){
      if((it.status||"diterima")==="diterima"){
        doTripBlnLap.push({tanggal:trip.tanggal,totalHPP:Number(it.qty||0)*Number(it.hppUnit||0)});
      }
    });
  });
  for(var d=1;d<=dim2;d++){
    var ds=bln+"-"+String(d).padStart(2,"0");
    var pp=penjAll.filter(x=>x.tanggal===ds);
    var oz=pp.reduce((a,x)=>a+(x.total||0),0);
    var mg=pp.reduce((a,x)=>a+(x.margin||0),0);
    var pn=(data.pengeluaran||[]).filter(x=>x.tanggal===ds).reduce((a,x)=>a+Number(x.nominal||0),0);
    var hpp=doBlnLap.filter(x=>x.tanggal===ds).reduce((a,x)=>a+Number(x.totalHPP||0),0)
           +doTripBlnLap.filter(x=>x.tanggal===ds).reduce((a,x)=>a+Number(x.totalHPP||0),0);
    var lb=mg-pn;
    cfKumul+=lb;
    chartData.push({tgl:String(d),omzet:oz,hpp,marginKotor:mg,pengeluaran:pn,labaBersih:lb,cashFlow:cfKumul});
  }
}

function exportExcel(){
var wb=XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([["Tgl","No.Inv","Konsumen","Sales","Total","Margin","Bayar"],...penjFilt.map(p=>{var emp=(data.employees||[]).find(e=>e.id===p.salesId);return[p.tanggal,p.noInv||"",p.konsumen,emp?.nama||"",p.total,p.margin,p.bayar];})]),"Penjualan");
// ── Detail Produk: 1 baris per item (Sales, Konsumen, Produk, Qty, Margin) ──
var detailRows=[];
penjFilt.forEach(p=>{
  var emp=(data.employees||[]).find(e=>e.id===p.salesId);
  var items=p.items||[];
  var totalQtyInv=items.reduce((a,it)=>a+Number(it.qty||0),0);
  items.forEach(it=>{
    var q=Number(it.qty||0);
    var marginItem;
    if(it.hppFIFO!=null){
      // Margin akurat dari HPP FIFO per unit
      marginItem=(Number(it.price||0)-Number(it.hppFIFO||0))*q;
    }else{
      // Fallback: alokasi proporsional dari margin invoice (transaksi sebelum FIFO aktif)
      marginItem=totalQtyInv>0?(p.margin||0)*(q/totalQtyInv):0;
    }
    detailRows.push([
      p.tanggal,p.noInv||"",emp?.nama||"",p.konsumen,
      it.ukuran+" — "+(it.jenis==="Tabung+Isi"?"Tabung+Isi":"Refill/Isi"),
      q,Number(it.price||0),q*Number(it.price||0),Math.round(marginItem)
    ]);
  });
});
var totalQtyAll=detailRows.reduce((a,r)=>a+r[5],0);
var totalMarginAll=detailRows.reduce((a,r)=>a+r[8],0);
var detailAOA=[["Tanggal","No.Invoice","Sales","Konsumen","Produk","Qty","Harga Satuan","Subtotal","Margin"],...detailRows,[],["","","","","TOTAL",totalQtyAll,"","",totalMarginAll]];
var wsDetail=XLSX.utils.aoa_to_sheet(detailAOA);
wsDetail["!cols"]=[{wch:12},{wch:20},{wch:14},{wch:18},{wch:24},{wch:8},{wch:13},{wch:14},{wch:13}];
XLSX.utils.book_append_sheet(wb,wsDetail,"Detail Produk");
XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([["Sales","Omzet","Margin","Trx","5.5kg","12kg","50kg"],...salesArr.map(s=>[s.nama,s.omzet,s.margin,s.trx,s.q55,s.q12,s.q50])]),"Per Sales");
XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([["Kategori","Omzet","Unik","Trx"],...katArr.map(k=>[k.kategori,k.omzet,k.unik,k.trx])]),"Per Kategori");
XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([["Pelanggan","Reg","Kategori","Omzet","Trx"],...plgArr.map(p=>[p.nama,p.regNo,p.kategori,p.omzet,p.trx])]),"Per Pelanggan");
XLSX.writeFile(wb,"Laporan_HTS_"+(mode==="bulanan"?bln:tgl)+".xlsx");
toast("✓ Excel didownload!");
}

var salesCols=[{key:"nama",label:"Sales",filterable:true},{key:"omzet",label:"Omzet",render:r=><b style={{color:C.blt}}>{fR(r.omzet)}</b>,sortVal:r=>r.omzet,filterable:false},{key:"margin",label:"Margin",render:r=><span style={{color:C.glt}}>{fR(r.margin)}</span>,sortVal:r=>r.margin,filterable:false},{key:"trx",label:"Trx",sortVal:r=>r.trx,filterable:false},{key:"q55",label:"5.5kg",sortVal:r=>r.q55,filterable:false},{key:"q12",label:"12kg",sortVal:r=>r.q12,filterable:false},{key:"q50",label:"50kg",sortVal:r=>r.q50,filterable:false}];
var katCols=[{key:"kategori",label:"Kategori",filterable:true},{key:"omzet",label:"Omzet",render:r=><b style={{color:C.blt}}>{fR(r.omzet)}</b>,sortVal:r=>r.omzet,filterable:false},{key:"unik",label:"Plg Unik",sortVal:r=>r.unik,filterable:false},{key:"trx",label:"Trx",sortVal:r=>r.trx,filterable:false},{key:"pct",label:"% Omzet",render:r=><span style={{color:C.gl2}}>{omzet>0?(r.omzet/omzet*100).toFixed(1):0}%</span>,filterable:false}];
var plgCols=[{key:"nama",label:"Pelanggan",filterable:true},{key:"regNo",label:"Reg",filterable:true},{key:"kategori",label:"Kategori",filterable:true},{key:"omzet",label:"Omzet",render:r=><b style={{color:C.blt}}>{fR(r.omzet)}</b>,sortVal:r=>r.omzet,filterable:false},{key:"trx",label:"Trx",sortVal:r=>r.trx,filterable:false}];
var skCols=[{key:"sales",label:"Sales",filterable:true},{key:"kategori",label:"Kategori",filterable:true},{key:"omzet",label:"Omzet",render:r=><b style={{color:C.blt}}>{fR(r.omzet)}</b>,sortVal:r=>r.omzet,filterable:false},{key:"trx",label:"Trx",sortVal:r=>r.trx,filterable:false}];
var detCols=[{key:"tanggal",label:"Tgl",render:r=>fDs(r.tanggal),sortVal:r=>r.tanggal,filterable:true},{key:"noInv",label:"Invoice",render:r=><span style={{fontSize:11,color:C.blt}}>{r.noInv||"-"}</span>,filterable:true},{key:"konsumen",label:"Konsumen",filterable:true},{key:"salesNama",label:"Sales",filterable:true},{key:"detail",label:"Produk",filterable:true},{key:"total",label:"Total",render:r=><b style={{color:C.wht}}>{fR(r.total)}</b>,sortVal:r=>r.total,filterable:false},{key:"bayar",label:"Bayar",render:r=>r.bayar==="bon"?<Bdg color="red">BON</Bdg>:r.bayar==="transfer"?<Bdg color="blue">TF</Bdg>:<Bdg color="green">Cash</Bdg>,filterable:true,filterType:"select",options:[{v:"cash",l:"Cash"},{v:"transfer",l:"Transfer"},{v:"bon",l:"BON"}]}];

return <div>
<STitle icon="📊" children="Laporan"/>
<Card>
<div style={{display:"flex",gap:8,marginBottom:10,flexWrap:"wrap"}}>{[["bulanan","📆 Bulanan"],["harian","📅 Harian"]].map(x=><button key={x[0]} onClick={()=>setMode(x[0])} style={{background:mode===x[0]?C.blu:C.nav,color:mode===x[0]?"white":C.wht,border:"1px solid "+(mode===x[0]?C.blt:C.bdr),borderRadius:8,padding:"7px 14px",fontWeight:700,fontSize:12,cursor:"pointer"}}>{x[1]}</button>)}</div>
{mode==="bulanan"?<MonthPicker value={bln} onChange={setBln} label=""/>:<Inp type="date" value={tgl} onChange={setTgl} label=""/>}
<Btn sm color="green" onClick={exportExcel}>📥 Export Excel</Btn>
</Card>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(160px,1fr))",gap:10,marginBottom:14}}>
{[["Omzet",omzet,C.wht,"📈"],["Laba Kotor",margin,C.blt,"💹"],["Pengeluaran",pengeluaran,C.rlt,"💸"],["Laba Bersih",labaBersih,labaBersih>=0?C.glt:C.rlt,"🏆"],["Transaksi",penjFilt.length+" trx",C.gl2,"🧾"]].map(x=><SC key={x[0]} label={x[0]} value={typeof x[1]==="number"?fR(x[1]):x[1]} icon={x[3]} color={x[2]}/>)}
</div>
<div style={{display:"flex",gap:5,marginBottom:14,flexWrap:"wrap"}}>{[["harian","🗓️ Lap. Harian"],["bukukas","📔 Buku Kas Harian"],["grafik","📈 Grafik"],["stok","📋 Stok Harian"],["pengeluaran","💸 Pengeluaran"],["sales","👤 Per Sales"],["produk","📦 Per Produk"],["pelanggan","👥 Per Pelanggan"],["matrix","📋 Sales×Kategori"],["detail","🔍 Detail"]].map(x=><button key={x[0]} onClick={()=>setTab(x[0])} style={{background:tab===x[0]?C.blu:C.nav,color:tab===x[0]?"white":C.wht,border:"1px solid "+(tab===x[0]?C.blt:C.bdr),borderRadius:8,padding:"6px 11px",fontWeight:700,fontSize:11,cursor:"pointer"}}>{x[1]}</button>)}</div>
{tab==="harian"&&(()=>{
// ── Data hari ini ──
var HARI_ID=["Minggu","Senin","Selasa","Rabu","Kamis","Jumat","Sabtu"];
var dtH=new Date(tglHarian+"T00:00:00");
var hariLabel=HARI_ID[dtH.getDay()]+", "+dtH.toLocaleDateString("id-ID",{day:"2-digit",month:"long",year:"numeric"});
var penjH=(data.penjualan||[]).filter(p=>p.tanggal===tglHarian);
var doH=(data.doList||[]).filter(d=>d.tanggal===tglHarian&&(d.status||"diterima")==="diterima");
var penH=(data.pengeluaran||[]).filter(p=>p.tanggal===tglHarian);
var bonBayarH=(data.bon||[]).flatMap(b=>(b.pembayaran||[]).filter(px=>px.tanggal===tglHarian).map(px=>({...px,konsumen:b.konsumen,salesId:b.salesId,salesNama:b.salesNama||""})));
// Stok dari buildStokHarian
var bulanH=tglHarian.slice(0,7);
var stokRows=buildStokHarian(data,bulanH).filter(r=>r.tgl<=tglHarian);
var stokRow=stokRows.length>0?stokRows[stokRows.length-1]:null;
// HPP per ukuran (dari modalHistory)
var getHpp=function(uk){var mh=(data.modalHistory||[]).filter(m=>m.ukuran===uk&&m.jenis==="Isi"&&m.tanggal<=tglHarian);return mh.length>0?mh[0].harga:0;};
// Kelompok per sales
var salesGroups={};
penjH.forEach(p=>{var empP=(data.employees||[]).find(e=>e.id===p.salesId);var sNama=empP?.nama||p.salesNama||p.sales||"";var key=sNama||("(Tanpa Sales)");if(!salesGroups[key])salesGroups[key]={nama:key,items:[],omzet:0,margin:0};salesGroups[key].items.push(p);salesGroups[key].omzet+=(p.total||0);salesGroups[key].margin+=(p.margin||0);});
// Pengeluaran per sales
var penPerSales={};
penH.forEach(p=>{var key=p.karyawanNama||"(Umum)";if(!penPerSales[key])penPerSales[key]=0;penPerSales[key]+=Number(p.nominal||0);});
var totalPenH=penH.reduce((a,p)=>a+Number(p.nominal||0),0);
var totalOmzetH=penjH.reduce((a,p)=>a+(p.total||0),0);
var totalMarginH=penjH.reduce((a,p)=>a+(p.margin||0),0);
var totalBonH=bonBayarH.reduce((a,b)=>a+Number(b.jumlah||b.nominal||0),0);
// Print styles
var PS={page:{background:"white",color:"#111",fontFamily:"'Plus Jakarta Sans',Arial,sans-serif",padding:20,maxWidth:900,margin:"0 auto"},h1:{fontSize:18,fontWeight:800,color:"#0a1f44",marginBottom:2},h2:{fontSize:13,fontWeight:700,color:"#0a1f44",margin:"16px 0 8px",borderBottom:"2px solid #0a1f44",paddingBottom:4},th:{background:"#0a1f44",color:"white",padding:"6px 8px",fontSize:10,fontWeight:700,textAlign:"center",border:"1px solid #ccc"},td:{padding:"5px 8px",fontSize:11,border:"1px solid #ddd",verticalAlign:"top"},tdL:{padding:"5px 8px",fontSize:11,border:"1px solid #ddd",textAlign:"left"},tbl:{width:"100%",borderCollapse:"collapse",marginBottom:10},sub:{fontSize:11,color:"#555",marginBottom:2}};

return <div>
<div style={{display:"flex",gap:10,alignItems:"flex-end",marginBottom:14,flexWrap:"wrap"}}>
<Inp label="Tanggal Laporan" type="date" value={tglHarian} onChange={setTglHarian} style={{maxWidth:200,marginBottom:0}}/>
<button onClick={()=>doPrint("_lap_harian")} style={{background:"#0a1f44",color:"white",border:"none",padding:"9px 18px",borderRadius:8,fontSize:13,cursor:"pointer",fontWeight:700}}>🖨️ Cetak / PDF</button>
<span style={{fontSize:10,color:C.gl2,fontStyle:"italic",alignSelf:"center"}}>💡 PDF: Lap-Harian-{tglHarian}.pdf</span>
</div>

<div id="_lap_harian" style={PS.page}>
{/* HEADER */}
<div style={{textAlign:"center",borderBottom:"3px solid #0a1f44",paddingBottom:12,marginBottom:16}}>
<div style={{fontSize:20,fontWeight:900,color:"#0a1f44"}}>LAPORAN HARIAN OPERASIONAL</div>
<div style={{fontSize:14,fontWeight:700,color:"#0a1f44"}}>{(data.company?.nama||"PT. HOE TRANG SA").toUpperCase()}</div>
<div style={{fontSize:12,color:"#555",marginTop:4}}>{hariLabel}</div>
</div>

{/* 1. MUTASI STOK */}
<div style={PS.h2}>1. MUTASI STOK TABUNG</div>
{stokRow?<table style={PS.tbl}>
<thead><tr>
<th style={{...PS.th,textAlign:"left"}}>Keterangan</th>
{SIZES.map(s=>[<th key={"isi"+s} style={PS.th}>isi {s}</th>,<th key={"tk"+s} style={PS.th}>TK {s}</th>])}
{SIZES.map(s=><th key={"tot"+s} style={PS.th}>Total {s}</th>)}
</tr></thead>
<tbody>
{[
["Stok Awal",SIZES.flatMap(s=>[stokRow.awalIsi[s],stokRow.awalTK[s]])],
["Tabung Masuk (DO+Return)",SIZES.flatMap(s=>[stokRow.masukIsi[s],stokRow.masukTK[s]])],
["Tabung Keluar (Penjualan)",SIZES.flatMap(s=>[stokRow.keluarIsi[s],Math.abs(stokRow.keluarTK[s])])],
["Stok Akhir",SIZES.flatMap(s=>[stokRow.akhirIsi[s],stokRow.akhirTK[s]])],
].map((r,i)=><tr key={i} style={{background:i===3?"#EFF6FF":"white"}}>
<td style={{...PS.tdL,fontWeight:i===3?700:400}}>{r[0]}</td>
{r[1].map((v,j)=><td key={j} style={{...PS.td,textAlign:"center",fontWeight:i===3?700:400}}>{v}</td>)}
{i===3?SIZES.map(s=><td key={s} style={{...PS.td,textAlign:"center",fontWeight:700,color:"#1D4ED8"}}>{stokRow.total[s]}</td>):<>{SIZES.map(s=><td key={s} style={{...PS.td,textAlign:"center",color:"#888"}}>—</td>)}</>}
</tr>)}
</tbody>
</table>:<div style={{color:"#888",fontSize:11,marginBottom:10}}>Tidak ada data stok untuk tanggal ini. Gunakan Inject Stok Awal terlebih dahulu.</div>}

{/* 2. DELIVERY ORDER */}
<div style={PS.h2}>2. DELIVERY ORDER MASUK</div>
{doH.length>0?<><table style={PS.tbl}>
<thead><tr>{["Trip","SPPBE","Ukuran","Qty","HPP/Unit","Total HPP","Driver"].map(h=><th key={h} style={PS.th}>{h}</th>)}</tr></thead>
<tbody>
{doH.map((d,i)=><tr key={i} style={{background:i%2===0?"white":"#f9f9f9"}}>
<td style={PS.td}>{d.trip}</td><td style={PS.td}>{d.sppbe}</td><td style={PS.td}>{d.ukuran}</td>
<td style={{...PS.td,textAlign:"center",fontWeight:700}}>{d.qty}</td>
<td style={{...PS.td,textAlign:"right"}}>{fR(d.hppUnit||0)}</td>
<td style={{...PS.td,textAlign:"right",fontWeight:700}}>{fR(d.totalHPP||0)}</td>
<td style={PS.td}>{d.dibuatOleh||"-"}</td>
</tr>)}
</tbody>
</table>
<div style={{fontSize:11,fontWeight:700,marginBottom:4}}>
Total DO: {SIZES.map(s=>{var q=doH.filter(d=>d.ukuran===s).reduce((a,d)=>a+Number(d.qty||0),0);return q>0?s+" = "+q+" tab":null;}).filter(Boolean).join(", ")} | Total HPP: {fR(doH.reduce((a,d)=>a+Number(d.totalHPP||0),0))}
</div></>:<div style={{color:"#888",fontSize:11,marginBottom:10}}>Tidak ada DO masuk hari ini.</div>}

{/* 3. PENJUALAN PER SALES */}
<div style={PS.h2}>3. PENJUALAN — {penjH.length} Invoice | Omzet: {fR(totalOmzetH)} | Margin: {fR(totalMarginH)}</div>
{Object.values(salesGroups).length>0?Object.values(salesGroups).map((sg,gi)=>{
var q55sg=sg.items.reduce((a,p)=>a+(p.items||[]).filter(it=>it.ukuran==="5.5 kg").reduce((b,it)=>b+Number(it.qty||0),0),0);
var q12sg=sg.items.reduce((a,p)=>a+(p.items||[]).filter(it=>it.ukuran==="12 kg").reduce((b,it)=>b+Number(it.qty||0),0),0);
var q50sg=sg.items.reduce((a,p)=>a+(p.items||[]).filter(it=>it.ukuran==="50 kg").reduce((b,it)=>b+Number(it.qty||0),0),0);
var sgCash=sg.items.filter(p=>(p.bayar||"").toLowerCase()==="cash").reduce((a,p)=>a+(p.total||0),0);
var sgTF=sg.items.filter(p=>(p.bayar||"").toLowerCase()==="transfer"||(p.bayar||"").toLowerCase()==="tf").reduce((a,p)=>a+(p.total||0),0);
var sgBon=sg.items.filter(p=>(p.bayar||"").toLowerCase()==="bon").reduce((a,p)=>a+(p.total||0),0);
return <div key={gi} style={{marginBottom:12}}>
<div style={{fontWeight:700,fontSize:12,color:"#0a1f44",background:"#EFF6FF",padding:"5px 8px",borderRadius:4,marginBottom:5}}>── {sg.nama} &nbsp;—&nbsp; Omzet: {fR(sg.omzet)} | Margin: {fR(sg.margin)}</div>
<table style={PS.tbl}>
<thead><tr>{["No. Invoice","Konsumen","5,5kg","12kg","50kg","Total","Bayar"].map(h=><th key={h} style={PS.th}>{h}</th>)}</tr></thead>
<tbody>
{sg.items.map((p,i)=>{
var q55p=(p.items||[]).filter(it=>it.ukuran==="5.5 kg").reduce((a,it)=>a+Number(it.qty||0),0);
var q12p=(p.items||[]).filter(it=>it.ukuran==="12 kg").reduce((a,it)=>a+Number(it.qty||0),0);
var q50p=(p.items||[]).filter(it=>it.ukuran==="50 kg").reduce((a,it)=>a+Number(it.qty||0),0);
var byr=(p.bayar||"").toLowerCase();
return <tr key={i} style={{background:i%2===0?"white":"#f9f9f9"}}>
<td style={PS.td}>{p.noInv}</td>
<td style={{...PS.tdL,fontWeight:600}}>{p.konsumen}</td>
<td style={{...PS.td,textAlign:"center"}}>{q55p||"—"}</td>
<td style={{...PS.td,textAlign:"center"}}>{q12p||"—"}</td>
<td style={{...PS.td,textAlign:"center"}}>{q50p||"—"}</td>
<td style={{...PS.td,textAlign:"right",fontWeight:700}}>{fR(p.total)}</td>
<td style={{...PS.td,textAlign:"center",color:byr==="cash"?"#15803D":byr==="bon"?"#DC2626":"#1D4ED8",fontWeight:600}}>{p.bayar}</td>
</tr>;})}
<tr style={{background:"#EFF6FF",fontWeight:700}}>
<td colSpan={2} style={{...PS.tdL,fontWeight:700,color:"#0a1f44"}}>Sub-total {sg.nama}</td>
<td style={{...PS.td,textAlign:"center",color:"#0a1f44"}}>{q55sg||"—"}</td>
<td style={{...PS.td,textAlign:"center",color:"#0a1f44"}}>{q12sg||"—"}</td>
<td style={{...PS.td,textAlign:"center",color:"#0a1f44"}}>{q50sg||"—"}</td>
<td colSpan={2} style={{...PS.td,textAlign:"left",color:"#0a1f44"}}>
Cash: {fR(sgCash)} &nbsp;|&nbsp; TF: {fR(sgTF)} &nbsp;|&nbsp; <span style={{color:"#888"}}>BON: {fR(sgBon)}</span> &nbsp;|&nbsp; Total: <b>{fR(sg.omzet)}</b>
</td>
</tr>
</tbody>
</table>
</div>;}):
<div style={{color:"#888",fontSize:11,marginBottom:10}}>Tidak ada penjualan hari ini.</div>}
{penjH.length>0&&(()=>{
var totCashH=penjH.filter(p=>(p.bayar||"").toLowerCase()==="cash").reduce((a,p)=>a+(p.total||0),0);
var totTFH=penjH.filter(p=>(p.bayar||"").toLowerCase()==="transfer"||(p.bayar||"").toLowerCase()==="tf").reduce((a,p)=>a+(p.total||0),0);
var totBonH2=penjH.filter(p=>(p.bayar||"").toLowerCase()==="bon").reduce((a,p)=>a+(p.total||0),0);
return <table style={{...PS.tbl,marginTop:6}}>
<thead><tr style={{background:"#0a1f44"}}><th colSpan={4} style={{...PS.th,textAlign:"left",fontSize:11}}>TOTAL PENJUALAN — {hariLabel}</th></tr>
<tr style={{background:"#0a1f44"}}>{["Total Cash","Total Transfer","BON / Piutang Baru","TOTAL OMZET"].map(h=><th key={h} style={PS.th}>{h}</th>)}</tr></thead>
<tbody><tr style={{background:"#EFF6FF",fontWeight:700}}>
<td style={{...PS.td,textAlign:"right",color:"#15803D"}}>{fR(totCashH)}</td>
<td style={{...PS.td,textAlign:"right",color:"#1D4ED8"}}>{fR(totTFH)}</td>
<td style={{...PS.td,textAlign:"right",color:"#888"}}>{fR(totBonH2)}</td>
<td style={{...PS.td,textAlign:"right",color:"#0a1f44",fontSize:13}}>{fR(totalOmzetH)}</td>
</tr></tbody>
</table>;})()}

{/* 4. PEMBAYARAN BON */}
<div style={PS.h2}>4. PEMBAYARAN BON / PIUTANG</div>
{bonBayarH.length>0?<><table style={PS.tbl}>
<thead><tr>{["Konsumen","Nominal Bayar","Metode","Sales Penerima"].map(h=><th key={h} style={PS.th}>{h}</th>)}</tr></thead>
<tbody>
{bonBayarH.map((b,i)=>{var metBon=(b.metode||"cash").toLowerCase();return <tr key={i} style={{background:i%2===0?"white":"#f9f9f9"}}>
<td style={PS.tdL}>{b.konsumen}</td>
<td style={{...PS.td,textAlign:"right",fontWeight:700}}>{fR(b.jumlah||b.nominal||0)}</td>
<td style={{...PS.td,textAlign:"center",color:metBon==="cash"?"#15803D":"#1D4ED8",fontWeight:600}}>{b.metode||"Cash"}</td>
<td style={PS.td}>{b.salesNama||"-"}</td>
</tr>;})}
</tbody>
</table>
{(()=>{
var bonCashH=bonBayarH.filter(b=>(b.metode||"cash").toLowerCase()==="cash").reduce((a,b)=>a+Number(b.jumlah||b.nominal||0),0);
var bonTFH=bonBayarH.filter(b=>(b.metode||"").toLowerCase()==="transfer"||(b.metode||"").toLowerCase()==="tf").reduce((a,b)=>a+Number(b.jumlah||b.nominal||0),0);
return <div style={{display:"flex",gap:12,marginTop:6,padding:"6px 10px",background:"#FEF3C7",borderRadius:4,fontSize:11,fontWeight:700}}>
<span>Cash: <b style={{color:"#15803D"}}>{fR(bonCashH)}</b></span>
<span>Transfer: <b style={{color:"#1D4ED8"}}>{fR(bonTFH)}</b></span>
<span style={{marginLeft:"auto"}}>TOTAL: <b style={{color:"#0a1f44"}}>{fR(totalBonH)}</b></span>
</div>;})()} 
</>:<div style={{color:"#888",fontSize:11,marginBottom:10}}>Tidak ada pembayaran BON hari ini.</div>}

{/* 5. PENGELUARAN */}
<div style={PS.h2}>5. PENGELUARAN OPERASIONAL</div>
{penH.length>0?<>{(()=>{
var penCashH=penH.filter(p=>(p.metode||"cash").toLowerCase()==="cash").reduce((a,p)=>a+Number(p.nominal||0),0);
var penTFH=penH.filter(p=>(p.metode||"").toLowerCase()==="transfer"||(p.metode||"").toLowerCase()==="tf").reduce((a,p)=>a+Number(p.nominal||0),0);
return <>
<table style={PS.tbl}>
<thead><tr>{["Kategori","Keterangan","Atas Nama","Metode","Nominal"].map(h=><th key={h} style={PS.th}>{h}</th>)}</tr></thead>
<tbody>
{penH.map((p,i)=>{var met=(p.metode||"cash").toLowerCase();return <tr key={i} style={{background:i%2===0?"white":"#f9f9f9"}}>
<td style={PS.td}>{p.kategori}</td>
<td style={PS.tdL}>{p.ket||"-"}</td>
<td style={PS.td}>{p.karyawanNama||"-"}</td>
<td style={{...PS.td,textAlign:"center",color:met==="cash"?"#15803D":"#1D4ED8",fontWeight:600}}>{p.metode||"Cash"}</td>
<td style={{...PS.td,textAlign:"right",fontWeight:700,color:"#DC2626"}}>{fR(p.nominal)}</td>
</tr>;})}
</tbody>
</table>
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginTop:6}}>
<div style={{background:"#FEF2F2",borderRadius:4,padding:"6px 10px",border:"1px solid #FECACA"}}>
<div style={{fontSize:9,color:"#6B7280"}}>Total Cash ({penH.filter(p=>(p.metode||"cash").toLowerCase()==="cash").length} trx)</div>
<div style={{fontWeight:700,fontSize:12,color:"#DC2626"}}>{fR(penCashH)}</div>
</div>
<div style={{background:"#EFF6FF",borderRadius:4,padding:"6px 10px",border:"1px solid #BFDBFE"}}>
<div style={{fontSize:9,color:"#6B7280"}}>Total Transfer ({penH.filter(p=>(p.metode||"").toLowerCase()==="transfer"||(p.metode||"").toLowerCase()==="tf").length} trx)</div>
<div style={{fontWeight:700,fontSize:12,color:"#1D4ED8"}}>{fR(penTFH)}</div>
</div>
<div style={{background:"#FEE2E2",borderRadius:4,padding:"6px 10px",border:"1px solid #FCA5A5"}}>
<div style={{fontSize:9,color:"#6B7280"}}>TOTAL PENGELUARAN ({penH.length} trx)</div>
<div style={{fontWeight:700,fontSize:13,color:"#DC2626"}}>{fR(totalPenH)}</div>
</div>
</div>
</>;})()} 
</>:<div style={{color:"#888",fontSize:11,marginBottom:10}}>Tidak ada pengeluaran hari ini.</div>}

{/* 6. RINGKASAN PER SALES */}
<div style={PS.h2}>6. RINGKASAN PER SALES</div>
<div style={{overflowX:"auto"}}>
<table style={{...PS.tbl,minWidth:700}}>
<thead>
<tr style={{background:"#0a1f44"}}>
<th rowSpan={2} style={{...PS.th,textAlign:"left",fontSize:8}}>Sales</th>
<th colSpan={2} style={{...PS.th,fontSize:8}}>5,5 kg</th>
<th colSpan={2} style={{...PS.th,fontSize:8}}>12 kg</th>
<th colSpan={2} style={{...PS.th,fontSize:8}}>50 kg</th>
<th colSpan={3} style={{...PS.th,fontSize:8}}>Penjualan</th>
<th colSpan={2} style={{...PS.th,fontSize:8}}>Bayar BON</th>
<th rowSpan={2} style={{...PS.th,fontSize:8}}>Pemasukan Kas</th>
<th rowSpan={2} style={{...PS.th,fontSize:8}}>Margin</th>
<th rowSpan={2} style={{...PS.th,fontSize:8}}>Pengeluaran</th>
</tr>
<tr style={{background:"#0a1f44"}}>
{["Qty","Nominal","Qty","Nominal","Qty","Nominal","Cash","TF","BON*","Cash","TF"].map(h=><th key={h} style={{...PS.th,fontSize:7,padding:"3px 4px"}}>{h}</th>)}
</tr>
</thead>
<tbody>
{Object.values(salesGroups).map((sg,i)=>{
var q55s=sg.items.reduce((a,p)=>a+(p.items||[]).filter(it=>it.ukuran==="5.5 kg").reduce((b,it)=>b+Number(it.qty||0),0),0);
var q12s=sg.items.reduce((a,p)=>a+(p.items||[]).filter(it=>it.ukuran==="12 kg").reduce((b,it)=>b+Number(it.qty||0),0),0);
var q50s=sg.items.reduce((a,p)=>a+(p.items||[]).filter(it=>it.ukuran==="50 kg").reduce((b,it)=>b+Number(it.qty||0),0),0);
var nom55s=sg.items.reduce((a,p)=>a+(p.items||[]).filter(it=>it.ukuran==="5.5 kg").reduce((b,it)=>b+Number(it.qty||0)*Number(it.price||0),0),0);
var nom12s=sg.items.reduce((a,p)=>a+(p.items||[]).filter(it=>it.ukuran==="12 kg").reduce((b,it)=>b+Number(it.qty||0)*Number(it.price||0),0),0);
var nom50s=sg.items.reduce((a,p)=>a+(p.items||[]).filter(it=>it.ukuran==="50 kg").reduce((b,it)=>b+Number(it.qty||0)*Number(it.price||0),0),0);
var sgCashS=sg.items.filter(p=>(p.bayar||"").toLowerCase()==="cash").reduce((a,p)=>a+(p.total||0),0);
var sgTFS=sg.items.filter(p=>(p.bayar||"").toLowerCase()==="transfer"||(p.bayar||"").toLowerCase()==="tf").reduce((a,p)=>a+(p.total||0),0);
var sgBonS=sg.items.filter(p=>(p.bayar||"").toLowerCase()==="bon").reduce((a,p)=>a+(p.total||0),0);
var bonCashS=bonBayarH.filter(b=>b.salesNama===sg.nama&&(b.metode||"cash").toLowerCase()==="cash").reduce((a,b)=>a+Number(b.jumlah||b.nominal||0),0);
var bonTFS=bonBayarH.filter(b=>b.salesNama===sg.nama&&((b.metode||"").toLowerCase()==="transfer"||(b.metode||"").toLowerCase()==="tf")).reduce((a,b)=>a+Number(b.jumlah||b.nominal||0),0);
var pemasukanKasS=sgCashS+sgTFS+bonCashS+bonTFS;
var penSales=(penPerSales[sg.nama]||0);
return <tr key={i} style={{background:i%2===0?"white":"#f9f9f9"}}>
<td style={{...PS.tdL,fontWeight:700,fontSize:9,padding:"3px 5px"}}>{sg.nama}</td>
<td style={{...PS.td,textAlign:"center",fontSize:9,padding:"3px 4px"}}>{q55s||"—"}</td>
<td style={{...PS.td,textAlign:"right",fontSize:9,padding:"3px 4px"}}>{nom55s?fR(nom55s):"—"}</td>
<td style={{...PS.td,textAlign:"center",fontSize:9,padding:"3px 4px"}}>{q12s||"—"}</td>
<td style={{...PS.td,textAlign:"right",fontSize:9,padding:"3px 4px"}}>{nom12s?fR(nom12s):"—"}</td>
<td style={{...PS.td,textAlign:"center",fontSize:9,padding:"3px 4px"}}>{q50s||"—"}</td>
<td style={{...PS.td,textAlign:"right",fontSize:9,padding:"3px 4px"}}>{nom50s?fR(nom50s):"—"}</td>
<td style={{...PS.td,textAlign:"right",fontSize:9,padding:"3px 4px"}}>{fR(sgCashS)}</td>
<td style={{...PS.td,textAlign:"right",fontSize:9,padding:"3px 4px"}}>{fR(sgTFS)}</td>
<td style={{...PS.td,textAlign:"right",color:"#888",fontSize:9,padding:"3px 4px"}}>{sgBonS?fR(sgBonS):"—"}</td>
<td style={{...PS.td,textAlign:"right",fontSize:9,padding:"3px 4px"}}>{bonCashS?fR(bonCashS):"—"}</td>
<td style={{...PS.td,textAlign:"right",fontSize:9,padding:"3px 4px"}}>{bonTFS?fR(bonTFS):"—"}</td>
<td style={{...PS.td,textAlign:"right",fontWeight:700,color:"#0a1f44",fontSize:9,padding:"3px 4px"}}>{fR(pemasukanKasS)}</td>
<td style={{...PS.td,textAlign:"right",color:"#15803D",fontWeight:700,fontSize:9,padding:"3px 4px"}}>{fR(sg.margin)}</td>
<td style={{...PS.td,textAlign:"right",color:"#DC2626",fontSize:9,padding:"3px 4px"}}>{penSales?fR(penSales):"—"}</td>
</tr>;})}
{(()=>{
var tq55=Object.values(salesGroups).reduce((a,sg)=>a+sg.items.reduce((b,p)=>b+(p.items||[]).filter(it=>it.ukuran==="5.5 kg").reduce((c,it)=>c+Number(it.qty||0),0),0),0);
var tq12=Object.values(salesGroups).reduce((a,sg)=>a+sg.items.reduce((b,p)=>b+(p.items||[]).filter(it=>it.ukuran==="12 kg").reduce((c,it)=>c+Number(it.qty||0),0),0),0);
var tq50=Object.values(salesGroups).reduce((a,sg)=>a+sg.items.reduce((b,p)=>b+(p.items||[]).filter(it=>it.ukuran==="50 kg").reduce((c,it)=>c+Number(it.qty||0),0),0),0);
var tCash=penjH.filter(p=>(p.bayar||"").toLowerCase()==="cash").reduce((a,p)=>a+(p.total||0),0);
var tTF=penjH.filter(p=>(p.bayar||"").toLowerCase()==="transfer"||(p.bayar||"").toLowerCase()==="tf").reduce((a,p)=>a+(p.total||0),0);
var tBonPenj=penjH.filter(p=>(p.bayar||"").toLowerCase()==="bon").reduce((a,p)=>a+(p.total||0),0);
var bonCashTotal=bonBayarH.filter(b=>(b.metode||"cash").toLowerCase()==="cash").reduce((a,b)=>a+Number(b.jumlah||b.nominal||0),0);
var bonTFTotal=bonBayarH.filter(b=>(b.metode||"").toLowerCase()==="transfer"||(b.metode||"").toLowerCase()==="tf").reduce((a,b)=>a+Number(b.jumlah||b.nominal||0),0);
var totalPemasukan=tCash+tTF+bonCashTotal+bonTFTotal;
return <tr style={{background:"#0a1f44",color:"white",fontWeight:700}}>
<td style={{...PS.tdL,color:"white",fontSize:9,padding:"4px 5px"}}>TOTAL</td>
<td style={{...PS.td,textAlign:"center",color:"white",fontSize:9,padding:"4px"}}>{tq55||"—"}</td>
<td style={{...PS.td,color:"white",fontSize:9,padding:"4px"}}>—</td>
<td style={{...PS.td,textAlign:"center",color:"white",fontSize:9,padding:"4px"}}>{tq12||"—"}</td>
<td style={{...PS.td,color:"white",fontSize:9,padding:"4px"}}>—</td>
<td style={{...PS.td,textAlign:"center",color:"white",fontSize:9,padding:"4px"}}>{tq50||"—"}</td>
<td style={{...PS.td,color:"white",fontSize:9,padding:"4px"}}>—</td>
<td style={{...PS.td,textAlign:"right",color:"white",fontSize:9,padding:"4px"}}>{fR(tCash)}</td>
<td style={{...PS.td,textAlign:"right",color:"white",fontSize:9,padding:"4px"}}>{fR(tTF)}</td>
<td style={{...PS.td,textAlign:"right",color:"#aaa",fontSize:9,padding:"4px"}}>{tBonPenj?fR(tBonPenj):"—"}</td>
<td style={{...PS.td,textAlign:"right",color:"white",fontSize:9,padding:"4px"}}>{fR(bonCashTotal)}</td>
<td style={{...PS.td,textAlign:"right",color:"white",fontSize:9,padding:"4px"}}>{fR(bonTFTotal)}</td>
<td style={{...PS.td,textAlign:"right",color:"#86EFAC",fontSize:10,padding:"4px",fontWeight:700}}>{fR(totalPemasukan)}</td>
<td style={{...PS.td,textAlign:"right",color:"#86EFAC",fontSize:9,padding:"4px"}}>{fR(totalMarginH)}</td>
<td style={{...PS.td,textAlign:"right",color:"#FCA5A5",fontSize:9,padding:"4px"}}>{fR(totalPenH)}</td>
</tr>;})()} 
</tbody>
</table>
</div>
<div style={{fontSize:9,color:"#888",marginTop:4}}>* BON = piutang baru, belum masuk kas &nbsp;|&nbsp; Pemasukan Kas = Cash Penj + TF Penj + Bayar BON Cash + Bayar BON TF</div>

{/* 7. REKAP TABUNG AKHIR */}
<div style={PS.h2}>7. REKAP TABUNG AKHIR HARI</div>
{stokRow?<table style={PS.tbl}>
<thead><tr>{["Keterangan","5,5 kg - Isi","5,5 kg - TK","12 kg - Isi","12 kg - TK","50 kg - Isi","50 kg - TK"].map(h=><th key={h} style={PS.th}>{h}</th>)}</tr></thead>
<tbody>
{[["Stok Akhir Isi & TK",SIZES.flatMap(s=>[stokRow.akhirIsi[s],stokRow.akhirTK[s]])],
["Titip di Konsumen",SIZES.flatMap(s=>[stokRow.titipSnap[s],"—"])],
["TOTAL KESELURUHAN",SIZES.flatMap(s=>[stokRow.total[s],"="+stokRow.total[s]])]].map((r,i)=><tr key={i} style={{background:i===2?"#DBEAFE":"white",fontWeight:i===2?700:400}}>
<td style={PS.tdL}>{r[0]}</td>
{r[1].map((v,j)=><td key={j} style={{...PS.td,textAlign:"center"}}>{v}</td>)}
</tr>)}
</tbody>
</table>:<div style={{color:"#888",fontSize:11}}>Data stok tidak tersedia.</div>}

<div style={{marginTop:20,borderTop:"2px solid #0a1f44",paddingTop:10,display:"flex",justifyContent:"space-between",fontSize:10,color:"#555"}}>
<span>Dicetak: {new Date().toLocaleString("id-ID")}</span>
<span>{data.company?.nama||"PT. HOE TRANG SA"}</span>
</div>
</div>
</div>;
})()}

{tab==="sales"&&<Card><div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>👤 Per Sales</div><FilterTbl columns={salesCols} data={salesArr} empty="Tidak ada data"/></Card>}
{tab==="produk"&&<Card><div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>📦 Per Produk</div><RTbl headers={["Ukuran","Qty","Omzet","% Omzet"]} rows={prodArr.map(p=>[<b style={{color:C.wht}}>{p.ukuran}</b>,<b style={{color:C.glt}}>{p.qty} tab</b>,<b style={{color:C.blt}}>{fR(p.omzet)}</b>,(omzet>0?(p.omzet/omzet*100).toFixed(1):0)+"%"])}/></Card>}
{tab==="pelanggan"&&<Card><div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>👥 Ranking Pelanggan per Omzet</div><FilterTbl columns={plgCols} data={plgArr} empty="Tidak ada data"/></Card>}
{tab==="matrix"&&<Card><div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>📋 Sales × Kategori</div><FilterTbl columns={skCols} data={skArr} empty="Tidak ada data"/></Card>}
{tab==="detail"&&<Card><div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>🔍 Detail Penjualan ({penjFilt.length})</div><FilterTbl columns={detCols} data={detailRows} empty="Tidak ada data" maxRows={300}/></Card>}

{tab==="stok"&&<div>
<Card>
<div style={{fontWeight:700,color:C.gl2,marginBottom:12,fontSize:13}}>📋 Laporan Stok Harian — {BULAN_ID[Number(bln.split("-")[1])-1]} {bln.split("-")[0]}</div>
<TabelStokBulanan data={data} bulan={bln}/>
</Card>
</div>}
{tab==="pengeluaran"&&<div>
{/* Summary */}
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(150px,210px))",gap:8,marginBottom:12}}>
<div style={{background:C.card,borderRadius:8,padding:"10px 12px",border:"1px solid "+C.bdr}}><div style={{fontSize:10,color:C.gl2}}>Total Pengeluaran</div><div style={{fontSize:15,fontWeight:900,color:C.rlt}}>{fR(totalPenBln)}</div></div>
<div style={{background:C.card,borderRadius:8,padding:"10px 12px",border:"1px solid "+C.bdr}}><div style={{fontSize:10,color:C.gl2}}>Jumlah Kategori</div><div style={{fontSize:15,fontWeight:900,color:C.wht}}>{katArr.length}</div></div>
<div style={{background:C.card,borderRadius:8,padding:"10px 12px",border:"1px solid "+C.bdr}}><div style={{fontSize:10,color:C.gl2}}>Jumlah Transaksi</div><div style={{fontSize:15,fontWeight:900,color:C.wht}}>{penBlnAll.length}</div></div>
</div>

{/* Tabel per kategori */}
{katArr.length>0?<Card style={{marginBottom:10}}>
<div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>💸 Total per Kategori — {BULAN_ID[Number(bln.split("-")[1])-1]} {bln.split("-")[0]}</div>
<RTbl headers={["Kategori","Total","% dari Total","Transaksi"]} rows={katArr.map(k=>{
var pct=totalPenBln>0?Math.round(k.total/totalPenBln*100):0;
var cnt=penBlnAll.filter(p=>(p.kategori||"Lainnya")===k.kategori).length;
return[
<b style={{color:C.wht}}>{k.kategori}</b>,
<b style={{color:C.rlt}}>{fR(k.total)}</b>,
<div style={{display:"flex",alignItems:"center",gap:6}}><div style={{width:Math.max(4,pct)+"px",maxWidth:80,height:6,background:C.rlt,borderRadius:3}}/><span style={{color:C.gl2,fontSize:11}}>{pct}%</span></div>,
<span style={{color:C.gl2}}>{cnt}x</span>
];})}/>
</Card>:<Card><div style={{color:C.gl2,textAlign:"center",padding:20}}>Tidak ada pengeluaran bulan ini</div></Card>}

{/* Grafik bar per kategori */}
{katArr.length>0&&<Card style={{marginBottom:10}}>
<div style={{fontWeight:700,color:C.gl2,marginBottom:8,fontSize:13}}>📊 Grafik per Kategori</div>
<ResponsiveContainer width="100%" height={Math.max(200,katArr.length*35)}><BarChart layout="vertical" data={katArr} margin={{top:4,right:60,bottom:0,left:10}}>
<CartesianGrid strokeDasharray="3 3" stroke={C.bdr}/>
<XAxis type="number" stroke={C.gl2} fontSize={9} tickFormatter={v=>(v/1e6).toFixed(1)+"jt"}/>
<YAxis type="category" dataKey="kategori" stroke={C.gl2} fontSize={9} width={110}/>
<Tooltip contentStyle={{background:C.card,border:"1px solid "+C.bdr,color:C.wht,fontSize:11}} formatter={v=>fR(v)}/>
<Bar dataKey="total" fill="#EF4444" name="Total" radius={[0,4,4,0]}/>
</BarChart></ResponsiveContainer>
</Card>}

{/* Tren 6 bulan */}
<Card>
<div style={{fontWeight:700,color:C.gl2,marginBottom:8,fontSize:13}}>📈 Tren Pengeluaran 6 Bulan Terakhir</div>
<ResponsiveContainer width="100%" height={180}><AreaChart data={trenPen} margin={{top:4,right:10,bottom:0,left:10}}>
<CartesianGrid strokeDasharray="3 3" stroke={C.bdr}/>
<XAxis dataKey="bulan" stroke={C.gl2} fontSize={9}/>
<YAxis stroke={C.gl2} fontSize={9} tickFormatter={v=>(v/1e6).toFixed(1)+"jt"}/>
<Tooltip contentStyle={{background:C.card,border:"1px solid "+C.bdr,color:C.wht,fontSize:11}} formatter={v=>fR(v)}/>
<Area type="monotone" dataKey="total" stroke="#EF4444" fill="#EF4444" fillOpacity={0.25} name="Total Pengeluaran"/>
</AreaChart></ResponsiveContainer>
</Card>

{/* Detail transaksi */}
{penBlnAll.length>0&&<Card>
<div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>📋 Detail Pengeluaran</div>
<RTbl headers={["Tgl","Kategori","Keterangan","Karyawan","Nominal"]} rows={penBlnAll.slice().sort((a,b)=>(b.tanggal||"").localeCompare(a.tanggal||"")).slice(0,100).map(p=>[
fDs(p.tanggal),
<Bdg color="red">{p.kategori}</Bdg>,
<span style={{color:C.gl2,fontSize:11}}>{p.ket||"-"}</span>,
p.karyawanNama||"-",
<b style={{color:C.rlt}}>{fR(p.nominal)}</b>
])}/>
</Card>}
</div>}

{tab==="grafik"&&<div>
{mode!=="bulanan"?<Card><div style={{color:C.gl2,fontSize:13,textAlign:"center",padding:20}}>📈 Grafik hanya tersedia untuk mode Bulanan</div></Card>:
!chartData.some(d=>d.omzet>0)?<Card><div style={{color:C.gl2,fontSize:13,textAlign:"center",padding:20}}>Belum ada data penjualan bulan ini</div></Card>:<>

{/* Ringkasan bulanan */}
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(140px,1fr))",gap:8,marginBottom:14}}>
{[["Total Omzet",omzet,C.wht,"📈"],["HPP/Modal DO",doBlnLap.reduce((a,d)=>a+Number(d.totalHPP||0),0)+doTripBlnLap.reduce((a,d)=>a+Number(d.totalHPP||0),0),C.olt,"🏭"],["Margin Kotor",margin,C.blt,"💹"],["Pengeluaran Ops",pengeluaran,C.rlt,"💸"],["Laba Bersih",labaBersih,labaBersih>=0?C.glt:C.rlt,"🏆"],["Cash Flow Akhir",chartData[chartData.length-1]?.cashFlow||0,(chartData[chartData.length-1]?.cashFlow||0)>=0?C.glt:C.rlt,"💰"]].map(x=><div key={x[0]} style={{background:C.card,borderRadius:8,padding:"10px 12px",border:"1px solid "+C.bdr}}><div style={{fontSize:9,color:C.gl2,marginBottom:2}}>{x[3]} {x[0]}</div><div style={{fontSize:13,fontWeight:800,color:x[2]}}>{fR(x[1])}</div></div>)}
</div>

{/* Grafik 1: Omzet & HPP */}
<Card><div style={{fontWeight:700,color:C.gl2,marginBottom:8,fontSize:12}}>📈 1. Omzet Penjualan & Modal HPP per Hari</div>
<ResponsiveContainer width="100%" height={200}><AreaChart data={chartData} margin={{top:4,right:10,bottom:0,left:10}}>
<CartesianGrid strokeDasharray="3 3" stroke={C.bdr}/>
<XAxis dataKey="tgl" stroke={C.gl2} fontSize={9}/>
<YAxis stroke={C.gl2} fontSize={9} tickFormatter={v=>(v/1e6).toFixed(1)+"jt"}/>
<Tooltip contentStyle={{background:C.card,border:"1px solid "+C.bdr,color:C.wht,fontSize:11}} formatter={(v,n)=>[fR(v),n]}/>
<Legend wrapperStyle={{fontSize:11,color:C.gl2}}/>
<Area type="monotone" dataKey="omzet" stroke="#3B82F6" fill="#3B82F6" fillOpacity={0.25} name="Omzet"/>
<Area type="monotone" dataKey="hpp" stroke="#F59E0B" fill="#F59E0B" fillOpacity={0.2} name="HPP/Modal DO"/>
</AreaChart></ResponsiveContainer></Card>

{/* Grafik 2: Margin vs Pengeluaran */}
<Card style={{marginTop:10}}><div style={{fontWeight:700,color:C.gl2,marginBottom:8,fontSize:12}}>💹 2. Margin Kotor vs Pengeluaran Operasional per Hari</div>
<ResponsiveContainer width="100%" height={200}><BarChart data={chartData} margin={{top:4,right:10,bottom:0,left:10}}>
<CartesianGrid strokeDasharray="3 3" stroke={C.bdr}/>
<XAxis dataKey="tgl" stroke={C.gl2} fontSize={9}/>
<YAxis stroke={C.gl2} fontSize={9} tickFormatter={v=>(v/1e6).toFixed(1)+"jt"}/>
<Tooltip contentStyle={{background:C.card,border:"1px solid "+C.bdr,color:C.wht,fontSize:11}} formatter={(v,n)=>[fR(v),n]}/>
<Legend wrapperStyle={{fontSize:11,color:C.gl2}}/>
<Bar dataKey="marginKotor" fill="#3B82F6" name="Margin Kotor" opacity={0.85}/>
<Bar dataKey="pengeluaran" fill="#EF4444" name="Pengeluaran Ops" opacity={0.85}/>
</BarChart></ResponsiveContainer></Card>

{/* Grafik 3: Laba Bersih per hari */}
<Card style={{marginTop:10}}><div style={{fontWeight:700,color:C.gl2,marginBottom:8,fontSize:12}}>🏆 3. Laba Bersih per Hari (bar hijau=untung, merah=rugi)</div>
<ResponsiveContainer width="100%" height={180}><BarChart data={chartData} margin={{top:4,right:10,bottom:0,left:10}}>
<CartesianGrid strokeDasharray="3 3" stroke={C.bdr}/>
<XAxis dataKey="tgl" stroke={C.gl2} fontSize={9}/>
<YAxis stroke={C.gl2} fontSize={9} tickFormatter={v=>(v/1e6).toFixed(1)+"jt"}/>
<Tooltip contentStyle={{background:C.card,border:"1px solid "+C.bdr,color:C.wht,fontSize:11}} formatter={(v,n)=>[fR(v),n]}/>
{chartData.map((entry,index)=><Cell key={index} fill={entry.labaBersih>=0?"#22C55E":"#EF4444"}/>)}
<Bar dataKey="labaBersih" name="Laba Bersih">{chartData.map((entry,index)=><Cell key={index} fill={entry.labaBersih>=0?"#22C55E":"#EF4444"}/>)}</Bar>
</BarChart></ResponsiveContainer></Card>

{/* Grafik 4: Cash Flow Kumulatif — PALING PENTING */}
<Card style={{marginTop:10,border:"2px solid "+C.glt}}><div style={{fontWeight:700,color:C.glt,marginBottom:8,fontSize:12}}>💰 4. Cash Flow Kumulatif Bulanan (PALING PENTING)</div>
<div style={{fontSize:11,color:C.gl2,marginBottom:8}}>Akumulasi laba bersih dari hari 1 sampai akhir bulan — naik = bisnis sehat</div>
<ResponsiveContainer width="100%" height={220}><AreaChart data={chartData} margin={{top:4,right:10,bottom:0,left:10}}>
<CartesianGrid strokeDasharray="3 3" stroke={C.bdr}/>
<XAxis dataKey="tgl" stroke={C.gl2} fontSize={9}/>
<YAxis stroke={C.gl2} fontSize={9} tickFormatter={v=>(v/1e6).toFixed(1)+"jt"}/>
<Tooltip contentStyle={{background:C.card,border:"1px solid "+C.bdr,color:C.wht,fontSize:11}} formatter={(v,n)=>[fR(v),n]}/>
<defs><linearGradient id="cfGrad" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor="#22C55E" stopOpacity={0.4}/><stop offset="95%" stopColor="#22C55E" stopOpacity={0}/></linearGradient></defs>
<Area type="monotone" dataKey="cashFlow" stroke="#22C55E" fill="url(#cfGrad)" strokeWidth={2} name="Cash Flow Kumulatif"/>
<ReferenceLine y={0} stroke={C.rlt} strokeDasharray="4 4"/>
</AreaChart></ResponsiveContainer></Card>

{/* Grafik 5: Top 10 Konsumen */}
{plgArr.length>0&&<Card style={{marginTop:10}}><div style={{fontWeight:700,color:C.gl2,marginBottom:8,fontSize:12}}>👥 5. Top 10 Pelanggan bulan ini</div>
<ResponsiveContainer width="100%" height={220}><BarChart layout="vertical" data={plgArr.slice().sort((a,b)=>b.omzet-a.omzet).slice(0,10)} margin={{top:4,right:60,bottom:0,left:10}}>
<CartesianGrid strokeDasharray="3 3" stroke={C.bdr}/>
<XAxis type="number" stroke={C.gl2} fontSize={9} tickFormatter={v=>(v/1e6).toFixed(1)+"jt"}/>
<YAxis type="category" dataKey="nama" stroke={C.gl2} fontSize={9} width={100}/>
<Tooltip contentStyle={{background:C.card,border:"1px solid "+C.bdr,color:C.wht,fontSize:11}} formatter={v=>fR(v)}/>
<Bar dataKey="omzet" fill="#3B82F6" name="Omzet" radius={[0,4,4,0]}/>
</BarChart></ResponsiveContainer></Card>}

{/* Grafik 6: Per Sales */}
{salesArr.length>0&&<Card style={{marginTop:10}}><div style={{fontWeight:700,color:C.gl2,marginBottom:8,fontSize:12}}>👤 6. Omzet per Sales</div>
<ResponsiveContainer width="100%" height={Math.max(160,salesArr.length*45)}><BarChart layout="vertical" data={salesArr.slice().sort((a,b)=>b.omzet-a.omzet)} margin={{top:4,right:60,bottom:0,left:10}}>
<CartesianGrid strokeDasharray="3 3" stroke={C.bdr}/>
<XAxis type="number" stroke={C.gl2} fontSize={9} tickFormatter={v=>(v/1e6).toFixed(1)+"jt"}/>
<YAxis type="category" dataKey="nama" stroke={C.gl2} fontSize={9} width={100}/>
<Tooltip contentStyle={{background:C.card,border:"1px solid "+C.bdr,color:C.wht,fontSize:11}} formatter={v=>fR(v)}/>
<Bar dataKey="omzet" fill="#8B5CF6" name="Omzet" radius={[0,4,4,0]}/>
<Bar dataKey="margin" fill="#22C55E" name="Margin" radius={[0,4,4,0]}/>
</BarChart></ResponsiveContainer></Card>}

</>}
</div>}
{tab==="bukukas"&&<div>
{mode!=="bulanan"?<Card><div style={{textAlign:"center",padding:30,color:C.gl2,fontSize:12}}>📔 Buku Kas Harian hanya tersedia di mode Bulanan. Ubah filter di atas ke "📆 Bulanan".</div></Card>:(()=>{
var bkRows=buildBukuKasHarian(data,bln);
var totalRow=bkRows.reduce((a,r)=>({
  doQty:{"5.5 kg":a.doQty["5.5 kg"]+r.doQty["5.5 kg"],"12 kg":a.doQty["12 kg"]+r.doQty["12 kg"],"50 kg":a.doQty["50 kg"]+r.doQty["50 kg"]},
  doNilai:a.doNilai+r.doNilai,
  penjQty:{"5.5 kg":a.penjQty["5.5 kg"]+r.penjQty["5.5 kg"],"12 kg":a.penjQty["12 kg"]+r.penjQty["12 kg"],"50 kg":a.penjQty["50 kg"]+r.penjQty["50 kg"]},
  totalPenjualan:a.totalPenjualan+r.totalPenjualan,marginKotor:a.marginKotor+r.marginKotor,kasTF:a.kasTF+r.kasTF,
  bonBaru:a.bonBaru+r.bonBaru,bayarBon:a.bayarBon+r.bayarBon,pengeluaranOps:a.pengeluaranOps+r.pengeluaranOps,
  labaBersih:a.labaBersih+r.labaBersih,estimasiPPN:a.estimasiPPN+r.estimasiPPN
}),{doQty:{"5.5 kg":0,"12 kg":0,"50 kg":0},doNilai:0,penjQty:{"5.5 kg":0,"12 kg":0,"50 kg":0},totalPenjualan:0,marginKotor:0,kasTF:0,bonBaru:0,bayarBon:0,pengeluaranOps:0,labaBersih:0,estimasiPPN:0});

function exportBukuKas(){
var wb=XLSX.utils.book_new();
var header=["Hari","Tanggal","DO 5,5kg","DO 12kg","DO 50kg","Pembelian DO (Modal)","Penjualan 5,5kg","Penjualan 12kg","Penjualan 50kg","Total Penjualan","Margin Kotor/Laba Kotor","Penerimaan Kas & TF","BON Baru (Piutang)","Bayar BON (Piutang Dilunasi)","Pengeluaran Operasional","Laba Bersih","Estimasi PPN (internal)"];
var aoa=[header,...bkRows.map(r=>[r.dayName,fDs(r.tgl),r.doQty["5.5 kg"]||"",r.doQty["12 kg"]||"",r.doQty["50 kg"]||"",r.doNilai||"",r.penjQty["5.5 kg"]||"",r.penjQty["12 kg"]||"",r.penjQty["50 kg"]||"",r.totalPenjualan||"",r.marginKotor||"",r.kasTF||"",r.bonBaru||"",r.bayarBon||"",r.pengeluaranOps||"",r.labaBersih||"",r.estimasiPPN||""])];
aoa.push(["Jumlah","",totalRow.doQty["5.5 kg"],totalRow.doQty["12 kg"],totalRow.doQty["50 kg"],totalRow.doNilai,totalRow.penjQty["5.5 kg"],totalRow.penjQty["12 kg"],totalRow.penjQty["50 kg"],totalRow.totalPenjualan,totalRow.marginKotor,totalRow.kasTF,totalRow.bonBaru,totalRow.bayarBon,totalRow.pengeluaranOps,totalRow.labaBersih,totalRow.estimasiPPN]);
var ws=XLSX.utils.aoa_to_sheet(aoa);
ws["!cols"]=header.map(()=>({wch:14}));
XLSX.utils.book_append_sheet(wb,ws,"Buku Kas Harian");
XLSX.writeFile(wb,"Buku_Kas_Harian_"+bln+".xlsx");
toast("✓ Excel Buku Kas Harian didownload!");
}

return <>
<div style={{fontSize:11,color:C.gl2,marginBottom:12,fontStyle:"italic"}}>Semua angka otomatis dari data sistem (DO, Penjualan, BON, Pengeluaran). "Estimasi PPN (internal)" adalah perkiraan 11% dari margin kotor untuk catatan internal — bukan angka resmi pajak, konsultasikan ke akuntan/konsultan pajak untuk pelaporan resmi.</div>
<div style={{display:"flex",justifyContent:"flex-end",marginBottom:10}}>
<button onClick={exportBukuKas} style={{background:"#15803D",color:"white",border:"none",padding:"8px 16px",borderRadius:8,fontSize:12,cursor:"pointer",fontWeight:700}}>📥 Export Excel</button>
</div>
<Card>
<div style={{overflowX:"auto"}}>
<table style={{width:"100%",borderCollapse:"collapse",fontSize:10.5,minWidth:1500}}>
<thead>
<tr style={{background:C.nav}}>
<th rowSpan={2} style={{padding:"6px 8px",border:"1px solid "+C.bdr,color:C.gl2,fontSize:9}}>Hari</th>
<th rowSpan={2} style={{padding:"6px 8px",border:"1px solid "+C.bdr,color:C.gl2,fontSize:9}}>Tanggal</th>
<th colSpan={4} style={{padding:"6px 8px",border:"1px solid "+C.bdr,color:C.wht,fontSize:10,fontWeight:800}}>DO / Pembelian</th>
<th colSpan={4} style={{padding:"6px 8px",border:"1px solid "+C.bdr,color:C.wht,fontSize:10,fontWeight:800}}>Penjualan</th>
<th rowSpan={2} style={{padding:"6px 8px",border:"1px solid "+C.bdr,color:C.gl2,fontSize:9}}>Margin Kotor</th>
<th rowSpan={2} style={{padding:"6px 8px",border:"1px solid "+C.bdr,color:C.gl2,fontSize:9}}>Kas & TF</th>
<th rowSpan={2} style={{padding:"6px 8px",border:"1px solid "+C.bdr,color:C.gl2,fontSize:9}}>BON Baru</th>
<th rowSpan={2} style={{padding:"6px 8px",border:"1px solid "+C.bdr,color:C.gl2,fontSize:9}}>Bayar BON</th>
<th rowSpan={2} style={{padding:"6px 8px",border:"1px solid "+C.bdr,color:C.gl2,fontSize:9}}>Pengeluaran Ops</th>
<th rowSpan={2} style={{padding:"6px 8px",border:"1px solid "+C.bdr,color:C.glt,fontSize:9}}>Laba Bersih</th>
<th rowSpan={2} style={{padding:"6px 8px",border:"1px solid "+C.bdr,color:C.rlt,fontSize:9}}>Estimasi PPN</th>
</tr>
<tr style={{background:C.nav}}>
<th style={{padding:"4px 6px",border:"1px solid "+C.bdr,color:C.gl2,fontSize:8.5}}>5,5kg</th>
<th style={{padding:"4px 6px",border:"1px solid "+C.bdr,color:C.gl2,fontSize:8.5}}>12kg</th>
<th style={{padding:"4px 6px",border:"1px solid "+C.bdr,color:C.gl2,fontSize:8.5}}>50kg</th>
<th style={{padding:"4px 6px",border:"1px solid "+C.bdr,color:C.gl2,fontSize:8.5}}>Modal (Rp)</th>
<th style={{padding:"4px 6px",border:"1px solid "+C.bdr,color:C.gl2,fontSize:8.5}}>5,5kg</th>
<th style={{padding:"4px 6px",border:"1px solid "+C.bdr,color:C.gl2,fontSize:8.5}}>12kg</th>
<th style={{padding:"4px 6px",border:"1px solid "+C.bdr,color:C.gl2,fontSize:8.5}}>50kg</th>
<th style={{padding:"4px 6px",border:"1px solid "+C.bdr,color:C.gl2,fontSize:8.5}}>Total (Rp)</th>
</tr>
</thead>
<tbody>
{bkRows.map((r,i)=><tr key={r.tgl} style={{background:r.dayName==="Minggu"?(C.mode==="dark"?"#2A1515":"#FEE2E2"):i%2===0?C.bg:C.nav,opacity:r.adaData?1:.45}}>
<td style={{padding:"4px 6px",border:"1px solid "+C.bdr,color:C.wht,fontSize:9.5}}>{r.dayName}</td>
<td style={{padding:"4px 6px",border:"1px solid "+C.bdr,color:C.wht,fontSize:9.5,whiteSpace:"nowrap"}}>{fDs(r.tgl)}</td>
<td style={{padding:"4px 6px",border:"1px solid "+C.bdr,textAlign:"right",color:r.doQty["5.5 kg"]?C.glt:C.gry}}>{r.doQty["5.5 kg"]||"-"}</td>
<td style={{padding:"4px 6px",border:"1px solid "+C.bdr,textAlign:"right",color:r.doQty["12 kg"]?C.glt:C.gry}}>{r.doQty["12 kg"]||"-"}</td>
<td style={{padding:"4px 6px",border:"1px solid "+C.bdr,textAlign:"right",color:r.doQty["50 kg"]?C.glt:C.gry}}>{r.doQty["50 kg"]||"-"}</td>
<td style={{padding:"4px 6px",border:"1px solid "+C.bdr,textAlign:"right",color:r.doNilai?C.olt:C.gry,fontWeight:r.doNilai?700:400}}>{r.doNilai?fR(r.doNilai).replace("Rp ",""):"-"}</td>
<td style={{padding:"4px 6px",border:"1px solid "+C.bdr,textAlign:"right",color:r.penjQty["5.5 kg"]?C.wht:C.gry}}>{r.penjQty["5.5 kg"]||"-"}</td>
<td style={{padding:"4px 6px",border:"1px solid "+C.bdr,textAlign:"right",color:r.penjQty["12 kg"]?C.wht:C.gry}}>{r.penjQty["12 kg"]||"-"}</td>
<td style={{padding:"4px 6px",border:"1px solid "+C.bdr,textAlign:"right",color:r.penjQty["50 kg"]?C.wht:C.gry}}>{r.penjQty["50 kg"]||"-"}</td>
<td style={{padding:"4px 6px",border:"1px solid "+C.bdr,textAlign:"right",color:r.totalPenjualan?C.wht:C.gry,fontWeight:r.totalPenjualan?700:400}}>{r.totalPenjualan?fR(r.totalPenjualan).replace("Rp ",""):"-"}</td>
<td style={{padding:"4px 6px",border:"1px solid "+C.bdr,textAlign:"right",color:r.marginKotor?C.blt:C.gry}}>{r.marginKotor?fR(r.marginKotor).replace("Rp ",""):"-"}</td>
<td style={{padding:"4px 6px",border:"1px solid "+C.bdr,textAlign:"right",color:r.kasTF?C.glt:C.gry}}>{r.kasTF?fR(r.kasTF).replace("Rp ",""):"-"}</td>
<td style={{padding:"4px 6px",border:"1px solid "+C.bdr,textAlign:"right",color:r.bonBaru?C.olt:C.gry}}>{r.bonBaru?fR(r.bonBaru).replace("Rp ",""):"-"}</td>
<td style={{padding:"4px 6px",border:"1px solid "+C.bdr,textAlign:"right",color:r.bayarBon?C.glt:C.gry}}>{r.bayarBon?fR(r.bayarBon).replace("Rp ",""):"-"}</td>
<td style={{padding:"4px 6px",border:"1px solid "+C.bdr,textAlign:"right",color:r.pengeluaranOps?C.rlt:C.gry}}>{r.pengeluaranOps?fR(r.pengeluaranOps).replace("Rp ",""):"-"}</td>
<td style={{padding:"4px 6px",border:"1px solid "+C.bdr,textAlign:"right",color:r.labaBersih>0?C.glt:r.labaBersih<0?C.rlt:C.gry,fontWeight:700}}>{r.labaBersih?fR(r.labaBersih).replace("Rp ",""):"-"}</td>
<td style={{padding:"4px 6px",border:"1px solid "+C.bdr,textAlign:"right",color:r.estimasiPPN?C.rlt:C.gry,fontStyle:"italic"}}>{r.estimasiPPN?fR(Math.round(r.estimasiPPN)).replace("Rp ",""):"-"}</td>
</tr>)}
</tbody>
<tfoot>
<tr style={{background:C.olt}}>
<td colSpan={2} style={{padding:"8px 6px",border:"1px solid "+C.bdr,color:"white",fontWeight:800,fontSize:11}}>JUMLAH</td>
<td style={{padding:"8px 6px",border:"1px solid "+C.bdr,textAlign:"right",color:"white",fontWeight:900}}>{totalRow.doQty["5.5 kg"]}</td>
<td style={{padding:"8px 6px",border:"1px solid "+C.bdr,textAlign:"right",color:"white",fontWeight:900}}>{totalRow.doQty["12 kg"]}</td>
<td style={{padding:"8px 6px",border:"1px solid "+C.bdr,textAlign:"right",color:"white",fontWeight:900}}>{totalRow.doQty["50 kg"]}</td>
<td style={{padding:"8px 6px",border:"1px solid "+C.bdr,textAlign:"right",color:"white",fontWeight:700,fontSize:9.5}}>{fR(totalRow.doNilai).replace("Rp ","")}</td>
<td style={{padding:"8px 6px",border:"1px solid "+C.bdr,textAlign:"right",color:"white",fontWeight:900}}>{totalRow.penjQty["5.5 kg"]}</td>
<td style={{padding:"8px 6px",border:"1px solid "+C.bdr,textAlign:"right",color:"white",fontWeight:900}}>{totalRow.penjQty["12 kg"]}</td>
<td style={{padding:"8px 6px",border:"1px solid "+C.bdr,textAlign:"right",color:"white",fontWeight:900}}>{totalRow.penjQty["50 kg"]}</td>
<td style={{padding:"8px 6px",border:"1px solid "+C.bdr,textAlign:"right",color:"white",fontWeight:900,fontSize:10}}>{fR(totalRow.totalPenjualan).replace("Rp ","")}</td>
<td style={{padding:"8px 6px",border:"1px solid "+C.bdr,textAlign:"right",color:"white",fontWeight:900,fontSize:10}}>{fR(totalRow.marginKotor).replace("Rp ","")}</td>
<td style={{padding:"8px 6px",border:"1px solid "+C.bdr,textAlign:"right",color:"white",fontWeight:700,fontSize:9.5}}>{fR(totalRow.kasTF).replace("Rp ","")}</td>
<td style={{padding:"8px 6px",border:"1px solid "+C.bdr,textAlign:"right",color:"white",fontWeight:700,fontSize:9.5}}>{fR(totalRow.bonBaru).replace("Rp ","")}</td>
<td style={{padding:"8px 6px",border:"1px solid "+C.bdr,textAlign:"right",color:"white",fontWeight:700,fontSize:9.5}}>{fR(totalRow.bayarBon).replace("Rp ","")}</td>
<td style={{padding:"8px 6px",border:"1px solid "+C.bdr,textAlign:"right",color:"white",fontWeight:700,fontSize:9.5}}>{fR(totalRow.pengeluaranOps).replace("Rp ","")}</td>
<td style={{padding:"8px 6px",border:"1px solid "+C.bdr,textAlign:"right",color:"white",fontWeight:900,fontSize:10.5}}>{fR(totalRow.labaBersih).replace("Rp ","")}</td>
<td style={{padding:"8px 6px",border:"1px solid "+C.bdr,textAlign:"right",color:"white",fontWeight:700,fontSize:9.5}}>{fR(Math.round(totalRow.estimasiPPN)).replace("Rp ","")}</td>
</tr>
</tfoot>
</table>
</div>
</Card>
</>;
})()}
</div>}

</div>;
}

// ─── KARYAWAN & AMBILAN ───────────────────────────────────────────────────────
function KaryawanMod({data,setData,toast}){
var C=useTheme();
var blk={username:"",password:"",role:"sales_driver",nama:"",posisi:"Sales Driver",telepon:"",alamat:"",gajiPokok:"",uangMakan:"15000",uangMakanMode:"harian",aktif:true};
var[f,setF]=useState({...blk});var[edit,setEdit]=useState(null);var[delId,setDelId]=useState(null);
var posOpts=["Owner/Komisaris","Manajer","Admin","Kasir/Akuntan","Sales Driver","Sales Freelance","Sales Marketing","Checker","Driver Truck SPBE","Helper"];
function save(){if(!f.nama||!f.username)return;if(edit){setData(d=>({...d,employees:(d.employees||[]).map(e=>e.id===edit.id?{...e,...f,gajiPokok:Number(f.gajiPokok||0),uangMakan:Number(f.uangMakan||15000)}:e)}));setEdit(null);}else setData(d=>({...d,employees:[{id:uid(),...f,gajiPokok:Number(f.gajiPokok||0),uangMakan:Number(f.uangMakan||15000)},...(d.employees||[])]}));setF({...blk});toast("✓ Karyawan disimpan!");}
return <div>
<STitle icon="👤" children="Karyawan & Akun"/>
<Card style={{width:"fit-content",maxWidth:"100%",minWidth:660}}><div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(155px,210px))",gap:10}}><Inp label="Nama" value={f.nama} onChange={v=>setF(p=>({...p,nama:v}))}/><Sel label="Posisi" value={f.posisi} onChange={v=>setF(p=>({...p,posisi:v}))} opts={posOpts}/><Sel label="Role" value={f.role} onChange={v=>setF(p=>({...p,role:v}))} opts={Object.keys(ROLE_LBL).map(k=>({v:k,l:ROLE_LBL[k]}))}/><Inp label="Username" value={f.username} onChange={v=>setF(p=>({...p,username:v}))}/><Inp label="Password" type="password" value={f.password} onChange={v=>setF(p=>({...p,password:v}))}/><Inp label="Telepon" value={f.telepon} onChange={v=>setF(p=>({...p,telepon:v}))}/><Inp label="Alamat" value={f.alamat} onChange={v=>setF(p=>({...p,alamat:v}))}/><Inp label="Gaji Pokok" type="number" value={f.gajiPokok} onChange={v=>setF(p=>({...p,gajiPokok:v}))}/><Inp label="Uang Makan/Hari" type="number" value={f.uangMakan} onChange={v=>setF(p=>({...p,uangMakan:v}))}/><Sel label="Mode Uang Makan" value={f.uangMakanMode} onChange={v=>setF(p=>({...p,uangMakanMode:v}))} opts={[{v:"harian",l:"💰 Harian"},{v:"akhir_bulan",l:"📅 Akhir Bulan"}]}/></div><Btn color="green" onClick={save} dis={!f.nama||!f.username}>➕ Tambah Karyawan</Btn></Card>
<Card><RTbl headers={["Nama","Posisi","Role","Status","Absensi","Aksi"]} widths={[190,150,130,100,90,180]} rows={(data.employees||[]).map(e=>[<div><b style={{color:C.wht}}>{e.nama}</b><div style={{fontSize:11,color:C.gl2}}>{e.telepon}</div></div>,e.posisi,<Bdg color={["admin","owner"].includes(e.role)?"red":"blue"}>{ROLE_LBL[e.role]||e.role}</Bdg>,e.aktif?<Bdg color="green">Aktif</Bdg>:<Bdg color="gray">Non-aktif</Bdg>,<div style={{display:"flex",gap:5}}><button onClick={()=>setData(d=>({...d,employees:(d.employees||[]).map(x=>x.id===e.id?{...x,aktif:!x.aktif}:x)}))} style={{background:C.nav,border:"1px solid "+C.bdr,borderRadius:7,padding:"5px 9px",color:C.gl2,cursor:"pointer",fontSize:12}}>{e.aktif?"🔒":"🔓"}</button><button onClick={()=>setData(d=>({...d,employees:(d.employees||[]).map(x=>x.id===e.id?{...x,ikutAbsensi:!x.ikutAbsensi}:x)}))} title="Toggle Absensi" style={{background:e.ikutAbsensi?C.grn:C.nav,border:"1px solid "+(e.ikutAbsensi?C.glt:C.bdr),borderRadius:7,padding:"5px 9px",color:e.ikutAbsensi?C.glt:C.gl2,cursor:"pointer",fontSize:11,fontWeight:700}}>{e.ikutAbsensi?"📅":"—"}</button>
<ActBtns onEdit={()=>{setEdit(e);setF({...e,gajiPokok:String(e.gajiPokok||""),uangMakan:String(e.uangMakan||15000)});}} onDel={()=>setDelId(e)}/></div>])}/></Card>

{edit&&<Modal title={"Edit: "+edit.nama} onSave={save} onClose={()=>{setEdit(null);setF({...blk});}}><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}><Inp label="Nama" value={f.nama} onChange={v=>setF(p=>({...p,nama:v}))}/><Sel label="Posisi" value={f.posisi} onChange={v=>setF(p=>({...p,posisi:v}))} opts={posOpts}/><Sel label="Role" value={f.role} onChange={v=>setF(p=>({...p,role:v}))} opts={Object.keys(ROLE_LBL).map(k=>({v:k,l:ROLE_LBL[k]}))}/><Inp label="Username" value={f.username} onChange={v=>setF(p=>({...p,username:v}))}/><Inp label="Password" type="password" value={f.password} onChange={v=>setF(p=>({...p,password:v}))}/><Inp label="Gaji Pokok" type="number" value={f.gajiPokok} onChange={v=>setF(p=>({...p,gajiPokok:v}))}/><Inp label="Uang Makan" type="number" value={f.uangMakan} onChange={v=>setF(p=>({...p,uangMakan:v}))}/><Inp label="Alamat" value={f.alamat} onChange={v=>setF(p=>({...p,alamat:v}))}/></div><Sel label="Mode Uang Makan" value={f.uangMakanMode} onChange={v=>setF(p=>({...p,uangMakanMode:v}))} opts={[{v:"harian",l:"💰 Harian"},{v:"akhir_bulan",l:"📅 Akhir Bulan"}]}/></Modal>}
{delId&&<ConfirmDel msg={"Hapus \""+delId.nama+"\"?"} onCancel={()=>setDelId(null)} onConfirm={()=>{setData(d=>({...d,employees:(d.employees||[]).filter(x=>x.id!==delId.id)}));setDelId(null);}}/>}
</div>;
}

// ─── ABSENSI ──────────────────────────────────────────────────────────────────
function AbsensiPayrollMod({data,setData,toast}){
var C=useTheme();
var[tabAP,setTabAP]=useState("absensi");
var[viewBln,setViewBln]=useState(toMonth());
var STATUS_CYCLE=["H","I","C","A","L"];// Hadir, Izin, Cuti, Alpha, Libur
var STATUS_COLOR={H:"#15803D",I:"#B45309",C:"#1D4ED8",A:"#DC2626",L:"#6B7280"};
var STATUS_BG={H:"#DCFCE7",I:"#FEF3C7",C:"#DBEAFE",A:"#FEE2E2",L:"#E5E7EB"};
var STATUS_LABEL={H:"Hadir",I:"Izin",C:"Cuti",A:"Alpha",L:"Libur"};
var karList=sortEmp((data.employees||[]).filter(e=>e.aktif));
var karAbsensi=karList.filter(e=>e.ikutAbsensi);
var dim=daysInMonth(viewBln);
var days=Array.from({length:dim},(_,i)=>String(i+1).padStart(2,"0"));
function isMinggu(tgl){return new Date(viewBln+"-"+tgl+"T00:00:00").getDay()===0;}

function toggleAbsensi(karyawanId,tgl,current){
var next=current?STATUS_CYCLE[(STATUS_CYCLE.indexOf(current)+1)%STATUS_CYCLE.length]:STATUS_CYCLE[0];
var emp=(data.employees||[]).find(e=>e.id===karyawanId);
var fullTgl=viewBln+"-"+tgl;
var exists=(data.absensi||[]).find(a=>a.karyawanId===karyawanId&&a.tanggal===fullTgl);
if(next===null){// hapus
setData(d=>({...d,absensi:(d.absensi||[]).filter(a=>!(a.karyawanId===karyawanId&&a.tanggal===fullTgl))}));
}else{
var rec={id:exists?.id||uid(),karyawanId,karyawanNama:emp?.nama||"",tanggal:fullTgl,status:next,ket:""};
setData(d=>({...d,absensi:exists?(d.absensi||[]).map(a=>a.id===exists.id?rec:a):[rec,...(d.absensi||[])]}));
}
}

// Hapus langsung jadi kosong (klik kanan / tombol reset) — untuk salah klik
function clearAbsensi(karyawanId,tgl){
var fullTgl=viewBln+"-"+tgl;
setData(d=>({...d,absensi:(d.absensi||[]).filter(a=>!(a.karyawanId===karyawanId&&a.tanggal===fullTgl))}));
}

// Hitung rekap per karyawan
var rekapAbsensi=karAbsensi.map(emp=>{
var absBln=(data.absensi||[]).filter(a=>a.karyawanId===emp.id&&(a.tanggal||"").startsWith(viewBln));
var counts={H:0,I:0,C:0,A:0,L:0};
absBln.forEach(a=>{if(counts[a.status]!==undefined)counts[a.status]++;});
return{...emp,counts,absBln};
});

// Rekap ambilan dari pengeluaran
var ambilanList=(data.pengeluaran||[]).filter(p=>(p.kategori||"").toLowerCase().includes("kasbon")||(p.kategori||"").toLowerCase().includes("ambilan")).sort((a,b)=>b.tanggal.localeCompare(a.tanggal));

return <div>
<STitle icon="📅" children="Absensi & Payroll"/>
<div style={{display:"flex",gap:5,marginBottom:14,flexWrap:"wrap"}}>
{[["absensi","📅 Absensi"],["payroll","💰 Payroll"],["ambilan","💸 Rekap Ambilan"]].map(x=><button key={x[0]} onClick={()=>setTabAP(x[0])} style={{background:tabAP===x[0]?C.blu:C.nav,color:tabAP===x[0]?"white":C.wht,border:"1px solid "+(tabAP===x[0]?C.blt:C.bdr),borderRadius:8,padding:"7px 14px",fontWeight:700,fontSize:12,cursor:"pointer"}}>{x[1]}</button>)}
</div>

{/* ── TAB ABSENSI ── */}
{tabAP==="absensi"&&<div>
<Card style={{padding:"10px 12px",marginBottom:8}}>
<div style={{display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
<MonthPicker value={viewBln} onChange={setViewBln} label="Bulan"/>
<div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
{Object.entries(STATUS_LABEL).map(([k,v])=><div key={k} style={{display:"flex",alignItems:"center",gap:4}}>
<div style={{width:20,height:20,background:STATUS_BG[k],border:"1px solid "+STATUS_COLOR[k],borderRadius:4,display:"flex",alignItems:"center",justifyContent:"center",fontSize:10,fontWeight:700,color:STATUS_COLOR[k]}}>{k}</div>
<span style={{fontSize:11,color:C.gl2}}>{v}</span>
</div>)}
<span style={{fontSize:10,color:C.gl2,fontStyle:"italic"}}>Klik kiri: ubah status (H→I→C→A→L) &nbsp;|&nbsp; Klik kanan: kosongkan (salah klik)</span>
</div>
</div>
</Card>
<Card style={{padding:0,overflow:"hidden"}}>
<div style={{overflowX:"auto"}}>
<table style={{width:"100%",borderCollapse:"collapse",fontSize:10,minWidth:700}}>
<thead>
<tr style={{background:C.nav,borderBottom:"2px solid "+C.bdr}}>
<th style={{padding:"8px 12px",color:C.gl2,textAlign:"left",fontWeight:700,fontSize:11,position:"sticky",left:0,background:C.nav,minWidth:120,zIndex:1}}>Karyawan</th>
{days.map(d=>{var minggu=isMinggu(d);return <th key={d} style={{padding:"6px 3px",color:minggu?"#DC2626":C.gl2,textAlign:"center",fontWeight:700,minWidth:28,borderLeft:"1px solid "+C.bdr,background:minggu?"rgba(220,38,38,.12)":"transparent"}}>{Number(d)}</th>;})}
<th style={{padding:"6px 8px",color:"#15803D",textAlign:"center",fontWeight:700,borderLeft:"2px solid "+C.bdr,background:C.nav}}>H</th>
<th style={{padding:"6px 8px",color:"#B45309",textAlign:"center",fontWeight:700}}>I</th>
<th style={{padding:"6px 8px",color:"#1D4ED8",textAlign:"center",fontWeight:700}}>C</th>
<th style={{padding:"6px 8px",color:"#DC2626",textAlign:"center",fontWeight:700}}>A</th>
<th style={{padding:"6px 8px",color:"#6B7280",textAlign:"center",fontWeight:700}}>L</th>
</tr>
</thead>
<tbody>
{rekapAbsensi.map((emp,ri)=>{
return <tr key={emp.id} style={{borderBottom:"1px solid "+C.bdr,background:ri%2===0?C.bg:C.nav}}>
<td style={{padding:"6px 12px",color:C.wht,fontWeight:600,position:"sticky",left:0,background:ri%2===0?C.bg:C.nav,zIndex:1}}>{emp.nama}</td>
{days.map(d=>{
var fullTgl=viewBln+"-"+d;
var abs=emp.absBln.find(a=>a.tanggal===fullTgl);
var st=abs?.status||null;
var minggu=isMinggu(d);
return <td key={d} onClick={()=>toggleAbsensi(emp.id,d,st)} onContextMenu={e=>{e.preventDefault();if(st)clearAbsensi(emp.id,d);}} title={minggu?"Minggu — klik kanan untuk kosongkan":"klik kanan untuk kosongkan"} style={{padding:"3px 2px",textAlign:"center",cursor:"pointer",borderLeft:"1px solid "+C.bdr,background:st?STATUS_BG[st]:(minggu?"rgba(220,38,38,.06)":"transparent")}}>
{st&&<div style={{width:22,height:22,margin:"0 auto",background:STATUS_COLOR[st],borderRadius:4,display:"flex",alignItems:"center",justifyContent:"center",fontSize:9,fontWeight:700,color:"white"}}>{st}</div>}
{!st&&<div style={{width:22,height:22,margin:"0 auto",borderRadius:4,border:"1px dashed "+(minggu?"#DC2626":C.bdr)}}/>}
</td>;
})}
<td style={{padding:"6px 8px",textAlign:"center",fontWeight:700,color:"#15803D",borderLeft:"2px solid "+C.bdr,fontSize:12}}>{emp.counts.H||"-"}</td>
<td style={{padding:"6px 8px",textAlign:"center",fontWeight:700,color:"#B45309",fontSize:12}}>{emp.counts.I||"-"}</td>
<td style={{padding:"6px 8px",textAlign:"center",fontWeight:700,color:"#1D4ED8",fontSize:12}}>{emp.counts.C||"-"}</td>
<td style={{padding:"6px 8px",textAlign:"center",fontWeight:700,color:"#DC2626",fontSize:12}}>{emp.counts.A||"-"}</td>
<td style={{padding:"6px 8px",textAlign:"center",fontWeight:700,color:"#6B7280",fontSize:12}}>{emp.counts.L||"-"}</td>
</tr>;
})}
</tbody>
</table>
</div>
</Card>
</div>}

{/* ── TAB PAYROLL ── */}
{tabAP==="payroll"&&<PayrollMod data={data} setData={setData} toast={toast}/>}

{/* ── TAB REKAP AMBILAN ── */}
{tabAP==="ambilan"&&(()=>{
// Sumber 1: dari pengeluaran (kasbon/ambilan)
var ambPen=(data.pengeluaran||[]).filter(p=>(p.kategori||"").toLowerCase().includes("kasbon")||(p.kategori||"").toLowerCase().includes("ambilan")).map(p=>({tgl:p.tanggal,nama:p.karyawanNama||"—",empId:p.karyawanId||"",nominal:Number(p.nominal||0),metode:p.metode||"Cash",ket:p.ket||"—",sumber:"Pengeluaran"}));
// Sumber 2: dari kurang setor (data.ambilan)
var ambKurang=(data.ambilan||[]).map(a=>({tgl:a.tanggal,nama:a.karyawanNama||"—",empId:a.karyawanId||"",nominal:Number(a.nominal||0),metode:"Cash",ket:a.ket||"Kurang Setor",sumber:"Kurang Setor"}));
// Gabung dan sort
var allAmb=[...ambPen,...ambKurang].sort((a,b)=>b.tgl.localeCompare(a.tgl));
// Rekap per karyawan — pakai fungsi yang SAMA dengan Payroll
var rekapKar={};
(data.employees||[]).filter(e=>e.aktif).forEach(e=>{
var total=getTotalAmbilanKaryawan(e.id,data);
if(total>0)rekapKar[e.id]={nama:e.nama,total};
});
return <div>
<Card>
<div style={{fontWeight:700,color:C.gl2,marginBottom:6,fontSize:13}}>💸 Riwayat Ambilan / Kasbon</div>
<div style={{fontSize:11,color:C.gl2,marginBottom:10,fontStyle:"italic"}}>
  Sumber: 💸 Pengeluaran (Kategori "Kasbon / Ambilan Karyawan") + ⚠️ Kurang Setor
</div>
{allAmb.length===0
?<div style={{color:C.gl2,fontSize:12,fontStyle:"italic"}}>Belum ada data ambilan</div>
:<div style={{overflowX:"auto"}}>
<RTbl headers={["Tgl","Karyawan","Nominal","Metode","Keterangan","Sumber"]} widths={[100,180,150,110,250,140]} rows={allAmb.map(a=>[
fDs(a.tgl),
<b style={{color:C.wht}}>{a.nama}</b>,
<b style={{color:C.rlt,whiteSpace:"nowrap"}}>{fR(a.nominal)}</b>,
<Bdg color={(a.metode||"cash").toLowerCase()==="cash"?"green":"blue"}>{a.metode}</Bdg>,
a.ket,
<Bdg color={a.sumber==="Kurang Setor"?"red":"orange"}>{a.sumber}</Bdg>
])}/>
</div>}
</Card>
{Object.keys(rekapKar).length>0&&<Card>
<div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>👤 Total Ambilan per Karyawan</div>
<div style={{border:"1px solid "+C.bdr,borderRadius:8,overflow:"hidden"}}>
{Object.values(rekapKar).sort((a,b)=>a.nama.localeCompare(b.nama)).map((r,i)=><div key={i} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"10px 14px",background:i%2===0?C.bg:C.nav,borderBottom:"1px solid "+C.bdr}}>
<span style={{fontSize:13,color:C.wht,fontWeight:600}}>{r.nama}</span>
<b style={{fontSize:14,color:C.rlt}}>{fR(r.total)}</b>
</div>)}
<div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"12px 14px",background:C.olt}}>
<span style={{fontSize:13,color:"white",fontWeight:800}}>TOTAL SELURUH KARYAWAN</span>
<b style={{fontSize:16,color:"white"}}>{fR(Object.values(rekapKar).reduce((a,r)=>a+r.total,0))}</b>
</div>
</div>
</Card>}
</div>;
})()}
</div>;
}

// ─── SETTINGS ─────────────────────────────────────────────────────────────────
function SettingsMod({data,setData,toast,theme,setTheme}){
var C=useTheme();
var[f,setF]=useState({...data.company});
var logoRef=useRef();var logoPRef=useRef();var ttdKRef=useRef();var ttdDRef=useRef();var stempelRef=useRef();
var[hP,setHP]=useState({jenis:"Isi",ukuran:"5.5 kg",harga:""});
var[mP,setMP]=useState({jenis:"Isi",ukuran:"5.5 kg",harga:"",tgl:toDay()});
function saveCompany(){setData(d=>({...d,company:{...d.company,...f}}));toast("✓ Disimpan!");}
function handleUpload(e,key,label){var file=e.target.files[0];if(!file)return;if(file.size>2*1024*1024){toast("File terlalu besar (maks 2MB)","error");return;}var reader=new FileReader();reader.onload=ev=>{var dataUrl=ev.target.result;setF(p=>({...p,[key]:dataUrl}));setData(d=>({...d,company:{...d.company,[key]:dataUrl}}));toast("✓ "+label+" diupload!");};reader.readAsDataURL(file);}
function clearUpload(key,label){setF(p=>({...p,[key]:""}));setData(d=>({...d,company:{...d.company,[key]:""}}));toast("✓ "+label+" dihapus");}
function saveHET(){if(!hP.harga)return;setData(d=>{var hp={...d.hetPrices||{}};if(!hp[hP.jenis])hp[hP.jenis]={};hp[hP.jenis][hP.ukuran]=Number(hP.harga);return{...d,hetPrices:hp};});toast("✓ HET diperbarui!");}
function saveModal(){if(!mP.harga)return;setData(d=>({...d,modalHistory:[{id:uid(),tanggal:mP.tgl,jenis:mP.jenis,ukuran:mP.ukuran,harga:Number(mP.harga),sumber:"Manual Settings"},...(d.modalHistory||[])]}));toast("✓ Modal diperbarui!");}
function clearData(){if(window.confirm("⚠️ Hapus semua data transaksi? Tidak bisa dibatalkan!")){{setData(d=>({...d,penjualan:[],bon:[],pengeluaran:[],tutupBuku:[],setoranSales:[],setoranLog:[],stockLog:[],absensi:[],ambilan:[],payrollLog:[],titipList:[],counters:{inv:{},sg:{},reg:0}}));toast("✓ Data transaksi dihapus!");}}}
function backup(){try{var blob=new Blob([JSON.stringify(data,null,2)],{type:"application/json"});var a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download="backup_hoetrangsa_"+toDay()+".json";a.click();toast("✓ Backup didownload!");}catch(e){toast("Gagal","error");}}
function UploadSlot({label,value,refEl,onUpload,onClear,desc}){return <div style={{display:"flex",alignItems:"center",gap:12,padding:10,background:C.nav,borderRadius:8,border:"1px solid "+C.bdr,marginBottom:8}}>
{value?<img src={value} style={{height:60,maxWidth:120,objectFit:"contain",background:"white",borderRadius:4,padding:4}} alt={label}/>:<div style={{height:60,width:80,background:C.bg,borderRadius:6,border:"2px dashed "+C.bdr,display:"flex",alignItems:"center",justifyContent:"center",color:C.gl2,fontSize:10,textAlign:"center"}}>Belum<br/>upload</div>}
<div style={{flex:1}}><div style={{color:C.wht,fontWeight:700,fontSize:12}}>{label}</div><div style={{color:C.gl2,fontSize:10,marginBottom:6}}>{desc}</div><div style={{display:"flex",gap:6}}><input ref={refEl} type="file" accept="image/*" onChange={onUpload} style={{display:"none"}}/><Btn sm color="blue" onClick={()=>refEl.current?.click()}>📁 Upload</Btn>{value&&<Btn sm color="red" onClick={onClear}>🗑️ Hapus</Btn>}</div></div>
</div>;}
return <div>
<STitle icon="⚙️" children="Pengaturan"/>
<Card style={{border:"1px solid "+C.blt}}>
<div style={{fontWeight:700,color:C.blt,marginBottom:10,fontSize:13}}>🎨 Tema Aplikasi</div>
<div style={{display:"flex",gap:8}}>
<button onClick={()=>setTheme("light")} style={{flex:1,padding:"12px",background:theme==="light"?C.blu:C.nav,color:theme==="light"?"white":C.wht,border:"1px solid "+(theme==="light"?C.blt:C.bdr),borderRadius:8,fontWeight:700,fontSize:13,cursor:"pointer"}}>☀️ Terang Pertamina</button>
<button onClick={()=>setTheme("dark")} style={{flex:1,padding:"12px",background:theme==="dark"?C.blu:C.nav,color:theme==="dark"?"white":C.wht,border:"1px solid "+(theme==="dark"?C.blt:C.bdr),borderRadius:8,fontWeight:700,fontSize:13,cursor:"pointer"}}>🌙 Gelap Premium</button>
</div>
</Card>
<Card style={{border:"1px solid "+C.blt}}>
<div style={{fontWeight:700,color:C.blt,marginBottom:6,fontSize:13}}>📐 Lebar Tampilan Desktop</div>
<div style={{fontSize:11,color:C.gl2,marginBottom:12}}>Atur seberapa lebar konten aplikasi memenuhi layar (khusus PC/laptop — tampilan HP tidak terpengaruh). Berlaku permanen di semua device setelah disimpan.</div>
<div style={{display:"flex",alignItems:"center",gap:14}}>
<span style={{fontSize:10,color:C.gl2,whiteSpace:"nowrap"}}>Sempit</span>
<input type="range" min="1000" max="2200" step="50" value={f.appWidth||1600} onChange={e=>{var v=Number(e.target.value);setF(p=>({...p,appWidth:v}));setData(d=>({...d,company:{...d.company,appWidth:v}}));}} style={{flex:1}}/>
<span style={{fontSize:10,color:C.gl2,whiteSpace:"nowrap"}}>Lebar</span>
</div>
<div style={{textAlign:"center",margin:"8px 0 14px",fontSize:16,fontWeight:800,color:C.blt}}>{f.appWidth||1600}px</div>
<Btn color="blue" onClick={saveCompany}>💾 Simpan Lebar Tampilan</Btn>
</Card>
<Card><div style={{fontWeight:700,color:C.gl2,marginBottom:12,fontSize:13}}>🖼️ Logo & Tanda Tangan Elektronik</div>
<div style={{background:C.nav,border:"1px solid "+C.bdr,borderRadius:8,padding:"10px 12px",marginBottom:14,fontSize:11,color:C.gl2}}>ℹ️ Logo PT. Hoe Trang Sa & Pertamina sudah terpasang tetap (tidak perlu upload), otomatis tampil di semua dokumen & device.</div>
<UploadSlot label="TTD Kasir" value={f.ttdKasir} refEl={ttdKRef} onUpload={e=>handleUpload(e,"ttdKasir","TTD Kasir")} onClear={()=>clearUpload("ttdKasir","TTD Kasir")} desc="Tanda tangan di Invoice"/>
<UploadSlot label="TTD Direktur" value={f.ttdDirektur} refEl={ttdDRef} onUpload={e=>handleUpload(e,"ttdDirektur","TTD Direktur")} onClear={()=>clearUpload("ttdDirektur","TTD Direktur")} desc="Tanda tangan di Kwitansi Slip Gaji"/>
<UploadSlot label="Stempel LUNAS" value={f.stempelLunas} refEl={stempelRef} onUpload={e=>handleUpload(e,"stempelLunas","Stempel LUNAS")} onClear={()=>clearUpload("stempelLunas","Stempel LUNAS")} desc="Jika kosong, pakai teks LUNAS default"/>
</Card>
<Card><div style={{fontWeight:700,color:C.gl2,marginBottom:12,fontSize:13}}>🏢 Data Perusahaan</div>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(200px,1fr))",gap:10}}>
{[["nama","Nama Perusahaan"],["slogan","Slogan"],["alamat","Alamat"],["telepon","No. HP / WA"],["telepon2","No. Telp Kantor"],["email","Email"],["website","Website"],["npwp","NPWP"],["direkturNama","Nama Direktur (slip gaji)"],["kasirNama","Nama Kasir (invoice)"]].map(([k,l])=><Inp key={k} label={l} value={f[k]||""} onChange={v=>setF(p=>({...p,[k]:v}))}/>)}
</div>
<Btn color="blue" onClick={saveCompany}>💾 Simpan Data Perusahaan</Btn>
</Card>
<Card><div style={{fontWeight:700,color:C.gl2,marginBottom:12,fontSize:13}}>💳 Rekening Bank (tampil di Invoice)</div>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(180px,1fr))",gap:10}}><Inp label="Nama Bank" value={f.bankNama||""} onChange={v=>setF(p=>({...p,bankNama:v}))}/><Inp label="Atas Nama" value={f.bankAtasNama||""} onChange={v=>setF(p=>({...p,bankAtasNama:v}))}/><Inp label="No. Rekening" value={f.bankRekening||""} onChange={v=>setF(p=>({...p,bankRekening:v}))}/></div>
<Btn color="blue" onClick={saveCompany}>💾 Simpan Rekening</Btn></Card>
<Card><div style={{fontWeight:700,color:C.gl2,marginBottom:12,fontSize:13}}>💲 Update HET</div>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(140px,1fr))",gap:10}}><Sel label="Jenis" value={hP.jenis} onChange={v=>setHP(p=>({...p,jenis:v}))} opts={JENIS}/><Sel label="Ukuran" value={hP.ukuran} onChange={v=>setHP(p=>({...p,ukuran:v}))} opts={SIZES}/><Inp label={"HET ("+fR(getHET(data,hP.ukuran,hP.jenis))+")"} type="number" value={hP.harga} onChange={v=>setHP(p=>({...p,harga:v}))}/></div>
<Btn color="orange" onClick={saveHET} dis={!hP.harga}>💾 Update HET</Btn></Card>
<Card><div style={{fontWeight:700,color:C.gl2,marginBottom:12,fontSize:13}}>📦 Update Modal/HPP Manual</div>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(140px,1fr))",gap:10}}><Sel label="Jenis" value={mP.jenis} onChange={v=>setMP(p=>({...p,jenis:v}))} opts={JENIS}/><Sel label="Ukuran" value={mP.ukuran} onChange={v=>setMP(p=>({...p,ukuran:v}))} opts={SIZES}/><Inp label={"Modal ("+fR(getModal(data,mP.ukuran,mP.jenis))+")"} type="number" value={mP.harga} onChange={v=>setMP(p=>({...p,harga:v}))}/><Inp label="Berlaku Dari" type="date" value={mP.tgl} onChange={v=>setMP(p=>({...p,tgl:v}))}/></div>
<Btn color="green" onClick={saveModal} dis={!mP.harga}>💾 Update Modal</Btn></Card>
<Card><div style={{fontWeight:700,color:C.gl2,marginBottom:6,fontSize:13}}>🏭 HPP Fixed per SPPBE</div>
<div style={{fontSize:11,color:C.gl2,marginBottom:12}}>HPP untuk SPPBE KCR & MGL otomatis terkunci sesuai angka ini saat input DO. SPPBE Lainnya tetap bisa diisi manual per transaksi.</div>
{["SPPBE KCR","SPPBE MGL"].map(sp=><div key={sp} style={{marginBottom:14}}>
<div style={{fontSize:12,fontWeight:700,color:C.olt,marginBottom:8}}>{sp}</div>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(140px,1fr))",gap:10}}>
{SIZES.map(s=><Inp key={s} label={"HPP "+s} type="number" step="1000" value={(f.hppFixedSPPBE&&f.hppFixedSPPBE[sp]&&f.hppFixedSPPBE[sp][s])||""} onChange={v=>setF(p=>({...p,hppFixedSPPBE:{...(p.hppFixedSPPBE||{}),[sp]:{...((p.hppFixedSPPBE||{})[sp]||{}),[s]:Number(v)||0}}}))} placeholder="0"/>)}
</div>
</div>)}
<Btn color="blue" onClick={saveCompany}>💾 Simpan HPP Fixed</Btn>
</Card>
<Card><div style={{fontWeight:700,color:C.gl2,marginBottom:12,fontSize:13}}>🚛 Asset & Harga Tabung</div>
<div style={{fontSize:11,color:C.gl2,marginBottom:10}}>Untuk kalkulasi Asset Tabung Milik PT di Tutup Buku:</div>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(160px,1fr))",gap:10,marginBottom:12}}>
<Inp label="Asset Armada (Rp)" type="number" value={f.assetArmada||""} onChange={v=>setF(p=>({...p,assetArmada:Number(v)||0}))} placeholder="Nilai kendaraan operasional"/>
<Inp label="Harga Tbg Kosong 5,5kg" type="number" value={(f.hargaTbgKosong&&f.hargaTbgKosong["5.5 kg"])||""} onChange={v=>setF(p=>({...p,hargaTbgKosong:{...(p.hargaTbgKosong||{}),["5.5 kg"]:Number(v)||0}}))} placeholder="Harga jual Tbg+Isi - HPP isi"/>
<Inp label="Harga Tbg Kosong 12kg" type="number" value={(f.hargaTbgKosong&&f.hargaTbgKosong["12 kg"])||""} onChange={v=>setF(p=>({...p,hargaTbgKosong:{...(p.hargaTbgKosong||{}),["12 kg"]:Number(v)||0}}))} placeholder="Harga jual Tbg+Isi - HPP isi"/>
<Inp label="Harga Tbg Kosong 50kg" type="number" value={(f.hargaTbgKosong&&f.hargaTbgKosong["50 kg"])||""} onChange={v=>setF(p=>({...p,hargaTbgKosong:{...(p.hargaTbgKosong||{}),["50 kg"]:Number(v)||0}}))} placeholder="Harga jual Tbg+Isi - HPP isi"/>
</div>
<Btn color="blue" onClick={saveCompany}>💾 Simpan</Btn>
</Card>
<Card style={{width:"fit-content",maxWidth:"100%",minWidth:660}}><div style={{fontWeight:700,color:C.gl2,marginBottom:12,fontSize:13}}>🏭 Kode Pertamina</div>
<div style={{fontSize:11,color:C.gl2,marginBottom:10,fontStyle:"italic"}}>🔒 Kode penebusan (permanen — tidak berubah)</div>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(150px,210px))",gap:10,marginBottom:12}}>
{[["soldTo","Sold To"],["shipToKCR","Ship To KCR"],["shipToMGL","Ship To MGL"]].map(([k,l])=><Inp key={k} label={l} value={f[k]||""} onChange={v=>setF(p=>({...p,[k]:v}))}/>)}
</div>
<div style={{fontSize:11,color:C.olt,marginBottom:10,fontStyle:"italic"}}>🔄 Kode SA bulanan (update setiap bulan terima SA baru dari Pertamina)</div>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(150px,210px))",gap:10}}>
<Inp label="Bulan SA Aktif" value={f.saBulan||""} onChange={v=>setF(p=>({...p,saBulan:v}))} placeholder="Juni 2026"/>
<Inp label="SA 12kg — KCR" value={f.sa12KCR||""} onChange={v=>setF(p=>({...p,sa12KCR:v}))}/>
<Inp label="SA 5,5kg — KCR" value={f.sa55KCR||""} onChange={v=>setF(p=>({...p,sa55KCR:v}))}/>
<Inp label="SA 12kg — MGL" value={f.sa12MGL||""} onChange={v=>setF(p=>({...p,sa12MGL:v}))}/>
<Inp label="SA 5,5kg — MGL" value={f.sa55MGL||""} onChange={v=>setF(p=>({...p,sa55MGL:v}))}/>
</div>
<Btn color="blue" onClick={saveCompany} style={{marginTop:10}}>💾 Simpan Kode Pertamina</Btn>
</Card>
<Card><div style={{fontWeight:700,color:C.gl2,marginBottom:12,fontSize:13}}>💾 Backup & Data</div><div style={{display:"flex",gap:8,flexWrap:"wrap"}}><Btn color="green" onClick={backup}>📥 Export Backup JSON</Btn><Btn color="red" onClick={clearData}>🗑️ Hapus Semua Transaksi</Btn></div></Card>
</div>;
}

// ─── MAIN APP ─────────────────────────────────────────────────────────────────

var LOGO_HTS="data:image/png;base64,/9j/4AAQSkZJRgABAQAAAQABAAD/4gHYSUNDX1BST0ZJTEUAAQEAAAHIAAAAAAQwAABtbnRyUkdCIFhZWiAH4AABAAEAAAAAAABhY3NwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAA9tYAAQAAAADTLQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA...";

function KasBankMod({data,setData,user,toast}){
var C=useTheme();
var[tabK,setTabK]=useState("saldo");
var[bulanSaldo,setBulanSaldo]=useState(toMonth());
// Setup saldo awal
var[setupBSI,setSetupBSI]=useState({nominal:String((data.saldoAwalBank||{}).BSI?.nominal||""),tanggal:(data.saldoAwalBank||{}).BSI?.tanggal||""});
var[setupBCA,setSetupBCA]=useState({nominal:String((data.saldoAwalBank||{}).BCA?.nominal||""),tanggal:(data.saldoAwalBank||{}).BCA?.tanggal||""});
// Setoran cash states
var[bulanSetor,setBulanSetor]=useState(toMonth());
var[tglPilih,setTglPilih]=useState([]);// array tgl yang dipilih
var[bankSetor,setBankSetor]=useState("BSI");
var[penyetor,setPenyetor]=useState("Muhammad Haekal");
var[pecahInput,setPecahInput]=useState(()=>{var o={};DENOMS.forEach(d=>{o[d]="";});return o;});
var[showSlip,setShowSlip]=useState(false);
var[editSetor,setEditSetor]=useState(null);
var[editSetorF,setEditSetorF]=useState(null);
// Tarik/Setor TF states
var[tfJenis,setTfJenis]=useState("tarik");
var[tfTgl,setTfTgl]=useState(toDay());
var[tfBank,setTfBank]=useState("BSI");
var[tfNominal,setTfNominal]=useState("");
var[tfKet,setTfKet]=useState("");
var[tfEditId,setTfEditId]=useState(null);
var[tfDelId,setTfDelId]=useState(null);
// Filter states untuk Debit & Kredit tab
var[fKat,setFKat]=useState("");
var[fMetode,setFMetode]=useState("");
var[fOleh,setFOleh]=useState("");
var[fFrom,setFFrom]=useState("");
var[fTo,setFTo]=useState("");
var[fKet,setFKet]=useState("");

// ── Kalkulasi saldo bank ──
function getMutasiBank(bank){
  var list=[];
  // Saldo awal
  var sa=data.saldoAwalBank?.[bank]||{};
  var saldoAwal=Number(sa.nominal||0);
  var tglAwal=sa.tanggal||"";
  // TF Penjualan masuk ke bank ini
  (data.penjualan||[]).filter(p=>(!tglAwal||p.tanggal>=tglAwal)).forEach(p=>{
    var byr=(p.bayar||"").toLowerCase();
    if(byr==="transfer"&&(p.bank||"BSI")===bank){list.push({tanggal:p.tanggal,ket:"TF Penjualan — "+p.konsumen+" ("+p.noInv+")",jenis:"Masuk TF",masuk:p.total||0,keluar:0,bank});}
    else if(byr==="split"&&Number((p.splitDetail||{}).tf||0)>0&&(p.splitBank||"BSI")===bank){list.push({tanggal:p.tanggal,ket:"TF Penjualan (Split) — "+p.konsumen+" ("+p.noInv+")",jenis:"Masuk TF",masuk:Number(p.splitDetail.tf),keluar:0,bank});}
  });
  // Bayar BON TF masuk ke bank ini
  (data.bon||[]).forEach(b=>(b.pembayaran||[]).filter(px=>((px.metode||"").toLowerCase()==="transfer"||(px.metode||"").toLowerCase()==="tf")&&(px.bank||"BSI")===bank&&(!tglAwal||px.tanggal>=tglAwal)).forEach(px=>{
    list.push({tanggal:px.tanggal,ket:"Bayar BON — "+b.konsumen,jenis:"Masuk TF",masuk:Number(px.jumlah||px.nominal||0),keluar:0,bank});
  }));
  // Pengeluaran TF keluar dari bank ini
  (data.pengeluaran||[]).filter(p=>(p.metode||"").toLowerCase()==="transfer"&&(!tglAwal||p.tanggal>=tglAwal)).forEach(p=>{
    list.push({tanggal:p.tanggal,ket:"Pengeluaran — "+p.kategori+" ("+( p.ket||"")+")",jenis:"Keluar TF",masuk:0,keluar:Number(p.nominal||0),bank});
  });
  // DO selalu potong BSI
  if(bank==="BSI"){
    (data.doList||[]).filter(d=>(d.status||"diterima")==="diterima"&&(!tglAwal||d.tanggal>=tglAwal)).forEach(d=>{
      list.push({tanggal:d.tanggal,ket:"DO — "+d.sppbe+" ("+d.trip+")",jenis:"DO Bayar",masuk:0,keluar:Number(d.totalHPP||0),bank:"BSI"});
    });
  }
  // Setoran cash ke bank ini
  (data.setoranBank||[]).filter(s=>s.bank===bank&&(!tglAwal||s.tanggal>=tglAwal)).forEach(s=>{
    list.push({tanggal:s.tanggal,ket:"Setoran Cash ("+s.tglList?.join(",")+") ",jenis:"Setor Cash",masuk:s.nominal||0,keluar:0,bank});
  });
  // Tarik TF (cash laci -> bank ini): bank bertambah
  (data.kasBankTF||[]).filter(k=>k.jenis==="tarik"&&k.bank===bank&&(!tglAwal||k.tanggal>=tglAwal)).forEach(k=>{
    list.push({tanggal:k.tanggal,ket:"Tarik TF (Laci → "+bank+")"+(k.ket?" — "+k.ket:""),jenis:"Tarik TF",masuk:Number(k.nominal||0),keluar:0,bank});
  });
  // Setor TF (bank ini -> cash laci): bank berkurang
  (data.kasBankTF||[]).filter(k=>k.jenis==="setor"&&k.bank===bank&&(!tglAwal||k.tanggal>=tglAwal)).forEach(k=>{
    list.push({tanggal:k.tanggal,ket:"Setor TF ("+bank+" → Laci)"+(k.ket?" — "+k.ket:""),jenis:"Setor TF",masuk:0,keluar:Number(k.nominal||0),bank});
  });
  // Sort by tanggal
  list.sort((a,b)=>a.tanggal.localeCompare(b.tanggal));
  // Hitung running saldo
  var saldo=saldoAwal;
  list=list.map(x=>{saldo+=x.masuk-x.keluar;return{...x,saldo};});
  return{saldoAwal,list,saldoAkhir:saldo,tglAwal};
}

var mutBSI=getMutasiBank("BSI");
var mutBCA=getMutasiBank("BCA");
var totalBank=mutBSI.saldoAkhir+mutBCA.saldoAkhir;

// ── Kalkulasi per hari untuk kalender setoran ──
var dim=daysInMonth(bulanSetor);
var hariData={};
for(var di=1;di<=dim;di++){
  var tglD=bulanSetor+"-"+String(di).padStart(2,"0");
  // Cek ada tutup buku dengan pecahan
  var tb=(data.tutupBuku||[]).find(t=>t.tanggal===tglD);
  var hasPecah=tb&&(
  (typeof tb.totalPecah==="number"&&tb.totalPecah>0)||
  (typeof tb.totalPecah==="object"&&tb.totalPecah!==null&&Object.values(tb.totalPecah).some(v=>v>0))
);
  var wajibSetor=0;
  if(tb){
    var cashPenj=(data.penjualan||[]).filter(p=>p.tanggal===tglD&&(p.bayar||"").toLowerCase()==="cash").reduce((a,p)=>a+(p.total||0),0)+(data.penjualan||[]).filter(p=>p.tanggal===tglD&&(p.bayar||"").toLowerCase()==="split").reduce((a,p)=>a+Number((p.splitDetail||{}).cash||0),0);
    var bonCashD=(data.bon||[]).reduce((a,b)=>{var px=(b.pembayaran||[]).filter(p=>p.tanggal===tglD&&(p.metode||"cash").toLowerCase()==="cash");return a+px.reduce((s,p)=>s+Number(p.jumlah||p.nominal||0),0);},0);
    var penCashD=(data.pengeluaran||[]).filter(p=>p.tanggal===tglD&&(p.metode||"cash").toLowerCase()==="cash").reduce((a,p)=>a+Number(p.nominal||0),0);
    var kurangSetorD=(data.ambilan||[]).filter(a=>a.tanggal===tglD&&(a.ket||"").toLowerCase().includes("kurang setor")).reduce((a,x)=>a+Number(x.nominal||0),0);
    var jualLainCashD=(data.jualanLain||[]).filter(j=>j.tanggal===tglD&&j.bayar==="cash").reduce((a,j)=>a+Number(j.total||0),0);
    var tarikTFD=(data.kasBankTF||[]).filter(k=>k.tanggal===tglD&&k.jenis==="tarik").reduce((a,k)=>a+Number(k.nominal||0),0);
    var setorTFD=(data.kasBankTF||[]).filter(k=>k.tanggal===tglD&&k.jenis==="setor").reduce((a,k)=>a+Number(k.nominal||0),0);
    wajibSetor=Math.max(0,cashPenj+bonCashD+jualLainCashD-penCashD-kurangSetorD-tarikTFD+setorTFD);
  }
  var disetor=(data.kas||{})[tglD]?.disetor;
  // pecah object: dari field pecah (DENOMS qty) atau estimasi dari totalPecah number
var pecahObj={};
if(tb?.pecah&&typeof tb.pecah==="object"){pecahObj=tb.pecah;}
else if(typeof tb?.totalPecah==="number"&&tb.totalPecah>0){
  // estimasi dari total — bagi ke pecahan besar
  var sisa=tb.totalPecah;
  DENOMS.forEach(d=>{var lbr=Math.floor(sisa/d);pecahObj[d]=lbr;sisa-=lbr*d;});
}
hariData[tglD]={hasTB:!!tb,hasPecah,wajibSetor,disetor,pecah:pecahObj,totalPecah:typeof tb?.totalPecah==="number"?tb.totalPecah:0,tglD};
}

// Total dari tgl yang dipilih
var totalPilih=tglPilih.reduce((a,tgl)=>a+(hariData[tgl]?.wajibSetor||0),0);
// Gabungkan pecahan dari tgl yang dipilih
var pecahGabung={};
DENOMS.forEach(d=>{pecahGabung[d]=tglPilih.reduce((a,tgl)=>a+Number((hariData[tgl]?.pecah||{})[d]||0),0);});
// Init pecahInput dari gabungan saat tglPilih berubah
var totalFisik=DENOMS.reduce((a,d)=>a+Number(pecahInput[d]||0)*d,0);
var selisihPecah=totalFisik-totalPilih;

function konfirmasiSetor(){
  if(tglPilih.length===0){toast("Pilih minimal 1 tanggal");return;}
  var newKas={...(data.kas||{})};
  tglPilih.forEach(tgl=>{newKas[tgl]={disetor:true,bank:bankSetor,nominal:hariData[tgl]?.wajibSetor||0,tglSetor:toDay()};});
  var logSetor={id:uid(),tanggal:toDay(),bank:bankSetor,nominal:totalPilih,tglList:tglPilih,penyetor,pecah:{...pecahGabung},pecahReal:{...pecahInput}};
  setData(d=>({...d,kas:newKas,setoranBank:[logSetor,...(d.setoranBank||[])]}));
  setShowSlip(true);
  setTglPilih([]);
  setPecahInput(()=>{var o={};DENOMS.forEach(d=>{o[d]="";});return o;});
  toast("✓ Setoran Rp "+fR(totalPilih)+" ke "+bankSetor+" berhasil dikonfirmasi!");
}

// Hari ini
var HARI_ID=["Min","Sen","Sel","Rab","Kam","Jum","Sab"];
var firstDay=new Date(bulanSetor+"-01T00:00:00").getDay();
var BULAN_LABEL=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"];
var blnIdx=Number(bulanSetor.split("-")[1])-1;
var thnIdx=bulanSetor.split("-")[0];

return <div>
<STitle icon="📒" children="Buku Kas"/>
<div style={{display:"flex",gap:5,marginBottom:14,flexWrap:"wrap"}}>
{[["saldo","📒 Debit & Kredit"],["setor","💵 Setoran Cash"],["tarikSetorTF","🔄 Tarik/Setor TF"],["setup","⚙️ Setup Saldo Awal"]].map(x=><button key={x[0]} onClick={()=>setTabK(x[0])} style={{background:tabK===x[0]?C.blu:C.nav,color:tabK===x[0]?"white":C.wht,border:"1px solid "+(tabK===x[0]?C.blt:C.bdr),borderRadius:8,padding:"7px 14px",fontWeight:700,fontSize:12,cursor:"pointer"}}>{x[1]}</button>)}
</div>

{/* ── TAB SALDO & MUTASI ── */}
{tabK==="saldo"&&(()=>{
// ── Kumpulkan semua transaksi dari semua sumber ──
var allTrx=[];
var tglAwalBSI=(data.saldoAwalBank||{}).BSI?.tanggal||"";
var tglAwalBCA=(data.saldoAwalBank||{}).BCA?.tanggal||"";

// 1. Penjualan
(data.penjualan||[]).forEach(p=>{
var emp=(data.employees||[]).find(e=>e.id===p.salesId);
var oleh=emp?.nama||p.salesNama||p.sales||"—";
var byr=(p.bayar||"").toLowerCase();
var sd=p.splitDetail||{};
if(byr==="cash"){allTrx.push({tgl:p.tanggal,kat:"Penjualan",ket:p.konsumen,oleh,debit:p.total||0,kredit:0,metode:"Cash",ref:p.noInv});}
else if(byr==="transfer"||byr==="tf"){allTrx.push({tgl:p.tanggal,kat:"Penjualan",ket:p.konsumen,oleh,debit:p.total||0,kredit:0,metode:"TF "+(p.bank||""),ref:p.noInv});}
else if(byr==="bon"){/* BON tidak masuk Debit & Kredit */}
else if(byr==="split"){
if(Number(sd.cash)>0)allTrx.push({tgl:p.tanggal,kat:"Penjualan",ket:p.konsumen+" (Split Cash)",oleh,debit:Number(sd.cash),kredit:0,metode:"Cash",ref:p.noInv});
if(Number(sd.tf)>0)allTrx.push({tgl:p.tanggal,kat:"Penjualan",ket:p.konsumen+" (Split TF)",oleh,debit:Number(sd.tf),kredit:0,metode:"TF "+(p.splitBank||""),ref:p.noInv});
/* BON portion split: tidak masuk */
}
});

// 2. DO — bayar BSI
(data.doList||[]).filter(d=>(d.status||"diterima")==="diterima"&&d.totalHPP>0).forEach(d=>{
allTrx.push({tgl:d.tanggal,kat:"Belanja Modal",ket:"DO — "+d.sppbe,oleh:"—",debit:0,kredit:Number(d.totalHPP||0),metode:"TF BSI",ref:d.noDO||""});
});

// 3. Pengeluaran operasional
(data.pengeluaran||[]).forEach(p=>{
var m=(p.metode||"cash").toLowerCase();
var BELANJA_MODAL_CATS=["belanja modal","pancung","belanja tabung"];
var katTrx=BELANJA_MODAL_CATS.some(k=>(p.kategori||"").toLowerCase().includes(k))?"Belanja Modal":"Pengeluaran Ops";
allTrx.push({tgl:p.tanggal,kat:katTrx,ket:p.kategori+(p.ket?" — "+p.ket:""),oleh:p.karyawanNama||"—",debit:0,kredit:Number(p.nominal||0),metode:m==="cash"?"Cash":"TF "+(p.bank||""),ref:""});
});

// 4. Bayar BON (pembayaran piutang masuk)
(data.bon||[]).forEach(b=>{
var emp=(data.employees||[]).find(e=>e.id===b.salesId);
(b.pembayaran||[]).forEach(px=>{
var m=(px.metode||"cash").toLowerCase();
var nom=Number(px.jumlah||px.nominal||0);
if(nom>0)allTrx.push({tgl:px.tanggal,kat:"Bayar BON",ket:b.konsumen,oleh:px.salesPenerimaNama||emp?.nama||"—",debit:nom,kredit:0,metode:m==="cash"?"Cash":"TF "+(px.bank||""),ref:b.noInv});
});
});

// 5. Setoran Cash ke Bank — 2 baris (net 0, hanya perpindahan)
(data.setoranBank||[]).forEach(s=>{
// Baris 1: Masuk ke rekening bank (Debit)
allTrx.push({tgl:s.tanggal,kat:"Setoran Bank",ket:"Masuk Rek "+s.bank,oleh:s.penyetor||"—",debit:s.nominal||0,kredit:0,metode:"TF Bank",ref:""});
// Baris 2: Keluar dari laci (Kredit)
allTrx.push({tgl:s.tanggal,kat:"Setoran Bank",ket:"Keluar Laci → "+s.bank,oleh:s.penyetor||"—",debit:0,kredit:s.nominal||0,metode:"Cash",ref:""});
});

// 6. Tarik TF (cash laci -> bank) & Setor TF (bank -> cash laci) — 2 baris net 0
(data.kasBankTF||[]).forEach(k=>{
var nom=Number(k.nominal||0);
if(k.jenis==="tarik"){
  // Laci berkurang, Bank bertambah
  allTrx.push({tgl:k.tanggal,kat:"Tarik/Setor TF",ket:"Tarik TF → Masuk "+k.bank+(k.ket?" — "+k.ket:""),oleh:k.oleh||"—",debit:nom,kredit:0,metode:"TF "+k.bank,ref:""});
  allTrx.push({tgl:k.tanggal,kat:"Tarik/Setor TF",ket:"Tarik TF → Keluar Laci"+(k.ket?" — "+k.ket:""),oleh:k.oleh||"—",debit:0,kredit:nom,metode:"Cash",ref:""});
}else{
  // Bank berkurang, Laci bertambah
  allTrx.push({tgl:k.tanggal,kat:"Tarik/Setor TF",ket:"Setor TF → Keluar "+k.bank+(k.ket?" — "+k.ket:""),oleh:k.oleh||"—",debit:0,kredit:nom,metode:"TF "+k.bank,ref:""});
  allTrx.push({tgl:k.tanggal,kat:"Tarik/Setor TF",ket:"Setor TF → Masuk Laci"+(k.ket?" — "+k.ket:""),oleh:k.oleh||"—",debit:nom,kredit:0,metode:"Cash",ref:""});
}
});

// Sort by tanggal asc
allTrx.sort((a,b)=>a.tgl.localeCompare(b.tgl));

// Running balance — mulai dari saldo awal BSI+BCA
var saldoAwalTotal=(Number((data.saldoAwalBank||{}).BSI?.nominal||0))+(Number((data.saldoAwalBank||{}).BCA?.nominal||0));
var runBalance=saldoAwalTotal;
allTrx=allTrx.map(t=>{runBalance+=t.debit-t.kredit;return{...t,saldo:runBalance};});
var totalDebit=allTrx.reduce((a,t)=>a+t.debit,0);
var totalKredit=allTrx.reduce((a,t)=>a+t.kredit,0);

// Filter states sudah di level atas komponen
var filtered=allTrx.filter(t=>{
if(fKat&&t.kat!==fKat)return false;
if(fMetode&&t.metode!==fMetode)return false;
if(fOleh&&t.oleh!==fOleh)return false;
if(fFrom&&t.tgl<fFrom)return false;
if(fTo&&t.tgl>fTo)return false;
if(fKet&&!t.ket.toLowerCase().includes(fKet.toLowerCase()))return false;
return true;
});
var filtTotalDebit=filtered.reduce((a,t)=>a+t.debit,0);
var filtTotalKredit=filtered.reduce((a,t)=>a+t.kredit,0);

// Unique values for filter dropdowns
var uniKat=[...new Set(allTrx.map(t=>t.kat))];
var uniMetode=[...new Set(allTrx.map(t=>t.metode))];
var uniOleh=[...new Set(allTrx.map(t=>t.oleh))].filter(x=>x&&x!=="—");

// Saldo ringkasan atas
var cashLaciTerakhir=(()=>{var tbs=(data.tutupBuku||[]).slice().sort((a,b)=>b.tanggal.localeCompare(a.tanggal));return tbs.length>0?Number(tbs[0].cashLaci||0):0;})();

return <div>
{/* ── Ringkasan Saldo ── */}
<div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:8,marginBottom:14}}>
{[["💵 Cash di Laci",cashLaciTerakhir,C.glt],["🏦 Saldo BSI",mutBSI.saldoAkhir,C.blt],["🏦 Saldo BCA",mutBCA.saldoAkhir,"#9CA3AF"],["💰 Total",cashLaciTerakhir+totalBank,C.wht]].map(x=><Card key={x[0]} style={{marginBottom:0,textAlign:"center",padding:"10px 12px"}}>
<div style={{fontSize:10,color:C.gl2,marginBottom:3}}>{x[0]}</div>
<div style={{fontSize:16,fontWeight:900,color:x[2]}}>{fR(x[1])}</div>
</Card>)}
</div>

{/* ── Filter Bar ── */}
<Card style={{marginBottom:8,padding:"10px 12px"}}>
<div style={{display:"flex",gap:6,flexWrap:"wrap",alignItems:"flex-end"}}>
<div style={{minWidth:100}}>
<div style={{fontSize:10,color:C.gl2,marginBottom:3}}>Dari</div>
<input type="date" value={fFrom} onChange={e=>setFFrom(e.target.value)} style={{background:C.bg,border:"1px solid "+C.bdr,borderRadius:6,padding:"5px 7px",color:C.wht,fontSize:11,outline:"none",width:"100%"}}/>
</div>
<div style={{minWidth:100}}>
<div style={{fontSize:10,color:C.gl2,marginBottom:3}}>Sampai</div>
<input type="date" value={fTo} onChange={e=>setFTo(e.target.value)} style={{background:C.bg,border:"1px solid "+C.bdr,borderRadius:6,padding:"5px 7px",color:C.wht,fontSize:11,outline:"none",width:"100%"}}/>
</div>
<div style={{minWidth:130}}>
<div style={{fontSize:10,color:C.gl2,marginBottom:3}}>Kategori</div>
<select value={fKat} onChange={e=>setFKat(e.target.value)} style={{background:C.bg,border:"1px solid "+C.bdr,borderRadius:6,padding:"5px 7px",color:C.wht,fontSize:11,outline:"none",width:"100%"}}>
<option value="">Semua</option>{uniKat.map(k=><option key={k} value={k}>{k}</option>)}
</select>
</div>
<div style={{minWidth:120}}>
<div style={{fontSize:10,color:C.gl2,marginBottom:3}}>Metode</div>
<select value={fMetode} onChange={e=>setFMetode(e.target.value)} style={{background:C.bg,border:"1px solid "+C.bdr,borderRadius:6,padding:"5px 7px",color:C.wht,fontSize:11,outline:"none",width:"100%"}}>
<option value="">Semua</option>{uniMetode.map(k=><option key={k} value={k}>{k}</option>)}
</select>
</div>
<div style={{minWidth:130}}>
<div style={{fontSize:10,color:C.gl2,marginBottom:3}}>Transaksi Oleh</div>
<select value={fOleh} onChange={e=>setFOleh(e.target.value)} style={{background:C.bg,border:"1px solid "+C.bdr,borderRadius:6,padding:"5px 7px",color:C.wht,fontSize:11,outline:"none",width:"100%"}}>
<option value="">Semua</option>{uniOleh.map(k=><option key={k} value={k}>{k}</option>)}
</select>
</div>
<div style={{minWidth:140}}>
<div style={{fontSize:10,color:C.gl2,marginBottom:3}}>Keterangan</div>
<input value={fKet} onChange={e=>setFKet(e.target.value)} placeholder="Cari..." style={{background:C.bg,border:"1px solid "+C.bdr,borderRadius:6,padding:"5px 7px",color:C.wht,fontSize:11,outline:"none",width:"100%"}}/>
</div>
{(fKat||fMetode||fOleh||fFrom||fTo||fKet)&&<button onClick={()=>{setFKat("");setFMetode("");setFOleh("");setFFrom("");setFTo("");setFKet("");}} style={{background:C.rdk,border:"1px solid "+C.rlt,borderRadius:6,padding:"5px 10px",color:"white",cursor:"pointer",fontSize:11,fontWeight:700,alignSelf:"flex-end"}}>✕ Reset</button>}
</div>
</Card>

{/* ── Tabel Debit & Credit ── */}
<Card style={{padding:0,overflow:"hidden"}}>
<div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"10px 14px",borderBottom:"1px solid "+C.bdr}}>
<div style={{fontWeight:700,color:C.gl2,fontSize:13}}>📒 Debit & Credit — {filtered.length} transaksi</div>
<button onClick={()=>doPrint("_buku_kas_tbl")} style={{background:C.nav,border:"1px solid "+C.bdr,borderRadius:7,padding:"5px 12px",color:C.wht,cursor:"pointer",fontSize:11,fontWeight:700}}>🖨️ Cetak</button>
</div>
<div id="_buku_kas_tbl" style={{overflowX:"auto"}}>
<table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
<thead>
<tr style={{background:C.nav,borderBottom:"2px solid "+C.bdr}}>
{["Tgl","Kategori","Keterangan","Transaksi Oleh","Metode","Debit (Masuk)","Kredit (Keluar)","Saldo"].map(h=><th key={h} style={{padding:"8px 10px",color:C.gl2,fontWeight:700,textAlign:["Debit (Masuk)","Kredit (Keluar)","Saldo"].includes(h)?"right":"left",whiteSpace:"nowrap",borderBottom:"1px solid "+C.bdr,fontSize:10}}>{h}</th>)}
</tr>
</thead>
<tbody>
{filtered.length===0?<tr><td colSpan={8} style={{padding:"20px",textAlign:"center",color:C.gl2,fontStyle:"italic"}}>Tidak ada transaksi</td></tr>:
filtered.map((t,i)=><tr key={i} style={{background:i%2===0?C.bg:C.nav,borderBottom:"1px solid "+C.bdr}}>
<td style={{padding:"6px 10px",color:C.gl2,whiteSpace:"nowrap"}}>{fDs(t.tgl)}</td>
<td style={{padding:"6px 10px"}}><Bdg color={t.kat.includes("Pengeluaran")||t.kat.includes("Belanja")||t.kat.includes("Setoran")?"red":t.kat.includes("BON")?"orange":"green"}>{t.kat}</Bdg></td>
<td style={{padding:"6px 10px",color:C.wht,maxWidth:180,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{t.ket}</td>
<td style={{padding:"6px 10px",color:C.gl2,whiteSpace:"nowrap"}}>{t.oleh}</td>
<td style={{padding:"6px 10px",whiteSpace:"nowrap"}}><span style={{fontSize:10,fontWeight:600,color:t.metode.includes("Cash")?"#86EFAC":t.metode.includes("TF")?"#93C5FD":"#FCD34D"}}>{t.metode}</span></td>
<td style={{padding:"6px 10px",textAlign:"right",fontWeight:t.debit>0?700:400,color:t.debit>0?C.glt:C.gl2}}>{t.debit>0?fR(t.debit):"—"}</td>
<td style={{padding:"6px 10px",textAlign:"right",fontWeight:t.kredit>0?700:400,color:t.kredit>0?C.rlt:C.gl2}}>{t.kredit>0?fR(t.kredit):"—"}</td>
<td style={{padding:"6px 10px",textAlign:"right",fontWeight:700,color:t.saldo>=0?C.wht:C.rlt,whiteSpace:"nowrap"}}>{fR(t.saldo)}</td>
</tr>)}
</tbody>
<tfoot>
<tr style={{background:C.nav,borderTop:"2px solid "+C.bdr}}>
<td colSpan={5} style={{padding:"8px 10px",fontWeight:700,color:C.wht,fontSize:12}}>
{filtered.length<allTrx.length?`Filter aktif: ${filtered.length} dari ${allTrx.length} transaksi`:`Total ${allTrx.length} transaksi`}
</td>
<td style={{padding:"8px 10px",textAlign:"right",fontWeight:900,color:C.glt,fontSize:13}}>{fR(filtTotalDebit)}</td>
<td style={{padding:"8px 10px",textAlign:"right",fontWeight:900,color:C.rlt,fontSize:13}}>{fR(filtTotalKredit)}</td>
<td style={{padding:"8px 10px",textAlign:"right",fontWeight:900,color:(filtTotalDebit-filtTotalKredit)>=0?C.glt:C.rlt,fontSize:13}}>{fR(filtTotalDebit-filtTotalKredit)}</td>
</tr>
</tfoot>
</table>
</div>
{/* Summary bawah */}
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,padding:"10px 14px",borderTop:"1px solid "+C.bdr}}>
{[["Total Debit (Masuk)",filtTotalDebit,C.glt],["Total Kredit (Keluar)",filtTotalKredit,C.rlt],["Selisih Bersih",filtTotalDebit-filtTotalKredit,(filtTotalDebit-filtTotalKredit)>=0?C.glt:C.rlt]].map(x=><div key={x[0]} style={{background:C.nav,borderRadius:8,padding:"8px 12px",border:"1px solid "+C.bdr}}>
<div style={{fontSize:10,color:C.gl2,marginBottom:3}}>{x[0]}</div>
<div style={{fontSize:15,fontWeight:900,color:x[2]}}>{fR(x[1])}</div>
</div>)}
</div>
</Card>
</div>;
})()}

{/* ── TAB SETORAN CASH ── */}
{tabK==="setor"&&<div>
<Card>
<div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>💵 Pilih Tanggal Setoran</div>
<div style={{display:"flex",gap:8,alignItems:"center",marginBottom:10}}>
<Inp label="Bulan" type="month" value={bulanSetor} onChange={v=>{setBulanSetor(v);setTglPilih([]);}} style={{maxWidth:160,marginBottom:0}}/>
<div style={{fontSize:11,color:C.gl2}}>
<span style={{display:"inline-block",width:12,height:12,borderRadius:3,background:"#F59E0B",marginRight:4}}/>Tersedia &nbsp;
<span style={{display:"inline-block",width:12,height:12,borderRadius:3,background:C.blt,marginRight:4}}/>Dipilih &nbsp;
<span style={{display:"inline-block",width:12,height:12,borderRadius:3,background:C.glt,marginRight:4}}/>Sudah Setor &nbsp;
<span style={{display:"inline-block",width:12,height:12,borderRadius:3,background:C.bdr,marginRight:4}}/>Tidak Ada
</div>
</div>
{/* Status total belum setor */}
{(()=>{
var totalBelumSetor=Object.entries(hariData).filter(([tgl,hd])=>hd.hasPecah&&!hd.disetor).reduce((a,[_,hd])=>a+(hd.wajibSetor||0),0);
var jmlHariBelum=Object.entries(hariData).filter(([_,hd])=>hd.hasPecah&&!hd.disetor).length;
return totalBelumSetor>0?<div style={{background:"#FEF3C7",border:"1px solid #F59E0B",borderRadius:8,padding:"8px 14px",marginBottom:10,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
<span style={{fontSize:12,color:"#92400E",fontWeight:700}}>⚠️ {jmlHariBelum} hari belum disetor</span>
<span style={{fontSize:15,fontWeight:900,color:"#92400E"}}>Total: {fR(totalBelumSetor)}</span>
</div>:null;
})()}
{/* Kalender */}
<div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:4,marginBottom:12}}>
{HARI_ID.map(h=><div key={h} style={{textAlign:"center",fontSize:10,color:C.gl2,fontWeight:700,padding:"4px 0"}}>{h}</div>)}
{Array.from({length:firstDay}).map((_,i)=><div key={"e"+i}/>)}
{Array.from({length:dim}).map((_,i)=>{
var d2=i+1;
var tglD2=bulanSetor+"-"+String(d2).padStart(2,"0");
var hd=hariData[tglD2]||{};
var isPilih=tglPilih.includes(tglD2);
// Warna: dipilih=biru, sudah setor=hijau, tersedia=kuning, tidak ada=abu
var bg=isPilih?"#1D4ED8":hd.disetor?"#15803D":hd.hasPecah?"#F59E0B":C.bdr;
var color=isPilih||hd.disetor||hd.hasPecah?"white":C.gl2;
var canClick=hd.hasPecah;// SEMUA tgl ada pecahan bisa diklik, termasuk yang sudah setor
return <div key={d2} onClick={()=>{if(!canClick)return;setTglPilih(prev=>prev.includes(tglD2)?prev.filter(x=>x!==tglD2):[...prev,tglD2]);setPecahInput(()=>{var o={};DENOMS.forEach(d=>{o[d]="";});return o;});}} style={{textAlign:"center",padding:"6px 3px",borderRadius:6,background:bg,color,fontSize:12,fontWeight:isPilih||hd.disetor?700:400,cursor:canClick?"pointer":"default",border:isPilih?"2px solid white":hd.disetor?"2px solid #86EFAC":"2px solid transparent",position:"relative"}}>
<div style={{fontSize:13,fontWeight:700}}>{d2}</div>
{hd.wajibSetor>0&&<div style={{fontSize:10,fontWeight:700,marginTop:1,opacity:.95}}>{(hd.wajibSetor/1000000)>=1?(hd.wajibSetor/1000000).toFixed(1)+"jt":(hd.wajibSetor/1000).toFixed(0)+"k"}</div>}
{hd.disetor&&<div style={{fontSize:8,opacity:.85}}>✓setor</div>}
</div>;})}
</div>
</Card>

{/* Tabel 1: Rincian per tgl - selalu tampil Opsi A */}
<Card>
<div style={{fontWeight:700,color:C.gl2,marginBottom:8,fontSize:13}}>📋 Tabel 1 — Rincian Pecahan per Tanggal (dari Tutup Buku)</div>
<div style={{overflowX:"auto"}}>
<table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
<thead><tr style={{background:C.nav}}>
<th style={{padding:"6px 10px",color:C.gl2,textAlign:"left",border:"1px solid "+C.bdr,fontSize:10}}>Pecahan</th>
{tglPilih.length>0?tglPilih.sort().map(tgl=><th key={tgl} style={{padding:"6px 8px",color:C.blt,textAlign:"center",border:"1px solid "+C.bdr,fontSize:10}}>{fDs(tgl)}</th>):<th style={{padding:"6px 8px",color:C.gl2,textAlign:"center",border:"1px solid "+C.bdr,fontSize:10,fontStyle:"italic"}}>(pilih tanggal)</th>}
<th style={{padding:"6px 10px",color:C.olt,textAlign:"center",border:"1px solid "+C.bdr,fontSize:10,fontWeight:700}}>TOTAL</th>
</tr></thead>
<tbody>
{DENOMS.map(d=>{
var totD=tglPilih.reduce((a,tgl)=>a+Number((hariData[tgl]?.pecah||{})[d]||0),0);
return <tr key={d} style={{borderBottom:"1px solid "+C.bdr}}>
<td style={{padding:"5px 10px",color:C.wht,fontWeight:600,border:"1px solid "+C.bdr}}>{fR(d)}</td>
{tglPilih.length>0?tglPilih.sort().map(tgl=>{var v=Number((hariData[tgl]?.pecah||{})[d]||0);return <td key={tgl} style={{padding:"5px 8px",textAlign:"center",color:v>0?C.wht:C.gl2,border:"1px solid "+C.bdr}}>{v||"—"}</td>;}):
<td style={{padding:"5px 8px",textAlign:"center",color:C.gl2,border:"1px solid "+C.bdr}}>—</td>}
<td style={{padding:"5px 10px",textAlign:"center",fontWeight:700,color:totD>0?C.olt:C.gl2,border:"1px solid "+C.bdr}}>{totD||"—"}</td>
</tr>;})}
<tr style={{background:C.nav,borderTop:"2px solid "+C.bdr,fontWeight:700}}>
<td style={{padding:"6px 10px",color:C.wht,border:"1px solid "+C.bdr}}>TOTAL</td>
{tglPilih.length>0?tglPilih.sort().map(tgl=><td key={tgl} style={{padding:"6px 8px",textAlign:"center",color:C.glt,border:"1px solid "+C.bdr}}>{fR(hariData[tgl]?.wajibSetor||0)}</td>):
<td style={{padding:"6px 8px",textAlign:"center",color:C.gl2,border:"1px solid "+C.bdr}}>—</td>}
<td style={{padding:"6px 10px",textAlign:"center",fontWeight:900,color:tglPilih.length>0?C.glt:C.gl2,fontSize:14,border:"1px solid "+C.bdr}}>{tglPilih.length>0?fR(totalPilih):"—"}</td>
</tr>
</tbody>
</table>
</div>
</Card>

{/* Tabel 2: selalu tampil */}
<Card>
<div style={{fontWeight:700,color:C.gl2,marginBottom:8,fontSize:13}}>✏️ Tabel 2 — Input Real Pecahan (bisa diedit)</div>
{tglPilih.length===0&&<div style={{fontSize:11,color:C.gl2,marginBottom:8,fontStyle:"italic"}}>Pilih tanggal di kalender → Tabel 1 akan muncul dan Tabel 2 terisi otomatis.</div>}
<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
<div>
<div style={{fontSize:11,color:C.gl2,marginBottom:8}}>{tglPilih.length>0?"Terisi dari Tabel 1 · edit sesuai fisik yang ada:":"Input manual pecahan:"}</div>
<table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
<thead><tr style={{background:C.nav}}>
{["Pecahan","T1 (sistem)","T2 (real)","Selisih"].map(h=><th key={h} style={{padding:"5px 8px",color:C.gl2,textAlign:h==="Pecahan"?"left":"center",border:"1px solid "+C.bdr,fontSize:10}}>{h}</th>)}
</tr></thead>
<tbody>
{DENOMS.map(d=>{
var sys=pecahGabung[d]||0;
var real=Number(pecahInput[d]||0);
var sel=real-sys;
return <tr key={d} style={{borderBottom:"1px solid "+C.bdr}}>
<td style={{padding:"4px 8px",color:C.wht,fontWeight:600,border:"1px solid "+C.bdr}}>{fR(d)}</td>
<td style={{padding:"4px 8px",textAlign:"center",color:sys>0?C.blt:C.gl2,border:"1px solid "+C.bdr}}>{sys||"—"}</td>
<td style={{padding:"4px 6px",border:"1px solid "+C.bdr}}>
<input type="number" value={pecahInput[d]} placeholder={String(sys||0)} onChange={e=>setPecahInput(u=>({...u,[d]:e.target.value}))} style={{background:C.nav,border:"none",borderRadius:4,padding:"3px 6px",color:C.wht,fontSize:11,outline:"none",width:"100%",textAlign:"center"}}/>
</td>
<td style={{padding:"4px 8px",textAlign:"center",color:sel>0?C.glt:sel<0?C.rlt:C.gl2,fontWeight:sel!==0?700:400,border:"1px solid "+C.bdr}}>{tglPilih.length>0?(sel>0?"+"+sel:sel||"—"):"—"}</td>
</tr>;})}
<tr style={{background:C.nav,borderTop:"2px solid "+C.bdr,fontWeight:700}}>
<td colSpan={2} style={{padding:"6px 8px",color:C.wht,border:"1px solid "+C.bdr}}>Total T2 (real)</td>
<td style={{padding:"6px 8px",textAlign:"center",color:C.glt,fontWeight:900,fontSize:13,border:"1px solid "+C.bdr}}>{totalFisik>0?fR(totalFisik):"—"}</td>
<td style={{padding:"6px 8px",textAlign:"center",color:Math.abs(selisihPecah)<1000&&totalFisik>0?C.glt:C.rlt,fontWeight:700,border:"1px solid "+C.bdr}}>{tglPilih.length>0&&totalFisik>0?(selisihPecah>=0?"+":"")+fR(selisihPecah):"—"}</td>
</tr>
</tbody>
</table>
</div>
<div>
<div style={{background:C.nav,borderRadius:8,padding:14,border:"1px solid "+C.bdr,marginBottom:8}}>
<div style={{fontSize:11,color:C.gl2,marginBottom:4}}>Total dari Tabel 1 (sistem)</div>
<div style={{fontSize:16,fontWeight:700,color:C.blt,marginBottom:10}}>{tglPilih.length>0?fR(totalPilih):"—"}</div>
<div style={{fontSize:11,color:C.gl2,marginBottom:4}}>Total dari Tabel 2 (real)</div>
<div style={{fontSize:16,fontWeight:700,color:C.glt,marginBottom:10}}>{totalFisik>0?fR(totalFisik):"—"}</div>
<div style={{height:1,background:C.bdr,marginBottom:10}}/>
<div style={{fontSize:11,color:Math.abs(selisihPecah)<1000?C.glt:C.olt,fontWeight:700,marginBottom:4}}>Selisih T1 vs T2</div>
<div style={{fontSize:18,fontWeight:900,color:Math.abs(selisihPecah)<1000&&tglPilih.length>0?C.glt:C.rlt}}>{tglPilih.length>0&&totalFisik>0?(selisihPecah>=0?"+":"")+fR(selisihPecah):"—"}</div>
</div>
<div style={{display:"flex",flexDirection:"column",gap:8}}>
<Sel label="Setor ke Bank" value={bankSetor} onChange={setBankSetor} opts={[{v:"BSI",l:"BSI — Bank Syariah Indonesia"},{v:"BCA",l:"BCA"}]}/>
<Inp label="Penyetor" value={penyetor} onChange={setPenyetor}/>
<Btn color="green" onClick={konfirmasiSetor}>✓ Konfirmasi Setoran {tglPilih.length>0?fR(totalPilih):"—"} → {bankSetor}</Btn>
</div>
</div>
</div>
</Card>

{/* Riwayat setoran bank */}
<Card>
<div style={{fontWeight:700,color:C.gl2,marginBottom:8,fontSize:13}}>📋 Riwayat Setoran Bank</div>
{(data.setoranBank||[]).length===0?<div style={{color:C.gl2,fontSize:12,fontStyle:"italic"}}>Belum ada setoran</div>:
<div>
{(data.setoranBank||[]).slice(0,20).map((r,ri)=><div key={r.id||ri} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"8px 12px",background:ri%2===0?C.nav:C.bg,borderRadius:6,marginBottom:4,border:"1px solid "+C.bdr}}>
<div>
<div style={{fontSize:12,fontWeight:700,color:C.wht}}>{fDs(r.tanggal)} — <Bdg color="blue">{r.bank}</Bdg> — <b style={{color:C.glt}}>{fR(r.nominal)}</b></div>
<div style={{fontSize:10,color:C.gl2,marginTop:2}}>Penyetor: {r.penyetor} · Tgl: {r.tglList?.map(t=>fDs(t)).join(", ")}</div>
</div>
<div style={{display:"flex",gap:5}}>
<button onClick={()=>{setEditSetor(r);setShowSlip(true);}} style={{background:C.nav,border:"1px solid "+C.blt,borderRadius:6,padding:"4px 8px",color:C.blt,cursor:"pointer",fontSize:11,fontWeight:700}}>👁️ View</button>
<button onClick={()=>{setEditSetor(r);setEditSetorF({...r});}} style={{background:"#78350F",border:"1px solid #F59E0B",borderRadius:6,padding:"4px 8px",color:"#FCD34D",cursor:"pointer",fontSize:11,fontWeight:700}}>✏️ Edit</button>
<button onClick={()=>{setEditSetor(r);setShowSlip(true);setTimeout(()=>doPrint("_slip_setor"),400);}} style={{background:C.nav,border:"1px solid "+C.gl2,borderRadius:6,padding:"4px 8px",color:C.gl2,cursor:"pointer",fontSize:11,fontWeight:700}}>🖨️ Cetak</button>
</div>
</div>)}
</div>}
</Card>
</div>}

{/* ── TAB TARIK/SETOR TF ── */}
{tabK==="tarikSetorTF"&&(()=>{
function resetTfForm(){setTfJenis("tarik");setTfTgl(toDay());setTfBank("BSI");setTfNominal("");setTfKet("");setTfEditId(null);}
function simpanTF(){
  var nom=Number(tfNominal)||0;
  if(nom<=0){toast("⚠️ Nominal harus diisi");return;}
  var rec={id:tfEditId||uid(),tanggal:tfTgl,jenis:tfJenis,bank:tfBank,nominal:nom,ket:tfKet.trim(),oleh:user?.nama||""};
  if(tfEditId){
    setData(d=>({...d,kasBankTF:(d.kasBankTF||[]).map(x=>x.id===tfEditId?rec:x)}));
    toast("✓ Transaksi TF diperbarui!");
  }else{
    setData(d=>({...d,kasBankTF:[rec,...(d.kasBankTF||[])]}));
    toast((tfJenis==="tarik"?"✓ Tarik TF":"✓ Setor TF")+" Rp "+fR(nom)+" tersimpan!");
  }
  resetTfForm();
}
function mulaiEditTF(x){setTfEditId(x.id);setTfTgl(x.tanggal);setTfJenis(x.jenis);setTfBank(x.bank);setTfNominal(String(x.nominal));setTfKet(x.ket||"");}
var riwayatTF=(data.kasBankTF||[]).slice().sort((a,b)=>(b.tanggal||"").localeCompare(a.tanggal||"")).slice(0,50);
var totalTarik=(data.kasBankTF||[]).filter(k=>k.jenis==="tarik").reduce((a,k)=>a+Number(k.nominal||0),0);
var totalSetorTF=(data.kasBankTF||[]).filter(k=>k.jenis==="setor").reduce((a,k)=>a+Number(k.nominal||0),0);
return <div>
<Card style={{border:"1px solid "+(tfEditId?C.olt:C.bdr)}}>
<div style={{fontWeight:700,color:tfEditId?C.olt:C.gl2,marginBottom:10,fontSize:13}}>{tfEditId?"✏️ Edit Transaksi TF":"🔄 Tarik / Setor TF"}</div>
<div style={{display:"flex",gap:8,marginBottom:12}}>
<button onClick={()=>setTfJenis("tarik")} style={{flex:1,background:tfJenis==="tarik"?C.rdk:C.nav,color:tfJenis==="tarik"?"white":C.wht,border:"1px solid "+(tfJenis==="tarik"?C.rdk:C.bdr),borderRadius:8,padding:"10px",fontWeight:700,fontSize:12,cursor:"pointer"}}>📤 Tarik TF<div style={{fontSize:10,fontWeight:400,marginTop:2,opacity:.9}}>Laci → Bank</div></button>
<button onClick={()=>setTfJenis("setor")} style={{flex:1,background:tfJenis==="setor"?C.grn:C.nav,color:tfJenis==="setor"?"white":C.wht,border:"1px solid "+(tfJenis==="setor"?C.grn:C.bdr),borderRadius:8,padding:"10px",fontWeight:700,fontSize:12,cursor:"pointer"}}>📥 Setor TF<div style={{fontSize:10,fontWeight:400,marginTop:2,opacity:.9}}>Bank → Laci</div></button>
</div>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(180px,260px))",gap:10}}>
<Inp label="Tanggal" type="date" value={tfTgl} onChange={setTfTgl}/>
<Sel label={tfJenis==="tarik"?"Bank Tujuan":"Bank Asal"} value={tfBank} onChange={setTfBank} opts={[{v:"BSI",l:"BSI"},{v:"BCA",l:"BCA"}]}/>
<Inp label="Nominal" type="number" value={tfNominal} onChange={setTfNominal} placeholder="0"/>
<Inp label="Keterangan (opsional)" value={tfKet} onChange={setTfKet} placeholder="—"/>
</div>
<div style={{background:C.nav,borderRadius:8,padding:"10px 14px",border:"1px solid "+C.bdr,marginBottom:10,fontSize:12,color:C.gl2}}>
{tfJenis==="tarik"
  ?<>📤 Cash Laci <b style={{color:C.rlt}}>−{fR(Number(tfNominal)||0)}</b> → Saldo {tfBank} <b style={{color:C.glt}}>+{fR(Number(tfNominal)||0)}</b></>
  :<>📥 Saldo {tfBank} <b style={{color:C.rlt}}>−{fR(Number(tfNominal)||0)}</b> → Cash Laci <b style={{color:C.glt}}>+{fR(Number(tfNominal)||0)}</b></>}
</div>
<div style={{display:"flex",gap:8}}>
<Btn color={tfEditId?"orange":(tfJenis==="tarik"?"red":"green")} onClick={simpanTF}>{tfEditId?"💾 Simpan Perubahan":(tfJenis==="tarik"?"📤 Konfirmasi Tarik TF":"📥 Konfirmasi Setor TF")}</Btn>
{tfEditId&&<Btn color="gray" onClick={resetTfForm}>✕ Batal</Btn>}
</div>
</Card>

<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(180px,260px))",gap:10,marginBottom:12}}>
<div style={{background:C.card,borderRadius:10,border:"1px solid "+C.bdr,padding:12}}>
<div style={{fontSize:11,color:C.gl2,marginBottom:4}}>Total Tarik TF (Laci → Bank)</div>
<div style={{fontSize:16,fontWeight:900,color:C.rlt}}>{fR(totalTarik)}</div>
</div>
<div style={{background:C.card,borderRadius:10,border:"1px solid "+C.bdr,padding:12}}>
<div style={{fontSize:11,color:C.gl2,marginBottom:4}}>Total Setor TF (Bank → Laci)</div>
<div style={{fontSize:16,fontWeight:900,color:C.glt}}>{fR(totalSetorTF)}</div>
</div>
</div>

<Card>
<div style={{fontWeight:700,color:C.gl2,marginBottom:10,fontSize:13}}>📋 Riwayat Tarik/Setor TF</div>
<RTbl headers={["Tanggal","Jenis","Bank","Nominal","Keterangan","Aksi"]} rows={riwayatTF.map(r=>[
fDs(r.tanggal),
r.jenis==="tarik"?"📤 Tarik (Laci→Bank)":"📥 Setor (Bank→Laci)",
r.bank,
fR(r.nominal),
r.ket||"—",
<div style={{display:"flex",gap:4}}>
<button onClick={()=>mulaiEditTF(r)} style={{background:C.blu,border:"none",borderRadius:5,padding:"3px 8px",color:"white",cursor:"pointer",fontSize:10,fontWeight:700}}>✏️</button>
<button onClick={()=>setTfDelId(r)} style={{background:C.rdk,border:"none",borderRadius:5,padding:"3px 8px",color:"white",cursor:"pointer",fontSize:10,fontWeight:700}}>🗑️</button>
</div>
])} empty="Belum ada transaksi tarik/setor TF"/>
</Card>
{tfDelId&&<ConfirmDel msg={"Hapus transaksi "+(tfDelId.jenis==="tarik"?"Tarik":"Setor")+" TF "+fR(tfDelId.nominal)+"?"} onCancel={()=>setTfDelId(null)} onConfirm={()=>{setData(d=>({...d,kasBankTF:(d.kasBankTF||[]).filter(x=>x.id!==tfDelId.id)}));setTfDelId(null);toast("✓ Dihapus");}}/>}
</div>;
})()}

{/* ── TAB SETUP SALDO AWAL ── */}
{tabK==="setup"&&<div style={{display:"flex",gap:14,flexWrap:"wrap",alignItems:"flex-start"}}>
{["BSI","BCA"].map(bank=>{
var st=bank==="BSI"?setupBSI:setupBCA;
var setSt=bank==="BSI"?setSetupBSI:setSetupBCA;
return <div key={bank} style={{flex:"0 1 320px",minWidth:280}}>
<Card style={{width:"100%"}}>
<div style={{fontWeight:700,color:C.blt,marginBottom:12,fontSize:13}}>🏦 Bank {bank}</div>
<div style={{display:"grid",gridTemplateColumns:"1fr",gap:10,marginBottom:10}}>
<Inp label="Tanggal Efektif" type="date" value={st.tanggal} onChange={v=>setSt(p=>({...p,tanggal:v}))}/>
<Inp label="Saldo Awal (Rp)" type="number" value={st.nominal} onChange={v=>setSt(p=>({...p,nominal:v}))} placeholder="0"/>
</div>
<Btn color="blue" onClick={()=>{
var newSaldo={...(data.saldoAwalBank||{})};
newSaldo[bank]={nominal:Number(st.nominal||0),tanggal:st.tanggal};
setData(d=>({...d,saldoAwalBank:newSaldo}));
toast("✓ Saldo awal "+bank+" disimpan!");
}}>💾 Simpan Saldo Awal {bank}</Btn>
<div style={{marginTop:8,fontSize:11,color:C.gl2}}>Saldo saat ini: <b style={{color:C.glt}}>{fR(bank==="BSI"?mutBSI.saldoAkhir:mutBCA.saldoAkhir)}</b></div>
</Card>
</div>;})}
</div>}

{/* ── MODAL EDIT SETORAN ── */}
{editSetor&&!showSlip&&<div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.7)",zIndex:999,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
<div style={{background:C.card,borderRadius:12,width:"100%",maxWidth:500,padding:20,border:"1px solid "+C.bdr}}>
<div style={{fontWeight:700,color:C.wht,marginBottom:14,fontSize:14}}>✏️ Edit Setoran — {fDs(editSetor.tanggal)}</div>
{(()=>{
var ef=editSetorF||editSetor;
var setEf=setEditSetorF;
return <div>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(180px,260px))",gap:8,marginBottom:10}}>
<Inp label="Tanggal Setor" type="date" value={ef.tanggal||""} onChange={v=>setEf({...ef,tanggal:v})}/>
<Sel label="Bank" value={ef.bank||"BSI"} onChange={v=>setEf({...ef,bank:v})} opts={[{v:"BSI",l:"BSI"},{v:"BCA",l:"BCA"}]}/>
<Inp label="Nominal (Rp)" type="number" value={String(ef.nominal||"")} onChange={v=>setEf({...ef,nominal:Number(v)})}/>
<Inp label="Penyetor" value={ef.penyetor||""} onChange={v=>setEf({...ef,penyetor:v})}/>
</div>
<div style={{fontSize:11,fontWeight:700,color:C.gl2,marginBottom:6}}>Pecahan Real (Tabel 2):</div>
<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(180px,260px))",gap:6,marginBottom:12}}>
{DENOMS.map(d=><div key={d} style={{display:"flex",alignItems:"center",gap:6,background:C.nav,borderRadius:6,padding:"4px 8px",border:"1px solid "+C.bdr}}>
<span style={{fontSize:11,color:C.gl2,minWidth:70}}>{fR(d)}</span>
<input type="number" value={(ef.pecahReal||ef.pecah||{})[d]||""} placeholder="0" onChange={e=>{var pr={...(ef.pecahReal||ef.pecah||{})};pr[d]=Number(e.target.value)||0;setEf({...ef,pecahReal:pr});}} style={{background:"transparent",border:"none",color:C.wht,fontSize:11,outline:"none",width:60,textAlign:"center"}}/>
</div>)}
</div>
<div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
<button onClick={()=>{setEditSetor(null);setEditSetorF(null);}} style={{background:C.nav,border:"1px solid "+C.bdr,borderRadius:8,padding:"8px 14px",color:C.gl2,cursor:"pointer",fontWeight:700,fontSize:12}}>Batal</button>
<button onClick={()=>{
var updated=(data.setoranBank||[]).map(x=>x.id===editSetor.id?{...editSetor,...ef,id:editSetor.id}:x);
setData(d=>({...d,setoranBank:updated}));
setEditSetor(null);setEditSetorF(null);
toast("✓ Setoran diperbarui!");
}} style={{background:C.glt,border:"none",borderRadius:8,padding:"8px 14px",color:"white",cursor:"pointer",fontWeight:700,fontSize:12}}>💾 Simpan</button>
<button onClick={()=>setShowSlip(true)} style={{background:C.blt,border:"none",borderRadius:8,padding:"8px 14px",color:"white",cursor:"pointer",fontWeight:700,fontSize:12}}>🖨️ Cetak Slip</button>
</div>
</div>;
})()}
</div>
</div>}

{/* ── MODAL SLIP SETORAN ── */}
{showSlip&&(editSetor||(data.setoranBank||[]).length>0)&&(()=>{
var last=editSetor||(data.setoranBank||[])[0];
var tglSetor=new Date(last.tanggal+"T00:00:00");
var hariLabel=["Minggu","Senin","Selasa","Rabu","Kamis","Jumat","Sabtu"][tglSetor.getDay()];
var tglLabel=tglSetor.toLocaleDateString("id-ID",{day:"2-digit",month:"long",year:"numeric"});
var noRek=last.bank==="BSI"?"812 69 2121 8":"—";
return <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.7)",zIndex:999,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
<div style={{background:"white",borderRadius:12,width:"100%",maxWidth:420,boxShadow:"0 20px 60px rgba(0,0,0,.5)",overflow:"hidden"}}>
<div id="_slip_setor" style={{background:"white",padding:20,fontFamily:"Arial,sans-serif",color:"#111",borderRadius:12}}>
{/* Header — format invoice */}
<div style={{background:"linear-gradient(135deg,#0a1f44 0%,#122d5e 100%)",padding:"14px 18px 12px",display:"flex",justifyContent:"space-between",alignItems:"center",margin:-20,marginBottom:0}}>
<div style={{display:"flex",alignItems:"center",gap:10}}>
<CompanyLogo h={40} variant="dark"/>
</div>
<div style={{display:"flex",alignItems:"center",gap:5}}>
<PertaminaLogo h={26} variant="dark"/>
</div>
</div>
{/* 3-color divider */}
<div style={{height:3,display:"flex",margin:"0 -20px 14px"}}>
<div style={{flex:1,background:"#1565c0"}}/><div style={{flex:1,background:"#6ab04c"}}/><div style={{flex:1,background:"#e53935"}}/>
</div>
<div style={{textAlign:"center",marginBottom:12}}>
<div style={{fontSize:13,fontWeight:700,color:"#0a1f44",letterSpacing:.5}}>RINCIAN SETORAN</div>
<div style={{fontSize:12,color:"#555"}}>{hariLabel} &nbsp;|&nbsp; {tglLabel}</div>
</div>
{/* Tabel pecahan */}
<table style={{width:"90%",margin:"0 auto",borderCollapse:"collapse",fontSize:13,marginBottom:16}}>
<thead>
<tr style={{background:"#0a1f44"}}>
<th style={{padding:"8px 14px",color:"white",textAlign:"left",border:"1px solid #1e3a5f",fontSize:11,letterSpacing:.8}}>PECAHAN</th>
<th style={{padding:"8px 10px",color:"white",textAlign:"center",border:"1px solid #1e3a5f",fontSize:11,letterSpacing:.8}}>LBR</th>
<th style={{padding:"8px 14px",color:"white",textAlign:"right",border:"1px solid #1e3a5f",fontSize:11,letterSpacing:.8}}>JUMLAH</th>
</tr>
</thead>
<tbody>
{(()=>{
var totalReal=DENOMS.reduce((a,d)=>a+Number(last.pecahReal?.[d]||last.pecah?.[d]||0)*d,0);
return <>
{DENOMS.map(d=>{
var lbr=Number(last.pecahReal?.[d]||last.pecah?.[d]||0);
var jml=lbr*d;
return <tr key={d} style={{borderBottom:"1px solid #E5E7EB",background:lbr>0?"white":"#FAFAFA"}}>
<td style={{padding:"7px 14px",color:"#111",border:"1px solid #E2E8F0",fontSize:13}}>Rp {Number(d).toLocaleString("id-ID")}</td>
<td style={{padding:"7px 10px",textAlign:"center",color:lbr>0?"#0a1f44":"#CBD5E1",fontWeight:lbr>0?700:400,border:"1px solid #E2E8F0",fontSize:14}}>{lbr||"—"}</td>
<td style={{padding:"7px 14px",textAlign:"right",color:jml>0?"#0a1f44":"#CBD5E1",fontWeight:jml>0?700:400,border:"1px solid #E2E8F0",fontSize:13}}>{jml>0?"Rp "+jml.toLocaleString("id-ID"):"—"}</td>
</tr>;})}
<tr style={{background:"#0a1f44",fontWeight:700}}>
<td colSpan={2} style={{padding:"9px 14px",color:"white",border:"1px solid #1e3a5f",fontSize:13,letterSpacing:.5}}>TOTAL</td>
<td style={{padding:"9px 14px",textAlign:"right",color:"white",fontSize:16,border:"1px solid #1e3a5f",fontWeight:900}}>Rp {totalReal.toLocaleString("id-ID")}</td>
</tr>
</>;
})()}
</tbody>
</table>
{/* Footer info rekening — desain lebih menarik */}
<div style={{width:"90%",margin:"0 auto 12px",background:"#F0F7FF",borderRadius:10,border:"1.5px solid #BFDBFE",overflow:"hidden"}}>
<div style={{background:"#0a1f44",padding:"8px 16px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
<span style={{fontSize:11,color:"#93C5FD",fontWeight:700,letterSpacing:1}}>SETORAN LPG</span>
<span style={{fontSize:16,fontWeight:900,color:"white",letterSpacing:.5}}>{last.bank}</span>
</div>
<div style={{padding:"12px 16px"}}>
{[["Penyetor",last.penyetor||penyetor,14],["Nama Nasabah","PT. HOE TRANG SA",14],["Nomor Rekening",noRek,16]].map(x=><div key={x[0]} style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:x[2]===16?0:8,paddingBottom:x[2]===16?0:8,borderBottom:x[2]===16?"none":"1px solid #DBEAFE"}}>
<span style={{fontSize:11,color:"#475569"}}>{x[0]}</span>
<span style={{fontSize:x[2],fontWeight:700,color:"#0a1f44"}}>{x[1]}</span>
</div>)}
</div>
</div>
<div style={{textAlign:"center",fontSize:9,color:"#94A3B8",marginBottom:4}}>
Jl. Jendral Sudirman No. 80 · Banda Aceh &nbsp;|&nbsp; 081269002121 / (0651) 21221
</div>
</div>
{/* Buttons */}
<div style={{padding:"10px 16px",display:"flex",gap:8,flexWrap:"wrap",justifyContent:"center"}}>
<button onClick={()=>doPrint("_slip_setor")} style={{background:"#0a1f44",color:"white",border:"none",padding:"8px 16px",borderRadius:7,fontSize:12,cursor:"pointer",fontWeight:700}}>🖨️ Cetak / PDF</button>
<button onClick={()=>doDownloadPNG("_slip_setor","Slip-Setor-"+last.bank+"-"+last.tanggal+".png")} style={{background:"#1D6A96",color:"white",border:"none",padding:"8px 14px",borderRadius:7,fontSize:12,cursor:"pointer",fontWeight:700}}>💾 Download PNG</button>
<button onClick={()=>doCopyPNG("_slip_setor")} style={{background:"#145A32",color:"white",border:"none",padding:"8px 14px",borderRadius:7,fontSize:12,cursor:"pointer",fontWeight:700}}>📋 Copy PNG</button>
<button onClick={()=>setShowSlip(false)} style={{background:"#6B7280",color:"white",border:"none",padding:"8px 14px",borderRadius:7,fontSize:12,cursor:"pointer",fontWeight:700}}>✕ Tutup</button>
</div>
</div>
</div>;
})()}

</div>;
}


var ALL_TABS=[
{id:"dashboard",label:"Dashboard",icon:"🏠"},{id:"penjualan",label:"Penjualan",icon:"🧾"},
{id:"piutang",label:"Piutang",icon:"💳"},{id:"setoran",label:"Setoran",icon:"💰"},
{id:"laporan",label:"Laporan",icon:"📊"},{id:"stok",label:"Stok",icon:"📦"},
{id:"pengeluaran",label:"Pengeluaran",icon:"💸"},{id:"tutupbuku",label:"Tutup Buku",icon:"📒"},
{id:"pelanggan",label:"Pelanggan",icon:"👥"},{id:"karyawan",label:"Karyawan",icon:"👤"},
{id:"absensi",label:"Absensi & Payroll",icon:"📅"},{id:"jualanlain",label:"Jualan Lain",icon:"🛒"},
{id:"kas",label:"Buku Kas",icon:"📒"},{id:"do",label:"DO",icon:"🚚"},{id:"invoicemanual",label:"Invoice Manual",icon:"🧾"},{id:"fifo",label:"FIFO Detail",icon:"🔬",adminOnly:true},{id:"settings",label:"Pengaturan",icon:"⚙️"},
];
function getVisibleTabs(user){if(!user)return[];var rt=ROLE_TABS[user.role];var base=(rt===null||rt===undefined)?ALL_TABS:ALL_TABS.filter(t=>rt.includes(t.id));return base.filter(t=>!t.adminOnly||user.role==="admin");}

export default function App(){
var[theme,setTheme]=useState(()=>{try{var s=localStorage.getItem("hts_theme");return s||"light";}catch(e){return"light";}});
var C=THEMES[theme]||THEMES.light;
useEffect(()=>{try{localStorage.setItem("hts_theme",theme);}catch(e){}},[theme]);
var[data,setData]=useState(()=>{
try{var s=localStorage.getItem("lpg_mgmt_v4");
if(s){var d=JSON.parse(s);

if(!d.stokKosong||typeof d.stokKosong!=="object")d.stokKosong={"5.5 kg":0,"12 kg":0,"50 kg":0};
["5.5 kg","12 kg","50 kg"].forEach(function(s){if(typeof d.stokKosong[s]==="undefined")d.stokKosong[s]=0;});
if(!d.totalTabung)d.totalTabung={"5.5 kg":0,"12 kg":0,"50 kg":0};
// Recalc totalTabung = isi + kosong + titip (logika benar)
(function(){var SIZES_M=["5.5 kg","12 kg","50 kg"];SIZES_M.forEach(function(s){var isi=(d.stock||{})[s]||0;var kosong=(d.stokKosong||{})[s]||0;var titip=0;(d.titipList||[]).forEach(function(t){var items=t.items&&t.items.length>0?t.items:(t.ukuran===s?[{qty:t.qty}]:[]);var m=t.tipe==="titip"?1:t.tipe==="tarik"?-1:0;items.forEach(function(it){if(!t.items||t.items.length===0?true:it.ukuran===s)titip+=m*Number(it.qty||0);});});d.totalTabung[s]=isi+kosong+Math.max(0,titip);});})();
if(d.pelanggan)d.pelanggan=d.pelanggan.map(p=>({...p,hargaKhusus:Array.isArray(p.hargaKhusus)?p.hargaKhusus:[]}));
// Migration titipan flat format → items array
if(d.titipList)d.titipList=d.titipList.map(t=>{
  if(t.items&&t.items.length>0)return t;
  if(t.ukuran&&t.qty)return{...t,items:[{ukuran:t.ukuran,qty:Number(t.qty||0)}]};
  return t;
});
// Migration DO lama → status diterima (stok sudah masuk)
if(d.doList)d.doList=d.doList.map(x=>x.status?x:{...x,status:"diterima"});
if(!d.titipList)d.titipList=[];
if(!d.counters)d.counters={inv:{},sg:{},reg:0};
if(!d.company)d.company={...INIT.company};
if(!d.setoranLog)d.setoranLog=[];
Object.keys(INIT.company).forEach(k=>{if(d.company[k]==null)d.company[k]=INIT.company[k];});
if(!d.employees||d.employees.length<5)d.employees=DEF_EMP.slice();
d.employees=d.employees.map(e=>e.uangMakanMode?e:{...e,uangMakanMode:"harian"});
d.employees=d.employees.map(e=>e.uangMakan===15000?{...e,uangMakan:20000}:e);// update default uang makan 15k→20k
if(d.pelanggan&&d.pelanggan.length){var maxReg=d.counters.reg||0;d.pelanggan=d.pelanggan.map(p=>{if(p.regNo)return p;maxReg++;return{...p,regNo:"HTS/CST/"+String(maxReg).padStart(3,"0")};});d.counters.reg=Math.max(d.counters.reg||0,maxReg);}
var merged={...INIT,...d};
return ensureStokBatchInit(merged);}
// Try migrate from v3
try{var s3=localStorage.getItem("lpg_mgmt_v3");if(s3){var d3=JSON.parse(s3);return ensureStokBatchInit({...INIT,...d3,setoranLog:[]});}}catch(e){}
return ensureStokBatchInit({...INIT});}catch(e){return{...INIT};}
});
var[user,setUser]=useState(null);var[tab,setTab]=useState("dashboard");var[inv,setInv]=useState(null);var[sideOpen,setSideOpen]=useState(false);
var[syncStatus,setSyncStatus]=useState("idle");// idle|pushing|pulling|ok|error
var[syncMsg,setSyncMsg]=useState("");

async function handlePush(){
  setSyncMsg("Menyimpan ke cloud...");
  var ok=await pushAll(data,setSyncStatus);
  setSyncMsg(ok?"✅ Tersimpan ke Google Sheets!":"❌ Gagal push. Cek koneksi.");
  setTimeout(()=>setSyncMsg(""),4000);
}
async function handlePull(){
  setSyncMsg("Mengambil data dari cloud...");
  var pulled=await pullAll(setSyncStatus);
  if(pulled){
    setData(prev=>({...prev,...pulled,
      employees:pulled.employees&&pulled.employees.length>0?pulled.employees:prev.employees
    }));
    setSyncMsg("✅ Data berhasil diambil dari cloud!");
  }else{
    setSyncMsg("❌ Gagal pull. Cek koneksi.");
  }
  setTimeout(()=>setSyncMsg(""),4000);
}
var{toasts,toast}=useToast();var mobile=useMobile();
useEffect(()=>{try{localStorage.setItem("lpg_mgmt_v4",JSON.stringify(data));}catch(e){console.warn("Storage full");}},[data]);
// Auto-pull dari cloud sekali setiap kali aplikasi dibuka, supaya device baru/browser baru selalu dapat data terbaru
useEffect(()=>{
(async()=>{
  setSyncMsg("☁️ Menyinkronkan data dari cloud...");
  var pulled=await pullAll(setSyncStatus);
  if(pulled){
    setData(prev=>{
      var merged={...prev};
      Object.keys(pulled).forEach(k=>{
        var cloudVal=pulled[k];
        if(Array.isArray(cloudVal)){
          merged[k]=cloudVal;// array kosong dari cloud tetap dipakai (artinya memang belum ada data)
        }else if(cloudVal&&typeof cloudVal==="object"&&Object.keys(cloudVal).length>0){
          merged[k]={...merged[k],...cloudVal};
        }
      });
      if(pulled.employees&&pulled.employees.length===0)merged.employees=prev.employees;// jaga employees default jika cloud belum pernah diisi
      return merged;
    });
    setSyncMsg("✅ Data tersinkron dari cloud");
  }else{
    setSyncMsg("⚠️ Gagal sinkron — pakai data lokal terakhir");
  }
  setTimeout(()=>setSyncMsg(""),3000);
})();
// eslint-disable-next-line react-hooks/exhaustive-deps
},[]);
var setDataP=useCallback(updater=>{setData(prev=>typeof updater==="function"?updater(prev):updater);},[]);
var themeToggle=()=>setTheme(t=>t==="light"?"dark":"light");
if(!user)return <ThemeCtx.Provider value={C}><LoginScreen employees={data.employees||DEF_EMP} onLogin={u=>{setUser(u);setTab("dashboard");}} themeToggle={themeToggle} theme={theme}/><Toast toasts={toasts}/>
{syncMsg&&<div style={{position:"fixed",top:60,left:"50%",transform:"translateX(-50%)",background:syncMsg.startsWith("✅")?C.grn:syncMsg.startsWith("❌")?C.rdk:C.org,color:"#fff",padding:"8px 20px",borderRadius:20,fontSize:13,fontWeight:700,zIndex:9998,boxShadow:"0 4px 16px rgba(0,0,0,.2)",whiteSpace:"nowrap"}}>{syncMsg}</div>}</ThemeCtx.Provider>;
var tabs=getVisibleTabs(user);var curTab=tabs.find(t=>t.id===tab)||tabs[0];
var bonAlerts=(data.bon||[]).filter(b=>b.status!=="lunas"&&b.deadline&&dLeft(b.deadline)!==null&&dLeft(b.deadline)<=3).length;
function renderContent(){
var t=curTab?.id;var props={data,setData:setDataP,user,toast};
if(t==="dashboard")return <Dashboard data={data} setTab={setTab} user={user}/>;
if(t==="penjualan")return <PenjualanMod {...props} setInv={setInv}/>;
if(t==="piutang")return <PiutangMod {...props} setInv={setInv}/>;
if(t==="setoran")return <SetoranMod {...props}/>;
if(t==="laporan")return <LaporanMod data={data} toast={toast}/>;
if(t==="stok")return <StokMod {...props}/>;
if(t==="pengeluaran")return <PengeluaranMod {...props}/>;
if(t==="tutupbuku")return <TutupBukuMod {...props}/>;
if(t==="pelanggan")return <PelangganMod {...props}/>;
if(t==="karyawan")return <KaryawanMod {...props}/>;
if(t==="absensi")return <AbsensiPayrollMod {...props}/>;
if(t==="absensi")return <PayrollMod {...props}/>;
if(t==="jualanlain")return <JualanLainMod {...props}/>;
if(t==="fifo")return user?.role==="admin"?<FIFODetailMod {...props}/>:null;
if(t==="kas")return <KasBankMod {...props}/>;
if(t==="do")return <DOMod {...props}/>;
if(t==="invoicemanual")return <InvoiceManualMod {...props} setInv={setInv}/>;
if(t==="settings")return <SettingsMod data={data} setData={setDataP} toast={toast} theme={theme} setTheme={setTheme}/>;
return null;}
return <ThemeCtx.Provider value={C}>
<div style={{minHeight:"100vh",background:C.bg,color:C.wht,fontFamily:"'Segoe UI',system-ui,sans-serif",display:"flex",flexDirection:"column"}}>
<Toast toasts={toasts}/>
<div style={{background:C.card,borderBottom:"1px solid "+C.bdr,padding:"0 12px",display:"flex",alignItems:"center",justifyContent:"space-between",height:54,position:"sticky",top:0,zIndex:100,flexShrink:0,boxShadow:C.mode==="light"?"0 1px 3px rgba(0,0,0,0.05)":"none"}}>
<div style={{display:"flex",alignItems:"center",gap:10}}>
<button onClick={()=>setSideOpen(!sideOpen)} style={{background:"none",border:"none",color:C.gl2,fontSize:22,cursor:"pointer",padding:"8px",borderRadius:6}}>☰</button>
<div style={{display:"flex",alignItems:"center",gap:8}}>
{<CompanyLogo h={32} variant={C.mode==="dark"?"dark":"light"}/>}
</div>
</div>
<div style={{display:"flex",alignItems:"center",gap:12}}>
{!mobile&&<PertaminaLogo h={22} variant={C.mode==="dark"?"dark":"light"}/>}
<div style={{display:"flex",alignItems:"center",gap:8}}>
{bonAlerts>0&&<div style={{background:C.mode==="dark"?"#3D1A05":"#FFEDD5",border:"1px solid "+C.olt,borderRadius:20,padding:"3px 9px",fontSize:11,color:C.olt,fontWeight:700}}>⚠️ {bonAlerts}</div>}
<button onClick={themeToggle} style={{background:C.nav,border:"1px solid "+C.bdr,color:C.gl2,borderRadius:8,padding:"6px 10px",cursor:"pointer",fontSize:14}}>{theme==="light"?"🌙":"☀️"}</button>
<button onClick={handlePush} title="Simpan ke Google Sheets" style={{background:syncStatus==="pushing"?C.org:syncStatus==="ok"?C.grn:syncStatus==="error"?C.rdk:C.nav,border:"1px solid "+C.bdr,color:syncStatus==="idle"?C.gl2:"#fff",borderRadius:8,padding:"6px 10px",cursor:"pointer",fontSize:13,fontWeight:700}}>{syncStatus==="pushing"?"⏳":syncStatus==="pulling"?"⏳":"☁️"}</button>
<button onClick={handlePull} title="Ambil data dari Google Sheets" style={{background:C.nav,border:"1px solid "+C.bdr,color:C.gl2,borderRadius:8,padding:"6px 10px",cursor:"pointer",fontSize:13}}>⬇️</button>
<div style={{textAlign:"right"}}><div style={{fontSize:12,fontWeight:700,color:C.wht}}>{user.nama.split(" ")[0]}</div>{!mobile&&<div style={{fontSize:10,color:C.gl2}}>{user.posisi}</div>}</div>
<button onClick={()=>setUser(null)} style={{background:C.rdk,border:"none",borderRadius:8,color:"#FFF",fontSize:11,padding:"8px 10px",cursor:"pointer",fontWeight:700}}>⏻</button>
</div>
</div>
</div>
<div style={{flex:1,position:"relative",minHeight:0}}>
{sideOpen&&<div onClick={()=>setSideOpen(false)} style={{position:"fixed",inset:0,background:"rgba(0,0,0,.5)",zIndex:90}}/>}
<div style={{width:220,background:C.card,borderRight:"1px solid "+C.bdr,position:"fixed",top:54,bottom:0,left:0,transform:sideOpen?"translateX(0)":"translateX(-100%)",transition:"transform .22s",zIndex:95,overflowY:"auto"}}>
<div style={{padding:"10px 14px 4px",fontSize:10,color:C.gry,fontWeight:700,letterSpacing:1}}>MENU</div>
{tabs.map(t=><button key={t.id} onClick={()=>{setTab(t.id);setSideOpen(false);}} style={{display:"flex",alignItems:"center",gap:10,width:"100%",padding:"12px 16px",background:curTab?.id===t.id?(C.mode==="dark"?"rgba(41,128,185,.2)":"#DBEAFE"):"none",border:"none",color:curTab?.id===t.id?C.blt:C.gl2,fontSize:13,fontWeight:curTab?.id===t.id?700:400,cursor:"pointer",textAlign:"left",borderLeft:curTab?.id===t.id?"3px solid "+C.blt:"3px solid transparent"}}><span style={{fontSize:15}}>{t.icon}</span>{t.label}</button>)}
<div style={{padding:"10px 14px",marginTop:8,borderTop:"1px solid "+C.bdr}}><div style={{fontSize:11,color:C.gl2}}>Login sebagai</div><div style={{fontSize:13,fontWeight:700,color:C.wht}}>{user.nama}</div><div style={{fontSize:10,color:C.gl2}}>{user.posisi}</div></div>
</div>
<div style={{overflowY:"auto",height:"100%",padding:mobile?"10px":"14px",boxSizing:"border-box"}}>
<div style={{display:"flex",gap:5,marginBottom:14,overflowX:"auto",paddingBottom:4}}>
{tabs.map(t=><button key={t.id} onClick={()=>setTab(t.id)} style={{background:curTab?.id===t.id?C.blu:C.card,color:curTab?.id===t.id?"#FFF":C.gl2,border:"1px solid "+(curTab?.id===t.id?C.blt:C.bdr),borderRadius:8,padding:mobile?"8px 10px":"6px 13px",fontSize:mobile?11:12,fontWeight:curTab?.id===t.id?700:400,cursor:"pointer",whiteSpace:"nowrap",flexShrink:0}}>{t.icon}{!mobile&&" "+t.label}</button>)}
</div>
<div style={{maxWidth:mobile?980:(data.company?.appWidth||1600),margin:"0 auto"}}>{renderContent()}</div>
</div>
</div>
{inv&&<InvoiceView inv={inv} company={data.company} onClose={()=>setInv(null)}/>}
</div>
</ThemeCtx.Provider>;
}
// === SELESAI BAGIAN 4 — FILE v4 LENGKAP ===
console.log("[HTS-APP] Build: v4.3-fix-stok-titip — DO/Penj reverse, TitipTab fixed");

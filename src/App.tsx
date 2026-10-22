/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useRef } from 'react';
import {
  Monitor,
  Tablet,
  Smartphone,
  RotateCw,
  ExternalLink,
  Maximize2,
  Minimize2,
  KeyRound,
  FileCode2,
  CheckCircle2,
  Copy,
  Receipt,
  Printer,
  Table,
  HelpCircle,
  Bluetooth,
  AlertTriangle,
  Flame,
  ArrowRight
} from 'lucide-react';

const FULL_APPS_SCRIPT_TEMPLATE = `/**
 * GOOGLE APPS SCRIPT BACKEND FOR ²AL JAYA FROZEN FOOD POS
 * Paste kode ini di Google Sheets > Ekstensi > Apps Script
 * Lalu klik Terapkan (Deploy) > Deployment Baru > Jenis: Aplikasi Web
 * Jalankan sebagai: Saya (Me)
 * Akses: Siapa saja (Anyone)
 */

function getOrCreateSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers && headers.length) {
      sheet.appendRow(headers);
      var headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#0A4D7A').setFontColor('#FFFFFF').setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'all';
    
    var prodSheet = getOrCreateSheet(ss, 'Produk', ['ID', 'Nama', 'Kategori', 'Harga', 'Stok', 'Satuan', 'Barcode', 'Deskripsi', 'Gambar', 'Update_Terakhir']);
    var trxSheet = getOrCreateSheet(ss, 'Transaksi', ['ID_Transaksi', 'Tanggal', 'Pelanggan', 'Metode_Bayar', 'Kasir', 'Item_Ringkasan', 'Total', 'Diskon', 'Uang_Diterima', 'Kembalian']);
    var userSheet = getOrCreateSheet(ss, 'Users', ['ID', 'Nama', 'Username', 'Password', 'Role']);

    var result = { ok: true, timestamp: new Date().toISOString() };

    if (action === 'all' || action === 'products') {
      var prodData = prodSheet.getDataRange().getValues();
      var products = [];
      for (var i = 1; i < prodData.length; i++) {
        var row = prodData[i];
        if (row[0]) {
          products.push({
            id: Number(row[0]) || row[0],
            name: String(row[1] || ''),
            cat: String(row[2] || 'Lainnya'),
            price: Number(row[3]) || 0,
            stock: Number(row[4]) || 0,
            unit: String(row[5] || 'pack'),
            barcode: String(row[6] || ''),
            desc: String(row[7] || ''),
            img: String(row[8] || '')
          });
        }
      }
      result.products = products;
    }

    if (action === 'all' || action === 'transactions') {
      var trxData = trxSheet.getDataRange().getValues();
      var transactions = [];
      for (var j = 1; j < trxData.length; j++) {
        var r = trxData[j];
        if (r[0]) {
          transactions.push({
            id: String(r[0]),
            date: String(r[1]),
            customer: String(r[2]),
            pay: String(r[3]),
            kasir: String(r[4]),
            itemsDesc: String(r[5]),
            total: Number(r[6]) || 0,
            disc: Number(r[7]) || 0,
            cash: Number(r[8]) || 0,
            change: Number(r[9]) || 0
          });
        }
      }
      result.transactions = transactions;
    }

    if (action === 'all' || action === 'users') {
      var uData = userSheet.getDataRange().getValues();
      var users = [];
      for (var k = 1; k < uData.length; k++) {
        var ur = uData[k];
        if (ur[0]) {
          users.push({
            id: Number(ur[0]) || ur[0],
            name: String(ur[1]),
            username: String(ur[2]),
            pass: String(ur[3]),
            role: String(ur[4])
          });
        }
      }
      result.users = users;
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var payload = {};
    if (e && e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    }
    var action = payload.action || 'sync';

    var prodSheet = getOrCreateSheet(ss, 'Produk', ['ID', 'Nama', 'Kategori', 'Harga', 'Stok', 'Satuan', 'Barcode', 'Deskripsi', 'Gambar', 'Update_Terakhir']);
    var trxSheet = getOrCreateSheet(ss, 'Transaksi', ['ID_Transaksi', 'Tanggal', 'Pelanggan', 'Metode_Bayar', 'Kasir', 'Item_Ringkasan', 'Total', 'Diskon', 'Uang_Diterima', 'Kembalian']);
    var userSheet = getOrCreateSheet(ss, 'Users', ['ID', 'Nama', 'Username', 'Password', 'Role']);

    if (action === 'sync') {
      if (payload.products && Array.isArray(payload.products)) {
        var lastRow = prodSheet.getLastRow();
        if (lastRow > 1) {
          prodSheet.getRange(2, 1, lastRow - 1, 10).clearContent();
        }
        var prodRows = payload.products.map(function(p) {
          return [p.id, p.name, p.cat, p.price, p.stock, p.unit, p.barcode || '', p.desc || '', p.img || '', new Date()];
        });
        if (prodRows.length > 0) {
          prodSheet.getRange(2, 1, prodRows.length, 10).setValues(prodRows);
        }
      }

      if (payload.transactions && Array.isArray(payload.transactions)) {
        var lastTrxRow = trxSheet.getLastRow();
        if (lastTrxRow > 1) {
          trxSheet.getRange(2, 1, lastTrxRow - 1, 10).clearContent();
        }
        var trxRows = payload.transactions.map(function(t) {
          var itemsDesc = Array.isArray(t.items) ? t.items.map(function(it){ return it.name + ' (' + it.qty + 'x)'; }).join(', ') : (t.itemsDesc || '');
          return [t.id, t.date, t.customer, t.pay, t.kasir, itemsDesc, t.total, t.disc || 0, t.cash || t.total, t.change || 0];
        });
        if (trxRows.length > 0) {
          trxSheet.getRange(2, 1, trxRows.length, 10).setValues(trxRows);
        }
      }

      if (payload.users && Array.isArray(payload.users)) {
        var lastUserRow = userSheet.getLastRow();
        if (lastUserRow > 1) {
          userSheet.getRange(2, 1, lastUserRow - 1, 5).clearContent();
        }
        var userRows = payload.users.map(function(u) {
          return [u.id, u.name, u.username, u.password || u.pass, u.role];
        });
        if (userRows.length > 0) {
          userSheet.getRange(2, 1, userRows.length, 5).setValues(userRows);
        }
      }

      return ContentService.createTextOutput(JSON.stringify({ ok: true, message: 'Sync berhasil tersimpan di Google Sheets' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({ ok: true, message: 'Aksi diproses' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}`;

export default function App() {
  const [deviceMode, setDeviceMode] = useState<'desktop' | 'tablet' | 'mobile'>('desktop');
  const [isFullscreen, setIsFullscreen] = useState(false);
  const [showInfoModal, setShowInfoModal] = useState(false);
  const [showSheetsModal, setShowSheetsModal] = useState(false);
  const [showPrinterModal, setShowPrinterModal] = useState(false);
  const [copiedKey, setCopiedKey] = useState<string | null>(null);
  const iframeRef = useRef<HTMLIFrameElement>(null);

  const fileAssetUrl = '/assets/aistudio/index_c585c777.html';

  const handleReload = () => {
    if (iframeRef.current) {
      iframeRef.current.src = `${fileAssetUrl}?t=${Date.now()}`;
    }
  };

  const copyToClipboard = (text: string, keyId: string) => {
    navigator.clipboard.writeText(text);
    setCopiedKey(keyId);
    setTimeout(() => setCopiedKey(null), 2000);
  };

  return (
    <div className="flex flex-col h-screen w-screen bg-slate-900 text-slate-100 font-sans overflow-hidden select-none">
      {/* Top Navigation & Controls Bar */}
      {!isFullscreen && (
        <header className="h-14 bg-slate-950 border-b border-slate-800 px-3 sm:px-4 flex items-center justify-between shrink-0 z-30 shadow-md">
          {/* Logo & Info */}
          <div className="flex items-center gap-2 sm:gap-3">
            <span className="w-8 h-8 rounded-lg bg-sky-600 flex items-center justify-center font-black text-white text-base shadow-sm">
              ²AL
            </span>
            <div className="leading-tight hidden sm:block">
              <h1 className="text-sm font-bold text-slate-100 tracking-wide flex items-center gap-2">
                ²AL JAYA Frozen Food
                <span className="bg-emerald-500/20 text-emerald-400 text-[10px] font-semibold px-2 py-0.5 rounded-full border border-emerald-500/30">
                  Sheets & Bluetooth AT-58X
                </span>
              </h1>
              <p className="text-[11px] text-slate-400">
                POS, Manajemen Stok, Cetak Struk & Cloud Sync
              </p>
            </div>
          </div>

          {/* Device & Viewport Switchers */}
          <div className="flex items-center bg-slate-900 border border-slate-800 rounded-lg p-0.5 shadow-inner">
            <button
              onClick={() => setDeviceMode('desktop')}
              className={`flex items-center gap-1.5 px-3 py-1.5 rounded-md text-xs font-medium transition-colors ${
                deviceMode === 'desktop'
                  ? 'bg-sky-600 text-white shadow-sm'
                  : 'text-slate-400 hover:text-slate-200 hover:bg-slate-800/60'
              }`}
              title="Tampilan Desktop / POS Monitor"
            >
              <Monitor className="w-3.5 h-3.5" />
              <span className="hidden md:inline">Desktop</span>
            </button>
            <button
              onClick={() => setDeviceMode('tablet')}
              className={`flex items-center gap-1.5 px-3 py-1.5 rounded-md text-xs font-medium transition-colors ${
                deviceMode === 'tablet'
                  ? 'bg-sky-600 text-white shadow-sm'
                  : 'text-slate-400 hover:text-slate-200 hover:bg-slate-800/60'
              }`}
              title="Tampilan Tablet POS (768px)"
            >
              <Tablet className="w-3.5 h-3.5" />
              <span className="hidden md:inline">Tablet</span>
            </button>
            <button
              onClick={() => setDeviceMode('mobile')}
              className={`flex items-center gap-1.5 px-3 py-1.5 rounded-md text-xs font-medium transition-colors ${
                deviceMode === 'mobile'
                  ? 'bg-sky-600 text-white shadow-sm'
                  : 'text-slate-400 hover:text-slate-200 hover:bg-slate-800/60'
              }`}
              title="Tampilan Mobile / Smartphone POS (420px)"
            >
              <Smartphone className="w-3.5 h-3.5" />
              <span className="hidden md:inline">Mobile</span>
            </button>
          </div>

          {/* Quick Guide & Action Buttons */}
          <div className="flex items-center gap-1.5 sm:gap-2">
            {/* Google Sheets Guide Button */}
            <button
              onClick={() => setShowSheetsModal(true)}
              className="flex items-center gap-1.5 px-2.5 py-1.5 bg-emerald-950/60 hover:bg-emerald-900/80 text-emerald-300 rounded-md text-xs font-semibold border border-emerald-800/60 transition shadow-sm"
              title="Panduan Google Sheets"
            >
              <Table className="w-3.5 h-3.5 text-emerald-400" />
              <span className="hidden lg:inline">Panduan Sheets</span>
            </button>

            {/* AT-58X Printer Guide Button */}
            <button
              onClick={() => setShowPrinterModal(true)}
              className="flex items-center gap-1.5 px-2.5 py-1.5 bg-sky-950/60 hover:bg-sky-900/80 text-sky-300 rounded-md text-xs font-semibold border border-sky-800/60 transition shadow-sm"
              title="Panduan Printer AT-58X"
            >
              <Printer className="w-3.5 h-3.5 text-sky-400" />
              <span className="hidden lg:inline">Printer AT-58X</span>
            </button>

            {/* Demo Accounts Button */}
            <button
              onClick={() => setShowInfoModal(true)}
              className="flex items-center gap-1.5 px-2.5 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-200 rounded-md text-xs font-medium border border-slate-700 transition"
              title="Informasi Akun Demo"
            >
              <KeyRound className="w-3.5 h-3.5 text-amber-400" />
              <span className="hidden sm:inline">Akun Demo</span>
            </button>

            <button
              onClick={handleReload}
              className="p-1.5 text-slate-400 hover:text-slate-200 hover:bg-slate-800 rounded-md transition"
              title="Muat Ulang Pratinjau"
            >
              <RotateCw className="w-4 h-4" />
            </button>

            <a
              href={fileAssetUrl}
              target="_blank"
              rel="noreferrer"
              className="p-1.5 text-slate-400 hover:text-slate-200 hover:bg-slate-800 rounded-md transition"
              title="Buka Aplikasi di Tab Penuh Baru"
            >
              <ExternalLink className="w-4 h-4" />
            </a>

            <button
              onClick={() => setIsFullscreen(true)}
              className="p-1.5 text-slate-400 hover:text-slate-200 hover:bg-slate-800 rounded-md transition"
              title="Mode Layar Penuh"
            >
              <Maximize2 className="w-4 h-4" />
            </button>
          </div>
        </header>
      )}

      {/* Fullscreen Floating Exit Button */}
      {isFullscreen && (
        <button
          onClick={() => setIsFullscreen(false)}
          className="fixed top-3 right-3 z-50 bg-slate-900/90 hover:bg-slate-900 text-white px-3 py-1.5 rounded-full text-xs font-semibold backdrop-blur border border-slate-700 shadow-xl flex items-center gap-1.5 transition"
        >
          <Minimize2 className="w-3.5 h-3.5" />
          Keluar Layar Penuh
        </button>
      )}

      {/* Main Preview Container */}
      <main className="flex-1 bg-slate-950/80 p-2 sm:p-4 flex items-center justify-center overflow-auto relative">
        <div
          className={`transition-all duration-300 ease-in-out h-full flex flex-col bg-white shadow-2xl overflow-hidden ${
            isFullscreen
              ? 'w-full h-full rounded-none'
              : deviceMode === 'desktop'
              ? 'w-full h-full rounded-xl border border-slate-800'
              : deviceMode === 'tablet'
              ? 'w-[768px] h-full rounded-2xl border-4 border-slate-800 my-auto'
              : 'w-[420px] h-full rounded-3xl border-8 border-slate-800 my-auto'
          }`}
        >
          {/* Subtle Mobile/Tablet Header Bar Decoration */}
          {!isFullscreen && (deviceMode === 'mobile' || deviceMode === 'tablet') && (
            <div className="h-6 bg-slate-800 flex items-center justify-center shrink-0">
              <div className="w-16 h-1 bg-slate-600 rounded-full"></div>
            </div>
          )}

          {/* Direct Iframe to index_c585c777.html */}
          <iframe
            ref={iframeRef}
            src={fileAssetUrl}
            title="Preview ²AL JAYA Frozen Food POS"
            className="w-full h-full border-0 flex-1 bg-[#F2F8FC]"
            allow="camera; microphone; bluetooth; serial; usb; clipboard-write; display-capture; fullscreen"
          />
        </div>
      </main>

      {/* MODAL 1: GOOGLE SHEETS SETUP GUIDE */}
      {showSheetsModal && (
        <div className="fixed inset-0 z-50 bg-black/75 backdrop-blur-sm flex items-center justify-center p-4 overflow-y-auto">
          <div className="bg-slate-900 border border-slate-800 rounded-2xl max-w-2xl w-full p-6 text-slate-100 shadow-2xl relative my-8">
            <button
              onClick={() => setShowSheetsModal(false)}
              className="absolute top-4 right-4 text-slate-400 hover:text-white text-lg font-bold w-8 h-8 rounded-full flex items-center justify-center hover:bg-slate-800"
            >
              ✕
            </button>

            <div className="flex items-center gap-3 mb-4">
              <div className="w-10 h-10 rounded-xl bg-emerald-500/20 text-emerald-400 flex items-center justify-center text-xl">
                <Table className="w-5 h-5" />
              </div>
              <div>
                <h3 className="text-base font-bold text-white flex items-center gap-2">
                  Panduan Sinkronisasi Database Google Sheets
                </h3>
                <p className="text-xs text-slate-400">
                  Langkah-langkah menghubungkan toko ²AL JAYA langsung dengan Google Spreadsheet Anda.
                </p>
              </div>
            </div>

            {/* Step-by-Step Instructions */}
            <div className="space-y-4 text-xs text-slate-300">
              {/* Step 1 */}
              <div className="flex gap-3 items-start bg-slate-800/40 p-3 rounded-xl border border-slate-800">
                <span className="w-6 h-6 rounded-full bg-emerald-600 text-white font-bold flex items-center justify-center shrink-0">1</span>
                <div>
                  <b className="text-white text-sm">Buat Spreadsheet Baru</b>
                  <p className="text-slate-400 mt-0.5">Buka link Google Spreadsheet di tab baru:</p>
                  <a
                    href="https://sheets.new"
                    target="_blank"
                    rel="noreferrer"
                    className="inline-flex items-center gap-1.5 bg-emerald-600/30 hover:bg-emerald-600/50 text-emerald-300 border border-emerald-500/40 font-bold px-3 py-1 rounded-md text-xs mt-1.5 transition"
                  >
                    Buka: sheets.new <ExternalLink className="w-3 h-3" />
                  </a>
                </div>
              </div>

              {/* Step 2 */}
              <div className="flex gap-3 items-start bg-slate-800/40 p-3 rounded-xl border border-slate-800">
                <span className="w-6 h-6 rounded-full bg-emerald-600 text-white font-bold flex items-center justify-center shrink-0">2</span>
                <div>
                  <b className="text-white text-sm">Buka Apps Script</b>
                  <p className="text-slate-400 mt-0.5">
                    Pada menu atas Spreadsheet Anda, klik <strong className="text-white">Ekstensi (Extensions)</strong> &gt; <strong className="text-white">Apps Script</strong>.
                  </p>
                </div>
              </div>

              {/* Step 3 */}
              <div className="flex gap-3 items-start bg-slate-800/40 p-3 rounded-xl border border-slate-800">
                <span className="w-6 h-6 rounded-full bg-emerald-600 text-white font-bold flex items-center justify-center shrink-0">3</span>
                <div className="flex-1">
                  <div className="flex items-center justify-between">
                    <b className="text-white text-sm">Tempel Kode Apps Script</b>
                    <button
                      onClick={() => copyToClipboard(FULL_APPS_SCRIPT_TEMPLATE, 'scriptTemplate')}
                      className="px-2.5 py-1 bg-emerald-600 hover:bg-emerald-500 text-white rounded-md text-xs font-semibold flex items-center gap-1 transition"
                    >
                      {copiedKey === 'scriptTemplate' ? (
                        <>
                          <CheckCircle2 className="w-3.5 h-3.5" />
                          <span>Tersalin!</span>
                        </>
                      ) : (
                        <>
                          <Copy className="w-3.5 h-3.5" />
                          <span>Salin Kode Script</span>
                        </>
                      )}
                    </button>
                  </div>
                  <p className="text-slate-400 mt-1">
                    Hapus seluruh isi <code className="bg-slate-800 px-1 py-0.5 rounded text-amber-300 font-mono">myFunction()</code> di editor Apps Script, lalu tempel kode yang sudah disalin di atas.
                  </p>
                  <pre className="bg-slate-950 p-2.5 rounded-lg text-[10px] font-mono text-emerald-400 h-24 overflow-y-auto mt-2 border border-slate-800">
                    {FULL_APPS_SCRIPT_TEMPLATE.slice(0, 420)}...
                  </pre>
                </div>
              </div>

              {/* Step 4 */}
              <div className="flex gap-3 items-start bg-slate-800/40 p-3 rounded-xl border border-slate-800">
                <span className="w-6 h-6 rounded-full bg-emerald-600 text-white font-bold flex items-center justify-center shrink-0">4</span>
                <div>
                  <b className="text-white text-sm">Deploy sebagai Aplikasi Web (Web App)</b>
                  <p className="text-slate-400 mt-0.5">
                    Klik tombol biru <strong className="text-white">Terapkan (Deploy)</strong> &gt; <strong className="text-white">Deployment baru (New deployment)</strong>:
                  </p>
                  <ul className="list-disc ml-5 mt-1 text-slate-300 space-y-0.5">
                    <li>Pilih jenis roda gigi: <strong className="text-white">Aplikasi Web (Web App)</strong></li>
                    <li>Jalankan sebagai: <strong className="text-white">Saya (Me)</strong></li>
                    <li>
                      Siapa yang memiliki akses: <strong className="text-emerald-400">Siapa saja (Anyone)</strong> *(Wajib agar tidak error)*
                    </li>
                  </ul>
                </div>
              </div>

              {/* Step 5 */}
              <div className="flex gap-3 items-start bg-slate-800/40 p-3 rounded-xl border border-slate-800">
                <span className="w-6 h-6 rounded-full bg-emerald-600 text-white font-bold flex items-center justify-center shrink-0">5</span>
                <div>
                  <b className="text-white text-sm">Masukkan URL Web App ke Aplikasi</b>
                  <p className="text-slate-400 mt-0.5">
                    Salin URL yang berakhiran <code className="bg-slate-800 text-amber-300 px-1 py-0.5 rounded font-mono">/exec</code>.
                    Buka menu <strong className="text-white">Google Sheets</strong> di dalam aplikasi, tempel URL-nya di kolom input dan klik <strong className="text-emerald-400">Simpan &amp; Sync</strong>.
                  </p>
                </div>
              </div>
            </div>

            <div className="mt-5 flex justify-end">
              <button
                onClick={() => setShowSheetsModal(false)}
                className="bg-emerald-600 hover:bg-emerald-500 text-white px-5 py-2 rounded-xl font-bold text-xs transition"
              >
                Tutup &amp; Mulai Menggunakan
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL 2: AT-58X PRINTER PAIRING GUIDE */}
      {showPrinterModal && (
        <div className="fixed inset-0 z-50 bg-black/75 backdrop-blur-sm flex items-center justify-center p-4 overflow-y-auto">
          <div className="bg-slate-900 border border-slate-800 rounded-2xl max-w-xl w-full p-6 text-slate-100 shadow-2xl relative my-8">
            <button
              onClick={() => setShowPrinterModal(false)}
              className="absolute top-4 right-4 text-slate-400 hover:text-white text-lg font-bold w-8 h-8 rounded-full flex items-center justify-center hover:bg-slate-800"
            >
              ✕
            </button>

            <div className="flex items-center gap-3 mb-4">
              <div className="w-10 h-10 rounded-xl bg-sky-500/20 text-sky-400 flex items-center justify-center text-xl">
                <Printer className="w-5 h-5" />
              </div>
              <div>
                <h3 className="text-base font-bold text-white flex items-center gap-2">
                  Panduan Pairing Printer Thermal Bluetooth AT-58X
                </h3>
                <p className="text-xs text-slate-400">
                  Langkah koneksi printer kasir thermal 58mm via Bluetooth ESC/POS &amp; RawBT.
                </p>
              </div>
            </div>

            {/* Printer Details */}
            <div className="space-y-3.5 text-xs text-slate-300">
              <div className="bg-sky-950/40 border border-sky-800/60 rounded-xl p-3.5">
                <p className="font-bold text-sky-300 flex items-center gap-1.5 mb-1">
                  <Bluetooth className="w-4 h-4 text-sky-400" />
                  Kredensial &amp; Detail Printer AT-58X:
                </p>
                <div className="grid grid-cols-2 gap-2 mt-2 font-mono text-[11px]">
                  <div className="bg-slate-900/80 p-2 rounded-lg">
                    <span className="text-slate-400 block">Nama Bluetooth:</span>
                    <b className="text-white">AT-58X / POS-58 / MPT-II</b>
                  </div>
                  <div className="bg-slate-900/80 p-2 rounded-lg">
                    <span className="text-slate-400 block">PIN Pairing Default:</span>
                    <b className="text-amber-400">1234 atau 0000</b>
                  </div>
                  <div className="bg-slate-900/80 p-2 rounded-lg">
                    <span className="text-slate-400 block">Lebar Kertas:</span>
                    <b className="text-white">58mm (Roll Thermal)</b>
                  </div>
                  <div className="bg-slate-900/80 p-2 rounded-lg">
                    <span className="text-slate-400 block">Protokol Cetak:</span>
                    <b className="text-white">ESC/POS Command</b>
                  </div>
                </div>
              </div>

              {/* Instructions */}
              <div className="space-y-2">
                <b className="text-white block">Cara Menghubungkan Printer:</b>
                <ol className="list-decimal ml-4 space-y-1.5 text-slate-300">
                  <li>
                    <strong className="text-white">Nyalakan Printer:</strong> Tekan tombol Power hingga lampu LED indikator hijau/biru menyala stabil.
                  </li>
                  <li>
                    <strong className="text-white">Pasang Kertas Thermal:</strong> Buka tutup printer, pasang rol kertas thermal ukuran 58mm (sisi licin menghadap ke atas/head printer).
                  </li>
                  <li>
                    <strong className="text-white">Pairing di Perangkat:</strong> Buka pengaturan Bluetooth di HP atau laptop, cari perangkat bernama <code className="bg-slate-800 text-sky-300 px-1 py-0.5 rounded font-mono">AT-58X</code> atau <code className="bg-slate-800 text-sky-300 px-1 py-0.5 rounded font-mono">POS-58</code>, lalu masukkan PIN <code className="bg-slate-800 text-amber-300 px-1 py-0.5 rounded font-mono">1234</code> atau <code className="bg-slate-800 text-amber-300 px-1 py-0.5 rounded font-mono">0000</code>.
                  </li>
                  <li>
                    <strong className="text-white">Cetak dari Aplikasi:</strong> Di aplikasi kasir, klik tombol <strong className="text-sky-300">"Setting Printer"</strong> di kasir atau tombol <strong className="text-sky-300">"PRINT AT-58X"</strong> saat transaksi selesai.
                  </li>
                  <li>
                    <strong className="text-white">Opsi Tambahan di Android (RawBT):</strong> Jika browser membatasi Web Bluetooth, Anda juga dapat menginstal aplikasi gratis <strong className="text-white">RawBT Print Service</strong> di Play Store dan menggunakan tombol <strong className="text-emerald-400">"RawBT App"</strong>.
                  </li>
                </ol>
              </div>
            </div>

            <div className="mt-5 flex justify-end">
              <button
                onClick={() => setShowPrinterModal(false)}
                className="bg-sky-600 hover:bg-sky-500 text-white px-5 py-2 rounded-xl font-bold text-xs transition"
              >
                Tutup Panduan Printer
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL 3: DEMO CREDENTIALS MODAL */}
      {showInfoModal && (
        <div className="fixed inset-0 z-50 bg-black/70 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-800 rounded-2xl max-w-lg w-full p-6 text-slate-100 shadow-2xl relative">
            <button
              onClick={() => setShowInfoModal(false)}
              className="absolute top-4 right-4 text-slate-400 hover:text-white text-lg font-bold w-8 h-8 rounded-full flex items-center justify-center hover:bg-slate-800"
            >
              ✕
            </button>

            <div className="flex items-center gap-3 mb-4">
              <div className="w-10 h-10 rounded-xl bg-sky-500/20 text-sky-400 flex items-center justify-center">
                <KeyRound className="w-5 h-5" />
              </div>
              <div>
                <h3 className="text-base font-bold text-white">
                  Akun Pengguna &amp; Hak Akses
                </h3>
                <p className="text-xs text-slate-400">
                  Gunakan kredensial berikut untuk login atau menguji hak akses:
                </p>
              </div>
            </div>

            {/* Credentials List */}
            <div className="space-y-2.5 my-4">
              {/* Admin */}
              <div className="bg-slate-800/80 rounded-xl p-3 border border-slate-700/60 flex items-center justify-between">
                <div>
                  <div className="flex items-center gap-2">
                    <span className="font-semibold text-sm text-sky-300">Admin Jaya</span>
                    <span className="text-[10px] bg-sky-500/20 text-sky-300 px-2 py-0.5 rounded-full font-bold">
                      Owner / Admin
                    </span>
                  </div>
                  <p className="text-xs text-slate-300 mt-1 font-mono">
                    User: <strong className="text-white">admin</strong> | Pass: <strong className="text-white">admin123</strong>
                  </p>
                </div>
                <button
                  onClick={() => copyToClipboard('admin\tadmin123', 'admin')}
                  className="px-2.5 py-1.5 bg-slate-700 hover:bg-slate-600 rounded-lg text-xs font-medium text-slate-200 flex items-center gap-1.5 transition"
                >
                  {copiedKey === 'admin' ? (
                    <>
                      <CheckCircle2 className="w-3.5 h-3.5 text-emerald-400" />
                      <span>Tersalin</span>
                    </>
                  ) : (
                    <>
                      <Copy className="w-3.5 h-3.5" />
                      <span>Salin</span>
                    </>
                  )}
                </button>
              </div>

              {/* Kasir */}
              <div className="bg-slate-800/80 rounded-xl p-3 border border-slate-700/60 flex items-center justify-between">
                <div>
                  <div className="flex items-center gap-2">
                    <span className="font-semibold text-sm text-emerald-300">Siti Kasir</span>
                    <span className="text-[10px] bg-emerald-500/20 text-emerald-300 px-2 py-0.5 rounded-full font-bold">
                      Kasir Toko
                    </span>
                  </div>
                  <p className="text-xs text-slate-300 mt-1 font-mono">
                    User: <strong className="text-white">kasir</strong> | Pass: <strong className="text-white">kasir123</strong>
                  </p>
                </div>
                <button
                  onClick={() => copyToClipboard('kasir\tkasir123', 'kasir')}
                  className="px-2.5 py-1.5 bg-slate-700 hover:bg-slate-600 rounded-lg text-xs font-medium text-slate-200 flex items-center gap-1.5 transition"
                >
                  {copiedKey === 'kasir' ? (
                    <>
                      <CheckCircle2 className="w-3.5 h-3.5 text-emerald-400" />
                      <span>Tersalin</span>
                    </>
                  ) : (
                    <>
                      <Copy className="w-3.5 h-3.5" />
                      <span>Salin</span>
                    </>
                  )}
                </button>
              </div>

              {/* Gudang */}
              <div className="bg-slate-800/80 rounded-xl p-3 border border-slate-700/60 flex items-center justify-between">
                <div>
                  <div className="flex items-center gap-2">
                    <span className="font-semibold text-sm text-amber-300">Budi Gudang</span>
                    <span className="text-[10px] bg-amber-500/20 text-amber-300 px-2 py-0.5 rounded-full font-bold">
                      Staf Gudang
                    </span>
                  </div>
                  <p className="text-xs text-slate-300 mt-1 font-mono">
                    User: <strong className="text-white">budi</strong> | Pass: <strong className="text-white">budi123</strong>
                  </p>
                </div>
                <button
                  onClick={() => copyToClipboard('budi\tbudi123', 'budi')}
                  className="px-2.5 py-1.5 bg-slate-700 hover:bg-slate-600 rounded-lg text-xs font-medium text-slate-200 flex items-center gap-1.5 transition"
                >
                  {copiedKey === 'budi' ? (
                    <>
                      <CheckCircle2 className="w-3.5 h-3.5 text-emerald-400" />
                      <span>Tersalin</span>
                    </>
                  ) : (
                    <>
                      <Copy className="w-3.5 h-3.5" />
                      <span>Salin</span>
                    </>
                  )}
                </button>
              </div>
            </div>

            <div className="mt-5">
              <button
                onClick={() => setShowInfoModal(false)}
                className="w-full bg-sky-600 hover:bg-sky-500 text-white py-2.5 rounded-xl font-bold text-xs transition"
              >
                Tutup &amp; Mulai Uji Coba
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

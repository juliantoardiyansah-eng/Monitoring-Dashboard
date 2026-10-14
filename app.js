/**
 * Project Monitoring Dashboard - Main Application Script
 * Dynamic Excel Parsing, Chart.js Visualizations & Filtering Engine
 */

// Initial Sample Data (extracted directly from user's image)
const INITIAL_SAMPLE_DATA = [
  { Project: 'POBSI', "Task Name": 'Pembuatan SRS System', Status: 'IN PROGRESS', Type: 'Document', PIC: 'Andri P Ku', From: '-', "Due Date": '23 Sep 2024' },
  { Project: 'POBSI', "Task Name": 'sirkulasi POBSI BRD', Status: 'IN PROGRESS', Type: 'Document', PIC: 'Andri P Ku', From: '-', "Due Date": '07 Nov 2024' },
  { Project: 'POBSI', "Task Name": 'CR untuk Fitur Tambahan', Status: 'IN PROGRESS', Type: 'Document', PIC: 'Andri P Ku', From: '-', "Due Date": '25 Nov 2024' },
  { Project: 'POBSI', "Task Name": 'Perbaikan bug & enhancement', Status: 'IN PROGRESS', Type: 'System', PIC: 'Andri P Ku', From: '-', "Due Date": '28 Nov 2024' },
  { Project: 'POBSI', "Task Name": 'Menunggu UAT Sign-off', Status: 'IN PROGRESS', Type: 'System', PIC: 'Andri P Ku', From: '28 Jul 2024', "Due Date": '04 Aug 2024' },
  { Project: 'POBSI', "Task Name": 'source code migration', Status: 'BACKLOG', Type: 'System', PIC: 'Andri P Ku', From: '-', "Due Date": '-' },
  { Project: 'CCSA', "Task Name": 'sirkulasi dokumen', Status: 'TODO', Type: 'People', PIC: 'Dedik Mui', From: '-', "Due Date": '03 Oct 2024' },
  { Project: 'CCSA', "Task Name": 'Menunggu persetujuan BA', Status: 'IN PROGRESS', Type: 'People', PIC: 'Dedik Mui', From: '-', "Due Date": '12 May 2024' },
  { Project: 'CCSA', "Task Name": 'go live preparation', Status: 'IN PROGRESS', Type: 'People', PIC: 'Dedik Mui', From: '-', "Due Date": '20 May 2024' },
  { Project: 'CCSA', "Task Name": 'UAT dengan tim business', Status: 'IN PROGRESS', Type: 'System', PIC: 'Dedik Mui', From: '15 Jun 2024', "Due Date": '04 Aug 2024' },
  { Project: 'Carpool U', "Task Name": 'Info ke user terkait fitur', Status: 'BACKLOG', Type: 'People', PIC: 'Dedik Mui', From: '-', "Due Date": '21 Oct 2024' },
  { Project: 'Carpool U', "Task Name": 'Analisa kebutuhan teknis', Status: 'IN PROGRESS', Type: 'System', PIC: 'Dedik Mui', From: '12 May 2024', "Due Date": '19 May 2024' },
  { Project: 'Carpool U', "Task Name": 'SIT testing tahap 1', Status: 'TODO', Type: 'System', PIC: 'Dedik Mui', From: '-', "Due Date": '02 Jun 2024' },
  { Project: 'Carpool U', "Task Name": 'Testing uat bersama user', Status: 'IN PROGRESS', Type: 'System', PIC: 'Dedik Mui', From: '-', "Due Date": '02 Jun 2024' },
  { Project: 'Carpool U', "Task Name": 'Info ke Pak Director', Status: 'IN PROGRESS', Type: 'People', PIC: 'Dedik Mui', From: '26 May 2024', "Due Date": '02 Jun 2024' },
  { Project: 'Carpool U', "Task Name": 'Konfirmasi ketersediaan', Status: 'TODO', Type: 'People', PIC: 'Dedik Mui', From: '28 Jul 2024', "Due Date": '31 Jul 2024' },
  { Project: 'Carpool U', "Task Name": 'requirement gathering', Status: 'BACKLOG', Type: 'System', PIC: 'Dedik Mui', From: '-', "Due Date": '-' },
  { Project: 'Jual Aset', "Task Name": 'Design app wireframe', Status: 'DONE', Type: 'Document', PIC: 'Andri P Ku', From: '-', "Due Date": '04 Nov 2024' },
  { Project: 'Jual Aset', "Task Name": 'sirkulasi dokumen BA', Status: 'DONE', Type: 'Document', PIC: 'Andri P Ku', From: '-', "Due Date": '07 Nov 2024' },
  { Project: 'Jual Aset', "Task Name": 'sirkulasi dokumen FSD', Status: 'DONE', Type: 'Document', PIC: 'Andri P Ku', From: '-', "Due Date": '07 Nov 2024' },
  { Project: 'Jual Aset', "Task Name": 'Sirkulasi GSD ke User', Status: 'DONE', Type: 'People', PIC: 'Andri P Ku', From: '-', "Due Date": '07 Nov 2024' },
  { Project: 'Metabase', "Task Name": 'Dashboard Sales', Status: 'BACKLOG', Type: 'System', PIC: 'Hendra Tr', From: '-', "Due Date": '28 Dec 2024' },
  { Project: 'Metabase', "Task Name": 'Dashboard Executive', Status: 'BACKLOG', Type: 'System', PIC: 'Hendra Tr', From: '-', "Due Date": '23 Dec 2024' },
  { Project: 'Metabase', "Task Name": 'Menampilkan widget chart', Status: 'BACKLOG', Type: 'System', PIC: 'Hendra Tr', From: '-', "Due Date": '24 Apr 2024' },
  { Project: 'Metabase', "Task Name": 'Visualisasi Data', Status: 'TODO', Type: 'System', PIC: 'Hendra Tr', From: '09 Jun 2024', "Due Date": '15 Jun 2024' },
  { Project: 'MNC Anin', "Task Name": 'Pembuatan System Admin', Status: 'DONE', Type: 'Document', PIC: 'Dedik Mui', From: '-', "Due Date": '27 Jun 2024' },
  { Project: 'MNC Anin', "Task Name": 'Pembuatan Database Schema', Status: 'DONE', Type: 'System', PIC: 'Dedik Mui', From: '19 May 2024', "Due Date": '02 Jun 2024' },
  { Project: 'MNC Anin', "Task Name": 'Update inisial data', Status: 'TODO', Type: 'System', PIC: 'Dedik Mui', From: '28 Jul 2024', "Due Date": '28 Jul 2024' },
  { Project: 'Policy Learn', "Task Name": 'pembuatan modul baru', Status: 'IN PROGRESS', Type: 'Document', PIC: 'Jihad Mau', From: '-', "Due Date": '14 Apr 2024' },
  { Project: 'Policy Learn', "Task Name": 'CR Deploy Backend', Status: 'DONE', Type: 'System', PIC: 'Jihad Mau', From: '09 Jun 2024', "Due Date": '10 Jun 2024' },
  { Project: 'Policy Learn', "Task Name": 'Konfirmasi ke Business', Status: 'TODO', Type: 'People', PIC: 'Jihad Mau', From: '28 Jul 2024', "Due Date": '31 Jul 2024' },
  { Project: 'Policy Learn', "Task Name": 'follow up hasil testing', Status: 'BACKLOG', Type: 'System', PIC: '-', From: '-', "Due Date": '-' },
  { Project: 'Policy Learn', "Task Name": 'meeting koordinasi', Status: 'BACKLOG', Type: 'System', PIC: '-', From: '-', "Due Date": '-' },
  { Project: 'Policy Learn', "Task Name": 'Persiapan go live', Status: 'BACKLOG', Type: 'System', PIC: 'Jihad Mau', From: '-', "Due Date": '-' },
  { Project: 'QR Advert', "Task Name": 'Sirkulasi BRD final', Status: 'DONE', Type: 'Document', PIC: 'Dedik Mui', From: '-', "Due Date": '14 Apr 2024' },
  { Project: 'QR Advert', "Task Name": 'Create BA testing', Status: 'DONE', Type: 'Document', PIC: 'Dedik Mui', From: '-', "Due Date": '06 May 2024' },
  { Project: 'QR Advert', "Task Name": 'Sirkulasi dokumen SRS', Status: 'IN PROGRESS', Type: 'People', PIC: 'Dedik Mui', From: '-', "Due Date": '29 May 2024' },
  { Project: 'QR Advert', "Task Name": 'Persentasi ke VP', Status: 'DONE', Type: 'People', PIC: 'Dedik Mui', From: '-', "Due Date": '02 Jun 2024' },
  { Project: 'QR Advert', "Task Name": 'Sirkulasi FSD', Status: 'DONE', Type: 'Document', PIC: 'Dedik Mui', From: '-', "Due Date": '14 Jul 2024' },
  { Project: 'QR Advert', "Task Name": 'Re-Konfirmasi fitur', Status: 'TODO', Type: 'Document', PIC: 'Dedik Mui', From: '28 Jul 2024', "Due Date": '04 Aug 2024' },
  { Project: 'Document', "Task Name": 'maping pengajuan akses', Status: 'DONE', Type: 'System', PIC: 'Dedik Mui', From: '-', "Due Date": '30 Apr 2024' },
  { Project: 'Document', "Task Name": 'Konfirmasi ke PIC', Status: 'IN PROGRESS', Type: 'People', PIC: 'Dedik Mui', From: '26 May 2024', "Due Date": '31 Jul 2024' },
  { Project: 'MNC Char', "Task Name": 'CR deploy to Staging', Status: 'DONE', Type: 'Document', PIC: 'Jihad Mau', From: '-', "Due Date": '05 May 2024' },
  { Project: 'MNC Char', "Task Name": 'Follow up isu prod', Status: 'IN PROGRESS', Type: 'System', PIC: 'Jihad Mau', From: '26 May 2024', "Due Date": '31 Jul 2024' },
  { Project: 'MNC Char', "Task Name": 'Sirkulasi dokumen FSD', Status: 'IN PROGRESS', Type: 'People', PIC: 'Jihad Mau', From: '-', "Due Date": '04 Aug 2024' },
  { Project: 'Dashboard', "Task Name": 'Tarik data log user', Status: 'DONE', Type: 'System', PIC: 'Ridwan N', From: '-', "Due Date": '19 May 2024' },
  { Project: 'Dashboard', "Task Name": 'Pembuatan Widget KPI', Status: 'TODO', Type: 'Document', PIC: 'Dedik Mui', From: '-', "Due Date": '02 Jun 2024' },
  { Project: 'Dashboard', "Task Name": 'Penambahan Server Node', Status: 'DONE', Type: 'System', PIC: 'Andri P Ku', From: '09 Jun 2024', "Due Date": '16 Jun 2024' },
  { Project: 'Dashboard', "Task Name": 'Penambahan Index DB', Status: 'BACKLOG', Type: 'System', PIC: 'Ridwan N', From: '09 Jun 2024', "Due Date": '16 Jun 2024' },
  { Project: 'Dashboard', "Task Name": 'Perubahan Query API', Status: 'DONE', Type: 'System', PIC: 'Ridwan N', From: '09 Jun 2024', "Due Date": '16 Jun 2024' },
  { Project: 'Dashboard', "Task Name": 'Migrasi ke Cloud', Status: 'TODO', Type: 'System', PIC: 'Ridwan N', From: '26 May 2024', "Due Date": '30 Jun 2024' },
  { Project: 'Dashboard', "Task Name": 'Sirkulasi BRD Update', Status: 'IN PROGRESS', Type: 'System', PIC: 'Jihad Mau', From: '26 May 2024', "Due Date": '31 Jul 2024' },
  { Project: 'Inventory', "Task Name": 'Sirkulasi Dokumen CR', Status: 'DONE', Type: 'Document', PIC: 'Dedik Mui', From: '-', "Due Date": '19 May 2024' },
  { Project: 'Inventory', "Task Name": 'Dashboard Stock', Status: 'DONE', Type: 'Document', PIC: 'Hendra Tr', From: '09 Jun 2024', "Due Date": '12 Jun 2024' },
  { Project: 'Inventory', "Task Name": 'Revisi Dokumen SRS', Status: 'TODO', Type: 'Document', PIC: 'Dedik Mui', From: '28 Jul 2024', "Due Date": '28 Jul 2024' },
  { Project: 'Inventory', "Task Name": 'Perbaikan Bug Form', Status: 'TODO', Type: 'System', PIC: 'Cecep Yus', From: '28 Jul 2024', "Due Date": '31 Jul 2024' },
  { Project: 'MNC Cont', "Task Name": 'Presentasi Proposal', Status: 'TODO', Type: 'People', PIC: 'Jihad Mau', From: '05 May 2024', "Due Date": '20 May 2024' },
  { Project: 'MNC Cont', "Task Name": 'Presentasi Ke Client', Status: 'TODO', Type: 'People', PIC: 'Jihad Mau', From: '-', "Due Date": '26 May 2024' },
  { Project: 'MNC Cont', "Task Name": 'Eksekusi Deployment', Status: 'TODO', Type: 'System', PIC: 'Jihad Mau', From: '28 Jul 2024', "Due Date": '31 Jul 2024' },
  { Project: 'e-Form', "Task Name": 'Presentasi Wireframe', Status: 'TODO', Type: 'People', PIC: 'Jihad Mau', From: '-', "Due Date": '25 May 2024' },
  { Project: 'e-Form', "Task Name": 'Sirkulasi dokumen SRS', Status: 'IN PROGRESS', Type: 'System', PIC: 'Jihad Mau', From: '-', "Due Date": '04 Aug 2024' },
  { Project: 'e-Form', "Task Name": 'Sirkulasi dokumen FSD', Status: 'IN PROGRESS', Type: 'Document', PIC: 'Jihad Mau', From: '-', "Due Date": '04 Aug 2024' },
  { Project: 'e-Form', "Task Name": 'Sirkulasi dokumen BRD', Status: 'IN PROGRESS', Type: 'Document', PIC: 'Jihad Mau', From: '-', "Due Date": '04 Aug 2024' },
  { Project: 'e-Form', "Task Name": 'Testing admin portal', Status: 'TODO', Type: 'System', PIC: 'Dedik Mui', From: '28 Jul 2024', "Due Date": '04 Aug 2024' },
  { Project: 'e-Form', "Task Name": 'UAT perubahan ke BS', Status: 'SYSTEM', Type: 'System', PIC: '-', From: '-', "Due Date": '-' },
  { Project: 'e-Form', "Task Name": 'Persentasi Final', Status: 'BACKLOG', Type: 'People', PIC: 'Jihad Mau', From: '-', "Due Date": '-' },
  { Project: 'Messaging', "Task Name": 'Presentasi Ke Tim IT', Status: 'TODO', Type: 'People', PIC: 'Jihad Mau', From: '12 May 2024', "Due Date": '26 May 2024' },
  { Project: 'Messaging', "Task Name": 'Presentasi Ke Management', Status: 'TODO', Type: 'People', PIC: 'Jihad Mau', From: '-', "Due Date": '28 May 2024' },
  { Project: 'Website M', "Task Name": 'Permintaan fitur baru', Status: 'IN PROGRESS', Type: 'System', PIC: 'Andri P Ku', From: '-', "Due Date": '29 May 2024' },
  { Project: 'Website M', "Task Name": 'Sirkulasi GSD ke Tim', Status: 'IN PROGRESS', Type: 'Document', PIC: 'Jihad Mau', From: '19 May 2024', "Due Date": '02 Jun 2024' },
  { Project: 'Website M', "Task Name": 'Design banner baru', Status: 'TODO', Type: 'System', PIC: 'Ardha Raf', From: '28 Jul 2024', "Due Date": '28 Jul 2024' },
  { Project: 'Website M', "Task Name": 'Merubah layout home', Status: 'TODO', Type: 'System', PIC: 'Ardha Raf', From: '28 Jul 2024', "Due Date": '28 Jul 2024' },
  { Project: 'Website M', "Task Name": 'Penambahan meta SEO', Status: 'TODO', Type: 'System', PIC: '-', From: '28 Jul 2024', "Due Date": '28 Jul 2024' },
  { Project: 'Website K', "Task Name": 'Sirkulasi BRD ke Pak Dir', Status: 'DOCUMENT', Type: 'Document', PIC: 'Dedik Mui', From: '12 May 2024', "Due Date": '02 Jun 2024' },
  { Project: 'Website K', "Task Name": 'Masih terhambat isu prod', Status: 'TODO', Type: 'People', PIC: 'Julianto A', From: '21 Apr 2024', "Due Date": '30 Jun 2024' },
  { Project: 'MNC Pedi', "Task Name": 'Admin FMIN Testing', Status: 'IN PROGRESS', Type: 'System', PIC: 'Dedik Mui', From: '26 May 2024', "Due Date": '02 Jun 2024' },
  { Project: 'MNC Pedi', "Task Name": 'Masih ada bug login', Status: 'DONE', Type: 'System', PIC: 'Dedik Mui', From: '26 May 2024', "Due Date": '21 Jul 2024' },
  { Project: 'MNC Pedi', "Task Name": 'Pembuatan API Gateway', Status: 'DONE', Type: 'System', PIC: 'Dedik Mui', From: '28 Jul 2024', "Due Date": '28 Jul 2024' },
  { Project: 'MNC Pedi', "Task Name": 'Kordinasi tim QA', Status: 'DONE', Type: 'System', PIC: 'Dedik Mui', From: '28 Jul 2024', "Due Date": '31 Jul 2024' },
  { Project: 'Website S', "Task Name": 'Sirkulasi GSD ke User', Status: 'IN PROGRESS', Type: 'System', PIC: 'Jihad Mau', From: '19 May 2024', "Due Date": '05 Jun 2024' },
  { Project: 'Website S', "Task Name": 'Pembuatan Sitemap', Status: 'TODO', Type: 'Document', PIC: 'Jihad Mau', From: '09 Jun 2024', "Due Date": '12 Jun 2024' },
  { Project: 'Website S', "Task Name": 'Pembuatan Template Page', Status: 'DONE', Type: 'Document', PIC: 'Jihad Mau', From: '28 Apr 2024', "Due Date": '24 Jul 2024' },
  { Project: 'Website S', "Task Name": 'Penambahan Script Analytic', Status: 'TODO', Type: 'System', PIC: 'Riyadh M', From: '28 Jul 2024', "Due Date": '10 Aug 2024' },
  { Project: 'Acquisition System', "Task Name": 'Follow terkait meeting offline oleh Mbak Diaz', Status: 'TO DO', Type: 'People', PIC: 'Jihad', From: '-', "Due Date": '28 Jul 26' },
  { Project: 'Acquisition System', "Task Name": 'Update module Approval', Status: 'TO DO', Type: 'System', PIC: 'Wahyu', From: '-', "Due Date": '31 Jul 26' },
  { Project: 'Acquisition System', "Task Name": 'Testing module Approval di server staging', Status: 'TO DO', Type: 'System', PIC: 'Jihad', From: '-', "Due Date": '31 Jul 26' },
  { Project: 'Acquisition System', "Task Name": 'Follow up terkait approval hirarki untuk module IA', Status: 'TO DO', Type: 'People', PIC: 'Jihad', From: '-', "Due Date": '28 Jul 26' },
  { Project: 'GTV', "Task Name": 'Meeting - Sync Requirement', Status: 'DONE', Type: 'People', PIC: 'Jihad Mau', From: '09 Jun 2024', "Due Date": '09 Jun 2024' },
  { Project: 'GTV', "Task Name": 'Pembuatan Wireframe', Status: 'IN PROGRESS', Type: 'People', PIC: 'Riyadh M', From: '23 Jul 2024', "Due Date": '31 Aug 2024' },
  { Project: 'GTV', "Task Name": 'Review design system', Status: 'DONE', Type: 'People', PIC: 'Jihad Mau', From: '-', "Due Date": '-' },
  { Project: 'Talent Cas', "Task Name": 'Pembuatan Portal', Status: 'DONE', Type: 'System', PIC: 'Dedik Mui', From: '09 Jun 2024', "Due Date": '19 Jun 2024' },
  { Project: 'Talent Cas', "Task Name": 'Sirkulasi BRD ke User', Status: 'IN PROGRESS', Type: 'Document', PIC: 'Dedik Mui', From: '-', "Due Date": '31 Jul 2024' },
  { Project: 'User Auth', "Task Name": 'New sirkulasi dokumen', Status: 'IN PROGRESS', Type: 'Document', PIC: 'Jihad Mau', From: '-', "Due Date": '26 Jun 2024' },
  { Project: 'OMS - MN', "Task Name": 'Change Request DB', Status: 'DONE', Type: 'System', PIC: 'Septriyon', From: '26 May 2024', "Due Date": '30 Jun 2024' },
  { Project: 'OMS - MN', "Task Name": 'CR penambahan field', Status: 'TODO', Type: 'System', PIC: '-', From: '28 Jul 2024', "Due Date": '29 Jul 2024' },
  { Project: 'Talent Sch', "Task Name": 'Info untuk tim HR', Status: 'DONE', Type: 'People', PIC: 'Dedik Mui', From: '09 Jun 2024', "Due Date": '14 Jul 2024' },
  { Project: 'Talent Sch', "Task Name": 'TBA-Diskusi alur kerja', Status: 'TODO', Type: 'System', PIC: 'Dedik Mui', From: '28 Jul 2024', "Due Date": '-' },
  { Project: 'MNC ASIA', "Task Name": 'Documentation , Project plan', Status: 'DONE', Type: 'System', PIC: 'Dedik Mui', From: '19 May 2024', "Due Date": '28 Jul 2024' },
  { Project: 'MNC ASIA', "Task Name": 'CR untuk penyesuaian', Status: 'SYSTEM', Type: 'System', PIC: 'Riyadh M', From: '08 Jun 2024', "Due Date": '28 Jul 2024' }
];

// App State
let currentDataset = [...INITIAL_SAMPLE_DATA];
let filteredData = [...INITIAL_SAMPLE_DATA];
let currentSortColumn = 'Due Date';
let isSortAscending = true;
let isOverdueOnlyFilter = false;
let currentPage = 1;
const rowsPerPage = 12;
let lastUploadedFileName = null;
let lastUploadDateStr = null;
let selectedFilters = {
  Project: [],
  Status: [],
  Type: [],
  PIC: []
};
let filterInstances = {};

// Chart Instances
let chartStatusInstance = null;
let chartPICInstance = null;
let chartTypeInstance = null;
let chartProjectInstance = null;

// Color Palette Constants
const STATUS_COLORS = {
  'DONE': { bg: 'rgba(16, 185, 129, 0.85)', border: '#10b981' },
  'IN PROGRESS': { bg: 'rgba(245, 158, 11, 0.85)', border: '#f59e0b' },
  'TODO': { bg: 'rgba(6, 182, 212, 0.85)', border: '#06b6d4' },
  'BACKLOG': { bg: 'rgba(168, 85, 247, 0.85)', border: '#a855f7' },
  'OTHER': { bg: 'rgba(100, 116, 139, 0.85)', border: '#64748b' }
};

// Initialize Application
document.addEventListener('DOMContentLoaded', () => {
  setupEventListeners();
  populateDropdownFilters();
  applyFiltersAndRender();
});

// Setup Event Listeners (Drag & Drop, Inputs, Buttons)
function setupEventListeners() {
  const dropzone = document.getElementById('dropzone');
  const fileInput = document.getElementById('fileInput');

  // Drag and Drop Effects
  ['dragenter', 'dragover'].forEach(eventName => {
    dropzone.addEventListener(eventName, (e) => {
      e.preventDefault();
      dropzone.classList.add('dragover');
    }, false);
  });

  ['dragleave', 'drop'].forEach(eventName => {
    dropzone.addEventListener(eventName, (e) => {
      e.preventDefault();
      dropzone.classList.remove('dragover');
    }, false);
  });

  dropzone.addEventListener('drop', (e) => {
    const dt = e.dataTransfer;
    const files = dt.files;
    if (files.length) handleExcelUpload(files[0]);
  });

  fileInput.addEventListener('change', (e) => {
    if (e.target.files.length) handleExcelUpload(e.target.files[0]);
  });

  // Filter Event Listeners
  document.getElementById('searchInput').addEventListener('input', applyFiltersAndRender);
  initTagFilters();

  document.getElementById('clearFiltersBtn').addEventListener('click', clearAllFilters);
  document.getElementById('resetSampleBtn').addEventListener('click', resetToSampleData);

  // Overdue KPI Card Click Handler
  const overdueCard = document.getElementById('kpiCardOverdue');
  if (overdueCard) {
    overdueCard.addEventListener('click', () => {
      isOverdueOnlyFilter = !isOverdueOnlyFilter;
      if (isOverdueOnlyFilter) {
        overdueCard.classList.add('active');
      } else {
        overdueCard.classList.remove('active');
      }
      applyFiltersAndRender();
      document.getElementById('taskTable').scrollIntoView({ behavior: 'smooth' });
    });
  }

  // Export Buttons
  document.getElementById('exportPdfBtn').addEventListener('click', exportPdfWeeklyUpdateTemplate);
  document.getElementById('exportExcelBtn').addEventListener('click', exportToExcel);
  document.getElementById('exportCSVBtn').addEventListener('click', exportToCSV);

  // Pagination Controls
  document.getElementById('prevPageBtn').addEventListener('click', () => {
    if (currentPage > 1) { currentPage--; renderTable(); updatePagination(); }
  });

  document.getElementById('nextPageBtn').addEventListener('click', () => {
    const totalPages = Math.ceil(filteredData.length / rowsPerPage);
    if (currentPage < totalPages) { currentPage++; renderTable(); updatePagination(); }
  });

  // Table Column Sort Click
  document.querySelectorAll('#taskTable th[data-sort]').forEach(th => {
    th.addEventListener('click', () => {
      const col = th.getAttribute('data-sort');
      if (currentSortColumn === col) {
        isSortAscending = !isSortAscending;
      } else {
        currentSortColumn = col;
        isSortAscending = true;
      }
      sortFilteredData();
      renderTable();
    });
  });
}

// Handle Excel / CSV File Upload with SheetJS (Bullet-proof Multi-mode Parser)
function handleExcelUpload(file) {
  if (!file) return;

  const reader = new FileReader();

  reader.onload = (e) => {
    try {
      let workbook;
      const buffer = e.target.result;

      // Primary Try: ArrayBuffer mode
      try {
        const data = new Uint8Array(buffer);
        workbook = XLSX.read(data, { type: 'array', cellDates: true, dateNF: 'yyyy-mm-dd' });
      } catch (e1) {
        console.warn('ArrayBuffer read failed, trying string mode...', e1);
        // Fallback: Binary string or raw text mode for CSV
        workbook = XLSX.read(buffer, { type: 'binary', cellDates: true });
      }

      if (!workbook || !workbook.SheetNames || !workbook.SheetNames.length) {
        throw new Error('Sheet Excel tidak terdeteksi atau kosong.');
      }

      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];

      // Parse JSON from worksheet
      let rawJson = XLSX.utils.sheet_to_json(worksheet, { defval: '-', raw: false });

      // Fallback: If sheet_to_json object list is empty or invalid, try header:1 (array of arrays)
      if (!rawJson || !rawJson.length) {
        const rawAoA = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '-' });
        rawJson = parseArrayOfArrays(rawAoA);
      }

      if (!rawJson || !rawJson.length) {
        alert('File Excel kosong atau tidak terdeteksi data baris yang valid.');
        return;
      }

      // Filter empty rows & normalize row objects
      const normalizedData = rawJson
        .map(row => normalizeRowData(row))
        .filter(item => item.Project !== '-' || item["Task Name"] !== '-');

      if (!normalizedData.length) {
        alert('Data Excel berhasil dibaca tetapi tidak ada baris yang valid.');
        return;
      }

      currentDataset = normalizedData;

      const now = new Date();
      const day = String(now.getDate()).padStart(2, '0');
      const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
      lastUploadDateStr = `${day} ${monthNames[now.getMonth()]} ${now.getFullYear()}`;
      lastUploadedFileName = file.name;

      // Update UI Header Badge
      const badge = document.getElementById('lastUpdatedBadge');
      if (badge) {
        badge.innerHTML = `
          <i data-lucide="check-circle"></i>
          <span>Data File: ${escapeHtml(file.name)} (${normalizedData.length} baris)</span>
        `;
      }
      lucide.createIcons();

      populateDropdownFilters();
      clearAllFilters();
      alert(`Berhasil mengunggah file ${file.name}! Dashboard ter-update dengan ${normalizedData.length} data task.`);

    } catch (err) {
      console.error('Excel Parsing Error:', err);
      alert(`Gagal membaca file Excel/CSV:\n${err.message || 'Format file tidak didukung.'}\n\nTips: Simpan ulang file sebagai .xlsx atau .csv standar.`);
    }
  };

  reader.onerror = (err) => {
    console.error('FileReader Error:', err);
    alert('Gagal membaca file dari sistem lokal Anda.');
  };

  // Choose best FileReader strategy
  if (file.name.toLowerCase().endsWith('.csv') || file.type.includes('csv')) {
    reader.readAsText(file);
  } else {
    reader.readAsArrayBuffer(file);
  }
}

// Fallback helper for Array of Arrays format
function parseArrayOfArrays(aoa) {
  if (!aoa || aoa.length < 2) return [];
  const headers = aoa[0].map(h => String(h || '').trim());
  const rows = [];

  for (let i = 1; i < aoa.length; i++) {
    const rowObj = {};
    const rowData = aoa[i];
    if (!rowData || !rowData.length) continue;

    headers.forEach((h, idx) => {
      if (h) rowObj[h] = rowData[idx] !== undefined ? rowData[idx] : '-';
    });
    rows.push(rowObj);
  }
  return rows;
}

// Smart Header Normalization (Flexible for Indonesian & English headers)
function normalizeRowData(row) {
  const normalized = {
    Project: '-',
    "Task Name": '-',
    Status: 'TODO',
    Type: 'System',
    PIC: '-',
    From: '-',
    "Due Date": '-',
    "Actual Date": '-'
  };

  if (!row || typeof row !== 'object') return normalized;

  Object.keys(row).forEach(key => {
    const cleanKey = String(key).trim().toLowerCase();
    const rawVal = row[key];
    const val = rawVal !== null && rawVal !== undefined ? String(rawVal).trim() : '-';

    if (cleanKey.includes('proj') || cleanKey.includes('proyek') || cleanKey.includes('stream')) {
      normalized.Project = val;
    } 
    else if (cleanKey.includes('task') || cleanKey.includes('pekerjaan') || cleanKey.includes('nama') || cleanKey.includes('aktivitas') || cleanKey.includes('item')) {
      normalized["Task Name"] = val;
    } 
    else if (cleanKey.includes('stat') || cleanKey.includes('kondisi') || cleanKey.includes('state')) {
      normalized.Status = val.toUpperCase();
    } 
    else if (cleanKey.includes('type') || cleanKey.includes('jenis') || cleanKey.includes('kategori') || cleanKey.includes('tipe')) {
      normalized.Type = val;
    } 
    else if (cleanKey.includes('pic') || cleanKey.includes('resour') || cleanKey.includes('orang') || cleanKey.includes('assign') || cleanKey.includes('owner')) {
      normalized.PIC = val;
    } 
    else if (cleanKey.includes('from') || cleanKey.includes('mulai') || cleanKey.includes('start') || cleanKey.includes('tgl mulai')) {
      normalized.From = formatDateValue(rawVal);
    } 
    else if (cleanKey.includes('actual') || cleanKey.includes('realisasi') || cleanKey.includes('tgl selesai') || (cleanKey.includes('selesai') && !cleanKey.includes('target') && !cleanKey.includes('due') && !cleanKey.includes('tenggat') && !cleanKey.includes('deadline'))) {
      normalized["Actual Date"] = formatDateValue(rawVal);
    }
    else if (cleanKey.includes('due') || cleanKey.includes('tenggat') || cleanKey.includes('target') || cleanKey.includes('deadline') || cleanKey.includes('selesai')) {
      normalized["Due Date"] = formatDateValue(rawVal);
    }
  });

  // Standardize Status values
  const st = String(normalized.Status).toUpperCase();
  if (st.includes('DONE') || st.includes('SELESAI') || st.includes('FINISH')) {
    normalized.Status = 'DONE';
  } else if (st.includes('PROG') || st.includes('JALAN') || st.includes('PROSES') || st.includes('PROGRESS')) {
    normalized.Status = 'IN PROGRESS';
  } else if (st.includes('LOG') || st.includes('PENDING') || st.includes('HOLD')) {
    normalized.Status = 'BACKLOG';
  } else {
    normalized.Status = 'TODO';
  }

  return normalized;
}

// Helper to format dates cleanly (Handles JS Date, Excel Serial numbers 45xxx, and strings)
function formatDateValue(val) {
  if (val === null || val === undefined || val === '' || val === '-') return '-';

  if (val instanceof Date) {
    if (isNaN(val.getTime())) return '-';
    return val.toLocaleDateString('id-ID', { day: '2-digit', month: 'short', year: 'numeric' });
  }

  // Handle Excel Serial Number (e.g. 45532)
  if (typeof val === 'number' && val > 20000 && val < 60000) {
    const jsDate = new Date(Math.round((val - 25569) * 86400 * 1000));
    return jsDate.toLocaleDateString('id-ID', { day: '2-digit', month: 'short', year: 'numeric' });
  }

  const str = String(val).trim();
  if (!str || str === 'undefined' || str === 'null') return '-';

  // Handle string Excel Serial Number
  const num = Number(str);
  if (!isNaN(num) && num > 20000 && num < 60000) {
    const jsDate = new Date(Math.round((num - 25569) * 86400 * 1000));
    return jsDate.toLocaleDateString('id-ID', { day: '2-digit', month: 'short', year: 'numeric' });
  }

  return str;
}

// Populate Dynamic Filter Dropdowns
function populateDropdownFilters() {
  const projects = [...new Set(currentDataset.map(d => d.Project))].filter(p => p && p !== '-').sort();
  const defaultStatuses = ['DONE', 'IN PROGRESS', 'TODO', 'BACKLOG'];
  const datasetStatuses = currentDataset ? [...new Set(currentDataset.map(d => d.Status))].filter(s => s && s !== '-') : [];
  const statuses = [...new Set([...defaultStatuses, ...datasetStatuses])].sort();
  const types = [...new Set(currentDataset.map(d => d.Type))].filter(t => t && t !== '-').sort();
  const pics = [...new Set(currentDataset.map(d => d.PIC))].filter(p => p && p !== '-').sort();

  if (filterInstances.Project) filterInstances.Project.setOptions(projects);
  if (filterInstances.Status) filterInstances.Status.setOptions(statuses);
  if (filterInstances.Type) filterInstances.Type.setOptions(types);
  if (filterInstances.PIC) filterInstances.PIC.setOptions(pics);
}

// Reusable Tag Pill Multi-Select Filter Control (Select2 / Tagify style)
class TagMultiSelect {
  constructor({ key, containerId, inputBoxId, wrapperId, placeholderId, clearBtnId, dropdownId, placeholderText, pillTheme }) {
    this.key = key;
    this.container = document.getElementById(containerId);
    this.inputBox = document.getElementById(inputBoxId);
    this.wrapper = document.getElementById(wrapperId);
    this.placeholder = document.getElementById(placeholderId);
    this.clearBtn = document.getElementById(clearBtnId);
    this.dropdown = document.getElementById(dropdownId);
    this.placeholderText = placeholderText;
    this.pillTheme = pillTheme;
    this.options = [];
    this.searchTerm = '';

    this.init();
  }

  init() {
    if (!this.container || !this.inputBox) return;

    this.inputBox.addEventListener('click', (e) => {
      e.stopPropagation();
      const isOpen = this.container.classList.contains('open');
      document.querySelectorAll('.tags-select-container').forEach(el => {
        el.classList.remove('open');
      });
      if (!isOpen) {
        this.container.classList.add('open');
        this.searchTerm = '';
        this.renderDropdown();
        setTimeout(() => {
          const searchInput = this.dropdown ? this.dropdown.querySelector('.tags-dropdown-search-input') : null;
          if (searchInput) searchInput.focus();
        }, 50);
      }
    });

    if (this.clearBtn) {
      this.clearBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        this.clearAll();
      });
    }

    document.addEventListener('click', (e) => {
      if (!this.container.contains(e.target)) {
        this.container.classList.remove('open');
      }
    });

    this.render();
  }

  setOptions(optionsList) {
    this.options = optionsList;
    this.renderDropdown();
  }

  getSelected() {
    return selectedFilters[this.key] || [];
  }

  setSelected(arr) {
    selectedFilters[this.key] = arr;
    this.render();
    applyFiltersAndRender();
  }

  toggleOption(val) {
    let current = [...this.getSelected()];
    if (current.includes(val)) {
      current = current.filter(v => v !== val);
    } else {
      current.push(val);
    }
    this.setSelected(current);
  }

  removeOption(val) {
    let current = this.getSelected().filter(v => v !== val);
    this.setSelected(current);
  }

  clearAll() {
    this.setSelected([]);
  }

  render() {
    if (!this.wrapper) return;
    const selected = this.getSelected();

    this.wrapper.innerHTML = '';

    if (selected.length === 0) {
      if (this.placeholder) {
        this.placeholder.textContent = this.placeholderText;
        this.placeholder.style.display = 'inline';
        this.wrapper.appendChild(this.placeholder);
      }
      if (this.clearBtn) this.clearBtn.style.display = 'none';
    } else {
      if (this.placeholder) this.placeholder.style.display = 'none';
      if (this.clearBtn) this.clearBtn.style.display = 'inline-flex';

      selected.forEach(val => {
        const themeClass = this.pillTheme ? this.pillTheme(val) : 'tag-default';
        const pill = document.createElement('span');
        pill.className = `status-pill-tag ${themeClass}`;
        pill.innerHTML = `
          <span class="status-pill-remove">&times;</span>
          <span class="status-pill-text">${escapeHtml(val)}</span>
        `;

        pill.querySelector('.status-pill-remove').addEventListener('click', (e) => {
          e.stopPropagation();
          this.removeOption(val);
        });

        this.wrapper.appendChild(pill);
      });
    }

    this.renderDropdown();
  }

  renderDropdown() {
    if (!this.dropdown) return;
    const selected = this.getSelected();

    this.dropdown.innerHTML = '';

    if (!this.options.length) {
      this.dropdown.innerHTML = `<div style="padding: 12px; color: var(--text-muted); font-size: 0.82rem; text-align: center;">Tidak ada opsi tersedia</div>`;
      return;
    }

    // Top Search Bar inside Dropdown
    const searchWrapper = document.createElement('div');
    searchWrapper.className = 'tags-dropdown-search-wrapper';
    searchWrapper.innerHTML = `
      <div class="dropdown-search-box">
        <i data-lucide="search" class="dropdown-search-icon"></i>
        <input type="text" class="tags-dropdown-search-input" placeholder="Cari ${escapeHtml(this.key)}..." value="${escapeHtml(this.searchTerm)}">
      </div>
    `;

    const searchInput = searchWrapper.querySelector('input');
    searchInput.addEventListener('click', (e) => e.stopPropagation());
    searchInput.addEventListener('input', (e) => {
      this.searchTerm = e.target.value.toLowerCase();
      this.renderOptionsList(optionsContainer);
    });

    this.dropdown.appendChild(searchWrapper);

    // Options Container
    const optionsContainer = document.createElement('div');
    optionsContainer.className = 'tags-dropdown-options-list';
    this.dropdown.appendChild(optionsContainer);

    this.renderOptionsList(optionsContainer);
    lucide.createIcons();
  }

  renderOptionsList(container) {
    if (!container) return;
    container.innerHTML = '';
    const selected = this.getSelected();

    const filteredOptions = this.options.filter(val => 
      !this.searchTerm || String(val).toLowerCase().includes(this.searchTerm)
    );

    if (!filteredOptions.length) {
      container.innerHTML = `<div style="padding: 12px; color: var(--text-muted); font-size: 0.8rem; text-align: center;">Opsi "${escapeHtml(this.searchTerm)}" tidak ditemukan</div>`;
      return;
    }

    filteredOptions.forEach(val => {
      const isSelected = selected.includes(val);
      const item = document.createElement('div');
      item.className = `tags-option-item ${isSelected ? 'selected' : ''}`;

      let badgeHtml = escapeHtml(val);
      if (this.key === 'Status') {
        let badgeClass = 'badge-todo';
        if (val === 'DONE') badgeClass = 'badge-done';
        else if (val === 'IN PROGRESS') badgeClass = 'badge-inprog';
        else if (val === 'BACKLOG') badgeClass = 'badge-backlog';
        badgeHtml = `<span class="status-badge ${badgeClass}">${escapeHtml(val)}</span>`;
      }

      item.innerHTML = `
        <span class="option-label-text">${badgeHtml}</span>
        ${isSelected ? '<i data-lucide="check" class="check-icon"></i>' : ''}
      `;

      item.addEventListener('click', (e) => {
        e.stopPropagation();
        this.toggleOption(val);
      });

      container.appendChild(item);
    });

    lucide.createIcons();
  }
}

function initTagFilters() {
  filterInstances.Project = new TagMultiSelect({
    key: 'Project',
    containerId: 'projectTagsContainer',
    inputBoxId: 'projectTagsInputBox',
    wrapperId: 'projectTagsWrapper',
    placeholderId: 'projectTagsPlaceholder',
    clearBtnId: 'projectClearAllBtn',
    dropdownId: 'projectTagsDropdown',
    placeholderText: 'Semua Project...',
    pillTheme: () => 'tag-project'
  });

  filterInstances.Status = new TagMultiSelect({
    key: 'Status',
    containerId: 'statusTagsContainer',
    inputBoxId: 'statusTagsInputBox',
    wrapperId: 'statusTagsWrapper',
    placeholderId: 'statusTagsPlaceholder',
    clearBtnId: 'statusClearAllBtn',
    dropdownId: 'statusTagsDropdown',
    placeholderText: 'Semua Status...',
    pillTheme: (val) => {
      if (val === 'DONE') return 'tag-done';
      if (val === 'IN PROGRESS') return 'tag-inprog';
      if (val === 'TODO') return 'tag-todo';
      if (val === 'BACKLOG') return 'tag-backlog';
      return 'tag-default';
    }
  });

  filterInstances.Type = new TagMultiSelect({
    key: 'Type',
    containerId: 'typeTagsContainer',
    inputBoxId: 'typeTagsInputBox',
    wrapperId: 'typeTagsWrapper',
    placeholderId: 'typeTagsPlaceholder',
    clearBtnId: 'typeClearAllBtn',
    dropdownId: 'typeTagsDropdown',
    placeholderText: 'Semua Type...',
    pillTheme: () => 'tag-type'
  });

  filterInstances.PIC = new TagMultiSelect({
    key: 'PIC',
    containerId: 'picTagsContainer',
    inputBoxId: 'picTagsInputBox',
    wrapperId: 'picTagsWrapper',
    placeholderId: 'picTagsPlaceholder',
    clearBtnId: 'picClearAllBtn',
    dropdownId: 'picTagsDropdown',
    placeholderText: 'Semua PIC...',
    pillTheme: () => 'tag-pic'
  });
}

// Clear All Filters
function clearAllFilters() {
  document.getElementById('searchInput').value = '';
  Object.keys(selectedFilters).forEach(key => {
    selectedFilters[key] = [];
    if (filterInstances[key]) filterInstances[key].render();
  });

  isOverdueOnlyFilter = false;
  const overdueCard = document.getElementById('kpiCardOverdue');
  if (overdueCard) overdueCard.classList.remove('active');
  applyFiltersAndRender();
}

// Clear All Dataset (Kosongkan Data)
function resetToSampleData() {
  currentDataset = [];
  lastUploadedFileName = null;
  lastUploadDateStr = null;
  const badge = document.getElementById('lastUpdatedBadge');
  if (badge) {
    badge.innerHTML = `
      <i data-lucide="info"></i>
      <span>Data: Kosong (0 Task)</span>
    `;
  }
  lucide.createIcons();
  populateDropdownFilters();
  clearAllFilters();
}

// Filtering Engine
function applyFiltersAndRender() {
  const search = document.getElementById('searchInput').value.toLowerCase();
  
  const selProj = selectedFilters.Project || [];
  const selStatus = selectedFilters.Status || [];
  const selType = selectedFilters.Type || [];
  const selPIC = selectedFilters.PIC || [];

  filteredData = currentDataset.filter(item => {
    const matchesSearch = !search || 
      item.Project.toLowerCase().includes(search) ||
      item["Task Name"].toLowerCase().includes(search) ||
      item.PIC.toLowerCase().includes(search);

    const matchesProj = selProj.length === 0 || selProj.includes(item.Project);
    const matchesStatus = selStatus.length === 0 || selStatus.includes(item.Status);
    const matchesType = selType.length === 0 || selType.includes(item.Type);
    const matchesPIC = selPIC.length === 0 || selPIC.includes(item.PIC);
    const matchesOverdue = !isOverdueOnlyFilter || isOverdueItem(item);

    return matchesSearch && matchesProj && matchesStatus && matchesType && matchesPIC && matchesOverdue;
  });

  sortFilteredData();
  currentPage = 1;

  updateKPICards();
  renderCharts();
  renderTable();
  updatePagination();
}

// Helper to parse date string into JS Date object
function getParsedDate(dateVal) {
  if (!dateVal || dateVal === '-') return null;
  if (dateVal instanceof Date) {
    return isNaN(dateVal.getTime()) ? null : dateVal;
  }

  // Handle Excel Serial Number (e.g. 45532)
  if (typeof dateVal === 'number' && dateVal > 20000 && dateVal < 60000) {
    return new Date(Math.round((dateVal - 25569) * 86400 * 1000));
  }

  const str = String(dateVal).trim();
  if (!str || str === 'undefined' || str === 'null' || str === '-') return null;

  const num = Number(str);
  if (!isNaN(num) && num > 20000 && num < 60000) {
    return new Date(Math.round((num - 25569) * 86400 * 1000));
  }

  // Month translation map (Indonesian & English short/full names)
  const monthMap = {
    'jan': 0, 'januari': 0, 'january': 0,
    'feb': 1, 'februari': 1, 'february': 1,
    'mar': 2, 'maret': 2, 'march': 2,
    'apr': 3, 'april': 3,
    'mei': 4, 'may': 4,
    'jun': 5, 'juni': 5, 'june': 5,
    'jul': 6, 'juli': 6, 'july': 6,
    'ags': 7, 'agu': 7, 'agustus': 7, 'aug': 7, 'august': 7,
    'sep': 8, 'september': 8,
    'okt': 9, 'oktober': 9, 'oct': 9, 'october': 9,
    'nov': 10, 'november': 10,
    'des': 11, 'desember': 11, 'dec': 11, 'december': 11
  };

  const parts = str.split(/[\s\/\-]+/);
  if (parts.length === 3) {
    const mKey = parts[1].toLowerCase();
    if (monthMap[mKey] !== undefined) {
      const day = parseInt(parts[0], 10);
      const month = monthMap[mKey];
      let year = parseInt(parts[2], 10);
      if (year < 100) year += 2000;
      return new Date(year, month, day);
    } else if (!isNaN(parseInt(parts[0])) && !isNaN(parseInt(parts[1])) && !isNaN(parseInt(parts[2]))) {
      let day, month, year;
      if (parts[0].length === 4) {
        year = parseInt(parts[0], 10);
        month = parseInt(parts[1], 10) - 1;
        day = parseInt(parts[2], 10);
      } else {
        day = parseInt(parts[0], 10);
        month = parseInt(parts[1], 10) - 1;
        year = parseInt(parts[2], 10);
        if (year < 100) year += 2000;
      }
      return new Date(year, month, day);
    }
  }

  const parsed = Date.parse(str);
  if (!isNaN(parsed)) return new Date(parsed);

  return null;
}

// Calculate how many days overdue
function getOverdueDays(dateVal) {
  const dueDateObj = getParsedDate(dateVal);
  if (!dueDateObj) return 0;

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  const targetDate = new Date(dueDateObj);
  targetDate.setHours(0, 0, 0, 0);

  const diffTime = today.getTime() - targetDate.getTime();
  const diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 24));
  return diffDays > 0 ? diffDays : 0;
}

// Sorting logic
function sortFilteredData() {
  filteredData.sort((a, b) => {
    if (currentSortColumn === 'Due Date') {
      const dateA = getParsedDate(a["Due Date"]);
      const dateB = getParsedDate(b["Due Date"]);
      const timeA = dateA ? dateA.getTime() : (isSortAscending ? Infinity : -Infinity);
      const timeB = dateB ? dateB.getTime() : (isSortAscending ? Infinity : -Infinity);
      if (timeA < timeB) return isSortAscending ? -1 : 1;
      if (timeA > timeB) return isSortAscending ? 1 : -1;
      return 0;
    }

    if (currentSortColumn === 'Overdue Date') {
      const daysA = isOverdueItem(a) ? getOverdueDays(a["Due Date"]) : (a.Status === 'DONE' ? -2 : -1);
      const daysB = isOverdueItem(b) ? getOverdueDays(b["Due Date"]) : (b.Status === 'DONE' ? -2 : -1);
      if (daysA < daysB) return isSortAscending ? -1 : 1;
      if (daysA > daysB) return isSortAscending ? 1 : -1;
      return 0;
    }

    let valA = a[currentSortColumn] || '';
    let valB = b[currentSortColumn] || '';

    if (valA < valB) return isSortAscending ? -1 : 1;
    if (valA > valB) return isSortAscending ? 1 : -1;
    return 0;
  });
}

// Helper to check if task is overdue
function isOverdueItem(d) {
  if (d.Status === 'DONE' || !d["Due Date"] || d["Due Date"] === '-') return false;
  const days = getOverdueDays(d["Due Date"]);
  return days > 0;
}

// Update KPI Summary Cards & Tooltip Popovers
function updateKPICards() {
  const totalTasks = filteredData.length;
  const projectNamesList = [...new Set(filteredData.map(d => d.Project))].filter(p => p !== '-').sort();
  const uniqueProjectsCount = projectNamesList.length;
  const doneTasks = filteredData.filter(d => d.Status === 'DONE').length;
  const activeTasks = filteredData.filter(d => d.Status === 'IN PROGRESS' || d.Status === 'TODO').length;

  const compRate = totalTasks > 0 ? Math.round((doneTasks / totalTasks) * 100) : 0;

  const elProjects = document.getElementById('kpiTotalProjects');
  if (elProjects) elProjects.textContent = uniqueProjectsCount;

  const elTasks = document.getElementById('kpiTotalTasks');
  if (elTasks) elTasks.textContent = totalTasks;

  const elRate = document.getElementById('kpiCompletionRate');
  if (elRate) elRate.textContent = `${compRate}%`;

  const elBar = document.getElementById('kpiProgressBar');
  if (elBar) elBar.style.width = `${compRate}%`;

  const elActive = document.getElementById('kpiActiveTasks');
  if (elActive) elActive.textContent = activeTasks;

  // Populate Total Projects Popover (Names only, without tasks)
  const projectsListBody = document.getElementById('projectsListBody');
  const elProjCount = document.getElementById('projectsPopoverCount');
  if (elProjCount) elProjCount.textContent = uniqueProjectsCount;

  if (projectsListBody) {
    if (uniqueProjectsCount === 0) {
      projectsListBody.innerHTML = `<div style="text-align:center; color: var(--text-muted); font-size:0.8rem; padding:8px;">Tidak ada proyek.</div>`;
    } else {
      projectsListBody.innerHTML = '';
      projectNamesList.forEach(projName => {
        const pill = document.createElement('div');
        pill.className = 'project-pill-item';
        pill.innerHTML = `📁 ${escapeHtml(projName)}`;
        
        pill.addEventListener('click', (e) => {
          e.stopPropagation();
          if (filterInstances.Project) {
            filterInstances.Project.setSelected([projName]);
          }
          const tableEl = document.getElementById('taskTable');
          if (tableEl) tableEl.scrollIntoView({ behavior: 'smooth' });
        });

        projectsListBody.appendChild(pill);
      });
    }
  }
}

// Render Visualizations with Chart.js
function renderCharts() {
  renderChartStatus();
  renderChartPIC();
  renderChartType();
  renderChartProject();
}

// Chart 1: Status Breakdown (Donut Chart with Project List on Mouse Over)
function renderChartStatus() {
  const ctx = document.getElementById('chartStatus').getContext('2d');
  if (chartStatusInstance) chartStatusInstance.destroy();

  const statusCounts = { 'DONE': 0, 'IN PROGRESS': 0, 'TODO': 0, 'BACKLOG': 0 };
  const statusProjects = { 'DONE': {}, 'IN PROGRESS': {}, 'TODO': {}, 'BACKLOG': {} };

  filteredData.forEach(d => {
    const st = statusCounts[d.Status] !== undefined ? d.Status : 'TODO';
    statusCounts[st]++;

    const projName = d.Project || 'Unknown Project';
    if (!statusProjects[st][projName]) {
      statusProjects[st][projName] = 0;
    }
    statusProjects[st][projName]++;
  });

  chartStatusInstance = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels: Object.keys(statusCounts),
      datasets: [{
        data: Object.values(statusCounts),
        backgroundColor: [
          STATUS_COLORS['DONE'].bg,
          STATUS_COLORS['IN PROGRESS'].bg,
          STATUS_COLORS['TODO'].bg,
          STATUS_COLORS['BACKLOG'].bg
        ],
        borderColor: '#1e293b',
        borderWidth: 2
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'bottom', labels: { color: '#94a3b8', font: { family: 'Inter' } } },
        tooltip: {
          backgroundColor: 'rgba(15, 23, 42, 0.95)',
          titleColor: '#ffffff',
          titleFont: { family: 'Inter', weight: 'bold', size: 13 },
          bodyColor: '#cbd5e1',
          bodyFont: { family: 'Inter', size: 12 },
          borderColor: 'rgba(99, 102, 241, 0.4)',
          borderWidth: 1,
          padding: 14,
          boxPadding: 6,
          displayColors: true,
          callbacks: {
            label: function(context) {
              const statusName = context.label || '';
              const count = context.raw || 0;
              return ` ${statusName}: ${count} task`;
            },
            afterBody: function(context) {
              if (!context || !context.length) return [];
              const statusName = context[0].label;
              const projsObj = statusProjects[statusName] || {};
              const sortedProjs = Object.entries(projsObj).sort((a, b) => b[1] - a[1]);

              if (!sortedProjs.length) return [];

              const lines = ['', '📁 Proyek Terkait:'];
              sortedProjs.forEach(([proj, num]) => {
                lines.push(`  • ${proj} (${num} task)`);
              });
              return lines;
            }
          }
        }
      },
      cutout: '68%'
    }
  });
}

// Chart 2: Workload per PIC (Stacked Bar Chart)
function renderChartPIC() {
  const ctx = document.getElementById('chartPIC').getContext('2d');
  if (chartPICInstance) chartPICInstance.destroy();

  const pics = [...new Set(filteredData.map(d => d.PIC))].filter(p => p !== '-').slice(0, 10);
  const statuses = ['DONE', 'IN PROGRESS', 'TODO', 'BACKLOG'];

  const datasets = statuses.map(st => {
    const data = pics.map(pic => filteredData.filter(d => d.PIC === pic && d.Status === st).length);
    return {
      label: st,
      data: data,
      backgroundColor: STATUS_COLORS[st].bg,
      borderRadius: 4
    };
  });

  chartPICInstance = new Chart(ctx, {
    type: 'bar',
    data: { labels: pics, datasets: datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        x: { stacked: true, ticks: { color: '#94a3b8' }, grid: { display: false } },
        y: { stacked: true, ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.05)' } }
      },
      plugins: {
        legend: { position: 'top', labels: { color: '#94a3b8' } }
      }
    }
  });
}

// Chart 3: Kategori Pekerjaan (Type Breakdown)
function renderChartType() {
  const ctx = document.getElementById('chartType').getContext('2d');
  if (chartTypeInstance) chartTypeInstance.destroy();

  const typeCounts = {};
  filteredData.forEach(d => {
    typeCounts[d.Type] = (typeCounts[d.Type] || 0) + 1;
  });

  chartTypeInstance = new Chart(ctx, {
    type: 'pie',
    data: {
      labels: Object.keys(typeCounts),
      datasets: [{
        data: Object.values(typeCounts),
        backgroundColor: [
          'rgba(99, 102, 241, 0.85)',
          'rgba(6, 182, 212, 0.85)',
          'rgba(244, 63, 94, 0.85)',
          'rgba(245, 158, 11, 0.85)'
        ],
        borderColor: '#1e293b',
        borderWidth: 2
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'bottom', labels: { color: '#94a3b8' } }
      }
    }
  });
}

// Chart 4: Volume Task per Project (Horizontal Bar)
function renderChartProject() {
  const ctx = document.getElementById('chartProject').getContext('2d');
  if (chartProjectInstance) chartProjectInstance.destroy();

  const projCounts = {};
  filteredData.forEach(d => {
    projCounts[d.Project] = (projCounts[d.Project] || 0) + 1;
  });

  // Sort top 10 projects
  const sortedProjs = Object.entries(projCounts)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 8);

  chartProjectInstance = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: sortedProjs.map(p => p[0]),
      datasets: [{
        label: 'Total Tasks',
        data: sortedProjs.map(p => p[1]),
        backgroundColor: 'rgba(99, 102, 241, 0.75)',
        borderColor: '#6366f1',
        borderWidth: 1,
        borderRadius: 6
      }]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        x: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.05)' } },
        y: { ticks: { color: '#94a3b8' }, grid: { display: false } }
      },
      plugins: {
        legend: { display: false }
      }
    }
  });
}

// Render Data Table
function renderTable() {
  const tbody = document.getElementById('tableBody');
  tbody.innerHTML = '';

  const startIndex = (currentPage - 1) * rowsPerPage;
  const paginatedItems = filteredData.slice(startIndex, startIndex + rowsPerPage);

  document.getElementById('rowCountBadge').textContent = `Showing ${filteredData.length} tasks`;

  if (!paginatedItems.length) {
    tbody.innerHTML = `<tr><td colspan="8" style="text-align:center; padding: 32px; color: var(--text-muted);">Tidak ada data task yang memenuhi filter.</td></tr>`;
    return;
  }

  paginatedItems.forEach(item => {
    const tr = document.createElement('tr');
    const isOverdue = isOverdueItem(item);
    const overdueDays = isOverdue ? getOverdueDays(item["Due Date"]) : 0;
    
    if (isOverdue) {
      tr.className = 'tr-overdue';
    }

    let statusClass = 'badge-todo';
    if (item.Status === 'DONE') statusClass = 'badge-done';
    else if (item.Status === 'IN PROGRESS') statusClass = 'badge-inprog';
    else if (item.Status === 'BACKLOG') statusClass = 'badge-backlog';

    const dueDateHtml = isOverdue ? 
      `<span class="date-overdue"><i data-lucide="alert-triangle"></i> ${escapeHtml(item["Due Date"])}</span>` : 
      escapeHtml(item["Due Date"]);

    let overdueColHtml = `<span class="badge-ontrack">On Track</span>`;
    if (isOverdue) {
      overdueColHtml = `<span class="badge-overdue"><i data-lucide="alert-triangle"></i> Terlewat ${overdueDays} Hari (${escapeHtml(item["Due Date"])})</span>`;
    } else if (item.Status === 'DONE') {
      overdueColHtml = `<span class="badge-ontrack" style="color: var(--accent-emerald);"><i data-lucide="check-circle"></i> Selesai</span>`;
    }

    tr.innerHTML = `
      <td><strong>${escapeHtml(item.Project)}</strong></td>
      <td>${escapeHtml(item["Task Name"])}</td>
      <td><span class="status-badge ${statusClass}">${escapeHtml(item.Status)}</span></td>
      <td><span class="type-pill">${escapeHtml(item.Type)}</span></td>
      <td>${escapeHtml(item.PIC)}</td>
      <td>${escapeHtml(item.From)}</td>
      <td>${dueDateHtml}</td>
      <td>${overdueColHtml}</td>
    `;

    tbody.appendChild(tr);
  });

  lucide.createIcons();
}

// Pagination Info & Buttons update
function updatePagination() {
  const totalPages = Math.max(1, Math.ceil(filteredData.length / rowsPerPage));
  document.getElementById('paginationInfo').textContent = `Halaman ${currentPage} dari ${totalPages} (${filteredData.length} data)`;
  document.getElementById('currentPageNum').textContent = currentPage;

  document.getElementById('prevPageBtn').disabled = currentPage === 1;
  document.getElementById('nextPageBtn').disabled = currentPage >= totalPages;
}

// Export Filtered Data to Excel (.xlsx)
function exportToExcel() {
  if (!filteredData.length) {
    alert('Tidak ada data untuk diekspor.');
    return;
  }

  const exportRows = filteredData.map(item => {
    const isOverdue = isOverdueItem(item);
    const overdueDays = isOverdue ? getOverdueDays(item["Due Date"]) : 0;
    return {
      ...item,
      "Overdue Date": isOverdue ? `Terlewat ${overdueDays} Hari (${item["Due Date"]})` : (item.Status === 'DONE' ? 'Selesai' : 'On Track')
    };
  });

  const worksheet = XLSX.utils.json_to_sheet(exportRows);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Filtered Tasks');

  XLSX.writeFile(workbook, `Project_Monitoring_Report_${new Date().toISOString().slice(0, 10)}.xlsx`);
}

// Export Filtered Data to CSV
function exportToCSV() {
  if (!filteredData.length) {
    alert('Tidak ada data untuk diekspor.');
    return;
  }

  const exportRows = filteredData.map(item => {
    const isOverdue = isOverdueItem(item);
    const overdueDays = isOverdue ? getOverdueDays(item["Due Date"]) : 0;
    return {
      ...item,
      "Overdue Date": isOverdue ? `Terlewat ${overdueDays} Hari (${item["Due Date"]})` : (item.Status === 'DONE' ? 'Selesai' : 'On Track')
    };
  });

  const worksheet = XLSX.utils.json_to_sheet(exportRows);
  const csvOutput = XLSX.utils.sheet_to_csv(worksheet);

  const blob = new Blob([csvOutput], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `Project_Monitoring_Report_${new Date().toISOString().slice(0, 10)}.csv`;
  a.click();
}

// Utility: HTML Escape
function escapeHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// Predefined Project Details matching Weekly Update PDF template
const PROJECT_METADATA_PRESETS = {
  'Acquisition System': { section: '1.New Project', pm: 'Julianto Ardiyansah', sa: 'Jihad', dev: 'Wahyu, Angga', user: 'Dhika Shakiya', staging: 'https://', prod: 'https://', targetLive: '-', priority: 'Medium', stage: 'Requirement' },
  'MNC ASIA HOLDING': { section: '1.New Project', pm: 'Julianto Ardiyansah', sa: 'Jihad', dev: 'Riyadh', user: '-', staging: 'https://', prod: 'https://', targetLive: 'May-26', priority: 'Medium', stage: 'Go Live' },
  'Sewa Lokasi': { section: '1.New Project', pm: '-', sa: 'Jihad', dev: 'Riyadh', user: 'Eka, Riko', staging: 'https://rcsewastudio.mncgroup.com/', prod: 'https://sewalokasi.com', targetLive: 'April-26', priority: 'Medium', stage: 'Go Live' },
  'MNC Kids': { section: '1.New Project', pm: 'Julianto', sa: 'Jihad', dev: 'Ardha, Riyadh', user: 'Bagus', staging: '-', prod: '-', targetLive: '22 Juni atau 23 Juli 26', priority: 'Medium', stage: 'Requirement' },
  'Dashboard RnD': { section: '1.New Project', pm: '-', sa: 'Jihad', dev: 'Ridwan Normansyah', user: '-', staging: 'https://spro.mnc.my.id/login', prod: '-', targetLive: '-', priority: 'Medium', stage: 'Live' },
  'Talent Casting Revamp': { section: '1.New Project', pm: 'Julianto', sa: 'Jihad', dev: 'Ardha, Angga', user: 'Pak Bakhti, Garry', staging: 'https://devcasting.mncgroup.com/', prod: '-', targetLive: 'Apr-26', priority: 'Medium', stage: 'Pre Go Live' },
  'OMS MNC Content': { section: '1.New Project', pm: 'Julianto', sa: 'Julianto', dev: 'Wahyu, Angga', user: 'Sukma, Koh Redo', staging: 'https://omsdev_content.mncgroup.com/', prod: '-', targetLive: 'Apr-26', priority: 'High', stage: 'Live' },

  'MNC PEDULI': { section: '2.Continuous Development', pm: '-', sa: 'Dedik', dev: 'Cecep', user: 'Mba Aldini, Raka', staging: 'https://devmncpeduli.mncgroup.com/\nhttps://devwebcms.mncgroup.com/', prod: 'https://mncpeduli.org/\nhttps://webcms.mncgroup.com', targetLive: '-', priority: 'High', stage: 'Live' },
  'MNC Channels': { section: '2.Continuous Development', pm: '-', sa: 'Jihad', dev: 'Ardha, Riyadh', user: 'Firza, Bagus Nova Gyatna', staging: 'Web:rcmncchannels-v2.mncgroup.com\nCMS:https://www.rcmncchannels.mncgroup.com', prod: 'https://www.mncchannels.com', targetLive: '-', priority: 'Low', stage: 'Live' },
  'E-FORM': { section: '2.Continuous Development', pm: '-', sa: 'Jihad', dev: 'Angga', user: 'Bu Emilia', staging: 'https://deveform.mncgroup.com/', prod: 'https://eform.mncgroup.com/', targetLive: '-', priority: 'Medium', stage: 'Live' },
  'MNC Users': { section: '2.Continuous Development', pm: '-', sa: 'Jihad', dev: '-', user: '-', staging: 'devusers2.mncgroup.com', prod: 'https://msender.mncgroup.com', targetLive: '-', priority: 'Medium', stage: 'Live' },
  'Carpool Units': { section: '2.Continuous Development', pm: '-', sa: 'Dedik', dev: 'Cecep', user: 'Pak Yudi FTA', staging: 'https://rccarpoolunit.mncgroup.com/', prod: 'http://carpool.mncgroup.com/', targetLive: '-', priority: 'Medium', stage: 'Live' },
  'CCSA': { section: '2.Continuous Development', pm: '-', sa: 'Dedik', dev: 'Riyadh', user: 'Yordina, Rio Cahyo Sigit', staging: '-', prod: 'https://ccsa.mncgroup.com', targetLive: '-', priority: 'Medium', stage: 'Go Live Preparation' },
  'DMS – Document Management System': { section: '2.Continuous Development', pm: '-', sa: 'Dedik', dev: 'Cecep', user: 'Yordina, Rio Cahyo Sigit', staging: 'https://rcdms.mncgroup.com/', prod: 'https://dms.mncgroup.com/', targetLive: '-', priority: 'MEDIUM', stage: 'Go Live preparation' },
  'MNC Animation': { section: '2.Continuous Development', pm: '-', sa: 'Dedik', dev: 'Riyadh', user: 'Suhendra', staging: 'https://rcmncanimation-v2.mncgroup.com/', prod: 'https://www.mncanimation.com', targetLive: '-', priority: 'MEDIUM', stage: 'Live' },
  'Messaging Engine (MNC Notification Sender)': { section: '2.Continuous Development', pm: '-', sa: 'Jihad', dev: 'Cecep', user: '-', staging: '-', prod: 'https://msender.mncgroup.com', targetLive: '-', priority: 'Medium', stage: 'Go Live preparation' },
  'PLMS (Policy Learning Management System)': { section: '2.Continuous Development', pm: '-', sa: 'Dedik', dev: 'Cecep', user: 'Yordin, Rio Cahyo', staging: 'https://rcplms-v2.mncgroup.com/logi', prod: 'http://plms.mncgroup.com', targetLive: '-', priority: 'Medium', stage: 'Go Live preparation' },
  'IMS (Inventory Management System)': { section: '2.Continuous Development', pm: '-', sa: 'Moch.Isa, Jihad, Julianto', dev: 'Ridwan, Cecep, Hendra', user: '-', staging: '-', prod: '-', targetLive: '-', priority: 'HIGH', stage: 'Live' },
  'GTV': { section: '2.Continuous Development', pm: '-', sa: 'Jihad', dev: 'Riyadh', user: 'M. Zainul', staging: '-', prod: 'https://gtv.id/', targetLive: '-', priority: 'Medium', stage: 'Live' },
  'MNC Content E-commerce': { section: '2.Continuous Development', pm: '-', sa: 'Jihad', dev: 'Ardha, Riyadh', user: '-', staging: 'https://rcmnccontents-v4.mncgroup.com/access-gateway', prod: 'https://mnccontents.com/', targetLive: '-', priority: 'Medium', stage: 'Live' },
  'POBSI': { section: '2.Continuous Development', pm: '-', sa: 'Dedik', dev: 'Ardha, Riyadh', user: 'Pak Andri', staging: 'https://ardha.my.id/pobsi', prod: 'https://www.pobsi.org', targetLive: '-', priority: 'Medium', stage: 'Live' },
  'Jual Asset': { section: '2.Continuous Development', pm: '-', sa: 'Jihad', dev: 'Ardha, Riyadh', user: 'Ibu Henny', staging: 'https://staging.jualaset.id/', prod: '-', targetLive: '-', priority: 'Medium', stage: 'Live' },
  'Metabase': { section: '2.Continuous Development', pm: '-', sa: 'Hendra', dev: 'Hendra', user: '-', staging: '-', prod: '-', targetLive: '-', priority: 'Medium', stage: 'Live' },
  'QR Advertising': { section: '2.Continuous Development', pm: '-', sa: 'Dedik', dev: 'Ridwan', user: '-', staging: '-', prod: '-', targetLive: '-', priority: 'Medium', stage: 'Live' },
  'Website Miss Indonesia': { section: '2.Continuous Development', pm: '-', sa: 'Dedik', dev: 'Ardha, Riyadh', user: '-', staging: '-', prod: '-', targetLive: '-', priority: 'Medium', stage: 'Live' },

  'Odoo Finance Radio': { section: '3.Odoo ERP\n3.1 New Project', pm: 'Chaerul Ismanto', sa: 'Functional ERP', dev: 'Badra, Rahmat, Rizka, Sarip', user: 'User IMG', staging: '-', prod: '-', targetLive: '-', priority: 'High', stage: 'Preparation' },
  'Odoo Finance MDE': { section: '3.Odoo ERP\n3.1 New Project', pm: 'Chaerul Ismanto', sa: 'Functional ERP', dev: 'Badra, Rahmat, Rizka, Sarip', user: 'User FTA , IMG & MDE', staging: 'https://dev3tv.mncgroup.com', prod: '-', targetLive: '-', priority: 'High', stage: 'Preparation' },
  'Odoo Finance MNCH': { section: '3.Odoo ERP\n3.2 Continous Development', pm: 'Chaerul Ismanto', sa: 'Functional ERP', dev: 'Badra, Rahmat, Rizka, Sarip', user: 'User MNCH', staging: '-', prod: '-', targetLive: '-', priority: 'High', stage: 'Live' },
  'Odoo Finance FTA': { section: '3.Odoo ERP\n3.2 Continous Development', pm: 'Chaerul Ismanto', sa: 'Functional ERP', dev: 'Badra, Rahmat, Rizka, Sarip', user: 'User FTA', staging: '-', prod: '-', targetLive: '-', priority: 'High', stage: 'Live' },
  'Odoo Finance MNCP': { section: '3.Odoo ERP\n3.2 Continous Development', pm: 'Chaerul Ismanto', sa: 'Functional ERP', dev: 'Badra, Rahmat, Rizka, Sarip', user: 'User MNCP', staging: '-', prod: '-', targetLive: '-', priority: 'High', stage: 'Live' },
  'Odoo Finance IMG': { section: '3.Odoo ERP\n3.2 Continous Development', pm: 'Chaerul Ismanto', sa: 'Functional ERP', dev: 'Badra, Rahmat, Rizka, Sarip', user: 'User IMG', staging: '-', prod: 'https://imgodooerp.mncgroup.com/', targetLive: '-', priority: 'High', stage: 'Live' }
};

// Helper function to get header upload date
function getHeaderUploadDate() {
  if (lastUploadDateStr) return lastUploadDateStr;
  const now = new Date();
  const day = String(now.getDate()).padStart(2, '0');
  const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  return `${day} ${monthNames[now.getMonth()]} ${now.getFullYear()}`;
}

// Helper functions to determine metadata dynamically
function determinePM(projName, tasks, preset) {
  const projLower = String(projName || '').toLowerCase();
  if (projLower.includes('odoo')) {
    return 'Chaerul Ismanto';
  }

  for (const t of tasks) {
    if (t.PIC && t.PIC !== '-') {
      const picLower = t.PIC.toLowerCase();
      if (picLower.includes('chaerul') || picLower.includes('ismanto')) {
        return 'Chaerul Ismanto';
      }
    }
  }

  if (preset && preset.pm && preset.pm !== '-') {
    const pmLower = preset.pm.toLowerCase();
    if (pmLower.includes('chaerul') || pmLower.includes('ismanto')) {
      return 'Chaerul Ismanto';
    }
    if (pmLower.includes('julianto')) {
      return 'Julianto Ardiyansah';
    }
  }

  return 'Julianto Ardiyansah';
}

function determineSA(projName, tasks, preset) {
  for (const t of tasks) {
    if (t.PIC && t.PIC !== '-') {
      const picLower = t.PIC.toLowerCase();
      if (picLower.includes('jihad')) {
        return 'Jihad';
      }
      if (picLower.includes('dedik')) {
        return 'Dedik';
      }
    }
  }

  if (preset && preset.sa && preset.sa !== '-') {
    const saLower = preset.sa.toLowerCase();
    if (saLower.includes('jihad')) return 'Jihad';
    if (saLower.includes('dedik')) return 'Dedik';
    return preset.sa;
  }

  return '-';
}

function determineDevelopers(projName, tasks, pmName, saName, preset) {
  const devSet = new Set();

  tasks.forEach(t => {
    if (t.PIC && t.PIC !== '-' && t.PIC !== 'SYSTEM' && t.PIC !== 'DOCUMENT' && t.PIC !== 'PEOPLE') {
      const picTrimmed = t.PIC.trim();
      const picLower = picTrimmed.toLowerCase();

      const isPM = (pmName && pmName.toLowerCase().includes(picLower)) || picLower.includes('julianto') || picLower.includes('chaerul');
      const isSA = (saName && saName.toLowerCase().includes(picLower)) || picLower.includes('jihad') || picLower.includes('dedik');

      if (!isPM && !isSA) {
        devSet.add(picTrimmed);
      }
    }
  });

  if (devSet.size === 0 && preset && preset.dev && preset.dev !== '-') {
    preset.dev.split(',').map(d => d.trim()).filter(Boolean).forEach(d => devSet.add(d));
  }

  return devSet.size > 0 ? Array.from(devSet).join(', ') : '-';
}

function determineUserPic(projName, tasks, preset) {
  if (preset && preset.user && preset.user !== '-') {
    return preset.user;
  }
  return '-';
}

function determineTargetLive(projName, tasks, preset) {
  const validDates = tasks.map(t => t["Due Date"]).filter(d => d && d !== '-' && d !== 'undefined');
  if (validDates.length > 0) {
    return validDates[validDates.length - 1];
  }
  if (preset && preset.targetLive && preset.targetLive !== '-') {
    return preset.targetLive;
  }
  return '-';
}

function determinePriority(projName, preset) {
  if (preset && preset.priority && preset.priority !== '-') {
    return preset.priority;
  }
  return 'Medium';
}

function determineStage(projName, preset) {
  if (preset && preset.stage && preset.stage !== '-') {
    return preset.stage;
  }
  return 'Requirement';
}

// Export Filtered Data to PDF — opens a new tab with the exact layout matching attached 40-page reference
function exportToPDF() {
  if (!filteredData || !filteredData.length) {
    alert('Tidak ada data task yang memenuhi filter untuk diekspor ke PDF.');
    return;
  }

  const uploadDateHeader = getHeaderUploadDate();

  // Group tasks by project
  const projectsGrouped = {};
  filteredData.forEach(item => {
    const projName = item.Project || 'Unspecified Project';
    if (!projectsGrouped[projName]) projectsGrouped[projName] = [];
    projectsGrouped[projName].push(item);
  });

  let allPagesHtml = '';
  let currentCategoryHeader = '';

  Object.keys(projectsGrouped).forEach((projName) => {
    const tasks = projectsGrouped[projName];
    const presetKey = findMatchingPresetKey(projName);
    const preset = PROJECT_METADATA_PRESETS[projName] || PROJECT_METADATA_PRESETS[presetKey] || {};

    const pm = determinePM(projName, tasks, preset);
    const sa = determineSA(projName, tasks, preset);
    const dev = determineDevelopers(projName, tasks, pm, sa, preset);

    const userPic = determineUserPic(projName, tasks, preset);
    const stagingUrl = preset.staging || '-';
    const prodUrl = preset.prod || '-';
    const targetLive = determineTargetLive(projName, tasks, preset);
    const priority = determinePriority(projName, preset);
    const stage = determineStage(projName, preset);
    const section = preset.section || '';

    let sectionHtml = '';
    if (section && section !== currentCategoryHeader) {
      currentCategoryHeader = section;
      section.split('\n').forEach(st => {
        sectionHtml += `<div class="pdf-section-title">${escapeHtml(st)}</div>`;
      });
    }

    const formatUrl = (u) => {
      if (!u || u === '-') return '-';
      return u.split('\n').map(link => {
        const clean = link.trim();
        if (clean.startsWith('http://') || clean.startsWith('https://')) {
          return `<a href="${clean}" target="_blank" style="color:#0066cc;">${clean}</a>`;
        }
        return escapeHtml(clean);
      }).join('<br>');
    };

    let taskRowsHtml = '';
    tasks.forEach((t, tIdx) => {
      const actualDate = (t["Actual Date"] && t["Actual Date"] !== '-') ? t["Actual Date"] : (t.From && t.From !== '-' ? t.From : '');
      const noteStr = t.Note || (t.PreviousDate ? `Previous Date ${t.PreviousDate}` : '');
      taskRowsHtml += `<tr>
        <td class="col-no">${tIdx + 1}</td>
        <td class="col-task">${escapeHtml(t["Task Name"] || '')}</td>
        <td class="col-status"><strong>${escapeHtml(t.Status || '')}</strong></td>
        <td class="col-type">${escapeHtml(t.Type || '')}</td>
        <td class="col-pic">${escapeHtml(t.PIC || '-')}</td>
        <td class="col-due">${escapeHtml(t["Due Date"] || '-')}</td>
        <td class="col-actual">${escapeHtml(actualDate)}</td>
        <td class="col-note">${escapeHtml(noteStr)}</td>
      </tr>`;
    });

    const minRows = 5;
    if (tasks.length < minRows) {
      for (let r = tasks.length + 1; r <= minRows; r++) {
        taskRowsHtml += `<tr>
          <td class="col-no">${r}</td>
          <td class="col-task">&nbsp;</td>
          <td class="col-status">&nbsp;</td>
          <td class="col-type">&nbsp;</td>
          <td class="col-pic">&nbsp;</td>
          <td class="col-due">&nbsp;</td>
          <td class="col-actual">&nbsp;</td>
          <td class="col-note">&nbsp;</td>
        </tr>`;
      }
    }

    allPagesHtml += `
    <div class="page">
      <div class="header">
        <h1>Weekly Update ${escapeHtml(uploadDateHeader)}</h1>
        <p class="subtitle">IT Corporate Application</p>
      </div>
      ${sectionHtml}
      <table class="meta-table">
        <tr>
          <td colspan="2" class="lbl">Project Name</td>
          <td class="colon">:</td>
          <td class="val"><strong>${escapeHtml(projName)}</strong></td>
        </tr>
        <tr>
          <td colspan="2" class="lbl">PM</td>
          <td class="colon" rowspan="4" style="vertical-align:middle;text-align:center;">:</td>
          <td class="val">${escapeHtml(pm)}</td>
        </tr>
        <tr><td colspan="2" class="lbl">SA</td><td class="val">${escapeHtml(sa)}</td></tr>
        <tr><td colspan="2" class="lbl">Developer</td><td class="val">${escapeHtml(dev)}</td></tr>
        <tr><td colspan="2" class="lbl">User (PIC)</td><td class="val">${escapeHtml(userPic)}</td></tr>
        <tr>
          <td rowspan="2" class="lbl" style="vertical-align:middle;">URL Web</td>
          <td class="sublbl">Staging</td>
          <td class="colon" rowspan="2" style="vertical-align:middle;text-align:center;">:</td>
          <td class="val">${formatUrl(stagingUrl)}</td>
        </tr>
        <tr><td class="sublbl">Production</td><td class="val">${formatUrl(prodUrl)}</td></tr>
        <tr><td colspan="2" class="lbl">Target Live</td><td class="colon">:</td><td class="val">${escapeHtml(targetLive)}</td></tr>
        <tr><td colspan="2" class="lbl">Project Priority</td><td class="colon">:</td><td class="val">${escapeHtml(priority)}</td></tr>
        <tr><td colspan="2" class="lbl">Project Stage</td><td class="colon">:</td><td class="val">${escapeHtml(stage)}</td></tr>
      </table>
      <table class="task-table">
        <thead>
          <tr>
            <th class="col-no">No</th>
            <th class="col-task" style="text-align:left;">Task Name</th>
            <th class="col-status">Task Status</th>
            <th class="col-type">Task Type</th>
            <th class="col-pic">PIC</th>
            <th class="col-due">Due Date</th>
            <th class="col-actual">Actual Date</th>
            <th class="col-note" style="text-align:left;">Note</th>
          </tr>
        </thead>
        <tbody>${taskRowsHtml}</tbody>
      </table>
    </div>`;
  });

  // 4. Helpdesk Support Summary Pages (Pages 32-34)
  allPagesHtml += `
  <div class="page">
    <div class="header">
      <h1>Weekly Update ${escapeHtml(uploadDateHeader)}</h1>
      <p class="subtitle">IT Corporate Application</p>
    </div>
    <div class="pdf-section-title">4. Helpdesk Support:</div>
    <p style="font-weight:bold;margin-bottom:12px;">Oracle ERP & Odoo</p>
    <p style="margin-bottom:16px;">1. Support Functional Summary by Date</p>
    <div style="text-align:center;margin:20px 0;">
      <div style="background:#f8fafc;border:1px solid #cbd5e1;padding:20px;border-radius:8px;display:inline-block;width:80%;text-align:left;">
        <p style="font-weight:bold;color:#1e3a8a;margin-bottom:4px;">Last Week Requests by Create On and Site ERP</p>
        <p style="font-size:11px;color:#64748b;">Generated by Ticket Oracle on Jul 29, 2026 04:42 PM | Total records: 37</p>
        <p style="font-size:11px;color:#64748b;margin-bottom:15px;">Created Time: From Jul 20, 2026 12:00 AM To Jul 26, 2026 11:59 PM</p>
        <div style="height:220px;background:#fff;border:1px solid #e2e8f0;display:flex;align-items:flex-end;justify-content:space-around;padding:10px;">
          <div style="width:30px;height:40%;background:#3b82f6;border-radius:3px 3px 0 0;" title="Jul 20"></div>
          <div style="width:30px;height:70%;background:#10b981;border-radius:3px 3px 0 0;" title="Jul 21"></div>
          <div style="width:30px;height:95%;background:#f59e0b;border-radius:3px 3px 0 0;" title="Jul 22"></div>
          <div style="width:30px;height:55%;background:#ef4444;border-radius:3px 3px 0 0;" title="Jul 23"></div>
          <div style="width:30px;height:75%;background:#8b5cf6;border-radius:3px 3px 0 0;" title="Jul 24"></div>
          <div style="width:30px;height:30%;background:#ec4899;border-radius:3px 3px 0 0;" title="Jul 25"></div>
        </div>
      </div>
    </div>
    <p style="text-align:center;font-size:12px;">Source: <a href="https://helpdesk.mncgroup.com/" style="color:#0066cc;">https://helpdesk.mncgroup.com/</a></p>
  </div>`;

  // Release Period Summary (Page 35)
  allPagesHtml += `
  <div class="page">
    <div class="header">
      <h1>Weekly Update ${escapeHtml(uploadDateHeader)}</h1>
      <p class="subtitle">IT Corporate Application</p>
    </div>
    <p class="release-title">Release Period 1 Jan – 31 July 2026:</p>
    <table class="task-table" style="width:650px;margin-bottom:24px;">
      <thead><tr style="background:#475569;color:#fff;">
        <th style="width:40px;color:#fff;">No</th>
        <th style="text-align:left;color:#fff;">Website / Web Apps</th>
        <th style="text-align:left;color:#fff;">Details</th>
      </tr></thead>
      <tbody>
        <tr><td style="text-align:center;">1</td><td>OMS Contents</td><td>Order management system</td></tr>
        <tr><td style="text-align:center;">2</td><td>MNC ASIA HOLDING</td><td>Company Profile</td></tr>
        <tr><td style="text-align:center;">3</td><td>IMS</td><td>Inventory management system</td></tr>
        <tr><td style="text-align:center;">4</td><td>RnD</td><td>Dashboard</td></tr>
      </tbody>
    </table>
    <p class="release-title">Ready to Release Period 1 Jan – 30 July 2026:</p>
    <table class="task-table" style="width:650px;">
      <thead><tr style="background:#475569;color:#fff;">
        <th style="width:40px;color:#fff;">No</th>
        <th style="text-align:left;color:#fff;">Website / Web Apps</th>
        <th style="text-align:left;color:#fff;">Details</th>
      </tr></thead>
      <tbody>
        <tr><td style="text-align:center;">1</td><td>Talent Casting Revamp</td><td>&nbsp;</td></tr>
      </tbody>
    </table>
  </div>`;

  // Wiki Pages (Pages 37-40)
  const wikis = WEEKLY_UPDATE_WIKI_DATA;
  const wikisPerPage = 12;
  for (let wp = 0; wp < Math.ceil(wikis.length / wikisPerPage); wp++) {
    let wikiRowsHtml = '';
    wikis.slice(wp * wikisPerPage, (wp + 1) * wikisPerPage).forEach(w => {
      wikiRowsHtml += `<tr>
        <td style="text-align:center;width:35px;">${w.no}</td>
        <td style="width:200px;">${escapeHtml(w.proj)}</td>
        <td style="width:180px;word-break:break-all;">${w.staging !== '-' ? `<a href="${w.staging}" target="_blank" style="color:#0066cc;">${escapeHtml(w.staging)}</a>` : ''}</td>
        <td style="width:180px;word-break:break-all;">${w.prod !== '-' ? `<a href="${w.prod}" target="_blank" style="color:#0066cc;">${escapeHtml(w.prod)}</a>` : ''}</td>
        <td>${escapeHtml(w.info)}</td>
      </tr>`;
    });
    allPagesHtml += `
    <div class="page">
      <div class="header">
        <h1>Weekly Update ${escapeHtml(uploadDateHeader)}</h1>
        <p class="subtitle">IT Corporate Application</p>
      </div>
      ${wp === 0 ? '<div class="pdf-section-title">Wiki</div>' : ''}
      <table class="task-table">
        <thead><tr>
          <th style="width:35px;">No</th>
          <th style="width:200px;text-align:left;">Project Name</th>
          <th style="width:180px;text-align:left;">URL STAGING</th>
          <th style="width:180px;text-align:left;">URL PROD</th>
          <th style="text-align:left;">Project Info</th>
        </tr></thead>
        <tbody>${wikiRowsHtml}</tbody>
      </table>
    </div>`;
  }

  const printCSS = `
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: Arial, Helvetica, sans-serif; font-size: 13px; background: #fff; color: #000; }
    .toolbar {
      position: fixed; top: 0; left: 0; right: 0; z-index: 9999;
      background: #1e293b; padding: 10px 20px;
      display: flex; align-items: center; justify-content: space-between;
      box-shadow: 0 2px 10px rgba(0,0,0,0.4);
    }
    .toolbar .title { color: #fff; font-weight: 600; font-size: 14px; }
    .toolbar .btns { display: flex; gap: 10px; }
    .btn-print {
      background: #3b82f6; color: #fff; border: none; padding: 8px 20px;
      border-radius: 8px; font-size: 13px; font-weight: 600; cursor: pointer;
    }
    .btn-print:hover { background: #2563eb; }
    .pages-wrapper { padding-top: 60px; background: #525659; min-height: 100vh; }
    .page {
      width: 297mm; min-height: 210mm;
      padding: 15mm 18mm; background: #fff;
      margin: 16px auto; box-shadow: 0 2px 12px rgba(0,0,0,0.25);
    }
    .header { margin-bottom: 14px; }
    h1 { font-size: 24px; font-weight: 700; color: #000; margin-bottom: 2px; }
    .subtitle { font-size: 13px; color: #333; margin-bottom: 14px; }
    .pdf-section-title { font-size: 20px; font-weight: 700; margin: 12px 0 10px; }
    .release-title { font-weight: 700; font-style: italic; margin-bottom: 8px; }
    .meta-table { width: 580px; border-collapse: collapse; margin-bottom: 18px; border: 1px solid #000; font-size: 12px; }
    .meta-table td { border: 1px solid #000; padding: 3px 8px; vertical-align: top; }
    .lbl { font-weight: 400; background: #fff; width: 140px; }
    .sublbl { width: 90px; font-weight: 400; }
    .colon { width: 14px; text-align: center; }
    .val { word-break: break-word; }
    .val a { color: #0066cc; }
    .task-table { width: 100%; border-collapse: collapse; margin-bottom: 20px; border: 1px solid #000; font-size: 12px; }
    .task-table th { border: 1px solid #000; background: #fff; font-weight: 700; padding: 5px 7px; text-align: center; }
    .task-table td { border: 1px solid #000; padding: 5px 7px; vertical-align: middle; }
    .col-no { width: 36px; text-align: center; }
    .col-task { width: auto; }
    .col-status { width: 100px; text-align: center; font-weight:600; }
    .col-type { width: 85px; text-align: center; }
    .col-pic { width: 90px; text-align: center; }
    .col-due { width: 88px; text-align: center; }
    .col-actual { width: 88px; text-align: center; }
    .col-note { width: 150px; }
    @media print {
      .toolbar { display: none !important; }
      .pages-wrapper { background: #fff; padding-top: 0; }
      .page { margin: 0; box-shadow: none; page-break-after: always; width: 100%; }
    }
    @page { size: A4 landscape; margin: 0; }
  `;

  const newTab = window.open('', '_blank');
  if (!newTab) {
    alert('Popup diblokir browser. Harap izinkan popup untuk situs ini, lalu coba lagi.');
    return;
  }

  newTab.document.write(`<!DOCTYPE html>
<html lang="id">
<head>
  <meta charset="UTF-8">
  <title>Preview PDF — Weekly Update ${escapeHtml(uploadDateHeader)}</title>
  <style>${printCSS}</style>
</head>
<body>
  <div class="toolbar">
    <span class="title">📄 Weekly Update ${escapeHtml(uploadDateHeader)} — IT Corporate Application</span>
    <div class="btns">
      <button class="btn-print" onclick="window.print()">🖨 Print / Save as PDF</button>
    </div>
  </div>
  <div class="pages-wrapper">
    ${allPagesHtml}
  </div>
</body>
</html>`);
  newTab.document.close();
}



// Export Filtered Data to PDF Template
function exportPdfWeeklyUpdateTemplate() {
  exportToPDF();
}


// In-page PDF Preview Modal
function openPdfPreviewModal(blobUrl, filename) {
  let modal = document.getElementById('pdfPreviewModal');
  if (modal) modal.remove();

  modal = document.createElement('div');
  modal.id = 'pdfPreviewModal';
  modal.style.cssText = `
    position: fixed;
    top: 0; left: 0; right: 0; bottom: 0;
    background: rgba(15, 23, 42, 0.9);
    backdrop-filter: blur(8px);
    z-index: 9999999;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    padding: 20px;
    font-family: 'Inter', sans-serif;
  `;

  modal.innerHTML = `
    <div style="width:95%;max-width:1200px;height:92vh;background:#1e293b;border-radius:12px;border:1px solid rgba(255,255,255,0.1);display:flex;flex-direction:column;overflow:hidden;box-shadow:0 25px 50px -12px rgba(0,0,0,0.6);">
      <div style="padding:14px 20px;background:#0f172a;border-bottom:1px solid rgba(255,255,255,0.1);display:flex;justify-content:space-between;align-items:center;flex-shrink:0;">
        <div style="display:flex;align-items:center;gap:10px;color:#fff;font-weight:600;font-size:0.95rem;">
          <span>📄 Preview PDF Export — ${escapeHtml(filename)}</span>
        </div>
        <div style="display:flex;gap:10px;align-items:center;">
          <button id="openInNewTabBtn" style="background:#059669;color:#fff;padding:7px 16px;border:none;border-radius:8px;font-size:0.85rem;font-weight:600;cursor:pointer;display:inline-flex;align-items:center;gap:6px;">
            🖨 Buka Tab Baru / Print PDF
          </button>
          <a href="${blobUrl}" download="${filename}" style="background:#3b82f6;color:#fff;padding:7px 16px;border-radius:8px;text-decoration:none;font-size:0.85rem;font-weight:600;display:inline-flex;align-items:center;gap:6px;">
            ⬇ Download PDF
          </a>
          <button id="closePdfModalBtn" style="background:rgba(255,255,255,0.1);border:none;color:#94a3b8;padding:7px 14px;border-radius:8px;cursor:pointer;font-size:0.85rem;">
            ✕ Tutup
          </button>
        </div>
      </div>
      <div style="flex:1;background:#525659;overflow:hidden;">
        <iframe src="${blobUrl}" style="width:100%;height:100%;border:none;display:block;"></iframe>
      </div>
    </div>`;

  document.body.appendChild(modal);

  document.getElementById('openInNewTabBtn').addEventListener('click', () => {
    if (document.body.contains(modal)) document.body.removeChild(modal);
    exportToPDF();
  });

  document.getElementById('closePdfModalBtn').addEventListener('click', () => {
    if (document.body.contains(modal)) document.body.removeChild(modal);
  });
}




// Predefined Wiki Data from reference PDF (Pages 37-40)
const WEEKLY_UPDATE_WIKI_DATA = [
  { no: 1, proj: 'Website Jual Aset dan Resources (Website)', staging: 'https://staging.jualaset.id/', prod: '-', info: 'Web site that information asset MNC group to the visitor about Asset that can be sell' },
  { no: 2, proj: 'Website Billiard /POBSI (Website)', staging: '-', prod: 'https://www.pobsi.org', info: 'Web Company profile about pobsi organization (Persatuan Olahraga Billiard)' },
  { no: 3, proj: 'CCSA (Web Apps)', staging: '-', prod: 'https://ccsa.mncgroup.com', info: 'Web Application that support Unit MNC group to make unit Comply with SOP which is user can be input the System by methode Self Assement' },
  { no: 4, proj: 'MNC Channel (Website)', staging: '-', prod: 'https://mncchannels.com/', info: 'Web site company profile about MNC Channels' },
  { no: 5, proj: 'MNC Animations (Website)', staging: '-', prod: 'https://www.mncanimation.com', info: 'Web site company profile MNC Animation' },
  { no: 6, proj: 'Website GTV / Global TV (Website)', staging: '-', prod: 'https://gtv.id', info: 'Web site Company Profile GTV (Global TV)' },
  { no: 7, proj: 'Talent Casting (Web Apps)', staging: '-', prod: 'https://talentcasting.mncgroup.com', info: 'Web application that record registered talent for internal' },
  { no: 8, proj: 'IDMS (Identity management System) (Web Apps)', staging: '-', prod: 'https://idms.mncgroup.com/', info: 'Web Application that can handle automation employee resign and set inactive user email /domain triggered by HRD user' },
  { no: 9, proj: 'MNC Content eCommerce (Website)', staging: 'https://rcmnccontents-v4.mncgroup.com/access-gateway', prod: 'https://mnccontents.com/', info: 'Web site company profile about information content Inventory MNC group' },
  { no: 10, proj: 'Federasi Futsal Indonesia (FFI) (Web site)', staging: '-', prod: 'https://futsalindonesia.org/', info: 'Web site company profile about Futsal Organization' },
  { no: 11, proj: 'User Authentication Engine (MNC User) (Web Apps)', staging: '-', prod: 'https://users.mncgroup.com', info: 'Internal web application use for register Authentication user from others application which purpose' },
  { no: 12, proj: 'Messaging Engine (MNC Notification Sender) (Web Apps)', staging: '-', prod: 'https://msender.mncgroup.com', info: 'Internal web application that can otomatic notification that can be integrated with other System' },
  { no: 13, proj: 'e-Form (Phase1) (Web Apps)', staging: '-', prod: 'https://devhris.mncgroup.com/eform/', info: 'Web app that can be use to mnc user to request AFR (Access Request Forms) & HRF (Hardware Request Forms)' },
  { no: 14, proj: 'Carpool Holding (Web Apps)', staging: '-', prod: 'https://carpool.mncgroup.com', info: 'Carpool Holding (Web Apps)' },
  { no: 15, proj: 'Carpool Unit (Web Apps)', staging: '-', prod: '-', info: 'Application Support for Booking Car and Driver Unit MNC (3FTA, MNC Vision & Others unit)' },
  { no: 16, proj: 'DMS (Web Apps)', staging: '-', prod: 'https://dms.mncgroup.com/', info: 'Digital Document Application for all unit' },
  { no: 17, proj: 'Learning Management System (Website)', staging: '-', prod: '-', info: 'Web Apps to manage data learning management' },
  { no: 18, proj: 'Legal System (Web Apps)', staging: '-', prod: 'https://legalSystem.mncgroup.com/login', info: 'Web Apps for Tracking Legal Document sign off MNC TBK' },
  { no: 19, proj: 'MNC Content (Website)', staging: '-', prod: 'https://mnccontents.com/', info: 'Company profiles website for MNC Content' },
  { no: 20, proj: 'MNC Group (WebSite)', staging: '-', prod: 'https://www.mncgroup.com/', info: "Company profile's website" },
  { no: 21, proj: 'MNC Licensing (WebSite )', staging: '-', prod: 'https://mnclicensing.co.id', info: "Company profile's website about Unit MNC Licencing" },
  { no: 22, proj: 'OMS MDI (Web Apps)', staging: '-', prod: 'https://omsmdi.mncgroup.com', info: 'Web Apps Support for transaction Media Order to Billing and integration with Oracle Finance (RCTI+)' },
  { no: 23, proj: 'OMS MPI (Web Apps)', staging: '-', prod: 'https://omSystem.mncgroup.com/', info: 'Web Apps Support for transaction Media Order to Billing and integration with Oracle Finance (MNC Portal Indonesia)' },
  { no: 24, proj: 'MNC MI (Website)', staging: '-', prod: 'https://www.mnc.co.id/en', info: 'Company Profiles Website MNC Media Investment Ltd, saat ini web tersebut di alihkan/diredirect' },
  { no: 25, proj: 'MNC Peduli (Web site)', staging: '-', prod: 'https://www.mncpeduli.org/', info: 'Company profile Website and Donation online with Payment Gateway (Infokom & MotionPay)' },
  { no: 26, proj: 'MNC User (Web Apps)', staging: '-', prod: 'https://users.mncgroup.com', info: 'Integration Web Apps for all user Register in exchange domain integration with Web Apps existing' },
  { no: 27, proj: "MNC Studio's International (MSI) (Website)", staging: '-', prod: 'https://www.mncdigital.com/', info: "Company Profiles Website MNC Studio's International" },
  { no: 28, proj: 'MNC University (web Site)', staging: '-', prod: 'https://mncu.ac.id/', info: 'SIAKAD Web Apps : Web Apps Information System Academic' },
  { no: 29, proj: 'MEDIACOM (Website)', staging: '-', prod: 'https://www.mediacom.co.id', info: 'Company Profiles Website PT Global Mediacom Tbk.' },
  { no: 30, proj: 'MNC Tbk (Website)', staging: '-', prod: 'https://www.mnc.co.id', info: 'MNC Tbk (Website)' },
  { no: 31, proj: 'MNC Games / Esport Star Indonesia (Website)', staging: '-', prod: 'https://www.esportsstarindonesia.id/en', info: 'Company Profiles Website MNC Esport Star Indonesia' },
  { no: 32, proj: 'Odoo Finance (Web Apps)', staging: '-', prod: 'https://odooerp3tvpregolive.mncgroup.com/web/login', info: 'Finance System Web Apps Implemented : MNC Holding , 3FTV (RCTI, MNCTV, GTV), MNCP' },
  { no: 33, proj: 'Odoo HR (Web Apps)', staging: '-', prod: 'https://mnchodoohr.mncgroup.com', info: 'HR System Web Apps' },
  { no: 34, proj: 'SOS (Sales Order System) (Web Apps)', staging: '-', prod: '-', info: 'Web Apps questioner about practical policy activity in unit business' },
  { no: 35, proj: 'Visqual (Web Apps)', staging: '-', prod: 'https://visqual.mncgroup.com/', info: 'Project Management Tools (Internal Development)' },
  { no: 36, proj: 'PLMS(Web Apps)', staging: 'https://rcplms-v2.mncgroup.com/login', prod: 'http://plms.mncgroup.com/', info: 'Web Apps questioner about practical policy activity in unit business' },
  { no: 37, proj: 'RCTI (Web Apps)', staging: '-', prod: 'https://www.rcti.tv', info: 'Web site company profile Rcti (Rajawali Citra Televisi Indonesia)' },
  { no: 38, proj: 'Transmission Monitoring Center (TMC) (Web Apps)', staging: '-', prod: 'https://tmc.mncgroup.com', info: 'Web App Transmission monitoring Center for Inews Technic Team' },
  { no: 39, proj: 'MNC Group-VP (Website)', staging: '-', prod: 'https://www.mncgroup-vp.com', info: 'Website company Video Profile about MNC Group' },
  { no: 40, proj: 'MVN (MNC Vision Networks) (Web site)', staging: '-', prod: 'https://www.mncvisionnetwork.com', info: 'Website company Profile about MNC Vision Network' },
  { no: 41, proj: 'MNC Pictures (Web site)', staging: '-', prod: 'https://www.mncpictures.com', info: 'Website company Profile about MNC Pictures' },
  { no: 42, proj: 'Star Media Nusantara (SMN) (Web site)', staging: '-', prod: 'https://www.starmedianusantara.com', info: 'Website company Profile about Start Media Nusantara' },
  { no: 43, proj: 'OMS Contents', staging: 'oms-staging.mncgroup.com / devcontents_v2.mncgroup.com', prod: 'omscontents.mncgroup.com', info: 'Website application for process selling program to client.' },
  { no: 45, proj: 'MNC ASIA HOLDING', staging: '-', prod: 'mncasiaholding.com', info: 'Website company Profile about MNC Media & Entertainment focuses on four key business pillars: FTA TV, Pay TV and ICT, iNews Media Group, Digital and Content.' },
  { no: 46, proj: 'RnD', staging: '-', prod: '-', info: '-' }
];

// Helper to fuzzy match preset keys
function findMatchingPresetKey(name) {
  if (!name) return '';
  const lower = name.toLowerCase();
  return Object.keys(PROJECT_METADATA_PRESETS).find(key => 
    key.toLowerCase().includes(lower) || lower.includes(key.toLowerCase())
  ) || '';
}



// ============================================
// DASHBOARD SENAM - COMPLETE VERSION WITH GROUP EXPORT
// ============================================

let dataGlobal = [];
let filteredData = [];
let currentPage = 1;
let itemsPerPage = 10;
let filteredDataForChart = [];
let chartMain, chartYearly, chartMonthly;
let currentDetailData = null;
let activeYearFilter = "all";
let selectedFile = null;
let uploadInProgress = false;
let selectedEmployees = new Map();
let noAttendanceEmployees = [];
let attendanceEmployees = [];
let currentSort = { column: "nama", direction: "asc" };

// Filter states
let shiftStatus = "non_shift";
let dateRangeFilter = {
  start: null,
  end: null,
  active: false,
};

// NEW: Filter struktur lini
let strukturLiniFilter = "all";

// NEW: Filter shift untuk export kelompok
let groupExportShiftFilter = "all"; // all, shift, non_shift

// ============================================
// INITIALIZATION
// ============================================

document.addEventListener("DOMContentLoaded", function () {
  initializeApp();
  setupEventListeners();
  loadData();
});

function initializeApp() {
  const today = new Date();
  const startDate = new Date(2022, 0, 1);
  const endDate = new Date(2032, 11, 1);

  document.getElementById("dateRangeStart").value = formatDate(startDate);
  document.getElementById("dateRangeEnd").value = formatDate(endDate);

  updateDateRangeText();
  updateDataStatus();
}

function setupEventListeners() {
  // Search dengan debounce
  let searchTimeout;
  const searchInput = document.getElementById("searchInput");
  if (searchInput) {
    searchInput.addEventListener("input", function (e) {
      clearTimeout(searchTimeout);
      searchTimeout = setTimeout(() => {
        performSearch(e.target.value);
      }, 300);
    });
  }

  // File input
  const fileInput = document.getElementById("fileInput");
  if (fileInput) {
    fileInput.addEventListener("change", function (e) {
      if (e.target.files.length > 0) {
        handleSelectedFile(e.target.files[0]);
      }
    });
  }

  // Drag and drop
  setupDragAndDrop();

  // ESC key
  document.addEventListener("keydown", function (e) {
    if (e.key === "Escape") {
      closeUploadModal();
      closeDetailModal();
    }
  });

  // Click outside dropdowns
  document.addEventListener("click", function (e) {
    if (!e.target.closest(".dropdown")) {
      closeAllDropdowns();
    }
    if (!e.target.closest(".date-range-container")) {
      hideDateRangeDropdown();
    }
  });
}

function formatDate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  return `${year}-${month}`;
}

function formatDateDisplay(dateStr) {
  if (!dateStr) return "";
  const [year, month] = dateStr.split("-");
  const monthNames = [
    "Januari",
    "Februari",
    "Maret",
    "April",
    "Mei",
    "Juni",
    "Juli",
    "Agustus",
    "September",
    "Oktober",
    "November",
    "Desember",
  ];
  return `${monthNames[parseInt(month) - 1]} ${year}`;
}

// ============================================
// LOADING FUNCTIONS
// ============================================

function showLoading(message = "Memuat...") {
  const loading = document.getElementById("loadingOverlay");
  if (loading) {
    loading.querySelector("p").textContent = message;
    loading.style.display = "flex";
  }
}

function hideLoading() {
  const loading = document.getElementById("loadingOverlay");
  if (loading) {
    loading.style.display = "none";
  }
}

async function loadData() {
  try {
    showLoading("Memuat data...");
    const response = await fetch("/api/data");

    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }

    dataGlobal = await response.json();
    filteredData = [...dataGlobal];

    updateDataStatus();
    renderDashboard();
    populateFilters();
    hideLoading();

    if (dataGlobal.length === 0) {
      showToast("Belum ada data. Silakan upload file Excel.", "info");
    } else {
      showToast(
        `Data berhasil dimuat: ${dataGlobal.length} pegawai`,
        "success"
      );
    }
  } catch (error) {
    console.error("Error loading data:", error);
    hideLoading();
    showToast("Gagal memuat data: " + error.message, "error");

    document.getElementById("tableBody").innerHTML = `
      <tr>
        <td colspan="9" class="empty-table">
          <i class="fas fa-exclamation-triangle"></i>
          <p>Gagal memuat data. Periksa koneksi server.</p>
          <button class="btn btn-primary mt-2" onclick="loadData()">
            <i class="fas fa-sync-alt"></i> Coba Lagi
          </button>
        </td>
      </tr>
    `;
  }
}

function refreshData() {
  loadData();
}

function updateDataStatus() {
  const statusElement = document.getElementById("dataStatus");
  if (!statusElement) return;

  if (dataGlobal.length === 0) {
    statusElement.innerHTML =
      '<i class="fas fa-database"></i> <span>Belum ada data</span>';
    statusElement.className = "header-status empty";
  } else {
    statusElement.innerHTML = `<i class="fas fa-database"></i> <span>${dataGlobal.length} pegawai</span>`;
    statusElement.className = "header-status loaded";
  }
}

// ============================================
// DASHBOARD RENDERING
// ============================================

function renderDashboard() {
  renderTable();
  renderKPI();
  renderChart();
  renderTopList();
  updateTableInfo();
}

function renderKPI() {
  const totalPegawai = filteredData.length;

  let totalSenam = 0;
  let totalPegawaiIkut = 0;

  filteredData.forEach((d) => {
    let attendance = 0;

    if (activeYearFilter && activeYearFilter !== "all") {
      attendance = d.tahunan[activeYearFilter] || 0;
    } else if (
      dateRangeFilter.active &&
      dateRangeFilter.start &&
      dateRangeFilter.end
    ) {
      attendance = getAttendanceInRange(
        d,
        dateRangeFilter.start,
        dateRangeFilter.end
      );
    } else {
      attendance = d.total_all || 0;
    }

    totalSenam += attendance;

    if (attendance > 0) {
      totalPegawaiIkut++;
    }
  });

  const rataSenam =
    totalPegawai > 0 ? (totalSenam / totalPegawai).toFixed(1) : 0;

  const totalTidakIkut = calculateNoAttendanceEmployees();

  document.getElementById("totalPegawai").textContent =
    totalPegawaiIkut.toLocaleString();
  document.getElementById("totalSenam").textContent =
    totalSenam.toLocaleString();
  document.getElementById("rataSenam").textContent = rataSenam;
  document.getElementById("totalTidakIkut").textContent =
    totalTidakIkut.toLocaleString();

  const years = Object.keys(filteredData[0]?.tahunan || {});
  if (years.length > 0) {
    const minYear = Math.min(...years.map((y) => parseInt(y)));
    const maxYear = Math.max(...years.map((y) => parseInt(y)));
    document.getElementById(
      "periodeData"
    ).textContent = `${minYear}-${maxYear}`;
  }
}

function renderChart() {
  const ctx = document.getElementById("mainChart");
  if (!ctx) return;

  if (chartMain) {
    chartMain.destroy();
  }

  const years = [
    "2022",
    "2023",
    "2024",
    "2025",
    "2026",
    "2027",
    "2028",
    "2029",
    "2030",
    "2031",
    "2032",
  ].filter((year) =>
    dataGlobal.some((d) => d.tahunan[year] && d.tahunan[year] > 0)
  );

  if (years.length === 0) {
    ctx.parentElement.innerHTML =
      '<p class="text-center py-4">Tidak ada data untuk ditampilkan</p>';
    return;
  }

  const data = years.map((year) => {
    return filteredDataForChart.reduce(
      (sum, d) => sum + (d.tahunan[year] || 0),
      0
    );
  });

  const chartType = document
    .querySelector(".chart-btn.active")
    ?.textContent?.toLowerCase()
    .includes("line")
    ? "line"
    : document
        .querySelector(".chart-btn.active")
        ?.textContent?.toLowerCase()
        .includes("pie")
    ? "pie"
    : "bar";

  const backgroundColors = [
    "rgba(52, 152, 219, 0.8)",
    "rgba(46, 204, 113, 0.8)",
    "rgba(155, 89, 182, 0.8)",
    "rgba(241, 196, 15, 0.8)",
    "rgba(230, 126, 34, 0.8)",
    "rgba(231, 76, 60, 0.8)",
    "rgba(26, 188, 156, 0.8)",
    "rgba(52, 73, 94, 0.8)",
    "rgba(149, 165, 166, 0.8)",
    "rgba(243, 156, 18, 0.8)",
    "rgba(192, 57, 43, 0.8)",
  ];

  const chartData = {
    labels: years,
    datasets: [
      {
        label: "Total Senam",
        data: data,
        backgroundColor:
          chartType === "pie"
            ? backgroundColors.slice(0, years.length)
            : backgroundColors[0],
        borderColor: "rgba(41, 128, 185, 1)",
        borderWidth: 2,
        fill: chartType === "line",
        tension: 0.4,
      },
    ],
  };

  const options = {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: {
        display: chartType === "pie",
        position: "right",
      },
      tooltip: {
        callbacks: {
          label: (ctx) => `Total: ${ctx.raw.toLocaleString()}`,
        },
      },
    },
    scales:
      chartType !== "pie"
        ? {
            y: {
              beginAtZero: true,
              title: {
                display: true,
                text: "Jumlah Senam",
              },
              ticks: {
                callback: function (value) {
                  return value.toLocaleString();
                },
              },
            },
            x: {
              title: {
                display: true,
                text: "Tahun",
              },
            },
          }
        : {},
  };

  chartMain = new Chart(ctx, {
    type: chartType,
    data: chartData,
    options: options,
  });
}

function changeChartType(type) {
  document.querySelectorAll(".chart-btn").forEach((btn) => {
    btn.classList.remove("active");
  });
  event.currentTarget.classList.add("active");
  renderChart();
}

function renderTopList() {
  const topList = document.getElementById("topList");
  const subtitle = document.getElementById("topSubtitle");

  if (activeYearFilter && activeYearFilter !== "all") {
    subtitle.textContent = `Tahun ${activeYearFilter}`;
  } else {
    subtitle.textContent = "Semua Tahun";
  }

  if (filteredData.length === 0) {
    topList.innerHTML = `
      <div class="empty-top">
        <i class="fas fa-users"></i>
        <p>Tidak ada data</p>
      </div>
    `;
    return;
  }

  let sortedData = [...filteredData];
  if (activeYearFilter && activeYearFilter !== "all") {
    sortedData.sort((a, b) => {
      const valA = a.tahunan[activeYearFilter] || 0;
      const valB = b.tahunan[activeYearFilter] || 0;
      return valB - valA;
    });
  } else {
    sortedData.sort((a, b) => b.total_all - a.total_all);
  }

  const top5 = sortedData.slice(0, 5);

  topList.innerHTML = "";
  top5.forEach((item, index) => {
    const total =
      activeYearFilter && activeYearFilter !== "all"
        ? item.tahunan[activeYearFilter] || 0
        : item.total_all;

    const medalIcons = ["🥇", "🥈", "🥉", "4️⃣", "5️⃣"];

    const div = document.createElement("div");
    div.className = "top-item";
    div.innerHTML = `
      <div class="top-rank">${medalIcons[index] || index + 1}</div>
      <div class="top-info">
        <h5>${item.nama}</h5>
        <p>${item.jabatan || "-"}</p>
      </div>
      <div class="top-value">${total}</div>
    `;

    div.addEventListener("click", () => {
      const idx = filteredData.findIndex((d) => d.id === item.id);
      if (idx !== -1) showDetail(idx);
    });

    topList.appendChild(div);
  });
}

function renderTable() {
  const tbody = document.getElementById("tableBody");
  const startIdx = (currentPage - 1) * itemsPerPage;
  const endIdx = startIdx + itemsPerPage;
  const pageData = filteredData.slice(startIdx, endIdx);

  tbody.innerHTML = "";

  if (pageData.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="10" class="empty-table">
          <i class="fas fa-search"></i>
          <p>${
            filteredData.length === 0
              ? "Belum ada data. Silakan upload file."
              : "Data tidak ditemukan"
          }</p>
          ${
            filteredData.length === 0
              ? `
          <button class="btn btn-primary mt-2" onclick="openUploadModal()">
            <i class="fas fa-upload"></i> Upload File
          </button>`
              : ""
          }
        </td>
      </tr>
    `;
    return;
  }

  pageData.forEach((item, index) => {
    const rowNum = startIdx + index + 1;
    const globalIndex = startIdx + index; // Index dalam filteredData
    const total =
      activeYearFilter && activeYearFilter !== "all"
        ? item.tahunan[activeYearFilter] || 0
        : item.total_all;

    // Gunakan ID asli dari data
    const employeeId = item.id;
    const isSelected = selectedEmployees.has(employeeId);

    const tr = document.createElement("tr");
    if (isSelected) {
      tr.classList.add("table-row-selected");
    }

    tr.innerHTML = `
      <td>
        <input 
          type="checkbox" 
          class="row-checkbox" 
          data-id="${employeeId}"
          data-index="${globalIndex}"
          ${isSelected ? "checked" : ""}
          onchange="toggleRowSelection(this, '${employeeId}', ${globalIndex})"
        />
      </td>
      <td>${rowNum}</td>
      <td><strong>${item.nama}</strong></td>
      <td>${item.nik || "-"}</td>
      <td><span class="gender-badge ${item.jk === "L" ? "male" : "female"}">${
      item.jk || "-"
    }</span></td>
      <td>${item.jabatan || "-"}</td>
      <td>${item.struktur || "-"}</td>
      <td>${item.tempat || "-"}</td>
      <td><span class="badge">${total}</span></td>
      <td>
        <button class="btn btn-sm btn-info" onclick="showDetail(${globalIndex})" title="Detail">
          <i class="fas fa-eye"></i>
        </button>
      </td>
    `;
    tbody.appendChild(tr);
  });

  renderPagination();

  // Update select all checkbox state
  updateSelectAllState();
}

function renderPagination() {
  const pagination = document.getElementById("pagination");
  const totalPages = Math.ceil(filteredData.length / itemsPerPage);

  if (totalPages <= 1) {
    pagination.innerHTML = "";
    return;
  }

  let html = "";

  html += `<button class="page-btn ${
    currentPage === 1 ? "disabled" : ""
  }" onclick="changePage(${currentPage - 1})" ${
    currentPage === 1 ? "disabled" : ""
  }>
    <i class="fas fa-chevron-left"></i>
  </button>`;

  const maxVisiblePages = 5;
  let startPage = Math.max(1, currentPage - 2);
  let endPage = Math.min(totalPages, startPage + maxVisiblePages - 1);

  if (endPage - startPage + 1 < maxVisiblePages) {
    startPage = Math.max(1, endPage - maxVisiblePages + 1);
  }

  if (startPage > 1) {
    html += `<button class="page-btn" onclick="changePage(1)">1</button>`;
    if (startPage > 2) {
      html += `<span class="page-dots">...</span>`;
    }
  }

  for (let i = startPage; i <= endPage; i++) {
    html += `<button class="page-btn ${
      i === currentPage ? "active" : ""
    }" onclick="changePage(${i})">${i}</button>`;
  }

  if (endPage < totalPages) {
    if (endPage < totalPages - 1) {
      html += `<span class="page-dots">...</span>`;
    }
    html += `<button class="page-btn" onclick="changePage(${totalPages})">${totalPages}</button>`;
  }

  html += `<button class="page-btn ${
    currentPage === totalPages ? "disabled" : ""
  }" onclick="changePage(${currentPage + 1})" ${
    currentPage === totalPages ? "disabled" : ""
  }>
    <i class="fas fa-chevron-right"></i>
  </button>`;

  pagination.innerHTML = html;
}

function changePage(page) {
  const totalPages = Math.ceil(filteredData.length / itemsPerPage);
  if (page < 1 || page > totalPages) return;

  currentPage = page;
  renderTable();
  updateTableInfo();
  window.scrollTo({
    top: document.querySelector(".table-container").offsetTop - 20,
    behavior: "smooth",
  });
}

function changeRowsPerPage() {
  itemsPerPage = parseInt(document.getElementById("rowsPerPage").value);
  currentPage = 1;
  renderTable();
  updateTableInfo();
}

function updateTableInfo() {
  const start = (currentPage - 1) * itemsPerPage + 1;
  const end = Math.min(currentPage * itemsPerPage, filteredData.length);
  const total = filteredData.length;

  let filterText = "";
  if (activeYearFilter && activeYearFilter !== "all") {
    filterText = ` • Tahun ${activeYearFilter}`;
  }

  document.getElementById(
    "tableInfo"
  ).textContent = `Menampilkan ${start}-${end} dari ${total} pegawai${filterText}`;
}

function sortTable(column) {
  if (currentSort.column === column) {
    currentSort.direction = currentSort.direction === "asc" ? "desc" : "asc";
  } else {
    currentSort.column = column;
    currentSort.direction = "asc";
  }

  filteredData.sort((a, b) => {
    let aVal = a[column];
    let bVal = b[column];

    if (column === "total_all") {
      aVal =
        activeYearFilter && activeYearFilter !== "all"
          ? a.tahunan[activeYearFilter] || 0
          : a.total_all;
      bVal =
        activeYearFilter && activeYearFilter !== "all"
          ? b.tahunan[activeYearFilter] || 0
          : b.total_all;
    }

    if (typeof aVal === "string") {
      aVal = aVal.toLowerCase();
      bVal = bVal.toLowerCase();
    }

    if (currentSort.direction === "asc") {
      return aVal > bVal ? 1 : -1;
    } else {
      return aVal < bVal ? 1 : -1;
    }
  });

  currentPage = 1;
  renderTable();
  updateTableInfo();
}

// ============================================
// FILTER FUNCTIONS
// ============================================

function populateFilters() {
  const tempatSet = new Set();
  const kelompokSet = new Set();
  const statusSet = new Set();
  const strukturSet = new Set();

  dataGlobal.forEach((item) => {
    if (item.tempat) tempatSet.add(item.tempat);
    if (item.kelompok) kelompokSet.add(item.kelompok);
    if (item.status) statusSet.add(item.status);
    if (item.struktur) strukturSet.add(item.struktur);
  });

  const tempatSelect = document.getElementById("filterTempat");
  tempatSelect.innerHTML = '<option value="">Semua Tempat</option>';
  [...tempatSet].sort().forEach((tempat) => {
    const option = document.createElement("option");
    option.value = tempat;
    option.textContent = tempat;
    tempatSelect.appendChild(option);
  });

  const kelompokSelect = document.getElementById("filterKelompok");
  kelompokSelect.innerHTML = '<option value="">Semua Kelompok</option>';
  [...kelompokSet].sort().forEach((kelompok) => {
    const option = document.createElement("option");
    option.value = kelompok;
    option.textContent = kelompok;
    kelompokSelect.appendChild(option);
  });

  const statusSelect = document.getElementById("filterStatus");
  statusSelect.innerHTML = '<option value="">Semua Status</option>';
  [...statusSet].sort().forEach((status) => {
    const option = document.createElement("option");
    option.value = status;
    option.textContent = status;
    statusSelect.appendChild(option);
  });

  // NEW: Populate filter struktur lini
  const strukturSelect = document.getElementById("filterStruktur");
  if (strukturSelect) {
    strukturSelect.innerHTML = '<option value="all">Semua Struktur</option>';
    [...strukturSet].sort().forEach((struktur) => {
      const option = document.createElement("option");
      option.value = struktur;
      option.textContent = struktur;
      strukturSelect.appendChild(option);
    });
  }
}

function toggleSelectAll() {
  const selectAllCheckbox = document.getElementById("selectAll");
  const checkboxes = document.querySelectorAll(".row-checkbox");

  checkboxes.forEach((checkbox) => {
    checkbox.checked = selectAllCheckbox.checked;
    const employeeId = checkbox.dataset.id;
    const employeeIndex = parseInt(checkbox.dataset.index);

    if (selectAllCheckbox.checked) {
      // Simpan data lengkap pegawai ke Map
      const employee = filteredData[employeeIndex];
      selectedEmployees.set(employeeId, employee);
      checkbox.closest("tr").classList.add("table-row-selected");
    } else {
      selectedEmployees.delete(employeeId);
      checkbox.closest("tr").classList.remove("table-row-selected");
    }
  });

  updateSelectionUI();
}

function toggleFilterPanel() {
  const panel = document.getElementById("filterPanel");
  if (panel.style.display === "none" || !panel.style.display) {
    panel.style.display = "block";
    setTimeout(() => {
      panel.scrollIntoView({ behavior: "smooth", block: "nearest" });
    }, 100);
  } else {
    panel.style.display = "none";
  }
}

function toggleRowSelection(checkbox, employeeId, employeeIndex) {
  const row = checkbox.closest("tr");

  if (checkbox.checked) {
    // Simpan data lengkap pegawai
    const employee = filteredData[employeeIndex];
    selectedEmployees.set(employeeId, employee);
    row.classList.add("table-row-selected");
  } else {
    selectedEmployees.delete(employeeId);
    row.classList.remove("table-row-selected");
  }

  // Update select all checkbox
  const allCheckboxes = document.querySelectorAll(".row-checkbox");
  const checkedCheckboxes = document.querySelectorAll(".row-checkbox:checked");
  const selectAllCheckbox = document.getElementById("selectAll");

  if (selectAllCheckbox) {
    selectAllCheckbox.checked =
      allCheckboxes.length === checkedCheckboxes.length &&
      allCheckboxes.length > 0;
    selectAllCheckbox.indeterminate =
      checkedCheckboxes.length > 0 &&
      checkedCheckboxes.length < allCheckboxes.length;
  }

  updateSelectionUI();
}

function updateSelectionUI() {
  const selectedActions = document.getElementById("selectedActions");
  const selectedCount = document.getElementById("selectedCount");

  if (selectedActions && selectedCount) {
    selectedCount.textContent = selectedEmployees.size;
    selectedActions.style.display =
      selectedEmployees.size > 0 ? "flex" : "none";
  }
}

function clearSelection() {
  selectedEmployees.clear();

  const checkboxes = document.querySelectorAll(".row-checkbox");
  checkboxes.forEach((checkbox) => {
    checkbox.checked = false;
    checkbox.closest("tr").classList.remove("table-row-selected");
  });

  const selectAllCheckbox = document.getElementById("selectAll");
  if (selectAllCheckbox) {
    selectAllCheckbox.checked = false;
    selectAllCheckbox.indeterminate = false;
  }

  updateSelectionUI();
}

function applyFilters() {
  const tempat = document.getElementById("filterTempat").value;
  const kelompok = document.getElementById("filterKelompok").value;
  const status = document.getElementById("filterStatus").value;
  const tahun = document.getElementById("filterTahun").value;
  const struktur = document.getElementById("filterStruktur").value;
  const searchTerm = document
    .getElementById("searchInput")
    .value.toLowerCase()
    .trim();

  activeYearFilter = tahun;
  strukturLiniFilter = struktur;

  filteredData = [...dataGlobal];
  filteredDataForChart = [...dataGlobal];

  if (searchTerm) {
    filteredData = filteredData.filter(
      (item) =>
        (item.nama && item.nama.toLowerCase().includes(searchTerm)) ||
        (item.nik && item.nik.toLowerCase().includes(searchTerm)) ||
        (item.jabatan && item.jabatan.toLowerCase().includes(searchTerm)) ||
        (item.tempat && item.tempat.toLowerCase().includes(searchTerm))
    );

    filteredDataForChart = filteredDataForChart.filter(
      (item) =>
        (item.nama && item.nama.toLowerCase().includes(searchTerm)) ||
        (item.nik && item.nik.toLowerCase().includes(searchTerm)) ||
        (item.jabatan && item.jabatan.toLowerCase().includes(searchTerm)) ||
        (item.tempat && item.tempat.toLowerCase().includes(searchTerm))
    );
  }

  if (tempat) {
    filteredData = filteredData.filter((item) => item.tempat === tempat);
    filteredDataForChart = filteredDataForChart.filter(
      (item) => item.tempat === tempat
    );
  }
  if (kelompok) {
    filteredData = filteredData.filter((item) => item.kelompok === kelompok);
    filteredDataForChart = filteredDataForChart.filter(
      (item) => item.kelompok === kelompok
    );
  }
  if (status) {
    filteredData = filteredData.filter((item) => item.status === status);
    filteredDataForChart = filteredDataForChart.filter(
      (item) => item.status === status
    );
  }

  if (struktur && struktur !== "all") {
    filteredData = filteredData.filter((item) => item.struktur === struktur);
    filteredDataForChart = filteredDataForChart.filter(
      (item) => item.struktur === struktur
    );
  }

  currentPage = 1;
  renderDashboard();
  updateFilterInfo();
}

function updateSelectAllState() {
  const selectAllCheckbox = document.getElementById("selectAll");
  if (!selectAllCheckbox) return;

  const allCheckboxes = document.querySelectorAll(".row-checkbox");
  const checkedCheckboxes = document.querySelectorAll(".row-checkbox:checked");

  if (allCheckboxes.length === 0) {
    selectAllCheckbox.checked = false;
    selectAllCheckbox.indeterminate = false;
  } else if (checkedCheckboxes.length === allCheckboxes.length) {
    selectAllCheckbox.checked = true;
    selectAllCheckbox.indeterminate = false;
  } else if (checkedCheckboxes.length > 0) {
    selectAllCheckbox.checked = false;
    selectAllCheckbox.indeterminate = true;
  } else {
    selectAllCheckbox.checked = false;
    selectAllCheckbox.indeterminate = false;
  }

  updateSelectionUI();
}

function updateFilterInfo() {
  const filterInfo = document.getElementById("filterInfo");
  const activeFilters = [];

  const searchTerm = document.getElementById("searchInput").value.trim();
  if (searchTerm) {
    activeFilters.push(`Pencarian: "${searchTerm}"`);
  }

  if (document.getElementById("filterTempat").value) {
    activeFilters.push(
      `Tempat: ${document.getElementById("filterTempat").value}`
    );
  }
  if (document.getElementById("filterKelompok").value) {
    activeFilters.push(
      `Kelompok: ${document.getElementById("filterKelompok").value}`
    );
  }
  if (document.getElementById("filterStatus").value) {
    activeFilters.push(
      `Status: ${document.getElementById("filterStatus").value}`
    );
  }
  // NEW: Tampilkan filter struktur lini
  if (strukturLiniFilter !== "all") {
    activeFilters.push(`Struktur: ${strukturLiniFilter}`);
  }
  if (activeYearFilter !== "all") {
    activeFilters.push(`Tahun: ${activeYearFilter}`);
  }

  if (activeFilters.length === 0) {
    filterInfo.textContent = "Tidak ada filter aktif";
    filterInfo.className = "filter-info";
  } else {
    filterInfo.textContent = `Filter aktif: ${activeFilters.join(", ")}`;
    filterInfo.className = "filter-info active";
  }
}

function resetAllFilters() {
  document.getElementById("filterTempat").value = "";
  document.getElementById("filterKelompok").value = "";
  document.getElementById("filterStatus").value = "";
  document.getElementById("filterTahun").value = "all";
  document.getElementById("filterStruktur").value = "all";
  document.getElementById("searchInput").value = "";

  activeYearFilter = "all";
  strukturLiniFilter = "all";

  dateRangeFilter.active = false;
  const startDate = new Date(2022, 0, 1);
  const endDate = new Date(2032, 11, 1);
  dateRangeFilter.start = formatDate(startDate);
  dateRangeFilter.end = formatDate(endDate);
  document.getElementById("dateRangeStart").value = dateRangeFilter.start;
  document.getElementById("dateRangeEnd").value = dateRangeFilter.end;
  updateDateRangeText();

  filteredData = [...dataGlobal];
  filteredDataForChart = [...dataGlobal];

  currentPage = 1;
  renderDashboard();
  updateFilterInfo();
  showToast("Semua filter telah direset", "info");
}

function performSearch(keyword) {
  applyFilters();
}

function clearSearch() {
  document.getElementById("searchInput").value = "";
  applyFilters();
}

// ============================================
// DATE RANGE - MANUAL CONTROL
// ============================================

function toggleDateRangeDropdown() {
  const dropdown = document.getElementById("dateRangeDropdown");
  dropdown.classList.toggle("show");
}

function hideDateRangeDropdown() {
  const dropdown = document.getElementById("dateRangeDropdown");
  dropdown.classList.remove("show");
}

function updateDateRangeText() {
  const textElement = document.getElementById("dateRangeText");
  if (dateRangeFilter.active && dateRangeFilter.start && dateRangeFilter.end) {
    textElement.textContent = `${formatDateDisplay(
      dateRangeFilter.start
    )} - ${formatDateDisplay(dateRangeFilter.end)}`;
  } else {
    textElement.textContent = "Rentang Waktu";
  }
}

function applyDateRange() {
  const start = document.getElementById("dateRangeStart").value;
  const end = document.getElementById("dateRangeEnd").value;

  if (!start || !end) {
    showToast("Pilih rentang tanggal terlebih dahulu", "error");
    return;
  }

  if (start > end) {
    showToast(
      "Tanggal awal tidak boleh lebih besar dari tanggal akhir",
      "error"
    );
    return;
  }

  dateRangeFilter = {
    start: start,
    end: end,
    active: true,
  };

  updateDateRangeText();
  hideDateRangeDropdown();

  renderDashboard();

  if (currentDetailData) {
    updateMonthlyDateRangeInfo();
    renderMonthlyData();
  }

  showToast(
    `Filter rentang waktu diterapkan: ${formatDateDisplay(
      start
    )} - ${formatDateDisplay(end)}`,
    "success"
  );
}

function clearDateRange() {
  const startDate = new Date(2022, 0, 1);
  const endDate = new Date(2032, 0, 1);

  dateRangeFilter = {
    start: formatDate(startDate),
    end: formatDate(endDate),
    active: false,
  };

  document.getElementById("dateRangeStart").value = dateRangeFilter.start;
  document.getElementById("dateRangeEnd").value = dateRangeFilter.end;

  updateDateRangeText();
  hideDateRangeDropdown();

  renderDashboard();

  if (currentDetailData) {
    updateMonthlyDateRangeInfo();
    renderMonthlyData();
  }

  showToast("Filter rentang waktu direset", "info");
}

function toggleModalDateRangeDropdown() {
  const dropdown = document.getElementById("modalDateRangeDropdown");
  const isMobile = window.innerWidth <= 768;

  if (dropdown.classList.contains("show")) {
    dropdown.classList.remove("show");
    if (isMobile) {
      removeBackdrop();
    }
  } else {
    dropdown.classList.add("show");
    if (isMobile) {
      createBackdrop();
    }
  }
}

// Fungsi untuk membuat backdrop
function createBackdrop() {
  // Hapus backdrop lama jika ada
  removeBackdrop();

  const backdrop = document.createElement("div");
  backdrop.className = "date-range-backdrop show";
  backdrop.id = "modalDateRangeBackdrop";
  backdrop.onclick = function () {
    closeModalDateRangeDropdown();
  };
  document.body.appendChild(backdrop);
}

// Fungsi untuk menghapus backdrop
function removeBackdrop() {
  const backdrop = document.getElementById("modalDateRangeBackdrop");
  if (backdrop) {
    backdrop.remove();
  }
}

// Fungsi untuk menutup dropdown
function closeModalDateRangeDropdown() {
  const dropdown = document.getElementById("modalDateRangeDropdown");
  dropdown.classList.remove("show");
  removeBackdrop();
}

// ============================================
// SHIFT STATUS
// ============================================

function changeShiftStatus(status) {
  shiftStatus = status;
  const text = status === "shift" ? "Shift" : "Non-Shift";
  document.getElementById("shiftStatusText").textContent = text;

  if (currentDetailData) {
    document.getElementById("detailShiftStatus").textContent = text;
    renderMonthlyData();
  }

  closeAllDropdowns();
  showToast(`Status shift diubah ke: ${text}`, "info");
}

// ============================================
// DETAIL MODAL
// ============================================

function showDetail(index) {
  if (index < 0 || index >= filteredData.length) return;

  currentDetailData = filteredData[index];

  document.getElementById("detailName").textContent = currentDetailData.nama;
  document.getElementById("detailNik").textContent =
    currentDetailData.nik || "-";
  document.getElementById("detailJk").textContent = currentDetailData.jk || "-";
  document.getElementById("detailStatus").textContent =
    currentDetailData.status || "-";
  document.getElementById("detailKelompok").textContent =
    currentDetailData.kelompok || "-";
  document.getElementById("detailJabatan").textContent =
    currentDetailData.jabatan || "-";
  document.getElementById("detailStruktur").textContent =
    currentDetailData.struktur || "-";
  document.getElementById("detailTempat").textContent =
    currentDetailData.tempat || "-";
  document.getElementById("detailShiftStatus").textContent =
    shiftStatus === "shift" ? "Shift" : "Non-Shift";

  const summaryGrid = document.getElementById("summaryYears");
  summaryGrid.innerHTML = "";

  const years = Object.keys(currentDetailData.tahunan).sort();
  years.forEach((year) => {
    const value = currentDetailData.tahunan[year] || 0;
    const div = document.createElement("div");
    div.className = "summary-item";
    div.innerHTML = `
      <div class="summary-year">${year}</div>
      <div class="summary-value">${value}</div>
    `;
    if (activeYearFilter === year) {
      div.classList.add("active");
    }
    summaryGrid.appendChild(div);
  });

  renderYearlyChart();

  const monthYearSelect = document.getElementById("monthYearSelect");
  monthYearSelect.innerHTML = "";
  years.forEach((year) => {
    const option = document.createElement("option");
    option.value = year;
    option.textContent = year;
    monthYearSelect.appendChild(option);
  });

  if (years.length > 0) {
    monthYearSelect.value = years[years.length - 1];
  }

  updateMonthlyDateRangeInfo();
  renderMonthlyData();

  document.getElementById("detailModal").style.display = "flex";
  document.body.style.overflow = "hidden";
  openDetailTab("info");
}

function closeDetailModal() {
  document.getElementById("detailModal").style.display = "none";
  document.body.style.overflow = "auto";
}

function openDetailTab(tabName) {
  document.querySelectorAll(".tab-pane").forEach((tab) => {
    tab.classList.remove("active");
  });

  document.querySelectorAll(".tab-btn").forEach((btn) => {
    btn.classList.remove("active");
  });

  document.getElementById(tabName + "Tab").classList.add("active");
  event.currentTarget.classList.add("active");

  if (tabName === "monthly") {
    setTimeout(() => {
      if (chartMonthly) {
        chartMonthly.resize();
      }
    }, 100);
  }
}

function renderYearlyChart() {
  const ctx = document.getElementById("yearlyChart");
  if (!ctx) return;

  if (chartYearly) {
    chartYearly.destroy();
  }

  const years = Object.keys(currentDetailData.tahunan).sort();
  const data = years.map((year) => currentDetailData.tahunan[year] || 0);

  if (years.length === 0) {
    ctx.parentElement.innerHTML =
      '<p class="text-center py-4">Tidak ada data tahunan</p>';
    return;
  }

  chartYearly = new Chart(ctx, {
    type: "bar",
    data: {
      labels: years,
      datasets: [
        {
          label: "Jumlah Senam",
          data: data,
          backgroundColor: "rgba(52, 152, 219, 0.7)",
          borderColor: "rgba(41, 128, 185, 1)",
          borderWidth: 2,
          borderRadius: 4,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: (ctx) => `Senam: ${ctx.raw} kali`,
          },
        },
      },
      scales: {
        y: {
          beginAtZero: true,
          ticks: { stepSize: 1 },
          title: { display: true, text: "Jumlah Senam" },
        },
        x: {
          title: { display: true, text: "Tahun" },
        },
      },
    },
  });

  renderYearlyTable();
}

function renderYearlyTable() {
  const tbody = document.getElementById("yearlyTableBody");
  const years = Object.keys(currentDetailData.tahunan).sort();

  tbody.innerHTML = "";
  years.forEach((year) => {
    const value = currentDetailData.tahunan[year] || 0;
    const percentage = 80 > 0 ? ((value / 80) * 100).toFixed(1) : 0;

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td><strong>${year}</strong></td>
      <td>${value} kali</td>
      <td><span class="percentage-badge ${
        value >= 20 ? "good" : "bad"
      }">${percentage}%</span></td>
    `;
    tbody.appendChild(tr);
  });
}

function changeMonthYear() {
  renderMonthlyData();
}

function updateMonthlyDateRangeInfo() {
  const infoElement = document.getElementById("monthlyDateRangeInfo");
  if (dateRangeFilter.active && dateRangeFilter.start && dateRangeFilter.end) {
    infoElement.innerHTML = `<i class="fas fa-calendar-alt"></i> ${formatDateDisplay(
      dateRangeFilter.start
    )} - ${formatDateDisplay(dateRangeFilter.end)}`;
  } else {
    infoElement.innerHTML =
      '<i class="fas fa-calendar-alt"></i> Semua bulan (24 bulan terakhir)';
  }
}

function renderMonthlyData() {
  const ctx = document.getElementById("monthlyChart");
  const tbody = document.getElementById("monthlyTableBody");
  const summaryContainer = document.getElementById("monthlySummary");

  if (!ctx || !tbody || !summaryContainer) return;

  if (chartMonthly) {
    chartMonthly.destroy();
  }

  tbody.innerHTML = "";
  summaryContainer.innerHTML = "";

  // Kumpulkan SEMUA bulan dari SEMUA tahun
  let allMonths = [];
  for (const year in currentDetailData.bulanan) {
    for (const [monthKey, monthData] of Object.entries(
      currentDetailData.bulanan[year]
    )) {
      const [yearStr, monthStr] = monthKey.split("-");

      // Filter by date range jika aktif
      if (
        dateRangeFilter.active &&
        dateRangeFilter.start &&
        dateRangeFilter.end
      ) {
        const currentDate = `${yearStr}-${monthStr}`;

        if (
          currentDate < dateRangeFilter.start ||
          currentDate > dateRangeFilter.end
        ) {
          continue;
        }
      }

      allMonths.push({
        key: monthKey,
        year: yearStr,
        month: monthStr,
        name: monthData.nama || `Bulan ${monthStr}`,
        value: monthData.value || 0,
        status: monthData.status || "Tidak Hadir",
      });
    }
  }

  // Sort by date (YYYY-MM)
  allMonths.sort((a, b) => a.key.localeCompare(b.key));

  if (!dateRangeFilter.active) {
    const monthsWithData = allMonths.filter(
      (m) => m.value > 0 || m.status !== "Tidak Ada Data"
    );

    if (monthsWithData.length > 0) {
      const lastMonth = monthsWithData[monthsWithData.length - 1];
      const lastDate = lastMonth.key;

      allMonths = allMonths.filter((m) => {
        const [year, month] = lastDate.split("-");
        const endDate = new Date(parseInt(year), parseInt(month) - 1);
        const startDate = new Date(endDate);
        startDate.setMonth(startDate.getMonth() - 23);

        const [mYear, mMonth] = m.key.split("-");
        const mDate = new Date(parseInt(mYear), parseInt(mMonth) - 1);

        return mDate >= startDate && mDate <= endDate;
      });
    } else {
      allMonths = allMonths.slice(-24);
    }
  }

  let totalAttendance = allMonths.reduce((sum, m) => sum + m.value, 0);
  let totalMonths = allMonths.length;

  let targetTotal;
  if (shiftStatus === "shift") {
    targetTotal = 40;
  } else {
    targetTotal = 56;
  }

  // Populate table
  allMonths.forEach((month) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${month.name} ${month.year}</td>
      <td>${month.value}</td>
      <td><span class="status-badge ${
        month.value > 0 ? "hadir" : "tidak-hadir"
      }">${month.status}</span></td>
    `;
    tbody.appendChild(tr);
  });

  // Summary cards
  const percentage =
    totalMonths > 0 ? ((totalAttendance / totalMonths) * 100).toFixed(1) : 0;
  const statusClass = totalAttendance >= targetTotal ? "good" : "bad";

  summaryContainer.innerHTML = `
    <div class="summary-card">
      <div class="summary-icon">
        <i class="fas fa-calendar-alt"></i>
      </div>
      <div class="summary-content">
        <span class="summary-label">Bulan Ditampilkan</span>
        <h3>${totalMonths}</h3>
      </div>
    </div>
    <div class="summary-card">
      <div class="summary-icon">
        <i class="fas fa-running"></i>
      </div>
      <div class="summary-content">
        <span class="summary-label">Total Hadir</span>
        <h3>${totalAttendance}</h3>
      </div>
    </div>
    <div class="summary-card">
      <div class="summary-icon">
        <i class="fas fa-bullseye"></i>
      </div>
      <div class="summary-content">
        <span class="summary-label">Target (${
          shiftStatus === "shift" ? "50%" : "70%"
        })</span>
        <h3 class="${statusClass}">${targetTotal}</h3>
      </div>
    </div>
    <div class="summary-card">
      <div class="summary-icon">
        <i class="fas fa-percentage"></i>
      </div>
      <div class="summary-content">
        <span class="summary-label">Rata-rata</span>
        <h3>${
          totalMonths > 0 ? (totalAttendance / totalMonths).toFixed(1) : 0
        }</h3>
      </div>
    </div>
  `;

  // Chart
  if (allMonths.length > 0) {
    const chartLabels = allMonths.map(
      (m) => `${m.name.substring(0, 3)} ${m.year.substring(2)}`
    );
    const chartValues = allMonths.map((m) => m.value);

    chartMonthly = new Chart(ctx, {
      type: "bar",
      data: {
        labels: chartLabels,
        datasets: [
          {
            label: "Jumlah Senam",
            data: chartValues,
            backgroundColor: chartValues.map((val) =>
              val > 0 ? "rgba(46, 204, 113, 0.7)" : "rgba(231, 76, 60, 0.7)"
            ),
            borderColor: chartValues.map((val) =>
              val > 0 ? "rgba(39, 174, 96, 1)" : "rgba(192, 57, 43, 1)"
            ),
            borderWidth: 1,
            borderRadius: 3,
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: (ctx) => `Senam: ${ctx.raw} kali`,
            },
          },
        },
        scales: {
          y: {
            beginAtZero: true,
            ticks: { stepSize: 1 },
            title: { display: true, text: "Jumlah Senam" },
          },
          x: {
            ticks: {
              maxRotation: 45,
              minRotation: 45,
              font: {
                size: 10,
              },
            },
          },
        },
      },
    });
  } else {
    ctx.parentElement.innerHTML =
      '<p class="text-center py-4">Tidak ada data bulanan untuk periode ini</p>';
  }
}

async function exportNoAttendanceExcel() {
  try {
    if (noAttendanceEmployees.length === 0) {
      showToast("Tidak ada pegawai yang tidak ikut senam", "info");
      return;
    }

    if (
      !dateRangeFilter.active ||
      !dateRangeFilter.start ||
      !dateRangeFilter.end
    ) {
      showToast("Silakan pilih rentang waktu terlebih dahulu", "error");
      return;
    }

    showLoading(
      `Mengexport Excel untuk ${noAttendanceEmployees.length} pegawai yang tidak ikut senam...`
    );

    const strukturText =
      strukturLiniFilter !== "all" ? strukturLiniFilter : "Semua";

    const response = await fetch("/api/export-no-attendance-excel", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        employees: noAttendanceEmployees,
        date_range: dateRangeFilter,
        shift_status: shiftStatus,
        struktur_lini: strukturText,
      }),
    });

    if (response.ok) {
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      const filename = `pegawai_tidak_ikut_senam_${strukturText.replace(
        / /g,
        "_"
      )}_${new Date().toISOString().slice(0, 10)}.xlsx`;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);

      showToast(
        `✓ Excel berhasil diekspor untuk ${noAttendanceEmployees.length} pegawai yang tidak ikut senam!`,
        "success"
      );
    } else {
      const error = await response.json();
      showToast(error.message || "Gagal mengekspor Excel", "error");
    }
  } catch (error) {
    console.error("Export error:", error);
    showToast("Terjadi kesalahan saat mengekspor Excel", "error");
  } finally {
    hideLoading();
  }
}

// Fungsi untuk export PDF pegawai tidak ikut senam
async function exportNoAttendancePDF() {
  try {
    if (noAttendanceEmployees.length === 0) {
      showToast("Tidak ada pegawai yang tidak ikut senam", "info");
      return;
    }

    if (
      !dateRangeFilter.active ||
      !dateRangeFilter.start ||
      !dateRangeFilter.end
    ) {
      showToast("Silakan pilih rentang waktu terlebih dahulu", "error");
      return;
    }

    showLoading(
      `Membuat PDF untuk ${noAttendanceEmployees.length} pegawai yang tidak ikut senam...`
    );

    const strukturText =
      strukturLiniFilter !== "all" ? strukturLiniFilter : "Semua";

    const response = await fetch("/api/export-no-attendance-pdf", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        employees: noAttendanceEmployees,
        date_range: dateRangeFilter,
        shift_filter: groupExportShiftFilter || "all",
        struktur_lini: strukturText,
      }),
    });

    if (response.ok) {
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      const filename = `pegawai_tidak_ikut_senam_${strukturText.replace(
        / /g,
        "_"
      )}_${new Date().toISOString().slice(0, 10)}.pdf`;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);

      showToast(
        `✓ PDF berhasil diekspor untuk ${noAttendanceEmployees.length} pegawai yang tidak ikut senam!`,
        "success"
      );
    } else {
      const error = await response.json();
      showToast(error.message || "Gagal mengekspor PDF", "error");
    }
  } catch (error) {
    console.error("Export error:", error);
    showToast("Terjadi kesalahan saat mengekspor PDF", "error");
  } finally {
    hideLoading();
  }
}

async function exportAttendanceExcel() {
  try {
    if (attendanceEmployees.length === 0) {
      showToast("Tidak ada pegawai yang ikut senam", "info");
      return;
    }

    if (
      !dateRangeFilter.active ||
      !dateRangeFilter.start ||
      !dateRangeFilter.end
    ) {
      showToast("Silakan pilih rentang waktu terlebih dahulu", "error");
      return;
    }

    showLoading(
      `Mengexport Excel untuk ${attendanceEmployees.length} pegawai yang ikut senam...`
    );

    const strukturText =
      strukturLiniFilter !== "all" ? strukturLiniFilter : "Semua";

    const response = await fetch("/api/export-attendance-excel", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        employees: attendanceEmployees,
        date_range: dateRangeFilter,
        shift_status: shiftStatus,
        struktur_lini: strukturText,
      }),
    });

    if (response.ok) {
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `pegawai_ikut_senam_${strukturText.replace(
        / /g,
        "_"
      )}_${new Date().toISOString().slice(0, 10)}.xlsx`;
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);
      showToast(
        `✓ Excel berhasil diekspor untuk ${attendanceEmployees.length} pegawai yang ikut senam!`,
        "success"
      );
    } else {
      const error = await response.json();
      showToast(error.message || "Gagal mengekspor Excel", "error");
    }
  } catch (error) {
    console.error("Export error:", error);
    showToast("Terjadi kesalahan saat mengekspor Excel", "error");
  } finally {
    hideLoading();
  }
}

async function exportAttendancePDF() {
  try {
    if (attendanceEmployees.length === 0) {
      showToast("Tidak ada pegawai yang ikut senam", "info");
      return;
    }

    if (
      !dateRangeFilter.active ||
      !dateRangeFilter.start ||
      !dateRangeFilter.end
    ) {
      showToast("Silakan pilih rentang waktu terlebih dahulu", "error");
      return;
    }

    showLoading(
      `Membuat PDF untuk ${attendanceEmployees.length} pegawai yang ikut senam...`
    );

    const strukturText =
      strukturLiniFilter !== "all" ? strukturLiniFilter : "Semua";

    const response = await fetch("/api/export-attendance-pdf", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        employees: attendanceEmployees,
        date_range: dateRangeFilter,
        shift_filter: groupExportShiftFilter || "all",
        struktur_lini: strukturText,
      }),
    });

    if (response.ok) {
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `pegawai_ikut_senam_${strukturText.replace(
        / /g,
        "_"
      )}_${new Date().toISOString().slice(0, 10)}.pdf`;
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);
      showToast(
        `✓ PDF berhasil diekspor untuk ${attendanceEmployees.length} pegawai yang ikut senam!`,
        "success"
      );
    } else {
      const error = await response.json();
      showToast(error.message || "Gagal mengekspor PDF", "error");
    }
  } catch (error) {
    console.error("Export error:", error);
    showToast("Terjadi kesalahan saat mengekspor PDF", "error");
  } finally {
    hideLoading();
  }
}

async function exportEmployeePDF() {
  if (!currentDetailData) {
    showToast("Tidak ada data pegawai yang dipilih", "error");
    return;
  }

  try {
    showLoading("Membuat PDF...");

    // Kumpulkan SEMUA bulan dari SEMUA tahun sesuai filter
    let bulananExport = {};
    let allMonths = [];

    for (const year in currentDetailData.bulanan) {
      for (const [monthKey, monthData] of Object.entries(
        currentDetailData.bulanan[year]
      )) {
        const [yearStr, monthStr] = monthKey.split("-");

        // Filter by date range jika aktif
        if (
          dateRangeFilter.active &&
          dateRangeFilter.start &&
          dateRangeFilter.end
        ) {
          const currentDate = `${yearStr}-${monthStr}`;

          if (
            currentDate >= dateRangeFilter.start &&
            currentDate <= dateRangeFilter.end
          ) {
            allMonths.push({ key: monthKey, data: monthData });
          }
        } else {
          allMonths.push({ key: monthKey, data: monthData });
        }
      }
    }

    // Sort by date
    allMonths.sort((a, b) => a.key.localeCompare(b.key));

    // Ambil 24 bulan terakhir jika tidak ada filter
    if (!dateRangeFilter.active) {
      allMonths = allMonths.slice(-24);
    }

    // Convert to object
    allMonths.forEach((item) => {
      bulananExport[item.key] = item.data;
    });

    const exportData = {
      employee_data: currentDetailData,
      date_range: dateRangeFilter.active
        ? dateRangeFilter
        : { start: null, end: null },
      shift_status: shiftStatus,
      bulanan_data: bulananExport,
      selected_year:
        document.getElementById("monthYearSelect")?.value ||
        new Date().getFullYear().toString(),
    };

    const response = await fetch("/api/export-pdf", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(exportData),
    });

    if (response.ok) {
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `rekap_senam_${currentDetailData.nik}_${new Date()
        .toISOString()
        .slice(0, 10)}.pdf`;
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);

      showToast("PDF berhasil diekspor!", "success");
    } else {
      const error = await response.json();
      showToast(error.message || "Gagal mengekspor PDF", "error");
    }
  } catch (error) {
    console.error("Export PDF error:", error);
    showToast("Terjadi kesalahan saat mengekspor PDF", "error");
  } finally {
    hideLoading();
  }
}

async function exportSelectedExcel() {
  if (selectedEmployees.size === 0) {
    showToast("Pilih minimal 1 pegawai terlebih dahulu", "error");
    return;
  }

  try {
    showLoading(`Mengexport Excel untuk ${selectedEmployees.size} pegawai...`);

    // Ambil data lengkap dari Map
    const selectedData = Array.from(selectedEmployees.values());

    const response = await fetch("/api/export-group-excel", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        employees: selectedData,
        date_range: dateRangeFilter.active
          ? dateRangeFilter
          : {
              start: "2022-01",
              end: "2032-12",
            },
        shift_status: shiftStatus,
        struktur_lini: `${selectedEmployees.size} Pegawai Terpilih`,
      }),
    });

    if (response.ok) {
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `rekap_senam_terpilih_${
        selectedEmployees.size
      }pegawai_${new Date().toISOString().slice(0, 10)}.xlsx`;
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);

      showToast(
        `✓ Excel berhasil diekspor untuk ${selectedEmployees.size} pegawai!`,
        "success"
      );
      clearSelection();
    } else {
      const error = await response.json();
      showToast(error.message || "Gagal mengekspor Excel", "error");
    }
  } catch (error) {
    console.error("Export error:", error);
    showToast("Terjadi kesalahan saat mengekspor Excel", "error");
  } finally {
    hideLoading();
  }
}

// Export Selected PDF
async function exportSelectedPDF() {
  if (selectedEmployees.size === 0) {
    showToast("Pilih minimal 1 pegawai terlebih dahulu", "error");
    return;
  }

  if (
    !dateRangeFilter.active ||
    !dateRangeFilter.start ||
    !dateRangeFilter.end
  ) {
    showToast("Silakan pilih rentang waktu terlebih dahulu", "error");
    return;
  }

  try {
    showLoading(`Membuat PDF untuk ${selectedEmployees.size} pegawai...`);

    // Ambil data lengkap dari Map
    const selectedData = Array.from(selectedEmployees.values());

    const response = await fetch("/api/export-group-pdf", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        employees: selectedData,
        date_range: dateRangeFilter,
        shift_filter: "all",
        struktur_lini: `${selectedEmployees.size} Pegawai Terpilih`,
      }),
    });

    if (response.ok) {
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `rekap_senam_terpilih_${
        selectedEmployees.size
      }pegawai_${new Date().toISOString().slice(0, 10)}.pdf`;
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);

      showToast(
        `✓ PDF berhasil diekspor untuk ${selectedEmployees.size} pegawai!`,
        "success"
      );
      clearSelection();
    } else {
      const error = await response.json();
      showToast(error.message || "Gagal mengekspor PDF", "error");
    }
  } catch (error) {
    console.error("Export error:", error);
    showToast("Terjadi kesalahan saat mengekspor PDF", "error");
  } finally {
    hideLoading();
  }
}

function showSelectedEmployees() {
  console.log("=== SELECTED EMPLOYEES ===");
  console.log("Total:", selectedEmployees.size);
  selectedEmployees.forEach((emp, id) => {
    console.log(`${id}: ${emp.nama} - ${emp.tempat} - ${emp.struktur}`);
  });
  console.log("========================");
}

function calculateNoAttendanceEmployees() {
  noAttendanceEmployees = [];
  attendanceEmployees = []; // TAMBAHAN

  filteredData.forEach((emp) => {
    let totalAttendance = 0;

    if (activeYearFilter && activeYearFilter !== "all") {
      totalAttendance = emp.tahunan[activeYearFilter] || 0;
    } else if (
      dateRangeFilter.active &&
      dateRangeFilter.start &&
      dateRangeFilter.end
    ) {
      totalAttendance = getAttendanceInRange(
        emp,
        dateRangeFilter.start,
        dateRangeFilter.end
      );
    } else {
      totalAttendance = emp.total_all || 0;
    }

    if (totalAttendance === 0) {
      noAttendanceEmployees.push(emp);
    } else {
      attendanceEmployees.push(emp);
    }
  });

  return noAttendanceEmployees.length;
}

function getAttendanceInRange(employee, startDate, endDate) {
  let total = 0;

  for (const year in employee.bulanan) {
    for (const [monthKey, monthData] of Object.entries(
      employee.bulanan[year]
    )) {
      if (monthKey >= startDate && monthKey <= endDate) {
        total += monthData.value || 0;
      }
    }
  }

  return total;
}

// ============================================
// UPLOAD FUNCTIONS
// ============================================

function setupDragAndDrop() {
  const dropArea = document.getElementById("dropArea");
  const fileInput = document.getElementById("fileInput");

  if (!dropArea || !fileInput) return;

  ["dragenter", "dragover", "dragleave", "drop"].forEach((eventName) => {
    dropArea.addEventListener(eventName, preventDefaults, false);
    document.body.addEventListener(eventName, preventDefaults, false);
  });

  ["dragenter", "dragover"].forEach((eventName) => {
    dropArea.addEventListener(eventName, highlight, false);
  });

  ["dragleave", "drop"].forEach((eventName) => {
    dropArea.addEventListener(eventName, unhighlight, false);
  });

  dropArea.addEventListener("drop", handleDrop, false);

  function preventDefaults(e) {
    e.preventDefault();
    e.stopPropagation();
  }

  function highlight() {
    dropArea.classList.add("highlight");
  }

  function unhighlight() {
    dropArea.classList.remove("highlight");
  }

  function handleDrop(e) {
    const dt = e.dataTransfer;
    const files = dt.files;

    if (files.length > 0) {
      handleSelectedFile(files[0]);
    }
  }
}

function triggerFileInput() {
  document.getElementById("fileInput").click();
}

function handleSelectedFile(file) {
  const validExtensions = [".xlsx", ".xls", ".csv"];
  const maxSize = 10 * 1024 * 1024;

  const fileExt = "." + file.name.split(".").pop().toLowerCase();

  if (!validExtensions.includes(fileExt)) {
    showToast(
      "Format file tidak didukung. Gunakan .xlsx, .xls, atau .csv",
      "error"
    );
    return;
  }

  if (file.size > maxSize) {
    showToast("Ukuran file terlalu besar. Maksimal 10MB", "error");
    return;
  }

  selectedFile = file;

  document.getElementById("fileName").textContent = file.name;
  document.getElementById("fileSize").textContent = formatFileSize(file.size);
  document.getElementById("fileInfo").style.display = "block";
  document.getElementById("uploadBtn").disabled = false;
  document.getElementById("uploadProgress").style.display = "none";
  document.getElementById("validationResult").style.display = "none";

  validateFileStructure(file);
}

function formatFileSize(bytes) {
  if (bytes === 0) return "0 Bytes";
  const k = 1024;
  const sizes = ["Bytes", "KB", "MB", "GB"];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + " " + sizes[i];
}

async function validateFileStructure(file) {
  try {
    showLoading("Memvalidasi file...");

    const formData = new FormData();
    formData.append("file", file);

    const response = await fetch("/api/validate-template", {
      method: "POST",
      body: formData,
    });

    const result = await response.json();
    hideLoading();

    const validationDiv = document.getElementById("validationDetails");
    validationDiv.innerHTML = "";

    if (result.success) {
      if (result.valid) {
        validationDiv.innerHTML = `
          <p class="text-success"><i class="fas fa-check-circle"></i> ${result.message}</p>
          <p class="text-success"><i class="fas fa-check-circle"></i> ${result.data_rows} baris data ditemukan</p>
          <p class="text-success"><i class="fas fa-check-circle"></i> ${result.total_columns} kolom terdeteksi</p>
        `;
        document.getElementById("uploadBtn").disabled = false;
      } else {
        validationDiv.innerHTML = `
          <p class="text-danger"><i class="fas fa-times-circle"></i> ${
            result.message
          }</p>
          ${
            result.missing_columns
              ? result.missing_columns
                  .map(
                    (col) =>
                      `<p class="text-danger"><i class="fas fa-times-circle"></i> Kolom "${col}" tidak ditemukan</p>`
                  )
                  .join("")
              : ""
          }
        `;
        document.getElementById("uploadBtn").disabled = true;
      }
      document.getElementById("validationResult").style.display = "block";
    } else {
      showToast(result.message || "Gagal memvalidasi file", "error");
    }
  } catch (error) {
    hideLoading();
    console.error("Validation error:", error);
    showToast("Gagal memvalidasi file", "error");
  }
}

function openUploadModal() {
  document.getElementById("uploadModal").style.display = "flex";
  document.body.style.overflow = "hidden";
  resetUploadForm();
}

function closeUploadModal() {
  if (!uploadInProgress) {
    document.getElementById("uploadModal").style.display = "none";
    document.body.style.overflow = "auto";
    resetUploadForm();
  }
}

function removeFile() {
  selectedFile = null;
  resetUploadForm();
}

function resetUploadForm() {
  selectedFile = null;
  document.getElementById("fileInput").value = "";
  document.getElementById("fileInfo").style.display = "none";
  document.getElementById("validationResult").style.display = "none";
  document.getElementById("uploadProgress").style.display = "none";
  document.getElementById("uploadBtn").disabled = true;

  const dropArea = document.getElementById("dropArea");
  dropArea.classList.remove("highlight");
}

async function uploadFile() {
  if (!selectedFile || uploadInProgress) return;

  uploadInProgress = true;

  const progressFill = document.getElementById("progressFill");
  const progressText = document.getElementById("progressText");
  const uploadProgress = document.getElementById("uploadProgress");

  uploadProgress.style.display = "block";
  progressFill.style.width = "0%";
  progressText.textContent = "Menyiapkan upload...";
  document.getElementById("uploadBtn").disabled = true;

  try {
    const formData = new FormData();
    formData.append("file", selectedFile);

    let progress = 0;
    const progressInterval = setInterval(() => {
      if (progress < 90) {
        progress += 5;
        progressFill.style.width = `${progress}%`;
        progressText.textContent =
          progress < 50
            ? "Mengupload file..."
            : progress < 80
            ? "Memproses data..."
            : "Menyimpan data...";
      }
    }, 200);

    const response = await fetch("/api/upload", {
      method: "POST",
      body: formData,
    });

    clearInterval(progressInterval);
    progressFill.style.width = "100%";
    progressText.textContent = "Menyelesaikan...";

    const result = await response.json();

    if (result.success) {
      showToast(result.message || "Data berhasil diupload!", "success");

      setTimeout(async () => {
        await loadData();
        closeUploadModal();
      }, 1500);
    } else {
      showToast(result.message || "Upload gagal", "error");
      document.getElementById("uploadBtn").disabled = false;
    }
  } catch (error) {
    console.error("Upload error:", error);
    showToast("Terjadi kesalahan: " + error.message, "error");
    document.getElementById("uploadBtn").disabled = false;
  } finally {
    uploadInProgress = false;
  }
}

// ============================================
// EXPORT FUNCTIONS
// ============================================

async function exportExcel() {
  try {
    if (filteredData.length === 0) {
      showToast("Tidak ada data untuk diexport", "error");
      return;
    }

    showLoading("Mengexport data ke Excel...");

    const response = await fetch("/api/export-excel", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        data: filteredData,
        filters: {
          year: activeYearFilter,
          dateRange: dateRangeFilter,
          shiftStatus: shiftStatus,
        },
      }),
    });

    if (response.ok) {
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `rekap_senam_${new Date().toISOString().slice(0, 10)}.xlsx`;
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);

      showToast("Data berhasil diexport ke Excel!", "success");
    } else {
      const error = await response.json();
      showToast(error.message || "Export Excel gagal", "error");
    }
  } catch (error) {
    console.error("Export error:", error);
    showToast("Terjadi kesalahan saat export Excel", "error");
  } finally {
    hideLoading();
  }
}

async function exportPDF() {
  showToast("Silakan pilih pegawai untuk export PDF individu", "info");
}

// ============================================
// NEW: EXPORT GROUP FUNCTIONS
// ============================================

function openGroupExportModal() {
  // Validasi awal
  if (filteredData.length === 0) {
    showToast("Tidak ada data untuk diexport", "error");
    return;
  }

  if (
    !dateRangeFilter.active ||
    !dateRangeFilter.start ||
    !dateRangeFilter.end
  ) {
    showToast("Silakan pilih rentang waktu terlebih dahulu", "error");
    return;
  }

  // Tampilkan modal
  document.getElementById("groupExportModal").style.display = "flex";
  document.body.style.overflow = "hidden";

  // Set default shift filter
  groupExportShiftFilter = "all";
  updateGroupShiftButtons();
}

function closeGroupExportModal() {
  document.getElementById("groupExportModal").style.display = "none";
  document.body.style.overflow = "auto";
}

function setGroupShiftFilter(filter) {
  groupExportShiftFilter = filter;
  updateGroupShiftButtons();
}

function updateGroupShiftButtons() {
  // Reset semua button
  document.querySelectorAll(".group-shift-btn").forEach((btn) => {
    btn.classList.remove("active");
  });

  // Aktifkan button yang dipilih
  const activeBtn = document.querySelector(
    `.group-shift-btn[data-shift="${groupExportShiftFilter}"]`
  );
  if (activeBtn) {
    activeBtn.classList.add("active");
  }
}

async function exportGroupPDF() {
  try {
    closeGroupExportModal();
    showLoading("Membuat PDF kelompok...");

    const strukturText =
      strukturLiniFilter !== "all" ? strukturLiniFilter : "Semua";

    const exportData = {
      employees: filteredData,
      date_range: dateRangeFilter,
      shift_filter: groupExportShiftFilter, // NEW: Kirim filter shift
      struktur_lini: strukturText,
    };

    const response = await fetch("/api/export-group-pdf", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(exportData),
    });

    if (response.ok) {
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;

      const shiftText =
        groupExportShiftFilter === "all"
          ? "Semua"
          : groupExportShiftFilter === "shift"
          ? "Shift"
          : "NonShift";
      const filename = `rekap_senam_kelompok_${strukturText.replace(
        / /g,
        "_"
      )}_${shiftText}_${new Date().toISOString().slice(0, 10)}.pdf`;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);

      showToast("PDF kelompok berhasil diekspor!", "success");
    } else {
      const error = await response.json();
      showToast(error.message || "Gagal mengekspor PDF kelompok", "error");
    }
  } catch (error) {
    console.error("Export Group PDF error:", error);
    showToast("Terjadi kesalahan saat mengekspor PDF kelompok", "error");
  } finally {
    hideLoading();
  }
}

async function exportGroupExcel() {
  try {
    if (filteredData.length === 0) {
      showToast("Tidak ada data untuk diexport", "error");
      return;
    }

    if (
      !dateRangeFilter.active ||
      !dateRangeFilter.start ||
      !dateRangeFilter.end
    ) {
      showToast("Silakan pilih rentang waktu terlebih dahulu", "error");
      return;
    }

    showLoading("Mengexport Excel kelompok...");

    const strukturText =
      strukturLiniFilter !== "all" ? strukturLiniFilter : "Semua";

    const exportData = {
      employees: filteredData,
      date_range: dateRangeFilter,
      shift_status: shiftStatus,
      struktur_lini: strukturText,
    };

    const response = await fetch("/api/export-group-excel", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(exportData),
    });

    if (response.ok) {
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      const filename = `rekap_senam_kelompok_${strukturText.replace(
        / /g,
        "_"
      )}_${new Date().toISOString().slice(0, 10)}.xlsx`;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);

      showToast("Excel kelompok berhasil diekspor!", "success");
    } else {
      const error = await response.json();
      showToast(error.message || "Gagal mengekspor Excel kelompok", "error");
    }
  } catch (error) {
    console.error("Export Group Excel error:", error);
    showToast("Terjadi kesalahan saat mengekspor Excel kelompok", "error");
  } finally {
    hideLoading();
  }
}

// ============================================
// UTILITY FUNCTIONS
// ============================================

function showToast(message, type = "info") {
  const toast = document.getElementById("toast");
  if (!toast) return;

  toast.textContent = message;
  toast.className = `toast ${type}`;

  if (toast.timeoutId) {
    clearTimeout(toast.timeoutId);
  }

  setTimeout(() => {
    toast.classList.add("show");
  }, 10);

  toast.timeoutId = setTimeout(() => {
    toast.classList.remove("show");
  }, 4000);
}

function closeAllDropdowns() {
  document.querySelectorAll(".dropdown-menu").forEach((menu) => {
    menu.classList.remove("show");
  });
}

function applyModalDateRange() {
  const start = document.getElementById("modalDateRangeStart").value;
  const end = document.getElementById("modalDateRangeEnd").value;

  if (!start || !end) {
    showToast("Pilih rentang tanggal terlebih dahulu", "error");
    return;
  }

  if (start > end) {
    showToast(
      "Tanggal awal tidak boleh lebih besar dari tanggal akhir",
      "error"
    );
    return;
  }

  dateRangeFilter = {
    start: start,
    end: end,
    active: true,
  };

  updateMonthlyDateRangeInfo();
  renderMonthlyData();
  closeModalDateRangeDropdown();

  showToast(
    `Filter diterapkan: ${formatDateDisplay(start)} - ${formatDateDisplay(
      end
    )}`,
    "success"
  );
}

function clearModalDateRange() {
  const startDate = new Date(2022, 0, 1);
  const endDate = new Date(2032, 11, 1);

  dateRangeFilter = {
    start: formatDate(startDate),
    end: formatDate(endDate),
    active: false,
  };

  document.getElementById("modalDateRangeStart").value = dateRangeFilter.start;
  document.getElementById("modalDateRangeEnd").value = dateRangeFilter.end;

  updateMonthlyDateRangeInfo();
  renderMonthlyData();
  closeModalDateRangeDropdown();

  showToast("Filter rentang waktu direset", "info");
}

// ============================================
// WINDOW EXPORTS
// ============================================

window.applyModalDateRange = applyModalDateRange;
window.clearModalDateRange = clearModalDateRange;
window.openUploadModal = openUploadModal;
window.closeUploadModal = closeUploadModal;
window.triggerFileInput = triggerFileInput;
window.removeFile = removeFile;
window.uploadFile = uploadFile;
window.exportExcel = exportExcel;
window.exportPDF = exportPDF;
window.openGroupExportModal = openGroupExportModal;
window.closeGroupExportModal = closeGroupExportModal;
window.setGroupShiftFilter = setGroupShiftFilter;
window.exportGroupPDF = exportGroupPDF;
window.exportGroupExcel = exportGroupExcel;
window.exportEmployeePDF = exportEmployeePDF;
window.changeShiftStatus = changeShiftStatus;
window.toggleDateRangeDropdown = toggleDateRangeDropdown;
window.applyDateRange = applyDateRange;
window.clearDateRange = clearDateRange;
window.toggleFilterPanel = toggleFilterPanel;
window.resetAllFilters = resetAllFilters;
window.applyFilters = applyFilters;
window.changePage = changePage;
window.changeChartType = changeChartType;
window.showDetail = showDetail;
window.closeDetailModal = closeDetailModal;
window.openDetailTab = openDetailTab;
window.changeMonthYear = changeMonthYear;
window.refreshData = refreshData;
window.sortTable = sortTable;
window.clearSearch = clearSearch;
window.changeRowsPerPage = changeRowsPerPage;
window.toggleSelectAll = toggleSelectAll;
window.toggleRowSelection = toggleRowSelection;
window.clearSelection = clearSelection;
window.exportSelectedExcel = exportSelectedExcel;
window.exportSelectedPDF = exportSelectedPDF;
window.updateSelectAllState = updateSelectAllState;
window.showSelectedEmployees = showSelectedEmployees;
window.exportNoAttendanceExcel = exportNoAttendanceExcel;
window.exportNoAttendancePDF = exportNoAttendancePDF;
window.exportAttendanceExcel = exportAttendanceExcel;
window.exportAttendancePDF = exportAttendancePDF;

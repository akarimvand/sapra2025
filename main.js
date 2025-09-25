// Initialize form counts
window.formCounts = {
    formA: 0,
    formB: 0,
    formC: 0,
    formD: 0
};

const translations = {
    allSystems: "همه سیستم‌ها",
    dashboard: "داشبورد",
    system: "سیستم",
    subsystem: "زیرسیستم",
    loadingData: "در حال بارگذاری داده‌ها...",
    pleaseWait: "لطفا صبر کنید...",
    dataLoadError: "خطا در بارگذاری داده‌ها",
    totalItems: "کل آیتم‌ها",
    completed: "تکمیل شده",
    pending: "در انتظار",
    remaining: "باقی مانده",
    issues: "موانع",
    punch: "پانچ",
    holdPoint: "نقاط توقف",
    formA_title: "فرم A",
    formB_title: "فرم B",
    formC_title: "فرم C",
    formD_title: "فرم D",
    formA_desc: "ارسال شده به کارفرما جهت تایید تکمیل مکانیکی",
    formB_desc: "بازگشت توسط کارفرما با پانچ‌های پیش‌راه‌اندازی",
    formC_desc: "رفع پانچ‌ها و ارسال مجدد جهت تایید",
    formD_desc: "تایید نهایی کارفرما و تحویل زیرسیستم",
    itemDetails: "جزئیات آیتم‌ها",
    noDataToDisplay: "داده‌ای برای نمایش وجود ندارد.",
    overview: "نمای کلی",
    byDiscipline: "بر اساس دیسیپلین",
    bySystem: "بر اساس سیستم",
    generalStatus: "وضعیت کلی",
    selectSubsystemToViewDetails: "برای مشاهده جزئیات، یک زیرسیستم انتخاب کنید.",
    noDisciplineData: "داده‌ای برای این دیسیپلین موجود نیست.",
    noSystemData: "هیچ سیستم یا زیرسیستمی برای نمایش وجود ندارد.",
    table_system: "سیستم",
    table_subsystem: "زیرسیستم",
    table_discipline: "دیسیپلین",
    table_totalItems: "تعداد کل",
    table_completed: "تکمیل شده",
    table_pending: "در انتظار",
    table_punch: "پانچ",
    table_holdPoint: "نقاط توقف",
    table_status: "وضعیت",
    table_type: "نوع",
    table_description: "شرح",
    table_category: "دسته",
    table_pl_no: "شماره PL",
    table_hp_priority: "اولویت HP",
    table_hp_description: "شرح HP",
    table_hp_location: "مکان HP",
    exportNoData: "داده‌ای برای خروجی گرفتن در انتخاب فعلی وجود ندارد.",
    exportError: "خطا در خروجی گرفتن اکسل.",
    exportNoDataModal: "داده‌ای در این مودال برای خروجی گرفتن وجود ندارد.",
    downloadAllError: "خطا در دانلود همه داده‌ها.",
    downloadingFiles: "در حال دانلود فایل‌ها",
    creatingZip: "در حال ساخت فایل فشرده...",
    searchPlaceholder: "جستجوی سیستم...",
    actions: "عملیات",
    docWorkflow: "گردش کار اسناد",
    dbStorage: "انبار داده‌ها",
    reports: "گزارش‌ها",
    exportExcel: "خروجی اکسل",
    downloadAll: "دانلود همه داده‌ها",
    exit: "خروج",
    items: "آیتم‌ها",
    noMatchingItems: "موردی یافت نشد.",
    inSystem: "در سیستم",
    inSubsystem: "در زیرسیستم",
    // Modals
    itemDetailsModal_title: "جزئیات آیتم",
    punchDetailsModal_title: "جزئیات پانچ",
    holdDetailsModal_title: "جزئیات نقاط توقف",
    formDetailsModal_title: "جزئیات فرم",
    modal_filterPlaceholder: "فیلتر...",
    modal_close: "بستن",
    modal_exportVisible: "خروجی موارد قابل مشاهده به اکسل",
    activitiesModal_title: "لیست فعالیت‌ها",
    activitiesModal_tagNo: "شماره تگ",
    activitiesModal_formTitle: "عنوان فرم",
    activitiesModal_status: "وضعیت",
    activitiesModal_noActivity: "هیچ فعالیتی برای این تگ یافت نشد.",
    activitiesModal_progress: "پیشرفت کلی:",
    activitiesModal_export: "خروجی اکسل",
    invalidPassword: "رمز عبور نامعتبر است. لطفا دوباره تلاش کنید.",
    hos_subsystem_name: "نام زیرسیستم"
};

// --- Constants ---
const GITHUB_BASE_URL = "https://raw.githubusercontent.com/akarimvand/SAPRA2/refs/heads/main/dbcsv/";

const CSV_URL = GITHUB_BASE_URL + "DATA.CSV";
const ITEMS_CSV_URL = GITHUB_BASE_URL + "ITEMS.CSV";
const PUNCH_CSV_URL = GITHUB_BASE_URL + "PUNCH.CSV";
const HOLD_POINT_CSV_URL = GITHUB_BASE_URL + "HOLD_POINT.CSV";
const ACTIVITIES_CSV_URL = GITHUB_BASE_URL + "ACTIVITES.CSV";

const COLORS_STATUS_CHARTJS = {
    done: 'rgba(0, 135, 90, 0.8)',
    pending: 'rgba(255, 171, 0, 0.8)',
    remaining: 'rgba(0, 184, 217, 0.8)'
};

const ICONS = {
    Collection: '<i class="bi bi-collection" aria-hidden="true"></i>',
    Folder: '<i class="bi bi-folder" aria-hidden="true"></i>',
    Puzzle: '<i class="bi bi-puzzle" aria-hidden="true"></i>',
    ChevronRight: '<i class="bi bi-chevron-left chevron-toggle" aria-hidden="true"></i>',
    FileEarmarkText: '<i class="bi bi-file-earmark-text"></i>',
    FileEarmarkCheck: '<i class="bi bi-file-earmark-check"></i>',
    FileEarmarkMedical: '<i class="bi bi-file-earmark-medical"></i>',
    FileEarmarkSpreadsheet: '<i class="bi bi-file-earmark-spreadsheet"></i>',
    PieChartIcon: '<i class="bi bi-pie-chart-fill fs-1" aria-hidden="true"></i>'
};

// --- Global State ---
let processedData = { systemMap: {}, subSystemMap: {}, allRawData: [] };
let selectedView = { type: 'all', id: 'all', name: translations.allSystems };
let searchTerm = '';
let activeChartTab = 'Overview';
let aggregatedStats = { totalItems: 0, done: 0, pending: 0, punch: 0, hold: 0, remaining: 0 };
let detailedItemsData = [], punchItemsData = [], holdPointItemsData = [], activitiesData = [];
let displayedItemsInModal = [], currentModalDataType = null;

const chartInstances = { overview: null, disciplines: {}, systems: {} };
let bootstrapTabObjects = {};
let itemDetailsModal, activitiesModal;

// --- DOM Elements ---
const DOMElements = {
    sidebar: document.getElementById('sidebar'),
    sidebarToggle: document.getElementById('sidebarToggle'),
    sidebarOverlay: document.getElementById('sidebarOverlay'),
    mainContent: document.getElementById('mainContent'),
    treeView: document.getElementById('treeView'),
    searchInput: document.getElementById('searchInput'),
    dashboardTitle: document.getElementById('dashboardTitle'),
    totalItemsCounter: document.getElementById('totalItemsCounter'),
    dashboardHeaderContainer: document.getElementById('dashboardHeaderContainer'),
    processStepperContainer: document.getElementById('processStepperContainer'),
    chartTabs: document.getElementById('chartTabs'),
    overviewChartsContainer: document.getElementById('overviewChartsContainer'),
    disciplineChartsContainer: document.getElementById('disciplineChartsContainer'),
    systemChartsContainer: document.getElementById('systemChartsContainer'),
    dataTableHead: document.getElementById('dataTableHead'),
    dataTableBody: document.getElementById('dataTableBody'),
    exportExcelBtn: document.getElementById('exportExcelBtn'),
    errorMessage: document.getElementById('errorMessage'),
    downloadAllBtn: document.getElementById('downloadAllBtn'),
    exitBtn: document.getElementById('exitBtn'),
    itemsDetailsTitle: document.getElementById('itemsDetailsTitle'),
};

// --- Initialization ---
document.addEventListener('DOMContentLoaded', () => {
    initEventListeners();
    initModals();
    loadAndProcessData();
    DOMElements.sidebarToggle.setAttribute('aria-expanded', 'false');
    translateStaticUI();
});

function translateStaticUI() {
    DOMElements.searchInput.placeholder = translations.searchPlaceholder;
    DOMElements.itemsDetailsTitle.textContent = translations.itemDetails;
}

function initModals() {
    itemDetailsModal = new bootstrap.Modal(document.getElementById('itemDetailsModal'), {});
    activitiesModal = new bootstrap.Modal(document.getElementById('activitiesModal'), {});
    const exportDetailsExcelBtn = document.getElementById('exportDetailsExcelBtn');
    if (exportDetailsExcelBtn) {
        exportDetailsExcelBtn.addEventListener('click', handleDetailsExport);
    }
}

function initEventListeners() {
    DOMElements.sidebarToggle.addEventListener('click', () => {
        const isOpen = DOMElements.sidebar.classList.contains('open');
        DOMElements.sidebar.classList.toggle('open');
        DOMElements.mainContent.classList.toggle('sidebar-open');
        DOMElements.sidebarOverlay.style.display = isOpen ? 'none' : 'block';
        DOMElements.sidebarToggle.setAttribute('aria-expanded', String(!isOpen));
    });
    DOMElements.sidebarOverlay.addEventListener('click', () => {
        DOMElements.sidebar.classList.remove('open');
        DOMElements.mainContent.classList.remove('sidebar-open');
        DOMElements.sidebarOverlay.style.display = 'none';
        DOMElements.sidebarToggle.setAttribute('aria-expanded', 'false');
    });

    DOMElements.searchInput.addEventListener('input', (e) => {
        searchTerm = e.target.value.toLowerCase();
        renderSidebar();
    });

    DOMElements.exportExcelBtn.addEventListener('click', handleExport);
    DOMElements.exitBtn.addEventListener('click', () => { window.location.href = 'index.html'; });
    DOMElements.downloadAllBtn.addEventListener('click', handleDownloadAll);

    DOMElements.chartTabs.addEventListener('click', (e) => {
        const button = e.target.closest('button[data-bs-toggle="tab"]');
        if (button) {
            activeChartTab = button.dataset.tabName;
            renderCharts();
        }
    });

    DOMElements.dashboardHeaderContainer.addEventListener('click', handleDetailsClick);
    DOMElements.dataTableBody.addEventListener('click', handleDetailsClick);

    DOMElements.processStepperContainer.addEventListener('click', function(e) {
        const stepElement = e.target.closest('.step');
        if (stepElement) {
            const statusType = stepElement.dataset.status;
            const dataType = statusType.replace('_', '').toLowerCase();
            if (statusType) filterHOSItems(statusType, dataType);
        }
    });

    document.getElementById('itemDetailsModal').addEventListener('keyup', (e) => {
        if (e.target.tagName === 'INPUT' && e.target.closest('#modal-filter-row')) {
            filterModalTable();
        }
    });
}

function filterModalTable() {
    const tableBody = document.getElementById('itemDetailsTableBody');
    const rows = tableBody.getElementsByTagName('tr');
    const filters = Array.from(document.querySelectorAll('#modal-filter-row input')).map(input => input.value.toLowerCase());

    for (const row of rows) {
        let display = true;
        const cells = row.getElementsByTagName('td');
        for (let i = 0; i < filters.length; i++) {
            if (filters[i] && cells[i]) {
                const cellText = cells[i].textContent.toLowerCase();
                if (!cellText.includes(filters[i])) {
                    display = false;
                    break;
                }
            }
        }
        row.style.display = display ? '' : 'none';
    }
}

function handleDetailsClick(e) {
    let target = e.target;
    let statusType = null;
    let filterContext = null;
    let dataType = null;

    const statItem = target.closest('.stat-item');
    const dataTableCell = target.closest('#dataTableBody td, #dataTableBody th');

    if (statItem) {
        statusType = statItem.dataset.status;
        filterContext = { type: 'summary', status: statusType };
    } else if (dataTableCell) {
        const tableRow = dataTableCell.closest('tr');
        const cellIndex = Array.from(tableRow.children).indexOf(dataTableCell);
        const headerCell = DOMElements.dataTableHead.querySelector(`th:nth-child(${cellIndex + 1})`);
        if (!headerCell) return;
        const headerText = headerCell.textContent.trim();

        const headerMap = {
            [translations.table_completed]: 'DONE', [translations.table_pending]: 'PENDING',
            [translations.table_punch]: 'PUNCH', [translations.table_holdPoint]: 'HOLD',
            [translations.table_status]: 'OTHER', [translations.table_totalItems]: 'TOTAL'
        };
        statusType = headerMap[headerText];

        if (statusType) {
            const rowData = {};
            Array.from(tableRow.children).forEach((cell, idx) => {
                const accessorMap = ['system', 'subsystem', 'discipline', 'totalItems', 'completed', 'pending', 'punch', 'holdPoint', 'statusPercent'];
                rowData[accessorMap[idx]] = cell.textContent.trim();
            });
            filterContext = { type: 'table', rowData: rowData, status: statusType };
        }
    }

    if (statusType) {
        dataType = (statusType === 'PUNCH') ? 'punch' : (statusType === 'HOLD') ? 'hold' : 'items';
        let dataToDisplay = [];

        if (dataType === 'items') dataToDisplay = filterDetailedItems(filterContext);
        else if (dataType === 'punch') dataToDisplay = filterPunchItems(filterContext);
        else if (dataType === 'hold') dataToDisplay = filterHoldItems(filterContext);

        populateDetailsModal(dataToDisplay, filterContext, dataType);
        itemDetailsModal.show();
    }
}

function filterDetailedItems(context) {
    let filtered = detailedItemsData;
    let modalTitle = translations.itemDetailsModal_title;

    if (context.type === 'summary') {
        const statusText = context.status === 'DONE' ? translations.completed : context.status === 'PENDING' ? translations.pending : context.status === 'TOTAL' ? translations.totalItems : translations.remaining;
        if (selectedView.type === 'system' && selectedView.id) {
            const subSystemIds = processedData.systemMap[selectedView.id]?.subs.map(sub => sub.id.toLowerCase()) || [];
            filtered = filtered.filter(item => item.subsystem && item.subsystem.toLowerCase() && subSystemIds.includes(item.subsystem.toLowerCase()));
            modalTitle = `${statusText} ${translations.inSystem}: ${selectedView.name}`;
        } else if (selectedView.type === 'subsystem' && selectedView.id) {
            filtered = filtered.filter(item => item.subsystem && item.subsystem.toLowerCase() === selectedView.id.toLowerCase());
            modalTitle = `${statusText} ${translations.inSubsystem}: ${selectedView.name}`;
        } else {
            modalTitle = `${statusText} (${translations.allSystems})`;
        }

        if (context.status !== 'TOTAL') {
            if (context.status === 'OTHER') {
                filtered = filtered.filter(item => !item.status || (item.status.toLowerCase() !== 'done' && item.status.toLowerCase() !== 'pending'));
            } else {
                filtered = filtered.filter(item => item.status && item.status.toLowerCase() === context.status.toLowerCase());
            }
        }
    } else if (context.type === 'table') {
        const rowData = context.rowData;
        const clickedSubsystem = rowData.subsystem.split(' - ')[0].toLowerCase();
        const clickedDiscipline = rowData.discipline.toLowerCase();
        filtered = filtered.filter(item => item.subsystem && item.subsystem.toLowerCase() === clickedSubsystem && item.discipline && item.discipline.toLowerCase() === clickedDiscipline);
        if (context.status !== 'TOTAL') {
            if (context.status === 'OTHER') {
                filtered = filtered.filter(item => !item.status || (item.status.toLowerCase() !== 'done' && item.status.toLowerCase() !== 'pending'));
            } else {
                filtered = filtered.filter(item => item.status && item.status.toLowerCase() === context.status.toLowerCase());
            }
        }
        const statusText = context.status === 'DONE' ? translations.completed : context.status === 'TOTAL' ? translations.totalItems : context.status;
        modalTitle = `${statusText} ${translations.inSubsystem} ${rowData.subsystem.split(' - ')[0]} / ${rowData.discipline}`;
    }

    document.getElementById('itemDetailsModalLabel').textContent = modalTitle;
    return filtered;
}

function filterPunchItems(context) {
    let filtered = punchItemsData;
    let modalTitle = translations.punchDetailsModal_title;

    if (context.type === 'summary') {
        if (selectedView.type === 'system' && selectedView.id) {
            const system = processedData.systemMap[selectedView.id];
            if (system) {
                const subSystemIds = system.subs.map(sub => sub.id.toLowerCase());
                filtered = filtered.filter(item => item.SD_Sub_System && subSystemIds.includes(item.SD_Sub_System.trim().toLowerCase()));
            }
            modalTitle = `${translations.punch} ${translations.inSystem}: ${selectedView.name}`;
        } else if (selectedView.type === 'subsystem' && selectedView.id) {
            filtered = filtered.filter(item => item.SD_Sub_System && item.SD_Sub_System.trim().toLowerCase() === selectedView.id.toLowerCase());
            modalTitle = `${translations.punch} ${translations.inSubsystem}: ${selectedView.name}`;
        } else {
            modalTitle = `${translations.punch} (${translations.allSystems})`;
        }
    } else if (context.type === 'table') {
        const rowData = context.rowData;
        const clickedSubsystem = rowData.subsystem.split(' - ')[0].trim().toLowerCase();
        const clickedDiscipline = rowData.discipline.trim().toLowerCase();
        filtered = filtered.filter(item => item.SD_Sub_System && item.SD_Sub_System.trim().toLowerCase() === clickedSubsystem && item.Discipline_Name && item.Discipline_Name.trim().toLowerCase() === clickedDiscipline);
        modalTitle = `${translations.punch} ${translations.inSubsystem} ${rowData.subsystem.split(' - ')[0]} / ${rowData.discipline}`;
    }

    filtered = filtered.filter(item => item.SD_Sub_System && item.Discipline_Name && item.ITEM_Tag_NO && item.PL_Punch_Category);
    document.getElementById('itemDetailsModalLabel').textContent = modalTitle;
    return filtered;
}

function filterHoldItems(context) {
    let filtered = holdPointItemsData;
    let modalTitle = translations.holdDetailsModal_title;

    if (context.type === 'summary') {
        if (selectedView.type === 'system' && selectedView.id) {
            const subSystemIds = processedData.systemMap[selectedView.id]?.subs.map(sub => sub.id.toLowerCase()) || [];
            filtered = filtered.filter(item => item.subsystem.toLowerCase() && subSystemIds.includes(item.subsystem.toLowerCase()));
            modalTitle = `${translations.holdPoint} ${translations.inSystem}: ${selectedView.name}`;
        } else if (selectedView.type === 'subsystem' && selectedView.id) {
            filtered = filtered.filter(item => item.subsystem.toLowerCase() === selectedView.id.toLowerCase());
            modalTitle = `${translations.holdPoint} ${translations.inSubsystem}: ${selectedView.name}`;
        } else {
            modalTitle = `${translations.holdPoint} (${translations.allSystems})`;
        }
    } else if (context.type === 'table') {
        const rowData = context.rowData;
        const clickedSubsystem = rowData.subsystem.split(' - ')[0].toLowerCase();
        const clickedDiscipline = rowData.discipline.toLowerCase();
        filtered = filtered.filter(item => item.subsystem && item.subsystem.toLowerCase() === clickedSubsystem && item.discipline && item.discipline.toLowerCase() === clickedDiscipline);
        modalTitle = `${translations.holdPoint} ${translations.inSubsystem} ${rowData.subsystem.split(' - ')[0]} / ${rowData.discipline}`;
    }

    document.getElementById('itemDetailsModalLabel').textContent = modalTitle;
    return filtered;
}

function populateDetailsModal(items, context, dataType) {
    const tbody = document.getElementById('itemDetailsTableBody');
    const noDetailsMessage = document.getElementById('noDetailsMessage');
    tbody.innerHTML = '';
    displayedItemsInModal = items;
    currentModalDataType = dataType;

    let headers = [];
    if (dataType === 'items') {
        headers = ['#', translations.table_subsystem, translations.table_discipline, translations.activitiesModal_tagNo, translations.table_type, translations.table_description, translations.table_status];
    } else if (dataType === 'punch') {
        headers = ['#', translations.table_subsystem, translations.table_discipline, translations.activitiesModal_tagNo, translations.table_type, translations.table_category, translations.table_description, translations.table_pl_no];
    } else if (dataType === 'hold') {
        headers = ['#', translations.table_subsystem, translations.table_discipline, translations.activitiesModal_tagNo, translations.table_type, translations.table_hp_priority, translations.table_hp_description, translations.table_hp_location];
    }

    const theadRow = document.getElementById('itemDetailsModalHeader');
    theadRow.innerHTML = headers.map(h => `<th scope="col">${h}</th>`).join('');

    const filterRow = document.getElementById('modal-filter-row');
    filterRow.innerHTML = headers.map((h, i) => `<th><input type="text" class="form-control form-control-sm" placeholder="${translations.modal_filterPlaceholder}" data-col-index="${i}"></th>`).join('');

    if (items.length === 0) {
        noDetailsMessage.style.display = 'block';
        noDetailsMessage.textContent = translations.noMatchingItems;
    } else {
        noDetailsMessage.style.display = 'none';
        items.forEach((item, index) => {
            const row = document.createElement('tr');
            let rowContent = '';
            let rowClass = '';

            if (dataType === 'items') {
                rowContent = `
                    <td>${index + 1}</td>
                    <td>${item.subsystem}</td>
                    <td>${item.discipline}</td>
                    <td style="cursor: pointer; color: #007bff; text-decoration: underline;">${item.tagNo}</td>
                    <td>${item.typeCode}</td>
                    <td>${item.description}</td>
                    <td>${item.status}</td>
                `;
            } else if (dataType === 'punch') {
                const punchCat = item.PL_Punch_Category ? item.PL_Punch_Category.toLowerCase() : '';
                if (punchCat === 'a') rowClass = 'table-danger';
                else if (punchCat === 'b') rowClass = 'table-info';
                else if (punchCat === 'c') rowClass = 'table-success';
                rowContent = `
                    <td>${index + 1}</td>
                    <td>${item.SD_Sub_System || 'N/A'}</td>
                    <td>${item.Discipline_Name || 'N/A'}</td>
                    <td style="cursor: pointer; color: #007bff; text-decoration: underline;">${item.ITEM_Tag_NO || 'N/A'}</td>
                    <td>${item.ITEM_Type_Code || 'N/A'}</td>
                    <td style="${item.PL_Punch_Category === 'A' ? 'color: red; font-weight: bold;' : ''}">${item.PL_Punch_Category || 'N/A'}</td>
                    <td>${item.PL_Punch_Description || 'N/A'}</td>
                    <td>${item.PL_No || 'N/A'}</td>
                `;
            } else if (dataType === 'hold') {
                rowContent = `
                    <td>${index + 1}</td>
                    <td>${item.subsystem}</td>
                    <td>${item.discipline}</td>
                    <td style="cursor: pointer; color: #007bff; text-decoration: underline;">${item.tagNo}</td>
                    <td>${item.typeCode || 'N/A'}</td>
                    <td>${item.hpPriority || 'N/A'}</td>
                    <td>${item.hpDescription || 'N/A'}</td>
                    <td>${item.hpLocation || 'N/A'}</td>
                `;
            }
            row.innerHTML = rowContent;
            if (rowClass) row.classList.add(rowClass);
            tbody.appendChild(row);
        });
    }
}


function filterHOSItems(statusType, dataType) {
    fetch(GITHUB_BASE_URL + 'HOS.CSV')
        .then(response => response.text())
        .then(csvText => {
            Papa.parse(csvText, {
                header: true,
                skipEmptyLines: true,
                complete: (results) => {
                    let filteredData = [];
                    if (statusType === 'FORM_A') {
                        filteredData = results.data.filter(row => row.FormA && row.FormA.trim() !== '');
                    } else if (statusType === 'FORM_B') {
                        filteredData = results.data.filter(row => row.FormB && row.FormB.trim() !== '');
                    } else if (statusType === 'FORM_C') {
                        filteredData = results.data.filter(row => row.FormC && row.FormC.trim() !== '');
                    } else if (statusType === 'FORM_D') {
                        filteredData = results.data.filter(row => row.FormD && row.FormD.trim() !== '');
                    }

                    const modalData = filteredData.map(row => ({
                        subsystem: row.Sub_System,
                        subsystemName: row.Subsystem_Name,
                        formA: row.FormA || '',
                        formB: row.FormB || '',
                        formC: row.FormC || '',
                        formD: row.FormD || ''
                    }));
                    populateHOSDetailsModal(modalData, statusType, dataType);
                    itemDetailsModal.show();
                },
                error: (err) => console.error("PapaParse error for HOS CSV:", err)
            });
        })
        .catch(error => console.error("Error loading HOS CSV:", error));
}

function populateHOSDetailsModal(items, statusType, dataType) {
    const tbody = document.getElementById('itemDetailsTableBody');
    const noDetailsMessage = document.getElementById('noDetailsMessage');
    tbody.innerHTML = '';
    displayedItemsInModal = items;
    currentModalDataType = dataType;

    const theadRow = document.getElementById('itemDetailsModalHeader');
    theadRow.innerHTML = `
        <th scope="col">#</th>
        <th scope="col">${translations.table_subsystem}</th>
        <th scope="col">${translations.hos_subsystem_name}</th>
        <th scope="col">${translations.formA_title}</th>
        <th scope="col">${translations.formB_title}</th>
        <th scope="col">${translations.formC_title}</th>
        <th scope="col">${translations.formD_title}</th>
    `;

    document.getElementById('modal-filter-row').innerHTML = `<th colspan="7"></th>`; // No filters for this view

    if (items.length === 0) {
        noDetailsMessage.style.display = 'block';
        noDetailsMessage.textContent = translations.noMatchingItems;
    } else {
        noDetailsMessage.style.display = 'none';
        items.forEach((item, index) => {
            const row = document.createElement('tr');
            const formADate = item.formA ? new Date(item.formA).toLocaleDateString('fa-IR') : '';
            const formBDate = item.formB ? new Date(item.formB).toLocaleDateString('fa-IR') : '';
            const formCDate = item.formC ? new Date(item.formC).toLocaleDateString('fa-IR') : '';
            const formDDate = item.formD ? new Date(item.formD).toLocaleDateString('fa-IR') : '';
            row.innerHTML = `
                <td>${index + 1}</td>
                <td>${item.subsystem}</td>
                <td>${item.subsystemName}</td>
                <td>${formADate}</td>
                <td>${formBDate}</td>
                <td>${formCDate}</td>
                <td>${formDDate}</td>
            `;
            tbody.appendChild(row);
        });
    }
    document.getElementById('itemDetailsModalLabel').textContent = `${translations.formDetailsModal_title}: ${statusType.replace('_', ' ')}`;
}

async function loadAndProcessData() {
    const loadingModal = new bootstrap.Modal(document.getElementById('loadingModal'), {});
    loadingModal.show();
    setTimeout(() => { loadingModal.hide(); }, 1500);

    DOMElements.errorMessage.style.display = 'none';
    try {
        const [hosResponse, mainResponse, itemsResponse, punchResponse, holdResponse, activitiesResponse] = await Promise.all([
            fetch(GITHUB_BASE_URL + 'HOS.CSV'), fetch(CSV_URL), fetch(ITEMS_CSV_URL),
            fetch(PUNCH_CSV_URL), fetch(HOLD_POINT_CSV_URL), fetch(ACTIVITIES_CSV_URL)
        ]);

        if (!hosResponse.ok) throw new Error(`HOS.CSV: ${hosResponse.statusText}`);
        Papa.parse(await hosResponse.text(), {
            header: true, skipEmptyLines: true,
            complete: r => {
                window.formCounts.formA = r.data.filter(row => row.FormA?.trim()).length;
                window.formCounts.formB = r.data.filter(row => row.FormB?.trim()).length;
                window.formCounts.formC = r.data.filter(row => row.FormC?.trim()).length;
                window.formCounts.formD = r.data.filter(row => row.FormD?.trim()).length;
            }
        });

        if (!mainResponse.ok) throw new Error(`DATA.CSV: ${mainResponse.statusText}`);
        Papa.parse(await mainResponse.text(), {
            header: true, skipEmptyLines: true,
            complete: results => {
                const { data } = results;
                const systemMap = {};
                const subSystemMap = {};
                data.forEach(row => {
                    if (!row.SD_System || !row.SD_Sub_System || !row.discipline) return;
                    const systemId = row.SD_System.trim();
                    const systemName = (row.SD_System_Name || 'Unknown System').trim();
                    const subId = row.SD_Sub_System.trim();
                    const subName = (row.SD_Subsystem_Name || 'Unknown Subsystem').trim();
                    const discipline = row.discipline.trim();
                    if (!systemMap[systemId]) {
                        systemMap[systemId] = { id: systemId, name: systemName, subs: [] };
                    }
                    if (!systemMap[systemId].subs.find(s => s.id === subId)) {
                        systemMap[systemId].subs.push({ id: subId, name: subName });
                    }
                    if (!subSystemMap[subId]) {
                        subSystemMap[subId] = { id: subId, name: subName, systemId, title: `${subId} - ${subName}`, disciplines: {} };
                    }
                    const total = parseInt(row["TOTAL ITEM"]) || 0;
                    const done = parseInt(row["TOTAL DONE"]) || 0;
                    const pending = parseInt(row["TOTAL PENDING"]) || 0;
                    subSystemMap[subId].disciplines[discipline] = {
                        total, done, pending,
                        punch: parseInt(row["TOTAL NOT CLEAR PUNCH"]) || 0,
                        hold: parseInt(row["TOTAL HOLD POINT"]) || 0,
                        remaining: Math.max(0, total - done - pending)
                    };
                });
                processedData = { systemMap, subSystemMap, allRawData: data };
                updateView();
            }
        });

        if (!itemsResponse.ok) throw new Error(`ITEMS.CSV: ${itemsResponse.statusText}`);
        Papa.parse(await itemsResponse.text(), { header: true, skipEmptyLines: true, complete: r => { detailedItemsData = r.data.map(i => ({...i})); } });

        if (!punchResponse.ok) throw new Error(`PUNCH.CSV: ${punchResponse.statusText}`);
        Papa.parse(await punchResponse.text(), { header: true, skipEmptyLines: true, complete: r => { punchItemsData = r.data.map(i => ({...i})); } });

        if (!holdResponse.ok) throw new Error(`HOLD_POINT.CSV: ${holdResponse.statusText}`);
        Papa.parse(await holdResponse.text(), { header: true, skipEmptyLines: true, complete: r => { holdPointItemsData = r.data.map(i => ({...i})); } });

        if (!activitiesResponse.ok) throw new Error(`ACTIVITES.CSV: ${activitiesResponse.statusText}`);
        Papa.parse(await activitiesResponse.text(), { header: true, skipEmptyLines: true, complete: r => { activitiesData = r.data.map(i => ({...i})); } });

    } catch (e) {
        DOMElements.errorMessage.textContent = `${translations.dataLoadError}: ${e.message}`;
        DOMElements.errorMessage.style.display = 'block';
        console.error("Fetch error:", e);
    }
}

// --- Rendering Functions ---
function updateView() {
    aggregatedStats = _aggregateStatsForView(selectedView, processedData.systemMap, processedData.subSystemMap);
    let titleText = translations.dashboard;
    if (selectedView.type === 'system' && selectedView.id) {
        const systemName = processedData.systemMap[selectedView.id]?.name || selectedView.name;
        titleText = `${translations.system}: ${selectedView.id} - ${systemName}`;
    } else if (selectedView.type === 'subsystem' && selectedView.id) {
        const systemName = processedData.systemMap[selectedView.parentId]?.name || selectedView.parentId;
        const subsystemName = processedData.subSystemMap[selectedView.id]?.name || selectedView.name;
        titleText = `${translations.system}: ${selectedView.parentId} - ${systemName}<br>${translations.subsystem}: ${selectedView.id} - ${subsystemName}`;
    }
    DOMElements.dashboardTitle.innerHTML = titleText;
    DOMElements.totalItemsCounter.textContent = aggregatedStats.totalItems.toLocaleString('fa-IR');

    renderDashboardHeader();
    renderProcessStepper();
    renderCharts();
    renderDataTable();
    renderSidebar();
}

function renderDashboardHeader() {
    const stats = [
        { label: translations.completed, count: aggregatedStats.done, color: 'text-success', status: 'DONE' },
        { label: translations.pending, count: aggregatedStats.pending, color: 'text-warning', status: 'PENDING' },
        { label: translations.remaining, count: aggregatedStats.remaining, color: 'text-info', status: 'OTHER' },
        { label: translations.punch, count: aggregatedStats.punch, color: 'text-danger', status: 'PUNCH' },
        { label: translations.holdPoint, count: aggregatedStats.hold, color: 'text-danger', status: 'HOLD' }
    ];

    const headerHTML = `
        <div class="dashboard-header">
            ${stats.map(stat => `
                <div class="stat-item" data-status="${stat.status}" style="cursor: pointer;">
                    <h4>${stat.label}</h4>
                    <p class="count-display ${stat.color}">${stat.count.toLocaleString('fa-IR')}</p>
                </div>
            `).join('')}
        </div>
    `;
    DOMElements.dashboardHeaderContainer.innerHTML = headerHTML;
}

function renderProcessStepper() {
    const steps = [
        { title: translations.formA_title, count: window.formCounts.formA, id: 'form-a', status: 'FORM_A', icon: ICONS.FileEarmarkText },
        { title: translations.formB_title, count: window.formCounts.formB, id: 'form-b', status: 'FORM_B', icon: ICONS.FileEarmarkCheck },
        { title: translations.formC_title, count: window.formCounts.formC, id: 'form-c', status: 'FORM_C', icon: ICONS.FileEarmarkMedical },
        { title: translations.formD_title, count: window.formCounts.formD, id: 'form-d', status: 'FORM_D', icon: ICONS.FileEarmarkSpreadsheet }
    ];

    const stepperHTML = `
        <div class="process-stepper">
            ${steps.map(step => `
                <div class="step ${step.id}" data-status="${step.status}" style="cursor: pointer;">
                    <div class="step-icon">${step.icon}</div>
                    <div class="step-title">${step.title}</div>
                    <div class="step-count">${step.count.toLocaleString('fa-IR')}</div>
                </div>
            `).join('')}
        </div>
    `;
    DOMElements.processStepperContainer.innerHTML = stepperHTML;
}

function renderSidebar() {
    let html = '';
    const createNodeHTML = (node, level = 0, parentId = null) => {
        const isSelected = selectedView.type === node.type && selectedView.id === node.id;
        const hasChildren = node.children && node.children.length > 0;
        let childrenHTML = '';
        let isOpen = node.isOpen || false;
        if (searchTerm && node.children?.some(child => child.name.toLowerCase().includes(searchTerm))) {
            isOpen = true;
        }
        if (hasChildren && isOpen) {
            childrenHTML = `<div class="tree-children" role="group" style="display: block;">${node.children.map(child => createNodeHTML(child, level + 1, node.id)).join('')}</div>`;
        } else if (hasChildren) {
            childrenHTML = `<div class="tree-children" role="group" style="display: none;">${node.children.map(child => createNodeHTML(child, level + 1, node.id)).join('')}</div>`;
        }
        const paddingRight = level * 12 + 12; // Changed for RTL
        const nodeId = `tree-node-${node.type}-${String(node.id).replace(/[^a-zA-Z0-9-_]/g, '')}`;
        let subtitle = '';
        if (node.type === 'system' && processedData.systemMap[node.id]) {
            subtitle = `<div class='small' style='font-size:0.78em; color: #ced4da !important;'>${processedData.systemMap[node.id].name}</div>`;
        } else if (node.type === 'subsystem' && processedData.subSystemMap[node.id]) {
            subtitle = `<div class='small' style='font-size:0.78em; color: #ced4da !important;'>${processedData.subSystemMap[node.id].name}</div>`;
        }
        return `
            <div id="${nodeId}" class="tree-node ${isSelected ? 'selected' : ''} ${isOpen ? 'open' : ''}"
                 role="treeitem" aria-selected="${isSelected}" ${hasChildren ? `aria-expanded="${isOpen}"` : ''}
                 data-type="${node.type}" data-id="${node.id}" data-name="${node.name}"
                 data-parent-id="${parentId || ''}" style="padding-right: ${paddingRight}px;" tabindex="0">
                ${node.icon || ''}
                <span class="flex-grow-1 text-truncate ms-2">${node.name}${subtitle}</span>
                ${hasChildren ? ICONS.ChevronRight : ''}
            </div>
            ${childrenHTML}
        `;
    };

    const treeNodes = [
        { id: 'all', name: translations.allSystems, type: 'all', icon: ICONS.Collection, isOpen: true }
    ];
    Object.values(processedData.systemMap).forEach(system => {
        treeNodes.push({
            id: system.id, name: system.id, type: 'system', icon: ICONS.Folder,
            children: system.subs.map(sub => ({
                id: sub.id, name: sub.id, type: 'subsystem', icon: ICONS.Puzzle, parentId: system.id
            })),
            isOpen: selectedView.id === system.id || selectedView.parentId === system.id
        });
    });

    const filterNodes = (nodes) => {
        if (!searchTerm) return nodes;
        return nodes.map(node => {
            const isMatch = node.name.toLowerCase().includes(searchTerm);
            const filteredChildren = node.children ? filterNodes(node.children) : null;
            if (isMatch || (filteredChildren && filteredChildren.length > 0)) {
                return { ...node, children: filteredChildren, isOpen: true };
            }
            return null;
        }).filter(Boolean);
    };

    const finalTreeNodes = filterNodes(treeNodes);
    html = finalTreeNodes.map(node => createNodeHTML(node)).join('');
    if (finalTreeNodes.length === 0 && searchTerm) {
        html = `<p class="text-muted text-center small p-3">${translations.noMatchingItems}</p>`;
    }

    DOMElements.treeView.innerHTML = `<div role="tree" aria-label="System and Subsystem Navigation">${html}</div>`;
    attachSidebarEventListeners();
}

function attachSidebarEventListeners() {
    DOMElements.treeView.querySelectorAll('.tree-node').forEach(el => {
        el.addEventListener('click', function(e) {
            e.stopPropagation();
            const type = this.dataset.type;
            const id = this.dataset.id;
            const name = this.dataset.name;
            const parentId = this.dataset.parentId;
            const targetIsChevron = e.target.classList.contains('chevron-toggle') || e.target.closest('.chevron-toggle');
            if (targetIsChevron && this.hasAttribute('aria-expanded')) {
                const isOpen = this.classList.toggle('open');
                this.setAttribute('aria-expanded', isOpen);
                const childrenContainer = this.nextElementSibling;
                if (childrenContainer) childrenContainer.style.display = isOpen ? 'block' : 'none';
            } else {
                handleNodeSelect(type, id, name, parentId);
                if (window.innerWidth < 992) {
                    DOMElements.sidebar.classList.remove('open');
                    DOMElements.mainContent.classList.remove('sidebar-open');
                    DOMElements.sidebarOverlay.style.display = 'none';
                    DOMElements.sidebarToggle.setAttribute('aria-expanded', 'false');
                }
            }
        });
    });
}

function handleNodeSelect(type, id, name, parentId = null) {
    selectedView = { type, id, name, parentId };
    updateView();
}

function renderCharts() {
    destroyChart(chartInstances.overview);
    Object.values(chartInstances.disciplines).forEach(destroyChart);
    chartInstances.disciplines = {};
    Object.values(chartInstances.systems).forEach(destroyChart);
    chartInstances.systems = {};

    const activeTabPane = document.querySelector(`.tab-pane.active[role="tabpanel"]`);
    if (activeTabPane) {
        if (activeTabPane.id === 'overviewChartsContainer') renderOverviewCharts();
        else if (activeTabPane.id === 'disciplineChartsContainer') renderDisciplineCharts();
        else if (activeTabPane.id === 'systemChartsContainer') renderSystemSubsystemCharts();
    }
}

function renderOverviewCharts() {
    const overviewCanvas = document.getElementById('overviewChart');
    if (!overviewCanvas) return;
    const overviewParent = overviewCanvas.parentElement;
    overviewParent.innerHTML = '<canvas id="overviewChart" role="img" aria-label="General status doughnut chart"></canvas>';
    const overviewCtx = document.getElementById('overviewChart').getContext('2d');
    const overviewChartData = {
        labels: [translations.completed, translations.pending, translations.remaining],
        datasets: [{
            label: translations.generalStatus,
            data: [aggregatedStats.done, aggregatedStats.pending, aggregatedStats.remaining].filter(v => v >= 0),
            backgroundColor: [COLORS_STATUS_CHARTJS.done, COLORS_STATUS_CHARTJS.pending, COLORS_STATUS_CHARTJS.remaining],
            hoverOffset: 4,
            borderColor: '#F4F5F7',
            borderWidth: 2
        }]
    };
    if (aggregatedStats.totalItems === 0) {
        overviewParent.insertAdjacentHTML('beforeend', `<div class="text-center text-muted small p-5">${translations.noDataToDisplay}</div>`);
    } else {
        chartInstances.overview = new Chart(overviewCtx, { type: 'doughnut', data: overviewChartData, options: { responsive: true, maintainAspectRatio: false, cutout: '70%', plugins: { legend: { position: 'bottom', labels: {font: { family: "'Vazirmatn', sans-serif"}}}, tooltip: { callbacks: { label: (context) => `${context.label}: ${context.formattedValue} (${Math.round(context.parsed / aggregatedStats.totalItems * 100)}%)`}}}} });
    }
}

function renderDisciplineCharts() {
    const container = DOMElements.disciplineChartsContainer;
    container.innerHTML = '';
    if (selectedView.type !== 'subsystem' || !selectedView.id) {
        container.innerHTML = `<div class="col-12 text-center py-5 text-muted" role="status">${ICONS.PieChartIcon}<p class="mt-2">${translations.selectSubsystemToViewDetails}</p></div>`;
        return;
    }
    const subSystem = processedData.subSystemMap[selectedView.id];
    if (!subSystem || Object.keys(subSystem.disciplines).length === 0) {
        container.innerHTML = `<div class="col-12 text-center py-5 text-muted" role="status">${ICONS.PieChartIcon}<p class="mt-2">${translations.noDisciplineData}</p></div>`;
        return;
    }
    const row = document.createElement('div');
    row.className = 'row g-3';
    Object.entries(subSystem.disciplines).forEach(([name, data]) => {
        const col = document.createElement('div');
        col.className = 'col-12 col-md-6 col-lg-4 col-xl-3';
        const chartId = `disciplineChart-${name.replace(/\s+/g, '-')}`;
        col.innerHTML = `
            <div class="card h-100 shadow-sm">
                <div class="card-body text-center">
                    <h6 class="text-muted small fw-medium mb-1">${name}</h6>
                    <p class="text-muted small mb-2">${data.total.toLocaleString('fa-IR')} ${translations.items}</p>
                    <div class="chart-container" style="height: 200px;"><canvas id="${chartId}"></canvas></div>
                </div>
            </div>`;
        row.appendChild(col);
        if (data.total > 0) {
            setTimeout(() => {
                const ctx = document.getElementById(chartId).getContext('2d');
                const chartData = {
                    labels: [translations.completed, translations.pending, translations.remaining],
                    datasets: [{ label: name, data: [data.done, data.pending, data.remaining], backgroundColor: [COLORS_STATUS_CHARTJS.done, COLORS_STATUS_CHARTJS.pending, COLORS_STATUS_CHARTJS.remaining] }]
                };
                chartInstances.disciplines[name] = new Chart(ctx, { type: 'doughnut', data: chartData, options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: true, position: 'bottom', labels: { boxWidth: 10, font: { size: 10, family: "'Vazirmatn', sans-serif" } } }, tooltip: { callbacks: { label: (context) => `${context.label}: ${context.formattedValue} (${Math.round(context.parsed / data.total * 100)}%)` } } } } });
            }, 0);
        } else {
            setTimeout(() => { document.getElementById(chartId).parentElement.innerHTML = `<div class="text-center text-muted small p-5" style="height:100%; display:flex; align-items:center; justify-content:center;">${translations.noDataToDisplay}</div>`; }, 0);
        }
    });
    container.appendChild(row);
}

function renderSystemSubsystemCharts() {
    const container = DOMElements.systemChartsContainer;
    container.innerHTML = '';
    let itemsToDisplay = [];
    if (selectedView.type === 'all') {
        itemsToDisplay = Object.values(processedData.systemMap).map(system => ({
            id: system.id, name: `${system.id} - ${system.name}`, data: _aggregateStatsForSystem(system.id, processedData.systemMap, processedData.subSystemMap)
        }));
    } else if (selectedView.type === 'system' && selectedView.id) {
        const system = processedData.systemMap[selectedView.id];
        if (system) {
            itemsToDisplay = system.subs.map(subRef => {
                const subSystem = processedData.subSystemMap[subRef.id];
                return { id: subRef.id, name: `${subRef.id} - ${subSystem?.name || 'N/A'}`, data: _aggregateStatsForSubSystem(subRef.id, processedData.subSystemMap) };
            });
        }
    } else if (selectedView.type === 'subsystem' && selectedView.id) {
        const subSystem = processedData.subSystemMap[selectedView.id];
        if (subSystem) {
            itemsToDisplay = [{ id: subSystem.id, name: `${subSystem.id} - ${subSystem.name}`, data: _aggregateStatsForSubSystem(subSystem.id, processedData.subSystemMap) }];
        }
    }
    itemsToDisplay = itemsToDisplay.filter(item => item.data.totalItems > 0);
    if (itemsToDisplay.length === 0) {
        container.innerHTML = `<div class="col-12 text-center py-5 text-muted" role="status">${ICONS.PieChartIcon}<p class="mt-2">${translations.noSystemData}</p></div>`;
        return;
    }
    const row = document.createElement('div');
    row.className = 'row g-3';
    itemsToDisplay.forEach(item => {
        item.data.remaining = Math.max(0, item.data.totalItems - item.data.done - item.data.pending);
        const col = document.createElement('div');
        col.className = 'col-12 col-md-6 col-lg-4 col-xl-3';
        const chartId = `systemSubChart-${item.id.replace(/\s+/g, '-|')}`;
        col.innerHTML = `
            <div class="card h-100 shadow-sm">
                <div class="card-body text-center">
                    <h6 class="text-muted small fw-medium mb-1 text-truncate" title="${item.name}">${item.name}</h6>
                    <p class="text-muted small mb-2">${item.data.totalItems.toLocaleString('fa-IR')} ${translations.items}</p>
                    <div class="chart-container" style="height: 200px;"><canvas id="${chartId}"></canvas></div>
                </div>
            </div>`;
        row.appendChild(col);
        if (item.data.totalItems > 0) {
            setTimeout(() => {
                const ctx = document.getElementById(chartId).getContext('2d');
                const chartData = {
                    labels: [translations.completed, translations.pending, translations.remaining],
                    datasets: [{ label: item.name, data: [item.data.done, item.data.pending, item.data.remaining], backgroundColor: [COLORS_STATUS_CHARTJS.done, COLORS_STATUS_CHARTJS.pending, COLORS_STATUS_CHARTJS.remaining] }]
                };
                chartInstances.systems[item.id] = new Chart(ctx, { type: 'doughnut', data: chartData, options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: true, position: 'bottom', labels: { boxWidth: 10, font: { size: 10, family: "'Vazirmatn', sans-serif" } } }, tooltip: { callbacks: { label: (context) => `${context.label}: ${context.formattedValue} (${Math.round(context.parsed / item.data.totalItems * 100)}%)` } } } } });
            }, 0);
        } else {
            setTimeout(() => { document.getElementById(chartId).parentElement.innerHTML = `<div class="text-center text-muted small p-5" style="height:100%; display:flex; align-items:center; justify-content:center;">${translations.noDataToDisplay}</div>`; }, 0);
        }
    });
    container.appendChild(row);
}

function renderDataTable() {
    const columns = [
        { header: translations.table_system, accessor: 'system' }, { header: translations.table_subsystem, accessor: 'subsystem' },
        { header: translations.table_discipline, accessor: 'discipline' }, { header: translations.table_totalItems, accessor: 'totalItems' },
        { header: translations.table_completed, accessor: 'completed' }, { header: translations.table_pending, accessor: 'pending' },
        { header: translations.table_punch, accessor: 'punch' }, { header: translations.table_holdPoint, accessor: 'holdPoint' },
        { header: translations.table_status, accessor: 'statusPercent' },
    ];
    DOMElements.dataTableHead.innerHTML = columns.map(col => `<th scope="col">${col.header}</th>`).join('');
    const tableData = _generateTableDataForView(selectedView, processedData, aggregatedStats.totalItems === 0);
    let bodyHTML = '';
    if (tableData.length === 0) {
        bodyHTML = `<tr><td colspan="${columns.length}" class="text-center py-5 text-muted">${translations.selectSubsystemToViewDetails}</td></tr>`;
    } else {
        tableData.forEach(row => {
            bodyHTML += '<tr>';
            columns.forEach((col, index) => {
                let cellValue = row[col.accessor];
                if (col.accessor === 'statusPercent') {
                    const badgeClass = row.statusPercent > 80 ? 'bg-success-subtle text-success' : row.statusPercent > 50 ? 'bg-info-subtle text-info' : 'bg-warning-subtle text-warning';
                    cellValue = `<span class="badge ${badgeClass} rounded-pill">${row.statusPercent.toLocaleString('fa-IR')}%</span>`;
                } else if (col.accessor === 'subsystem') {
                    cellValue = `${row.subsystem} - ${row.subsystemName}`;
                } else {
                    cellValue = (typeof cellValue === 'number') ? cellValue.toLocaleString('fa-IR') : cellValue;
                }
                const cellTag = index === 0 ? `<th scope="row">${cellValue}</th>` : `<td>${cellValue}</td>`;
                bodyHTML += cellTag;
            });
            bodyHTML += '</tr>';
        });
    }
    DOMElements.dataTableBody.innerHTML = bodyHTML;
}

const _emptyStats = () => ({ totalItems: 0, done: 0, pending: 0, punch: 0, hold: 0, remaining: 0 });

function _aggregateStatsForSubSystem(subSystemId, subSystemMap) {
    const subSystem = subSystemMap[subSystemId];
    if (!subSystem) return _emptyStats();
    return Object.values(subSystem.disciplines).reduce((acc, discipline) => {
        acc.totalItems += discipline.total;
        acc.done += discipline.done;
        acc.pending += discipline.pending;
        acc.punch += discipline.punch;
        acc.hold += discipline.hold;
        return acc;
    }, _emptyStats());
}

function _aggregateStatsForSystem(systemId, systemMap, subSystemMap) {
    const system = systemMap[systemId];
    if (!system) return _emptyStats();
    return system.subs.reduce((acc, subRef) => {
        const subSystemStats = _aggregateStatsForSubSystem(subRef.id, subSystemMap);
        Object.keys(subSystemStats).forEach(key => acc[key] += subSystemStats[key]);
        return acc;
    }, _emptyStats());
}

function _aggregateStatsForAll(systemMap, subSystemMap) {
    return Object.keys(systemMap).reduce((acc, systemId) => {
        const systemStats = _aggregateStatsForSystem(systemId, systemMap, subSystemMap);
        Object.keys(systemStats).forEach(key => acc[key] += systemStats[key]);
        return acc;
    }, _emptyStats());
}

function _aggregateStatsForView(view, systemMap, subSystemMap) {
    let stats;
    if (view.type === 'all' || !view.id) stats = _aggregateStatsForAll(systemMap, subSystemMap);
    else if (view.type === 'system') stats = _aggregateStatsForSystem(view.id, systemMap, subSystemMap);
    else stats = _aggregateStatsForSubSystem(view.id, subSystemMap);
    stats.remaining = Math.max(0, stats.totalItems - stats.done - stats.pending);
    return stats;
}

function _generateTableDataForView(view, pData, isEmptyView, forExport = false) {
    const { systemMap, subSystemMap, allRawData } = pData;
    if (!forExport && isEmptyView && view.type !== 'all') return [];
    let relevantRawData = [];
    if (view.type === 'all') relevantRawData = allRawData;
    else if (view.type === 'system' && view.id) {
        const system = systemMap[view.id];
        if (system) {
            const subIdsInSystem = new Set(system.subs.map(s => s.id));
            relevantRawData = allRawData.filter(row => subIdsInSystem.has(row.SD_Sub_System.trim()));
        }
    } else if (view.type === 'subsystem' && view.id) {
        relevantRawData = allRawData.filter(row => row.SD_Sub_System.trim() === view.id);
    }
    if (!forExport && relevantRawData.length === 0 && view.type !== 'all') return [];
    return relevantRawData.map(row => {
        const totalItems = parseInt(row["TOTAL ITEM"]) || 0;
        const completed = parseInt(row["TOTAL DONE"]) || 0;
        return {
            system: row.SD_System.trim(), systemName: (row.SD_System_Name || 'N/A').trim(),
            subsystem: row.SD_Sub_System.trim(), subsystemName: (row.SD_Subsystem_Name || 'N/A').trim(),
            discipline: row.discipline.trim(), totalItems, completed,
            pending: parseInt(row["TOTAL PENDING"]) || 0,
            punch: parseInt(row["TOTAL NOT CLEAR PUNCH"]) || 0,
            holdPoint: parseInt(row["TOTAL HOLD POINT"]) || 0,
            statusPercent: totalItems > 0 ? Math.round((completed / totalItems) * 100) : 0,
        };
    });
}

function handleExport() {
    if (!processedData.allRawData || processedData.allRawData.length === 0) {
        Swal.fire({ title: 'توجه', text: translations.exportNoData, icon: 'warning', confirmButtonText: 'باشه' });
        return;
    }
    const dataToExportRaw = _generateTableDataForView(selectedView, processedData, false, true);
    if (dataToExportRaw.length === 0) {
        Swal.fire({ title: 'توجه', text: translations.exportNoData, icon: 'warning', confirmButtonText: 'باشه' });
        return;
    }
    const dataToExport = dataToExportRaw.map(row => ({
        [translations.table_system]: row.system, 'نام سیستم': row.systemName,
        [translations.table_subsystem]: row.subsystem, 'نام زیرسیستم': row.subsystemName,
        [translations.table_discipline]: row.discipline, [translations.table_totalItems]: row.totalItems,
        [translations.table_completed]: row.completed, [translations.table_pending]: row.pending,
        [translations.table_punch]: row.punch, [translations.table_holdPoint]: row.holdPoint,
        'درصد پیشرفت': `${row.statusPercent}%`
    }));
    const currentDate = new Date().toLocaleDateString('fa-IR').replace(/\//g, '-');
    let viewName = translations.allSystems;
    if (selectedView.type === 'system' && selectedView.id) viewName = `System_${selectedView.id.replace(/[^a-zA-Z0-9]/g, '_')}`;
    else if (selectedView.type === 'subsystem' && selectedView.id) viewName = `SubSystem_${selectedView.id.replace(/[^a-zA-Z0-9]/g, '_')}`;
    const fileName = `گزارش_سپرا_${viewName}_${currentDate}.xlsx`;
    try {
        const worksheet = XLSX.utils.json_to_sheet(dataToExport);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'گزارش سپرا');
        XLSX.writeFile(workbook, fileName);
    } catch (error) {
        console.error("Error exporting to Excel:", error);
        Swal.fire({ title: 'خطا', text: translations.exportError, icon: 'error', confirmButtonText: 'باشه' });
    }
}

async function handleDownloadAll() {
    if (typeof JSZip === 'undefined') {
        Swal.fire({ title: 'خطا', text: 'کتابخانه مورد نیاز (JSZip) بارگذاری نشده است.', icon: 'error', confirmButtonText: 'باشه' });
        return;
    }
    const loadingModal = new bootstrap.Modal(document.getElementById('loadingModal'), {});
    const loadingModalLabel = document.getElementById('loadingModalLabel');
    const originalLabel = loadingModalLabel.textContent;
    try {
        loadingModalLabel.textContent = `${translations.downloadingFiles} (0/7)...`;
        loadingModal.show();
        const zip = new JSZip();
        const csvFiles = ['ACTIVITES.CSV', 'DATA.CSV', 'HOLD_POINT.CSV', 'HOS.CSV', 'ITEMS.CSV', 'PUNCH.CSV', 'TRANS.CSV'];
        for (let i = 0; i < csvFiles.length; i++) {
            const file = csvFiles[i];
            loadingModalLabel.textContent = `${translations.downloadingFiles} (${i + 1}/7)...`;
            const response = await fetch(GITHUB_BASE_URL + file);
            if (!response.ok) throw new Error(`Failed to fetch ${file}.`);
            const content = await response.blob();
            zip.file(file, content);
        }
        loadingModalLabel.textContent = translations.creatingZip;
        const zipContent = await zip.generateAsync({ type: 'blob' });
        const currentDate = new Date().toLocaleDateString('fa-IR').replace(/\//g, '-');
        const fileName = `داده_های_سپرا_${currentDate}.zip`;
        const link = document.createElement('a');
        link.href = URL.createObjectURL(zipContent);
        link.download = fileName;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(link.href);
    } catch (error) {
        console.error("Download error:", error);
        Swal.fire({ title: 'خطا', text: `${translations.downloadAllError}: ${error.message}`, icon: 'error', confirmButtonText: 'باشه' });
    } finally {
        loadingModalLabel.textContent = originalLabel;
        loadingModal.hide();
    }
}

function handleDetailsExport() {
    if (displayedItemsInModal.length === 0) {
        Swal.fire({ title: 'توجه', text: translations.exportNoDataModal, icon: 'warning', confirmButtonText: 'باشه' });
        return;
    }
    const currentDate = new Date().toLocaleDateString('fa-IR').replace(/\//g, '-');
    const modalTitle = document.getElementById('itemDetailsModalLabel').textContent.replace(/[^a-zA-Z0-9آ-ی ]/g, '').replace(/ /g, '_');
    let exportConfig;

    if (currentModalDataType === 'punch') {
        exportConfig = {
            fileName: `جزئیات_پانچ_${modalTitle || 'همه'}_${currentDate}.xlsx`,
            sheetName: 'Punch Items',
            headers: ['#', 'زیرسیستم', 'دیسیپلین', 'شماره تگ', 'نوع', 'دسته', 'شرح', 'شماره PL'],
            data: displayedItemsInModal.map((item, index) => [index + 1, item.SD_Sub_System, item.Discipline_Name, item.ITEM_Tag_NO, item.ITEM_Type_Code, item.PL_Punch_Category, item.PL_Punch_Description, item.PL_No])
        };
    } else if (currentModalDataType === 'items') {
        exportConfig = {
            fileName: `جزئیات_آیتم_${modalTitle || 'همه'}_${currentDate}.xlsx`,
            sheetName: 'Item Details',
            headers: ['#', 'زیرسیستم', 'دیسیپلین', 'شماره تگ', 'نوع', 'شرح', 'وضعیت'],
            data: displayedItemsInModal.map((item, index) => [index + 1, item.subsystem, item.discipline, item.tagNo, item.typeCode, item.description, item.status])
        };
    } else if (currentModalDataType === 'hold') {
        exportConfig = {
            fileName: `جزئیات_نقاط_توقف_${modalTitle || 'همه'}_${currentDate}.xlsx`,
            sheetName: 'Hold Points',
            headers: ['#', 'زیرسیستم', 'دیسیپلین', 'شماره تگ', 'نوع', 'اولویت', 'شرح', 'مکان'],
            data: displayedItemsInModal.map((item, index) => [index + 1, item.subsystem, item.discipline, item.tagNo, item.typeCode, item.hpPriority, item.hpDescription, item.hpLocation])
        };
    } else if (currentModalDataType && currentModalDataType.startsWith('form')) {
        exportConfig = {
            fileName: `جزئیات_فرم_${modalTitle || 'فرم'}_${currentDate}.xlsx`,
            sheetName: 'Form Details',
            headers: ['#', 'زیرسیستم', 'نام زیرسیستم', 'فرم A', 'فرم B', 'فرم C', 'فرم D'],
            data: displayedItemsInModal.map((item, index) => [index + 1, item.subsystem, item.subsystemName, item.formA, item.formB, item.formC, item.formD])
        };
    }

    try {
        const ws = XLSX.utils.aoa_to_sheet([exportConfig.headers, ...exportConfig.data]);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, exportConfig.sheetName);
        XLSX.writeFile(wb, exportConfig.fileName);
    } catch (error) {
        console.error("Export error:", error);
        Swal.fire({ title: 'خطا', text: `${translations.exportError}: ${error.message}`, icon: 'error', confirmButtonText: 'باشه' });
    }
}

function loadActivitiesForTag(tagNo) {
    document.getElementById('activitiesTagTitle').textContent = `${translations.activitiesModal_tagNo}: ${tagNo}`;
    const filtered = activitiesData.filter(a => a.Tag_No === tagNo);
    const list = document.getElementById('activitiesList');
    list.innerHTML = '';
    let doneCount = 0;

    if (filtered.length === 0) {
        list.innerHTML = `<tr><td colspan="3" class="text-center p-4">${translations.activitiesModal_noActivity}</td></tr>`;
        document.getElementById('activitiesProgressText').textContent = '0%';
        document.getElementById('activitiesProgressFill').style.width = '0%';
        return;
    }

    filtered.forEach((act, index) => {
        const tr = document.createElement('tr');
        const status = act.Done === '1' ? '✅' : '❌';
        const cls = act.Done === '1' ? 'text-success' : 'text-danger';
        tr.innerHTML = `
            <td class="text-center">${(index + 1).toLocaleString('fa-IR')}</td>
            <td>${act.Form_Title}</td>
            <td class="text-center"><span class="${cls}">${status}</span></td>
        `;
        list.appendChild(tr);
        if (act.Done === '1') doneCount++;
    });

    const percent = Math.round((doneCount / filtered.length) * 100);
    document.getElementById('activitiesProgressFill').style.width = `${percent}%`;
    document.getElementById('activitiesProgressText').textContent = `${percent.toLocaleString('fa-IR')}% (${doneCount.toLocaleString('fa-IR')}/${filtered.length.toLocaleString('fa-IR')})`;
}
/**
 * Sistema de Gestión de Órdenes de Servicio - Pago de Viajes a Tercerizados (CDN & CCN)
 * Contexto: Pago a transportistas/fleteros por viajes de entrega de equipo frío, botellón, fletes y distribución.
 */

// Ejemplos precargados de Clientes (ID numérico <-> Nombre de persona con 2 nombres y 2 apellidos)
const DEFAULT_CLIENTS = [
  { id: "100245", name: "Carlos Alberto Mendoza Morales" },
  { id: "100246", name: "María Fernanda López Rostrán" },
  { id: "100247", name: "Juan Carlos Gutiérrez Castillo" },
  { id: "100248", name: "Ana Patricia Blandón Somarriba" },
  { id: "100249", name: "José Antonio Martínez Zamora" },
  { id: "100250", name: "Rosa Elena Hernández Jarquín" },
  { id: "100251", name: "Luis Enrique Silva Obando" },
  { id: "100252", name: "Claudia Marcela Toruño Espinoza" },
  { id: "100253", name: "Francisco Javier Ruiz Baldelomar" },
  { id: "100254", name: "Sandra Iveth Morales Centeno" },
  { id: "100255", name: "Jorge Luis Pineda Chamorro" },
  { id: "100256", name: "Karla Vanessa Reyes Amador" }
];

// Ejemplos precargados de Motivos de Viaje vinculados a Número de Cuenta Contable
const DEFAULT_REASONS = [
  { reason: "Viaje por entrega de equipo frío", account: "CTA-70101-EQFRIO" },
  { reason: "Viaje por retiro de equipo frío para taller", account: "CTA-70102-RETFRIO" },
  { reason: "Viaje por entrega de botellón", account: "CTA-70201-BOTELLON" },
  { reason: "Viaje por recolección y retorno de botellón vacío", account: "CTA-70202-RECBOT" },
  { reason: "Viaje por traslado de producto entre bodegas", account: "CTA-70301-TRASLADO" },
  { reason: "Viaje por flete foráneo a sucursales", account: "CTA-70302-FORANEO" },
  { reason: "Viaje por distribución en canal tradicional", account: "CTA-70401-CANALTRAD" },
  { reason: "Viaje por entrega a canal moderno (Supermercados)", account: "CTA-70402-SUPERMKT" },
  { reason: "Viaje express por reposición de emergencia", account: "CTA-70501-EXPRESS" },
  { reason: "Viaje por canje y mantenimiento en ruta", account: "CTA-70601-MNTRUTA" }
];

// Lista de personas autorizadas para liberar órdenes de pago (Supervisores) con su Centro de Costo (solo números)
const DEFAULT_APPROVERS = [
  { name: "Carlos Buchitín", costCenter: "700720056" },
  { name: "Alicia Ramírez", costCenter: "700720057" },
  { name: "José Alejandro", costCenter: "700720058" },
  { name: "Adolfo Jarquín", costCenter: "700720059" },
  { name: "Mario Castillo", costCenter: "700720060" },
  { name: "Alex Flores", costCenter: "700720061" }
];

// Estado global de la aplicación
const AppState = {
  currentScreen: "portal", // "portal" | "main"
  selectedCompany: null,   // "CDN" | "CCN"
  orders: [],
  clients: [],
  reasons: [],
  approvers: [],
  theme: "light",
  filterCompany: "all",
  searchQuery: ""
};

// Referencias del DOM
const dom = {
  // Pantallas
  portalScreen: document.getElementById('portalScreen'),
  mainScreen: document.getElementById('mainScreen'),

  // Botones del Portal
  btnSelectCDN: document.getElementById('btnSelectCDN'),
  btnSelectCCN: document.getElementById('btnSelectCCN'),
  portalCdnCount: document.getElementById('portalCdnCount'),
  portalCcnCount: document.getElementById('portalCcnCount'),
  portalBtnToggleTheme: document.getElementById('portalBtnToggleTheme'),
  portalBtnOpenCatalog: document.getElementById('portalBtnOpenCatalog'),

  // Encabezado Pantalla Principal
  btnBackToPortal: document.getElementById('btnBackToPortal'),
  activeCompanyBadge: document.getElementById('activeCompanyBadge'),
  btnToggleTheme: document.getElementById('btnToggleTheme'),
  btnOpenCatalog: document.getElementById('btnOpenCatalog'),

  // Formulario y Campos
  orderForm: document.getElementById('orderForm'),
  clientIdSelect: document.getElementById('clientIdSelect'),
  clientNameSelect: document.getElementById('clientNameSelect'),
  orderReasonSelect: document.getElementById('orderReasonSelect'),
  accountNumberSelect: document.getElementById('accountNumberSelect'),
  orderConcept: document.getElementById('orderConcept'),
  orderAmount: document.getElementById('orderAmount'),
  releasedBySelect: document.getElementById('releasedBySelect'),
  costCenterInput: document.getElementById('costCenterInput'),

  // Botones de acción
  btnSubmit: document.getElementById('btnSubmit'),
  btnGenerateExcel: document.getElementById('btnGenerateExcel'),
  btnClearForm: document.getElementById('btnClearForm'),
  btnClearAllOrders: document.getElementById('btnClearAllOrders'),

  // Estadísticas KPI
  statTotal: document.getElementById('statTotal'),
  statCdn: document.getElementById('statCdn'),
  statCcn: document.getElementById('statCcn'),
  statTotalAmount: document.getElementById('statTotalAmount'),
  tableCountBadge: document.getElementById('tableCountBadge'),

  // Tabla y Búsqueda
  ordersTableBody: document.getElementById('ordersTableBody'),
  emptyState: document.getElementById('emptyState'),
  tableSearch: document.getElementById('tableSearch'),
  filterChips: document.querySelectorAll('.filter-chip'),

  // Modal Catálogo
  catalogModal: document.getElementById('catalogModal'),
  btnCloseCatalog: document.getElementById('btnCloseCatalog'),
  btnSaveAndCloseCatalog: document.getElementById('btnSaveAndCloseCatalog'),
  catalogClientsTable: document.getElementById('catalogClientsTable'),
  catalogReasonsTable: document.getElementById('catalogReasonsTable'),
  catalogApproversTable: document.getElementById('catalogApproversTable'),
  newClientId: document.getElementById('newClientId'),
  newClientName: document.getElementById('newClientName'),
  btnAddClient: document.getElementById('btnAddClient'),
  newReason: document.getElementById('newReason'),
  newAccount: document.getElementById('newAccount'),
  btnAddReason: document.getElementById('btnAddReason'),
  newApproverName: document.getElementById('newApproverName'),
  newApproverCostCenter: document.getElementById('newApproverCostCenter'),
  btnAddApprover: document.getElementById('btnAddApprover'),
  catalogTabs: document.querySelectorAll('.tab-btn'),
  tabClientsContent: document.getElementById('tabClientsContent'),
  tabReasonsContent: document.getElementById('tabReasonsContent'),
  tabApproversContent: document.getElementById('tabApproversContent'),

  // Notificaciones
  toastContainer: document.getElementById('toastContainer')
};

// ==========================================
// INICIALIZACIÓN
// ==========================================
document.addEventListener('DOMContentLoaded', () => {
  loadData();
  setupEventListeners();
  populateDropdowns();
  renderStats();
  renderOrdersTable();
  updatePortalCounts();
});

// Cargar almacenamiento persistente
function loadData() {
  const savedClients = localStorage.getItem('cdn_ccn_clients_v4');
  AppState.clients = savedClients ? JSON.parse(savedClients) : [...DEFAULT_CLIENTS];

  const savedReasons = localStorage.getItem('cdn_ccn_reasons_v4');
  AppState.reasons = savedReasons ? JSON.parse(savedReasons) : [...DEFAULT_REASONS];

  const savedApprovers = localStorage.getItem('cdn_ccn_approvers_v5');
  if (savedApprovers) {
    try {
      const parsed = JSON.parse(savedApprovers);
      if (Array.isArray(parsed) && parsed.length > 0 && typeof parsed[0] === 'object' && parsed[0].name && parsed[0].costCenter) {
        AppState.approvers = parsed;
      } else {
        AppState.approvers = [...DEFAULT_APPROVERS];
      }
    } catch {
      AppState.approvers = [...DEFAULT_APPROVERS];
    }
  } else {
    AppState.approvers = [...DEFAULT_APPROVERS];
  }
  saveData('approvers');

  const savedOrders = localStorage.getItem('cdn_ccn_orders_v4');
  AppState.orders = savedOrders ? JSON.parse(savedOrders) : [];

  const savedTheme = localStorage.getItem('cdn_ccn_theme') || 'light';
  setTheme(savedTheme);
}

function saveData(key) {
  if (key === 'orders') {
    localStorage.setItem('cdn_ccn_orders_v4', JSON.stringify(AppState.orders));
  } else if (key === 'clients') {
    localStorage.setItem('cdn_ccn_clients_v4', JSON.stringify(AppState.clients));
  } else if (key === 'reasons') {
    localStorage.setItem('cdn_ccn_reasons_v4', JSON.stringify(AppState.reasons));
  } else if (key === 'approvers') {
    localStorage.setItem('cdn_ccn_approvers_v5', JSON.stringify(AppState.approvers));
  }
}

// ==========================================
// LLENADO DE LISTAS DESPLEGABLES (DROPDOWNS)
// ==========================================
function populateDropdowns() {
  // Guardar selecciones actuales si existen
  const currentClientId = dom.clientIdSelect.value;
  const currentClientName = dom.clientNameSelect.value;
  const currentReason = dom.orderReasonSelect.value;
  const currentApprover = dom.releasedBySelect.value;

  // 1. Dropdown: ID del Cliente (SOLO va el ID numérico con el que se identifica)
  dom.clientIdSelect.innerHTML = '<option value="" disabled selected>-- Seleccione ID del Cliente --</option>';
  AppState.clients.forEach(c => {
    const opt = document.createElement('option');
    opt.value = c.id;
    opt.textContent = c.id; // SOLO el número de identificación
    dom.clientIdSelect.appendChild(opt);
  });

  // 2. Dropdown: Nombre del Cliente (SOLO va los dos nombres y dos apellidos de la persona)
  dom.clientNameSelect.innerHTML = '<option value="" disabled selected>-- Seleccione Nombre del Cliente --</option>';
  AppState.clients.forEach(c => {
    const opt = document.createElement('option');
    opt.value = c.name;
    opt.textContent = c.name; // SOLO los dos nombres y dos apellidos
    dom.clientNameSelect.appendChild(opt);
  });

  // 3. Dropdown: Motivo de la Orden (Tipo de Viaje)
  dom.orderReasonSelect.innerHTML = '<option value="" disabled selected>-- Seleccione Motivo del Viaje --</option>';
  AppState.reasons.forEach(r => {
    const opt = document.createElement('option');
    opt.value = r.reason;
    opt.textContent = r.reason;
    dom.orderReasonSelect.appendChild(opt);
  });

  // 4. Dropdown: Número de Cuenta
  dom.accountNumberSelect.innerHTML = '<option value="" disabled selected>-- Seleccione Número de Cuenta --</option>';
  AppState.reasons.forEach(r => {
    const opt = document.createElement('option');
    opt.value = r.account;
    opt.textContent = `${r.account} (${r.reason})`;
    dom.accountNumberSelect.appendChild(opt);
  });

  // 5. Dropdown: Quién lo Libera (Supervisores solicitados: Carlos Buchitín, Alicia Ramírez y 4 más)
  dom.releasedBySelect.innerHTML = '<option value="" disabled selected>-- Seleccione Quién lo Libera --</option>';
  AppState.approvers.forEach(app => {
    const opt = document.createElement('option');
    opt.value = app.name;
    opt.textContent = app.name;
    dom.releasedBySelect.appendChild(opt);
  });

  // Restaurar selecciones si aún aplican
  if (currentClientId) dom.clientIdSelect.value = currentClientId;
  if (currentClientName) dom.clientNameSelect.value = currentClientName;
  if (currentReason) dom.orderReasonSelect.value = currentReason;
  if (currentApprover) {
    dom.releasedBySelect.value = currentApprover;
    const match = AppState.approvers.find(a => a.name === currentApprover);
    dom.costCenterInput.value = match ? match.costCenter : '';
  } else {
    dom.costCenterInput.value = '';
  }
}

// ==========================================
// EVENT LISTENERS & ENLACES BIDIRECCIONALES
// ==========================================
function setupEventListeners() {
  // Selección de Empresa en Pantalla Inicial
  dom.btnSelectCDN.addEventListener('click', () => enterDashboard('CDN'));
  dom.btnSelectCCN.addEventListener('click', () => enterDashboard('CCN'));

  // Volver a la Pantalla de Selección
  dom.btnBackToPortal.addEventListener('click', returnToPortal);

  // -------------------------------------------------------------
  // AMARRE 1: ID de Cliente (Número) <--> Nombre de Cliente (2 Nombres y 2 Apellidos)
  // -------------------------------------------------------------
  dom.clientIdSelect.addEventListener('change', () => {
    const selectedId = dom.clientIdSelect.value;
    const match = AppState.clients.find(c => c.id === selectedId);
    if (match) {
      dom.clientNameSelect.value = match.name;
      highlightField(dom.clientNameSelect);
    }
  });

  dom.clientNameSelect.addEventListener('change', () => {
    const selectedName = dom.clientNameSelect.value;
    const match = AppState.clients.find(c => c.name === selectedName);
    if (match) {
      dom.clientIdSelect.value = match.id;
      highlightField(dom.clientIdSelect);
    }
  });

  // -------------------------------------------------------------
  // AMARRE 2: Motivo de la Orden <--> Número de Cuenta (Listas desplegables)
  // -------------------------------------------------------------
  dom.orderReasonSelect.addEventListener('change', () => {
    const selectedReason = dom.orderReasonSelect.value;
    const match = AppState.reasons.find(r => r.reason === selectedReason);
    if (match) {
      dom.accountNumberSelect.value = match.account;
      highlightField(dom.accountNumberSelect);
    }
  });

  dom.accountNumberSelect.addEventListener('change', () => {
    const selectedAccount = dom.accountNumberSelect.value;
    const match = AppState.reasons.find(r => r.account === selectedAccount);
    if (match) {
      dom.orderReasonSelect.value = match.reason;
      highlightField(dom.orderReasonSelect);
    }
  });

  // -------------------------------------------------------------
  // AMARRE 3: Quién lo Libera -> Centro de Costo (Auto-completar sólo números)
  // -------------------------------------------------------------
  const syncCostCenter = () => {
    const selectedName = dom.releasedBySelect.value;
    const match = AppState.approvers.find(a => a.name === selectedName);
    if (match) {
      dom.costCenterInput.value = match.costCenter;
      highlightField(dom.costCenterInput);
    } else {
      dom.costCenterInput.value = '';
    }
  };
  dom.releasedBySelect.addEventListener('change', syncCostCenter);
  dom.releasedBySelect.addEventListener('input', syncCostCenter);

  // Envío del Formulario
  dom.orderForm.addEventListener('submit', handleFormSubmit);

  // Limpiar campos del formulario
  dom.btnClearForm.addEventListener('click', () => clearForm(false));

  // Generar Excel
  dom.btnGenerateExcel.addEventListener('click', generateExcel);

  // Vaciar registros
  dom.btnClearAllOrders.addEventListener('click', confirmClearAllOrders);

  // Buscador y Filtros
  dom.tableSearch.addEventListener('input', (e) => {
    AppState.searchQuery = e.target.value.toLowerCase();
    renderOrdersTable();
  });

  dom.filterChips.forEach(chip => {
    chip.addEventListener('click', () => {
      dom.filterChips.forEach(c => c.classList.remove('active'));
      chip.classList.add('active');
      AppState.filterCompany = chip.dataset.filter;
      renderOrdersTable();
    });
  });

  // Tema
  dom.btnToggleTheme.addEventListener('click', toggleTheme);
  dom.portalBtnToggleTheme.addEventListener('click', toggleTheme);

  // Modal de Catálogo
  dom.btnOpenCatalog.addEventListener('click', openCatalogModal);
  dom.portalBtnOpenCatalog.addEventListener('click', openCatalogModal);
  dom.btnCloseCatalog.addEventListener('click', closeCatalogModal);
  dom.btnSaveAndCloseCatalog.addEventListener('click', closeCatalogModal);
  dom.catalogModal.addEventListener('click', (e) => {
    if (e.target === dom.catalogModal) closeCatalogModal();
  });

  // Tabs del Catálogo
  dom.catalogTabs.forEach(tab => {
    tab.addEventListener('click', () => {
      dom.catalogTabs.forEach(t => t.classList.remove('active'));
      tab.classList.add('active');
      dom.tabClientsContent.style.display = tab.dataset.tab === 'clients' ? 'block' : 'none';
      dom.tabReasonsContent.style.display = tab.dataset.tab === 'reasons' ? 'block' : 'none';
      dom.tabApproversContent.style.display = tab.dataset.tab === 'approvers' ? 'block' : 'none';
    });
  });

  // Acciones en Catálogo
  dom.btnAddClient.addEventListener('click', addClientToCatalog);
  dom.btnAddReason.addEventListener('click', addReasonToCatalog);
  dom.btnAddApprover.addEventListener('click', addApproverToCatalog);
}

// Efecto visual al sincronizar campos
function highlightField(element) {
  element.style.borderColor = "var(--accent-green)";
  element.style.boxShadow = "0 0 10px rgba(16, 185, 129, 0.4)";
  setTimeout(() => {
    element.style.borderColor = "";
    element.style.boxShadow = "";
  }, 1000);
}

// ==========================================
// LOGOS VECTORIALES OFICIALES (Transparentes, 100% Nítidos y con Carga Inmediata)
const LOGO_SVGS = {
  CDN: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 320 180" width="70" height="38" class="header-logo-svg">
    <g transform="translate(10, 10)">
      <path d="M 125,50 C 122,25 105,10 82,10 C 55,10 38,32 38,62 C 38,92 55,114 82,114 C 105,114 122,99 125,74 L 98,74 C 95,84 89,91 80,91 C 67,91 59,79 59,62 C 59,45 67,33 80,33 C 89,33 95,40 98,50 Z" fill="#1d4ed8" />
      <path d="M 105,38 L 132,38 C 152,38 166,51 166,73 C 166,95 152,108 132,108 L 105,108 Z M 123,54 L 123,92 L 131,92 C 141,92 148,84 148,73 C 148,62 141,54 131,54 Z" fill="#1d4ed8" />
      <path d="M 172,58 L 187,58 L 187,68 C 192,60 200,56 211,56 C 226,56 235,66 235,82 L 235,108 L 217,108 L 217,85 C 217,76 212,71 204,71 C 196,71 190,76 190,85 L 190,108 L 172,108 Z" fill="#16a34a" />
      <path d="M 32,108 C 80,138 180,140 248,96 C 220,126 100,128 32,108 Z" fill="#16a34a" />
      <text x="150" y="142" text-anchor="middle" font-family="'Arial Black', 'Montserrat', sans-serif" font-size="15" font-weight="900" font-style="italic" letter-spacing="1.2" fill="#1d4ed8">COMPAÑÍA DISTRIBUIDORA</text>
      <text x="150" y="160" text-anchor="middle" font-family="'Arial Black', 'Montserrat', sans-serif" font-size="14" font-weight="900" font-style="italic" letter-spacing="1.8" fill="#1d4ed8">DE NICARAGUA S.A.</text>
    </g>
  </svg>`,
  CCN: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 320 180" width="70" height="38" class="header-logo-svg">
    <g transform="translate(10, 10)">
      <path d="M 175,34 C 172,14 158,2 136,2 C 114,2 99,17 97,37 L 118,37 C 120,27 126,20 136,20 C 146,20 152,26 154,34 Z" fill="#15803d" />
      <path d="M 148,68 C 145,46 130,32 108,32 C 84,32 68,52 68,80 C 68,108 84,128 108,128 C 130,128 145,114 148,92 L 126,92 C 123,103 117,110 108,110 C 96,110 89,97 89,80 C 89,63 96,50 108,50 C 117,50 123,57 126,68 Z" fill="#15803d" />
      <path d="M 144,48 L 165,48 L 195,96 L 195,48 L 214,48 L 214,124 L 193,124 L 163,76 L 163,124 L 144,124 Z" fill="#15803d" />
      <path d="M 210,105 C 225,85 248,50 282,18" stroke="#d97706" stroke-width="3.5" stroke-linecap="round" fill="none" />
      <path d="M 276,16 C 286,20 288,32 278,35 C 268,32 268,22 276,16 Z" fill="#d97706" />
      <path d="M 264,28 C 274,32 276,44 266,47 C 256,44 256,34 264,28 Z" fill="#d97706" />
      <path d="M 250,42 C 260,46 262,58 252,61 C 242,58 242,48 250,42 Z" fill="#d97706" />
      <path d="M 236,56 C 246,60 248,72 238,75 C 228,72 228,62 236,56 Z" fill="#d97706" />
      <path d="M 222,70 C 232,74 234,86 224,89 C 214,86 214,76 222,70 Z" fill="#d97706" />
      <path d="M 24,106 C 75,136 185,138 255,92 C 225,124 95,126 24,106 Z" fill="#d97706" />
      <text x="150" y="146" text-anchor="middle" font-family="'Arial Black', 'Montserrat', sans-serif" font-size="15" font-weight="900" letter-spacing="1.2" fill="#15803d">COMPAÑÍA CERVECERA</text>
      <text x="150" y="164" text-anchor="middle" font-family="'Arial Black', 'Montserrat', sans-serif" font-size="14" font-weight="900" letter-spacing="1.8" fill="#15803d">DE NICARAGUA S.A.</text>
    </g>
  </svg>`
};

// CONTROL DE PANTALLAS (TRANSICIONES)
// ==========================================
function enterDashboard(company) {
  AppState.selectedCompany = company;
  AppState.currentScreen = "main";

  const companyFullName = company === 'CDN'
    ? 'Compañía Distribuidora de Nicaragua S.A.'
    : 'Compañía Cervecera de Nicaragua S.A.';

  // Actualizar distintivo de la empresa activa y cabecera con línea corporativa
  const appHeader = document.querySelector('.app-header');
  if (appHeader) {
    appHeader.className = `app-header ${company.toLowerCase()}`;
  }

  // Recuadro prominente con fondo corporativo y letras blancas
  dom.activeCompanyBadge.className = `active-company-pill ${company.toLowerCase()}`;
  dom.activeCompanyBadge.innerHTML = `
    <div class="header-logo-badge">
      ${LOGO_SVGS[company]}
    </div>
    <div class="header-company-info">
      <span class="active-company-name">${companyFullName}</span>
    </div>
  `;

  // Mostrar pantalla principal y ocultar portal
  dom.portalScreen.style.display = 'none';
  dom.mainScreen.style.display = 'block';

  // Sincronizar filtro de tabla con la empresa seleccionada por defecto
  dom.filterChips.forEach(chip => {
    const isCompany = chip.dataset.filter === company;
    chip.classList.toggle('active', isCompany);
  });
  AppState.filterCompany = company;

  renderStats();
  renderOrdersTable();
  clearForm(false);

  showToast(`🏢 Módulo activo: <strong>${companyFullName}</strong>`, "info");
}

function returnToPortal() {
  AppState.currentScreen = "portal";
  dom.mainScreen.style.display = 'none';
  dom.portalScreen.style.display = 'flex';
  updatePortalCounts();
}

function updatePortalCounts() {
  const cdnCount = AppState.orders.filter(o => o.company === 'CDN').length;
  const ccnCount = AppState.orders.filter(o => o.company === 'CCN').length;
  dom.portalCdnCount.textContent = cdnCount;
  dom.portalCcnCount.textContent = ccnCount;
}

// ==========================================
// REGISTRO DE ÓRDENES
// ==========================================
function handleFormSubmit(e) {
  e.preventDefault();

  if (!AppState.selectedCompany) {
    showToast("⚠️ Empresa no seleccionada. Regrese al menú inicial.", "error");
    return;
  }

  const clientId = dom.clientIdSelect.value;
  const clientName = dom.clientNameSelect.value;
  const orderReason = dom.orderReasonSelect.value;
  const accountNumber = dom.accountNumberSelect.value;
  const orderConcept = dom.orderConcept.value.trim();
  const rawAmount = parseFloat(dom.orderAmount.value);
  const releasedBy = dom.releasedBySelect.value;

  if (!clientId || !clientName) {
    showToast("⚠️ Debe seleccionar el Cliente (ID y Nombre)", "warning");
    dom.clientIdSelect.focus();
    return;
  }

  if (!orderReason || !accountNumber) {
    showToast("⚠️ Debe seleccionar el Motivo del Viaje", "warning");
    dom.orderReasonSelect.focus();
    return;
  }

  if (!orderConcept) {
    showToast("⚠️ Escriba el Concepto o detalles del viaje", "warning");
    dom.orderConcept.focus();
    return;
  }

  if (isNaN(rawAmount) || rawAmount < 0) {
    showToast("⚠️ Ingrese un Monto de viaje válido", "warning");
    dom.orderAmount.focus();
    return;
  }

  if (!releasedBy) {
    showToast("⚠️ Seleccione el Supervisor que libera el viaje", "warning");
    dom.releasedBySelect.focus();
    return;
  }

  const costCenter = dom.costCenterInput.value.trim();
  if (!costCenter) {
    showToast("⚠️ No se ha determinado el Centro de Costo del supervisor", "warning");
    dom.releasedBySelect.focus();
    return;
  }

  const now = new Date();
  const newOrder = {
    id: `ORD-${Date.now().toString().slice(-6)}`,
    company: AppState.selectedCompany,
    clientId,
    clientName,
    orderReason,
    accountNumber,
    orderConcept,
    amount: rawAmount,
    formattedAmount: `C$ ${rawAmount.toLocaleString('es-NI', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`,
    releasedBy,
    costCenter,
    dateIso: now.toISOString(),
    formattedDate: now.toLocaleDateString('es-ES', { day: '2-digit', month: '2-digit', year: 'numeric' }),
    formattedTime: now.toLocaleTimeString('es-ES', { hour: '2-digit', minute: '2-digit', second: '2-digit' })
  };

  // Agregar al inicio
  AppState.orders.unshift(newOrder);
  saveData('orders');

  renderStats();
  renderOrdersTable();
  updatePortalCounts();
  clearForm(false);

  showToast(`✅ Viaje <strong>${newOrder.id}</strong> registrado exitosamente (${newOrder.formattedAmount})`, "success");

  // Listo para la siguiente orden
  dom.clientIdSelect.focus();
}

function clearForm(clearAll = false) {
  dom.clientIdSelect.selectedIndex = 0;
  dom.clientNameSelect.selectedIndex = 0;
  dom.orderReasonSelect.selectedIndex = 0;
  dom.accountNumberSelect.selectedIndex = 0;
  dom.orderConcept.value = '';
  dom.orderAmount.value = '';
  dom.releasedBySelect.selectedIndex = 0;
  dom.costCenterInput.value = '';
}

// ==========================================
// RENDERIZADO DE TABLA Y ESTADÍSTICAS
// ==========================================
function renderStats() {
  const total = AppState.orders.length;
  const cdn = AppState.orders.filter(o => o.company === 'CDN').length;
  const ccn = AppState.orders.filter(o => o.company === 'CCN').length;

  const totalAmountSum = AppState.orders.reduce((acc, curr) => acc + (curr.amount || 0), 0);

  dom.statTotal.textContent = total;
  dom.statCdn.textContent = cdn;
  dom.statCcn.textContent = ccn;
  dom.statTotalAmount.textContent = `C$ ${totalAmountSum.toLocaleString('es-NI', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
  dom.tableCountBadge.textContent = `${total} Registros`;
}

function renderOrdersTable() {
  let list = [...AppState.orders];

  // Filtrado por empresa si está activo
  if (AppState.filterCompany !== 'all') {
    list = list.filter(o => o.company === AppState.filterCompany);
  }

  // Filtrado por búsqueda en tiempo real
  if (AppState.searchQuery) {
    const q = AppState.searchQuery;
    list = list.filter(o => 
      o.id.toLowerCase().includes(q) ||
      o.clientId.toLowerCase().includes(q) ||
      o.clientName.toLowerCase().includes(q) ||
      o.orderReason.toLowerCase().includes(q) ||
      o.accountNumber.toLowerCase().includes(q) ||
      o.orderConcept.toLowerCase().includes(q) ||
      o.releasedBy.toLowerCase().includes(q) ||
      (o.costCenter && o.costCenter.toLowerCase().includes(q)) ||
      (o.formattedAmount && o.formattedAmount.toLowerCase().includes(q))
    );
  }

  if (list.length === 0) {
    dom.ordersTableBody.innerHTML = '';
    dom.emptyState.style.display = 'block';
    return;
  }

  dom.emptyState.style.display = 'none';

  dom.ordersTableBody.innerHTML = list.map(order => `
    <tr>
      <td>
        <span style="font-family: monospace; font-weight: 700; color: var(--text-muted);">${order.id}</span>
      </td>
      <td>
        <span class="badge-corp ${order.company.toLowerCase()}">${order.company}</span>
      </td>
      <td>
        <strong style="color: #60a5fa; font-family: monospace; font-size: 0.95rem;">${escapeHtml(order.clientId)}</strong>
      </td>
      <td>
        <span style="font-weight: 600;">${escapeHtml(order.clientName)}</span>
      </td>
      <td>${escapeHtml(order.orderReason)}</td>
      <td>
        <span style="font-family: monospace; background: rgba(255,255,255,0.06); padding: 3px 8px; border-radius: 4px;">
          ${escapeHtml(order.accountNumber)}
        </span>
      </td>
      <td class="cell-concept" title="${escapeHtml(order.orderConcept)}">
        ${escapeHtml(order.orderConcept)}
      </td>
      <td>
        <strong style="color: #10b981; font-family: monospace; font-size: 0.92rem;">
          ${order.formattedAmount || ('C$ ' + (order.amount || 0).toFixed(2))}
        </strong>
      </td>
      <td>
        <span style="color: #a78bfa; font-weight: 600;">${escapeHtml(order.releasedBy)}</span>
      </td>
      <td>
        <span style="font-family: monospace; font-weight: 700; color: #60a5fa; background: rgba(59, 130, 246, 0.1); padding: 3px 8px; border-radius: 4px;">
          ${escapeHtml(order.costCenter || '')}
        </span>
      </td>
      <td style="font-size: 0.76rem; color: var(--text-dim);">
        ${order.formattedDate} <br/>
        <span style="opacity:0.8">${order.formattedTime}</span>
      </td>
      <td style="text-align: center;">
        <button class="btn btn-outline btn-sm" onclick="deleteOrder('${order.id}')" title="Eliminar orden" style="color: #f43f5e; padding: 4px 8px;">
          ✕
        </button>
      </td>
    </tr>
  `).join('');
}

window.deleteOrder = function(orderId) {
  if (confirm(`¿Está seguro de eliminar la orden ${orderId}?`)) {
    AppState.orders = AppState.orders.filter(o => o.id !== orderId);
    saveData('orders');
    renderStats();
    renderOrdersTable();
    updatePortalCounts();
    showToast(`Orden eliminada`, "info");
  }
};

function confirmClearAllOrders() {
  if (AppState.orders.length === 0) {
    showToast("No hay órdenes registradas para eliminar", "info");
    return;
  }
  if (confirm(`⚠️ ¿Desea eliminar TODAS las ${AppState.orders.length} órdenes registradas? Esta acción no se puede deshacer.`)) {
    AppState.orders = [];
    saveData('orders');
    renderStats();
    renderOrdersTable();
    updatePortalCounts();
    showToast("Todas las órdenes han sido eliminadas", "warning");
  }
}

// ==========================================
// GENERACIÓN DE EXCEL (.xlsx)
// ==========================================
function generateExcel() {
  if (AppState.orders.length === 0) {
    showToast("⚠️ No hay órdenes registradas para exportar. Registre al menos una.", "warning");
    return;
  }

  showToast("⏳ Generando archivo oficial de Excel...", "info");

  // Requerimientos exactos para el Excel:
  // 1. Número de Orden: Primer recuadro, va vacío porque aún no está asignada
  // 2. Número de Recibo: Segundo recuadro, va vacío (no plasmado en app)
  // 3. Empresa, ID Cliente, Nombre, Motivo, Cuenta, Concepto, Monto, Liberador, Centro de Costo, Fecha, Hora
  const headers = [
    "Número de Orden",
    "Número de Recibo",
    "Empresa",
    "ID del Cliente",
    "Nombre del Cliente",
    "Motivo de la Orden",
    "Número de Cuenta",
    "Concepto de la Orden",
    "Monto de la Orden",
    "Quién lo Libera",
    "Centro de Costo",
    "Fecha Registro",
    "Hora Registro"
  ];

  const rows = AppState.orders.map(order => [
    "", // Primer recuadro: Número de Orden VACÍO (no está asignada todavía)
    "", // Segundo recuadro: Número de Recibo VACÍO (no plasmado en app)
    order.company,
    order.clientId,
    order.clientName,
    order.orderReason,
    order.accountNumber,
    order.orderConcept,
    order.amount || 0,
    order.releasedBy,
    order.costCenter || "",
    order.formattedDate,
    order.formattedTime
  ]);

  const data = [headers, ...rows];
  const timestamp = new Date().toISOString().slice(0, 10);
  const companyTag = AppState.selectedCompany || 'General';
  const filename = `Reporte_Pago_Viajes_Tercerizados_${companyTag}_${timestamp}.xlsx`;

  if (typeof XLSX !== 'undefined') {
    try {
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet(data);

      // Auto-dimensionar anchos de columnas
      const colWidths = headers.map((header, colIdx) => {
        let maxLen = header.length;
        rows.forEach(row => {
          const val = row[colIdx] ? String(row[colIdx]) : "";
          if (val.length > maxLen) maxLen = val.length;
        });
        return { wch: Math.min(Math.max(maxLen + 4, 16), 50) };
      });
      ws['!cols'] = colWidths;

      XLSX.utils.book_append_sheet(wb, ws, "Órdenes de Servicio");

      XLSX.writeFile(wb, filename);

      const exportedCount = AppState.orders.length;

      // Borrar automáticamente los viajes capturados tras generar el Excel
      AppState.orders = [];
      saveData('orders');
      renderStats();
      renderOrdersTable();
      updatePortalCounts();

      showToast(`🎉 ¡Excel descargado con éxito! Se exportaron y limpiaron ${exportedCount} viajes.`, "success");
      return;
    } catch (err) {
      console.warn("Fallo en XLSX, usando contingencia CSV:", err);
    }
  }

  fallbackExportCsv(data);
}

function fallbackExportCsv(data) {
  let csvContent = "\uFEFF"; // BOM para caracteres latinos en Excel
  data.forEach(row => {
    const formattedRow = row.map(item => {
      const str = String(item ?? "").replace(/"/g, '""');
      return `"${str}"`;
    });
    csvContent += formattedRow.join(";") + "\r\n";
  });

  const exportedCount = AppState.orders.length;
  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  const timestamp = new Date().toISOString().slice(0, 10);
  const companyTag = AppState.selectedCompany || 'General';
  a.href = url;
  a.download = `Reporte_Pago_Viajes_Tercerizados_${companyTag}_${timestamp}.csv`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);

  // Borrar automáticamente los viajes capturados tras generar el archivo
  AppState.orders = [];
  saveData('orders');
  renderStats();
  renderOrdersTable();
  updatePortalCounts();

  showToast(`📁 Archivo para Excel descargado (${exportedCount} viajes). Tabla limpiada.`, "success");
}

// ==========================================
// GESTOR DE CATÁLOGOS (MODAL)
// ==========================================
function openCatalogModal() {
  renderCatalogTables();
  dom.catalogModal.classList.add('open');
}

function closeCatalogModal() {
  dom.catalogModal.classList.remove('open');
  populateDropdowns();
}

function renderCatalogTables() {
  // Clientes
  dom.catalogClientsTable.innerHTML = AppState.clients.map((c, i) => `
    <tr>
      <td><strong style="color:#60a5fa; font-family: monospace;">${c.id}</strong></td>
      <td>${escapeHtml(c.name)}</td>
      <td style="text-align:right">
        <button class="btn btn-outline btn-sm" onclick="removeClientFromCatalog(${i})" style="color:#f43f5e; padding:2px 8px;">✕</button>
      </td>
    </tr>
  `).join('');

  // Motivos y Cuentas
  dom.catalogReasonsTable.innerHTML = AppState.reasons.map((r, i) => `
    <tr>
      <td>${escapeHtml(r.reason)}</td>
      <td><strong style="font-family:monospace">${escapeHtml(r.account)}</strong></td>
      <td style="text-align:right">
        <button class="btn btn-outline btn-sm" onclick="removeReasonFromCatalog(${i})" style="color:#f43f5e; padding:2px 8px;">✕</button>
      </td>
    </tr>
  `).join('');

  // Liberadores con su Centro de Costo (solo números)
  dom.catalogApproversTable.innerHTML = AppState.approvers.map((app, i) => `
    <tr>
      <td><span style="color:#a78bfa; font-weight:600;">${escapeHtml(app.name)}</span></td>
      <td><strong style="color:#60a5fa; font-family: monospace;">${escapeHtml(app.costCenter || '')}</strong></td>
      <td style="text-align:right">
        <button class="btn btn-outline btn-sm" onclick="removeApproverFromCatalog(${i})" style="color:#f43f5e; padding:2px 8px;">✕</button>
      </td>
    </tr>
  `).join('');
}

function addClientToCatalog() {
  const id = dom.newClientId.value.trim();
  const name = dom.newClientName.value.trim();
  if (!id || !name) {
    showToast("Ingrese el ID (número) y los dos nombres y dos apellidos", "warning");
    return;
  }
  AppState.clients.push({ id, name });
  saveData('clients');
  dom.newClientId.value = '';
  dom.newClientName.value = '';
  renderCatalogTables();
  populateDropdowns();
  showToast("Cliente agregado al catálogo", "success");
}

window.removeClientFromCatalog = function(index) {
  AppState.clients.splice(index, 1);
  saveData('clients');
  renderCatalogTables();
  populateDropdowns();
  showToast("Cliente eliminado del catálogo", "info");
};

function addReasonToCatalog() {
  const reason = dom.newReason.value.trim();
  const account = dom.newAccount.value.trim().toUpperCase();
  if (!reason || !account) {
    showToast("Ingrese Motivo de viaje y Número de Cuenta", "warning");
    return;
  }
  AppState.reasons.push({ reason, account });
  saveData('reasons');
  dom.newReason.value = '';
  dom.newAccount.value = '';
  renderCatalogTables();
  populateDropdowns();
  showToast("Motivo agregado al catálogo", "success");
}

window.removeReasonFromCatalog = function(index) {
  AppState.reasons.splice(index, 1);
  saveData('reasons');
  renderCatalogTables();
  populateDropdowns();
  showToast("Motivo eliminado del catálogo", "info");
};

function addApproverToCatalog() {
  const name = dom.newApproverName.value.trim();
  const costCenter = dom.newApproverCostCenter.value.trim();
  if (!name || !costCenter) {
    showToast("Ingrese el nombre del supervisor y su centro de costo (solo números)", "warning");
    return;
  }
  if (!/^\d+$/.test(costCenter)) {
    showToast("El centro de costo debe contener solo números (ej: 700720056)", "warning");
    return;
  }
  AppState.approvers.push({ name, costCenter });
  saveData('approvers');
  dom.newApproverName.value = '';
  dom.newApproverCostCenter.value = '';
  renderCatalogTables();
  populateDropdowns();
  showToast("Supervisor y Centro de Costo agregados al catálogo", "success");
}

window.removeApproverFromCatalog = function(index) {
  AppState.approvers.splice(index, 1);
  saveData('approvers');
  renderCatalogTables();
  populateDropdowns();
  showToast("Supervisor eliminado de la lista", "info");
};

// ==========================================
// TEMA Y UTILIDADES
// ==========================================
function setTheme(theme) {
  AppState.theme = theme;
  document.documentElement.setAttribute('data-theme', theme);
  localStorage.setItem('cdn_ccn_theme', theme);
  const label = theme === 'dark' ? '☀️ Modo Claro' : '🌙 Modo Oscuro';
  dom.btnToggleTheme.textContent = label;
  dom.portalBtnToggleTheme.textContent = label;
}

function toggleTheme() {
  const nextTheme = AppState.theme === 'dark' ? 'light' : 'dark';
  setTheme(nextTheme);
}

function showToast(message, type = "info") {
  const toast = document.createElement('div');
  toast.className = `toast ${type}`;
  toast.innerHTML = message;
  dom.toastContainer.appendChild(toast);

  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateY(10px)';
    setTimeout(() => toast.remove(), 300);
  }, 3500);
}

function escapeHtml(text) {
  if (!text) return "";
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

export const RAW_CSV_PATHS = [
  'data/busbar-webapp-embedded-v2.csv',
  'data/busbar-webapp-embedded-v2.1.csv',
  'data/busbar-webapp-embedded-v2.2.csv',
  'data/XAP_busbar_prisliste_ekstrakt_5W_clean.csv'
];

export const XAP_SERIES = 'XAP-B';
export const EPOXY_IP68_SERIES = 'RCP-IP68';
export const COMPARISON_ELIGIBLE_SERIES = Object.freeze(['XCM','XCP-S']);
export const ENABLE_XAP_COMPARISON = false;
export const LOCKED_LEDERE_BY_SERIES = Object.freeze({
  XCM: '3F+N+PE',
  [XAP_SERIES]: '3F+N+PE',
  [EPOXY_IP68_SERIES]: '3F+N'
});
export const CRT_FEED_ALLOWED_SERIES = Object.freeze(['XCP-S', XAP_SERIES]);
export const seriesLockedLedereValue = series => LOCKED_LEDERE_BY_SERIES[series] || '';
export const seriesLocksLedere = series => Boolean(seriesLockedLedereValue(series));
export const seriesSupportsCrtFeed = series => CRT_FEED_ALLOWED_SERIES.includes(series);
export const shouldCompareXap = series => ENABLE_XAP_COMPARISON && COMPARISON_ELIGIBLE_SERIES.includes(series);

export const USD_TO_NOK_RATE = 10.95;
export const MARKET_REFRESH_INTERVAL_MS = 5 * 60 * 1000;
export const MARKET_STATUS_DEFAULT = 'Oppdateres daglig';
export const MARKET_STATUS_MANUAL = 'Oppdateres manuelt';
export const DEFAULT_MARGIN_RATE = 0.20;
export const DEFAULT_MATERIAL_MARGIN_RATE = 0.25;
export const MAX_MARGIN_RATE = 0.95;
export const MAX_AUTO_MONTERING_MARGIN_RATE = 0.99;

export const AUTH_SESSION_KEY = 'busbar.auth.session.v1';
export const ADMIN_NAV_ALLOWED_EMAILS = Object.freeze(['hans.jakob.risnes@busbar.no']);
export const LEGACY_PROJECTS_STORAGE_KEY = 'busbar.projects.v1';
export const PROJECTS_STORAGE_KEY_PREFIX = 'busbar.projects.user.v2';
export const PROJECT_SYNC_DEBOUNCE_MS = 800;
export const EMAIL_REGEX = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

export const MICROSOFT_AUTH_DEFAULT_SCOPES = Object.freeze(['openid', 'profile', 'email']);
export const MICROSOFT_GRAPH_CALENDAR_SCOPES = Object.freeze(['Calendars.ReadWrite']);
export const MICROSOFT_GRAPH_OUTLOOK_CATEGORY_SCOPES = Object.freeze(['MailboxSettings.ReadWrite']);
export const MICROSOFT_GRAPH_MAIL_SCOPES = Object.freeze(['Mail.ReadWrite', 'Mail.Send', 'Mail.ReadWrite.Shared', 'Mail.Send.Shared']);
export const MICROSOFT_GRAPH_SHAREPOINT_SCOPES = Object.freeze(['Files.ReadWrite.All', 'Sites.Read.All']);
export const MICROSOFT_GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';
export const CALENDAR_PROJECT_EXTENDED_PROPERTY_ID = 'String {4f28b47f-5e6b-4f87-9fcb-4c1a9c0c9d9f} Name BusbarProjectId';
export const CALENDAR_PROJECT_FLOW_TASK_EXTENDED_PROPERTY_ID = 'String {4f28b47f-5e6b-4f87-9fcb-4c1a9c0c9d9f} Name BusbarProjectFlowTaskId';
export const CALENDAR_TODO_EXTENDED_PROPERTY_ID = 'String {4f28b47f-5e6b-4f87-9fcb-4c1a9c0c9d9f} Name BusbarTodoId';
export const CALENDAR_EVENT_TYPE_EXTENDED_PROPERTY_ID = 'String {4f28b47f-5e6b-4f87-9fcb-4c1a9c0c9d9f} Name BusbarCalendarEventType';
export const OUTLOOK_PROJECT_CATEGORY_NAME = 'Busbar Prosjekt';
export const OUTLOOK_PROJECT_CATEGORY_COLOR = 'preset3';
export const OUTLOOK_TODO_CATEGORY_NAME = 'To-Do';
export const OUTLOOK_TODO_CATEGORY_COLOR = 'preset0';
export const OUTLOOK_TODO_COMPLETED_CATEGORY_NAME = 'To-Do fullført';
export const OUTLOOK_TODO_COMPLETED_CATEGORY_COLOR = 'preset4';
export const PROJECT_MAILBOX_ADDRESS = 'prosjekt@busbar.no';

export const SHAREPOINT_FOLDER_CONFIG = Object.freeze({
  'busbar-folders': {
    title: 'Busbar',
    statusId: 'busbarFoldersStatus',
    listId: 'busbarFoldersList',
    refreshBtnId: 'refreshBusbarFoldersBtn',
    siteHost: 'mcselektrotavler.sharepoint.com',
    sitePath: '/sites/BCDokumentarkiv',
    folderPath: 'Drift/BUSBAR'
  },
  'project-folders': {
    title: 'Prosjektmapper',
    statusId: 'projectFoldersStatus',
    listId: 'projectFoldersList',
    refreshBtnId: 'refreshProjectFoldersBtn',
    siteHost: 'mcselektrotavler.sharepoint.com',
    sitePath: '/sites/BCDokumentarkiv',
    folderPath: 'Drift/BUSBAR/Prosjekter/2026'
  },
  'supplier-folders': {
    title: 'Leverandørmapper',
    statusId: 'supplierFoldersStatus',
    listId: 'supplierFoldersList',
    refreshBtnId: 'refreshSupplierFoldersBtn',
    siteHost: 'mcselektrotavler.sharepoint.com',
    sitePath: '/sites/BCDokumentarkiv',
    folderPath: 'Drift/BUSBAR/Strømskinner'
  }
});

export const PROJECT_FOLDER_TEMPLATE_NAME = '0-MAL';
export const DASHBOARD_SIDEBAR_COLLAPSED_KEY = 'busbar.dashboard.sidebar.collapsed.v1';
export const PROJECT_SORT_STORAGE_KEY = 'busbar.project.sort.v1';
export const LINE_SORT_STORAGE_KEY = 'busbar.line.sort.v1';
export const OFFER_SORT_STORAGE_KEY = 'busbar.offer.sort.v1';
export const PROJECT_FLOW_STORAGE_KEY = 'busbar.project.flow.v1';
export const PROJECT_FLOW_ALL_PROJECTS = '__all__';
export const PROJECT_SORT_OPTIONS = Object.freeze(['date_newest', 'date_oldest', 'alpha_asc', 'alpha_desc']);
export const LINE_SORT_OPTIONS = Object.freeze(['date_newest', 'date_oldest', 'alpha_asc', 'alpha_desc']);
export const PROJECT_STATUS_OPTIONS = Object.freeze([
  { id: 'unresolved', label: 'Uavklart', tone: 'idle' },
  { id: 'won', label: 'Vunnet', tone: 'success' },
  { id: 'lost', label: 'Tapt', tone: 'danger' },
  { id: 'finished', label: 'Ferdig', tone: 'done' }
]);
export const PROJECT_ARCHIVE_STATUS_IDS = Object.freeze(['lost', 'finished']);
// betaTarget er kun koblingsmetadata. Prosjektflytene deler foreløpig ingen status eller handlinger.
export const PROJECT_FLOW_PHASES = Object.freeze([
  { id: 'request', label: 'Forespørsel', betaTarget: { type: 'task', number: '1.01.0' } },
  { id: 'offer', label: 'Tilbud', betaTarget: { type: 'task', number: '1.02.0' } },
  { id: 'order', label: 'Bestilling', betaTarget: { type: 'task', number: '1.03.0' } },
  { id: 'engineering', label: 'Prosjektering', betaTarget: { type: 'phase', number: '2.00.0' } },
  { id: 'procurement', label: 'Innkjøp', betaTarget: { type: 'phase', number: '3.00.0' } },
  { id: 'production', label: 'Bekreftelse fra leverandør', betaTarget: { type: 'task', number: '3.04.0' } },
  { id: 'delivery', label: 'Leveranse strømskinne', betaTarget: { type: 'task', number: 'X.03.0' } },
  { id: 'finished', label: 'Fakturert', betaTarget: { type: 'task', number: '1.04.0' } }
]);
export const PROJECT_FLOW_VISIBLE_WEEK_LEVELS = Object.freeze([5, 4, 3, 2]);
export const PROJECT_FLOW_DEFAULT_ZOOM_INDEX = 1;

// ==UserScript==
// @name         CopilotKeep — M365 chat backup & export
// @namespace    dv64.m365.copilot.exporter
// @version      0.1.0
// @description  Export Copilot chats (current + archive) via centered modal. Settings page with English/Arabic/Chinese UI. Interception-first auth, fresh fetch, rate-limit backoff, deflate ZIP. No delete operations.
// @author       DV64
// @license      MIT
// @match        https://m365.cloud.microsoft/*
// @match        https://copilot.cloud.microsoft/*
// @match        https://*.cloud.microsoft/*
// @run-at       document-start
// @grant        none
// @noframes
// ==/UserScript==

(function () {
  'use strict';
  if (window.__dv64M365CopilotExporterLoaded) return;
  window.__dv64M365CopilotExporterLoaded = true;

  var VERSION = '0.1.0';
  var APP = 'CopilotKeep';
  var CLIENT_ID = 'c0ab8ce9-e9a0-42e7-b064-33d422df41f1';
  var SUBSTRATE = 'https://substrate.office.com/m365Copilot';
  var CHAT_WRAPPER = '[data-testid="m365-chat-llm-web-ui-chat-message"]';
  var MAX_LOG_ENTRIES = 400;
  var REFRESH_MARGIN_SEC = 120;
  var RAM_WARN_MB = 60;
  var SYMBOL_URL = Symbol('dv64_xhr_url');
  var SYMBOL_HDR = Symbol('dv64_xhr_headers');

  /* ═══════ SETTINGS ═══════ */
  var LANGS = ['en', 'ar', 'zh'];
  var FULL_FMTS = ['md', 'clean', 'gpt', 'raw'];
  var ARC_FMTS = ['zip', 'md', 'clean', 'gpt', 'raw'];

  var DEFAULT_SETTINGS = {
    lang: 'en',              // en | ar | zh
    theme: 'auto',           // auto | light | dark
    fetchDelayMs: 700,       // delay between archive fetches
    pageSize: 30,            // archive rows per page
    maxChats: 100,           // default archive fetch limit
    defaultFullFmt: 'md',    // current-chat full export format
    defaultArchiveFmt: 'zip',// archive export format
    zipCompress: true,       // deflate ZIP when supported
    includeInternal: false,  // keep progress/tool messages
    includeUser: true,
    includeAssistant: true,
    toastDurationMs: 3000
  };

  var SETTINGS_KEY = 'dv64Exporter.settings';

  function clampNum(v, min, max, dflt) {
    v = Number(v);
    if (!isFinite(v)) return dflt;
    return Math.min(max, Math.max(min, Math.round(v)));
  }

  function mergeSettings(raw) {
    var s = JSON.parse(JSON.stringify(DEFAULT_SETTINGS));
    if (!raw || typeof raw !== 'object') return s;
    Object.keys(DEFAULT_SETTINGS).forEach(function (k) {
      if (raw[k] == null) return;
      switch (k) {
        case 'lang': s.lang = LANGS.indexOf(raw.lang) !== -1 ? raw.lang : 'en'; break;
        case 'theme': s.theme = ['auto', 'light', 'dark'].indexOf(raw.theme) !== -1 ? raw.theme : 'auto'; break;
        case 'fetchDelayMs': s.fetchDelayMs = clampNum(raw.fetchDelayMs, 100, 5000, DEFAULT_SETTINGS.fetchDelayMs); break;
        case 'pageSize': s.pageSize = clampNum(raw.pageSize, 5, 100, DEFAULT_SETTINGS.pageSize); break;
        case 'maxChats': s.maxChats = clampNum(raw.maxChats, 1, 500, DEFAULT_SETTINGS.maxChats); break;
        case 'defaultFullFmt': s.defaultFullFmt = FULL_FMTS.indexOf(raw.defaultFullFmt) !== -1 ? raw.defaultFullFmt : 'md'; break;
        case 'defaultArchiveFmt': s.defaultArchiveFmt = ARC_FMTS.indexOf(raw.defaultArchiveFmt) !== -1 ? raw.defaultArchiveFmt : 'zip'; break;
        case 'zipCompress': case 'includeInternal': case 'includeUser': case 'includeAssistant':
          s[k] = !!raw[k]; break;
        case 'toastDurationMs': s.toastDurationMs = clampNum(raw.toastDurationMs, 1000, 15000, DEFAULT_SETTINGS.toastDurationMs); break;
      }
    });
    return s;
  }

  function loadSettings() {
    try {
      var raw = localStorage.getItem(SETTINGS_KEY);
      return mergeSettings(raw ? JSON.parse(raw) : null);
    } catch (e) { return mergeSettings(null); }
  }
  function saveSettings() {
    try { localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings)); } catch (e) { /* storage unavailable */ }
  }

  var settings = loadSettings();

  /* ═══════ I18N ═══════ */
  var STR = {
    en: {
      appTagline: 'v{v} \u00B7 fresh fetch \u00B7 rate-limit aware',
      tabCurrent: 'Current chat', tabArchive: 'Archive', tabLogs: 'Logs', tabSettings: 'Settings',
      close: 'Close', cancel: 'Cancel', confirm: 'Confirm',
      copy: 'Copy', clear: 'Clear',
      logAll: 'All', logInfo: 'Info', logWarn: 'Warn', logError: 'Error',
      loading: 'Loading\u2026',
      noLogsYet: 'No logs yet', noLogsLevel: 'No {level} logs yet',
      copied: 'Copied', copyFailed: 'Copy failed', cancelled: 'Cancelled',
      waitOrCancel: 'Wait or cancel the current operation first',
      userMessages: 'User messages', copilotReplies: 'Copilot replies', includeInternal: 'Include internal',
      copyPageBtn: 'Copy page messages', dlPageBtn: 'Download page ({fmt})',
      fullExportLabel: 'Full export:', exportFromMs: 'Export from Microsoft',
      currentHelp: 'Page methods read only the visible DOM (partial for long chats). Full export always fetches fresh history from Microsoft\u2019s API.',
      fmtMd: 'Markdown', fmtClean: 'Clean JSON', fmtGpt: 'ChatGPT JSON', fmtRaw: 'Raw JSON (\u26A0 sensitive)', fmtZip: 'ZIP (recommended)',
      openChatFirst: 'Open a Copilot chat first.', openChatFirstShort: 'Open a chat first',
      visibleMsgs: '{n} visible messages', cachedHistory: 'Full history cached ({n} msgs)',
      notCapturedYet: 'Full history not captured yet',
      partialVisible: 'Only {a} of {b} visible \u2014 use Full export',
      idShort: 'ID: {id}',
      extracting: 'Extracting page messages\u2026',
      copiedN: 'Copied {n} messages', downloadedN: 'Downloaded {n} messages',
      partialNote: 'partial: {a}/{b}',
      partialLong: 'Partial export: {a} of {b} messages. Use Full export for the complete history.',
      fetchLatest: 'Fetch latest', load: 'Load',
      colName: 'Name', colCreated: 'Created', colUpdated: 'Updated', colStatus: 'Status',
      searchPh: 'Search\u2026', sortLabel: 'Sort',
      sortUpdated: 'Newest updated', sortCreated: 'Newest created', sortName: 'Name A\u2013Z',
      selectAll: 'Select all', format: 'Format', exportSelected: 'Export selected',
      archiveNote: 'Archive uses Microsoft\u2019s internal Substrate API. May change without notice. No delete operations.',
      noConvs: 'No conversations loaded. Click \u201CLoad\u201D.',
      noMatch: 'No conversations match your search.',
      untitled: 'Untitled', stExporting: 'Exporting', stDone: 'Done', stError: 'Error',
      loadingArchive: 'Loading archive list\u2026', loadedN: 'Loaded {n} conversations',
      startingExport: 'Starting export\u2026', starting: 'Starting\u2026', allFailed: 'All exports failed.',
      buildingZip: 'Building {mb} MB ZIP. May take a moment.',
      exportedPart: 'Exported {a}/{b} ({c} failed)', exportedNConvs: 'Exported {n} conversations',
      exportedNMsgs: 'Exported {n} messages', cachedNote: '(cached)',
      rawConfirmT: 'Export Raw JSON?',
      rawConfirmBodyFull: 'Raw JSON may contain location data, user IDs, tenant IDs, and telemetry. Only for private backup.',
      rawConfirmBodyArc: 'Raw JSON may contain sensitive metadata. Only for private backup.',
      exportRaw: 'Export raw',
      preparingFull: 'Preparing full export\u2026', fetchingMs: 'Fetching from Microsoft\u2026',
      usingCache: 'Using cached copy (fresh fetch failed)',
      launchTitle: 'Open Copilot Exporter',
      secGeneral: 'General',
      rowLanguage: 'Language', rowLanguageDesc: 'Choose the interface language.',
      rowTheme: 'Theme', rowThemeDesc: 'Follow the system or force light/dark.',
      themeAuto: 'Auto (system)', themeLight: 'Light', themeDark: 'Dark',
      secDefaults: 'Export defaults',
      rowFullFmt: 'Full export format', rowFullFmtDesc: 'Default format in the Current chat tab.',
      rowArcFmt: 'Archive format', rowArcFmtDesc: 'Default format in the Archive tab.',
      rowZip: 'Compress ZIP (deflate)', rowZipDesc: 'Shrink archive exports using deflate.',
      rowInternal: 'Include internal messages', rowInternalDesc: 'Keep progress and tool messages.',
      secPerformance: 'Performance',
      rowFetchDelay: 'Delay between fetches (ms)', rowFetchDelayDesc: 'Wait between requests during batch export.',
      rowPageSize: 'Rows per page', rowPageSizeDesc: 'Conversations shown per Archive page.',
      rowMaxChats: 'Default chat limit', rowMaxChatsDesc: 'Default count when loading the archive.',
      rowToastDur: 'Toast duration (ms)', rowToastDurDesc: 'How long notifications stay visible.',
      secData: 'Data', resetBtn: 'Reset to defaults',
      resetConfirmT: 'Reset settings?', resetConfirmB: 'All settings return to their defaults.',
      settingsSaved: 'Settings saved', settingsReset: 'Settings reset to defaults',
      langEn: 'English', langAr: 'العربية', langZh: '中文',
      errEncCookieMissing: 'Encryption cookie missing. Reload the page.',
      errEncCookieBad: 'Invalid encryption cookie.',
      errMsalKey: 'MSAL key missing.',
      errNoToken: 'No valid substrate token found. Reload the page.',
      errBadConvId: 'Invalid conversation ID.',
      errBadJson: 'Response is not valid JSON.',
      errConvShape: 'Unexpected conversation structure.',
      errConvMismatch: 'Conversation ID mismatch.',
      errNoPageMsgs: 'No visible messages on the page.',
      errAuthExpired: 'Auth expired. Reload the page and try again.',
      errTokenUnchanged: 'Token unchanged after refresh \u2014 reload the page.'
    },
    ar: {
      appTagline: 'الإصدار {v} · جلب مباشر · مراعاة حدود المعدل',
      tabCurrent: 'المحادثة الحالية', tabArchive: 'الأرشيف', tabLogs: 'السجلات', tabSettings: 'الإعدادات',
      close: 'إغلاق', cancel: 'إلغاء', confirm: 'تأكيد',
      copy: 'نسخ', clear: 'مسح',
      logAll: 'الكل', logInfo: 'معلومات', logWarn: 'تحذير', logError: 'خطأ',
      loading: 'جارٍ التحميل…',
      noLogsYet: 'لا توجد سجلات بعد', noLogsLevel: 'لا توجد سجلات {level} بعد',
      copied: 'تم النسخ', copyFailed: 'فشل النسخ', cancelled: 'تم الإلغاء',
      waitOrCancel: 'انتظر انتهاء العملية الحالية أو ألغِها أولاً',
      userMessages: 'رسائل المستخدم', copilotReplies: 'ردود Copilot', includeInternal: 'تضمين الرسائل الداخلية',
      copyPageBtn: 'نسخ رسائل الصفحة', dlPageBtn: 'تنزيل الصفحة ({fmt})',
      fullExportLabel: 'التصدير الكامل:', exportFromMs: 'التصدير من Microsoft',
      currentHelp: 'قراءة الصفحة تقتصر على المحتوى الظاهر (جزئية في المحادثات الطويلة)، أما التصدير الكامل فيجلب دائماً السجل الكامل من واجهة Microsoft.',
      fmtMd: 'Markdown', fmtClean: 'JSON مبسّط', fmtGpt: 'JSON بصيغة ChatGPT', fmtRaw: 'JSON خام (⚠ حسّاس)', fmtZip: 'ZIP (موصى به)',
      openChatFirst: 'افتح محادثة Copilot أولاً.', openChatFirstShort: 'افتح محادثة أولاً',
      visibleMsgs: '{n} رسالة ظاهرة', cachedHistory: 'سجل كامل مخزّن ({n} رسالة)',
      notCapturedYet: 'لم يتم التقاط السجل الكامل بعد',
      partialVisible: 'يظهر {a} من {b} فقط — استخدم التصدير الكامل',
      idShort: 'المعرّف: {id}',
      extracting: 'جارٍ استخراج رسائل الصفحة…',
      copiedN: 'تم نسخ {n} رسالة', downloadedN: 'تم تنزيل {n} رسالة',
      partialNote: 'جزئي: {a}/{b}',
      partialLong: 'تصدير جزئي: {a} من أصل {b} رسالة. استخدم التصدير الكامل للحصول على السجل الكامل.',
      fetchLatest: 'جلب الأحدث', load: 'تحميل',
      colName: 'الاسم', colCreated: 'تاريخ الإنشاء', colUpdated: 'تاريخ التحديث', colStatus: 'الحالة',
      searchPh: 'بحث…', sortLabel: 'ترتيب',
      sortUpdated: 'الأحدث تحديثاً', sortCreated: 'الأحدث إنشاءً', sortName: 'الاسم أ–ي',
      selectAll: 'تحديد الكل', format: 'التنسيق', exportSelected: 'تصدير المحدد',
      archiveNote: 'يستخدم الأرشيف واجهة Substrate الداخلية من Microsoft وقد تتغير دون إشعار. لا توجد عمليات حذف.',
      noConvs: 'لا توجد محادثات محمّلة. اضغط "تحميل".',
      noMatch: 'لا توجد محادثات مطابقة لبحثك.',
      untitled: 'بدون عنوان', stExporting: 'جارٍ التصدير', stDone: 'تم', stError: 'خطأ',
      loadingArchive: 'جارٍ تحميل قائمة الأرشيف…', loadedN: 'تم تحميل {n} محادثة',
      startingExport: 'جارٍ بدء التصدير…', starting: 'بدء…', allFailed: 'فشلت جميع عمليات التصدير.',
      buildingZip: 'جارٍ إنشاء ملف ZIP بحجم {mb} ميغابايت. قد يستغرق بعض الوقت.',
      exportedPart: 'تم تصدير {a}/{b} ({c} فشلت)', exportedNConvs: 'تم تصدير {n} محادثة',
      exportedNMsgs: 'تم تصدير {n} رسالة', cachedNote: '(نسخة مخزّنة)',
      rawConfirmT: 'تصدير JSON الخام؟',
      rawConfirmBodyFull: 'قد يحتوي JSON الخام على بيانات الموقع ومعرّفات المستخدمين والمستأجرين وبيانات القياس عن بُعد. للاستخدام في النسخ الاحتياطي الخاص فقط.',
      rawConfirmBodyArc: 'قد يحتوي JSON الخام على بيانات وصفية حسّاسة. للاستخدام في النسخ الاحتياطي الخاص فقط.',
      exportRaw: 'تصدير الخام',
      preparingFull: 'جارٍ تحضير التصدير الكامل…', fetchingMs: 'جارٍ الجلب من Microsoft…',
      usingCache: 'استخدام نسخة مخزّنة (فشل الجلب المباشر)',
      launchTitle: 'فتح مُصدِّر Copilot',
      secGeneral: 'عام',
      rowLanguage: 'لغة الواجهة', rowLanguageDesc: 'اختر لغة عرض الواجهة.',
      rowTheme: 'المظهر', rowThemeDesc: 'يتبع النظام أو يفرض فاتحاً/غامقاً.',
      themeAuto: 'تلقائي (النظام)', themeLight: 'فاتح', themeDark: 'غامق',
      secDefaults: 'الإعدادات الافتراضية للتصدير',
      rowFullFmt: 'تنسيق التصدير الكامل', rowFullFmtDesc: 'التنسيق الافتراضي في تبويب "المحادثة الحالية".',
      rowArcFmt: 'تنسيق الأرشيف', rowArcFmtDesc: 'التنسيق الافتراضي في تبويب "الأرشيف".',
      rowZip: 'ضغط ملفات ZIP (deflate)', rowZipDesc: 'ضغط ملفات التصدير لتقليل الحجم.',
      rowInternal: 'تضمين الرسائل الداخلية', rowInternalDesc: 'الإبقاء على رسائل التقدّم والأدوات.',
      secPerformance: 'الأداء',
      rowFetchDelay: 'التأخير بين الطلبات (ملي ثانية)', rowFetchDelayDesc: 'مدة الانتظار بين الطلبات أثناء التصدير الدفعي.',
      rowPageSize: 'صفوف في كل صفحة', rowPageSizeDesc: 'عدد المحادثات المعروضة في كل صفحة بالأرشيف.',
      rowMaxChats: 'الحد الافتراضي للمحادثات', rowMaxChatsDesc: 'العدد الافتراضي عند تحميل الأرشيف.',
      rowToastDur: 'مدة الإشعارات (ملي ثانية)', rowToastDurDesc: 'المدة التي تظهر فيها التنبيهات.',
      secData: 'البيانات', resetBtn: 'استعادة الإعدادات الافتراضية',
      resetConfirmT: 'استعادة الإعدادات الافتراضية؟', resetConfirmB: 'ستعود جميع الإعدادات إلى قيمها الافتراضية.',
      settingsSaved: 'تم حفظ الإعدادات', settingsReset: 'تمت استعادة الإعدادات الافتراضية',
      langEn: 'English', langAr: 'العربية', langZh: '中文',
      errEncCookieMissing: 'ملف تعريف التشفير مفقود. أعد تحميل الصفحة.',
      errEncCookieBad: 'ملف تعريف التشفير غير صالح.',
      errMsalKey: 'مفتاح MSAL مفقود.',
      errNoToken: 'لم يُعثر على substrate token صالح. أعد تحميل الصفحة.',
      errBadConvId: 'معرّف محادثة غير صالح.',
      errBadJson: 'الاستجابة ليست JSON صالحاً.',
      errConvShape: 'بنية محادثة غير متوقعة.',
      errConvMismatch: 'عدم تطابق معرّف المحادثة.',
      errNoPageMsgs: 'لا توجد رسائل ظاهرة في الصفحة.',
      errAuthExpired: 'انتهت صلاحية المصادقة. أعد تحميل الصفحة وحاول مجدداً.',
      errTokenUnchanged: 'لم يتغير الرمز بعد التحديث — أعد تحميل الصفحة.'
    },
    zh: {
      appTagline: 'v{v} · 实时获取 · 速率限制感知',
      tabCurrent: '当前对话', tabArchive: '存档', tabLogs: '日志', tabSettings: '设置',
      close: '关闭', cancel: '取消', confirm: '确认',
      copy: '复制', clear: '清空',
      logAll: '全部', logInfo: '信息', logWarn: '警告', logError: '错误',
      loading: '加载中…',
      noLogsYet: '暂无日志', noLogsLevel: '暂无{level}日志',
      copied: '已复制', copyFailed: '复制失败', cancelled: '已取消',
      waitOrCancel: '请先等待当前操作完成或取消',
      userMessages: '用户消息', copilotReplies: 'Copilot 回复', includeInternal: '包含内部消息',
      copyPageBtn: '复制页面消息', dlPageBtn: '下载页面（{fmt}）',
      fullExportLabel: '完整导出：', exportFromMs: '从 Microsoft 导出',
      currentHelp: '页面方式仅读取可见 DOM（长对话不完整）；完整导出始终从 Microsoft API 获取最新历史。',
      fmtMd: 'Markdown', fmtClean: '精简 JSON', fmtGpt: 'ChatGPT JSON', fmtRaw: '原始 JSON（⚠ 敏感）', fmtZip: 'ZIP（推荐）',
      openChatFirst: '请先打开一个 Copilot 对话。', openChatFirstShort: '请先打开对话',
      visibleMsgs: '可见消息 {n} 条', cachedHistory: '已缓存完整历史（{n} 条）',
      notCapturedYet: '尚未捕获完整历史',
      partialVisible: '仅显示 {a}/{b} 条 — 请使用完整导出',
      idShort: 'ID：{id}',
      extracting: '正在提取页面消息…',
      copiedN: '已复制 {n} 条消息', downloadedN: '已下载 {n} 条消息',
      partialNote: '部分：{a}/{b}',
      partialLong: '部分导出：共 {b} 条消息，当前仅 {a} 条。请使用完整导出获取全部历史。',
      fetchLatest: '获取最近', load: '加载',
      colName: '名称', colCreated: '创建时间', colUpdated: '更新时间', colStatus: '状态',
      searchPh: '搜索…', sortLabel: '排序',
      sortUpdated: '按更新时间', sortCreated: '按创建时间', sortName: '按名称 A–Z',
      selectAll: '全选', format: '格式', exportSelected: '导出所选',
      archiveNote: '存档使用 Microsoft 内部 Substrate API，可能随时变更。不包含删除操作。',
      noConvs: '尚未加载对话。请点击“加载”。',
      noMatch: '没有符合条件的对话。',
      untitled: '无标题', stExporting: '导出中', stDone: '完成', stError: '错误',
      loadingArchive: '正在加载存档列表…', loadedN: '已加载 {n} 个对话',
      startingExport: '正在开始导出…', starting: '开始…', allFailed: '所有导出均失败。',
      buildingZip: '正在构建 {mb} MB 的 ZIP 文件，可能需要一些时间。',
      exportedPart: '已导出 {a}/{b}（{c} 个失败）', exportedNConvs: '已导出 {n} 个对话',
      exportedNMsgs: '已导出 {n} 条消息', cachedNote: '（缓存）',
      rawConfirmT: '导出原始 JSON？',
      rawConfirmBodyFull: '原始 JSON 可能包含位置数据、用户 ID、租户 ID 和遥测数据。仅限私人备份。',
      rawConfirmBodyArc: '原始 JSON 可能包含敏感元数据。仅限私人备份。',
      exportRaw: '导出原始数据',
      preparingFull: '正在准备完整导出…', fetchingMs: '正在从 Microsoft 获取…',
      usingCache: '使用缓存副本（实时抓取失败）',
      launchTitle: '打开 Copilot 导出器',
      secGeneral: '常规',
      rowLanguage: '界面语言', rowLanguageDesc: '选择界面显示语言。',
      rowTheme: '主题', rowThemeDesc: '跟随系统或强制浅色/深色。',
      themeAuto: '自动（跟随系统）', themeLight: '浅色', themeDark: '深色',
      secDefaults: '导出默认设置',
      rowFullFmt: '完整导出格式', rowFullFmtDesc: '“当前对话”标签页的默认格式。',
      rowArcFmt: '存档格式', rowArcFmtDesc: '“存档”标签页的默认格式。',
      rowZip: '压缩 ZIP（deflate）', rowZipDesc: '使用 deflate 压缩导出文件以减小体积。',
      rowInternal: '包含内部消息', rowInternalDesc: '保留进度和工具类消息。',
      secPerformance: '性能',
      rowFetchDelay: '请求间隔（毫秒）', rowFetchDelayDesc: '批量导出时相邻请求的等待时间。',
      rowPageSize: '每页行数', rowPageSizeDesc: '存档列表每页显示的对话数量。',
      rowMaxChats: '默认对话上限', rowMaxChatsDesc: '加载存档时的默认数量。',
      rowToastDur: '提示持续时间（毫秒）', rowToastDurDesc: '通知显示的时长。',
      secData: '数据', resetBtn: '恢复默认设置',
      resetConfirmT: '重置设置？', resetConfirmB: '所有设置将恢复为默认值。',
      settingsSaved: '设置已保存', settingsReset: '设置已恢复为默认值',
      langEn: 'English', langAr: 'العربية', langZh: '中文',
      errEncCookieMissing: '缺少加密 Cookie。请重新加载页面。',
      errEncCookieBad: '加密 Cookie 无效。',
      errMsalKey: '缺少 MSAL 密钥。',
      errNoToken: '未找到有效的 substrate 令牌。请重新加载页面。',
      errBadConvId: '对话 ID 无效。',
      errBadJson: '响应不是有效的 JSON。',
      errConvShape: '对话结构异常。',
      errConvMismatch: '对话 ID 不匹配。',
      errNoPageMsgs: '页面上没有可见消息。',
      errAuthExpired: '认证已过期。请重新加载页面后重试。',
      errTokenUnchanged: '刷新后令牌未变化 — 请重新加载页面。'
    }
  };

  function t(key, params, lang) {
    var dict = STR[lang || settings.lang] || STR.en;
    var s = dict[key];
    if (s == null) s = STR.en[key];
    if (s == null) return key;
    if (params) {
      s = s.replace(/\{(\w+)\}/g, function (m, k) {
        return params[k] != null ? String(params[k]) : m;
      });
    }
    return s;
  }

  /** Localized error: message stays English (for logs), toast translates via dvT/dvP. */
  function locErr(key, params) {
    var e = new Error(t(key, params, 'en'));
    e.dvT = key;
    e.dvP = params || null;
    return e;
  }
  function locMsg(e) {
    if (e && e.dvT) return t(e.dvT, e.dvP);
    return e && e.message ? e.message : String(e);
  }

  /* ═══════ LOGGING ═══════ */
  var lastConversation = null;
  var lastConversationId = null;
  var capturedAuth = null; // {token, claims} captured via network interception
  var chatCache = [];
  var activeAbort = null;
  var activeModal = null;
  var isBusy = false;

  var ceLogs = [];
  function ceLog(level) {
    var args = Array.prototype.slice.call(arguments, 1);
    var msg = args.map(function (a) {
      try {
        if (a instanceof Error) return a.stack || a.message;
        if (typeof a === 'string') return a;
        return JSON.stringify(a);
      } catch (e) { return String(a); }
    }).join(' ');
    ceLogs.push({ t: Date.now(), level: level, msg: msg.slice(0, 1500) });
    if (ceLogs.length > MAX_LOG_ENTRIES) ceLogs.shift();
    var fn = level === 'error' ? console.error : level === 'warn' ? console.warn : console.log;
    fn.apply(console, ['[' + APP + ']'].concat(args));
  }

  var sleep = function (ms) { return new Promise(function (r) { setTimeout(r, ms); }); };

  var _escDiv = document.createElement('div');
  function esc(s) { _escDiv.textContent = s == null ? '' : String(s); return _escDiv.innerHTML; }

  function el(tag, attrs) {
    var node = document.createElement(tag);
    if (attrs) {
      Object.keys(attrs).forEach(function (k) {
        var v = attrs[k];
        if (v == null || v === false) return;
        if (k === 'text') node.textContent = String(v);
        else if (k === 'class') node.className = v;
        else if (k === 'dataset') Object.keys(v).forEach(function (d) { node.dataset[d] = v[d]; });
        else if (k.slice(0, 2) === 'on' && typeof v === 'function') node.addEventListener(k.slice(2).toLowerCase(), v);
        else if (v === true) node.setAttribute(k, '');
        else node.setAttribute(k, String(v));
      });
    }
    for (var i = 2; i < arguments.length; i++) {
      var c = arguments[i];
      if (c == null || c === false) continue;
      if (Array.isArray(c)) {
        for (var j = 0; j < c.length; j++) if (c[j] != null) node.appendChild(c[j].nodeType ? c[j] : document.createTextNode(String(c[j])));
      } else node.appendChild(c.nodeType ? c : document.createTextNode(String(c)));
    }
    return node;
  }

  /* ═══════ THEME ═══════ */
  function isDarkMode() {
    if (settings.theme === 'dark') return true;
    if (settings.theme === 'light') return false;
    try { return matchMedia('(prefers-color-scheme: dark)').matches; } catch (e) { return false; }
  }
  function applyModalTheme() {
    if (activeModal) activeModal.classList.toggle('dv-dark', isDarkMode());
    var tc = document.getElementById('dv64-toasts');
    if (tc) tc.classList.toggle('dv-dark', isDarkMode());
  }
  try {
    matchMedia('(prefers-color-scheme: dark)').addEventListener('change', function () { applyModalTheme(); });
  } catch (e) { /* matchMedia listener unavailable */ }

  /* ═══════ CSS (light/dark via variables, RTL support) ═══════ */
  var CSS = [
    '#dv64-overlay{--bg:#fff;--fg:#0f172a;--line:rgba(0,0,0,.07);--line-soft:rgba(0,0,0,.05);--line-faint:rgba(0,0,0,.04);',
    '--hov:rgba(0,0,0,.06);--hov-soft:rgba(0,0,0,.05);--chip:rgba(0,0,0,.06);--tint:rgba(0,0,0,.02);--tint2:rgba(0,0,0,.05);',
    '--input-bg:#fff;--input-br:rgba(0,0,0,.12);--cb-br:rgba(0,0,0,.25);--scroll:rgba(0,0,0,.15);',
    '--shadow:0 24px 64px rgba(0,0,0,.35),0 0 0 1px rgba(0,0,0,.05);',
    'position:fixed;inset:0;z-index:2147483646;display:flex;align-items:center;justify-content:center;padding:20px;',
    'background:rgba(10,10,15,.55);backdrop-filter:blur(10px);-webkit-backdrop-filter:blur(10px);}',
    '#dv64-overlay.dv-dark{--bg:#131316;--fg:#e4e4e7;--line:rgba(255,255,255,.08);--line-soft:rgba(255,255,255,.05);--line-faint:rgba(255,255,255,.04);',
    '--hov:rgba(255,255,255,.08);--hov-soft:rgba(255,255,255,.06);--chip:rgba(255,255,255,.08);--tint:rgba(255,255,255,.02);--tint2:rgba(255,255,255,.05);',
    '--input-bg:#1a1a1e;--input-br:rgba(255,255,255,.1);--cb-br:rgba(255,255,255,.3);--scroll:rgba(255,255,255,.15);',
    '--shadow:0 24px 64px rgba(0,0,0,.7),0 0 0 1px rgba(255,255,255,.06);}',
    '#dv64-overlay.rtl{direction:rtl;}',
    '#dv64-overlay.rtl .dv-tbl thead th{text-align:right;}',
    '#dv64-overlay.rtl .dv-tbl thead th:first-child{padding-left:9px;padding-right:16px;}',
    '#dv64-overlay.rtl .dv-tbl tbody td:first-child{padding-left:12px;padding-right:16px;}',
    '#dv64-modal{width:100%;max-width:940px;max-height:90vh;display:flex;flex-direction:column;',
    'background:var(--bg);color:var(--fg);border-radius:14px;overflow:hidden;box-shadow:var(--shadow);',
    'font-family:system-ui,-apple-system,"Segoe UI","Noto Sans","Noto Sans Arabic","Microsoft YaHei",sans-serif;font-size:14px;line-height:1.5;}',
    '.dv-h{display:flex;align-items:center;justify-content:space-between;padding:16px 20px;border-bottom:1px solid var(--line);}',
    '.dv-h-l{display:flex;align-items:center;gap:10px;}',
    '.dv-h-ico{width:34px;height:34px;border-radius:9px;background:linear-gradient(135deg,#6366f1,#8b5cf6);display:flex;align-items:center;justify-content:center;color:#fff;font-size:17px;box-shadow:0 4px 14px rgba(99,102,241,.35);flex-shrink:0;}',
    '.dv-h h2{margin:0;font-size:16px;font-weight:650;letter-spacing:-.01em;}',
    '.dv-h p{margin:1px 0 0;font-size:11px;opacity:.62;}',
    '.dv-x{width:32px;height:32px;border-radius:7px;background:none;border:none;cursor:pointer;font-size:20px;line-height:1;color:inherit;opacity:.6;transition:.15s;}',
    '.dv-x:hover{opacity:1;background:var(--hov);}',
    '.dv-x:disabled{opacity:.3;cursor:not-allowed;}',
    '.dv-tabs{display:flex;padding:0 20px;border-bottom:1px solid var(--line);}',
    '.dv-tab{background:none;border:none;padding:10px 2px;margin:0 9px;font:inherit;font-size:13px;font-weight:550;cursor:pointer;color:inherit;opacity:.55;position:relative;transition:opacity .15s;display:flex;align-items:center;gap:6px;}',
    '.dv-tab:hover{opacity:.85;}',
    '.dv-tab.active{opacity:1;}',
    '.dv-tab.active::after{content:"";position:absolute;bottom:-1px;left:0;right:0;height:2px;background:linear-gradient(90deg,#6366f1,#8b5cf6);border-radius:2px 2px 0 0;}',
    '.dv-badge{font-size:10px;padding:1px 6px;border-radius:10px;background:rgba(239,68,68,.15);color:#ef4444;font-weight:650;}',
    '.dv-badge[data-n="0"]{display:none;}',
    '.dv-panel{display:none;flex-direction:column;overflow:hidden;flex:1;min-height:0;}',
    '.dv-panel.active{display:flex;}',
    '.dv-bar{display:flex;align-items:center;justify-content:space-between;gap:10px;padding:11px 20px;background:var(--tint);border-bottom:1px solid var(--line-soft);flex-wrap:wrap;}',
    '.dv-bar-l,.dv-bar-r{display:flex;align-items:center;gap:8px;flex-wrap:wrap;}',
    '.dv-chk{display:inline-flex;align-items:center;gap:6px;cursor:pointer;font-size:12px;padding:4px 8px;border-radius:5px;}',
    '.dv-chk:hover{background:var(--hov-soft);}',
    '.dv-count{font-size:11px;padding:2px 9px;border-radius:10px;background:var(--chip);opacity:.7;font-variant-numeric:tabular-nums;}',
    '.dv-input,.dv-select{height:30px;padding:0 9px;font:inherit;font-size:12px;color:inherit;background:var(--input-bg);border:1px solid var(--input-br);border-radius:6px;outline:none;}',
    '.dv-input:focus,.dv-select:focus{border-color:#6366f1;box-shadow:0 0 0 3px rgba(99,102,241,.15);}',
    '.dv-input[type=number]{width:72px;text-align:center;}',
    '.dv-input[type=search]{width:150px;}',
    '.dv-btn{display:inline-flex;align-items:center;justify-content:center;gap:5px;height:30px;padding:0 13px;font:inherit;font-size:12px;font-weight:550;border:1px solid transparent;border-radius:6px;cursor:pointer;transition:.15s;white-space:nowrap;}',
    '.dv-btn:disabled{opacity:.4;cursor:not-allowed;}',
    '.dv-btn-pri{background:linear-gradient(135deg,#6366f1,#8b5cf6);color:#fff;box-shadow:0 2px 8px rgba(99,102,241,.35);}',
    '.dv-btn-pri:hover:not(:disabled){box-shadow:0 4px 14px rgba(99,102,241,.5);}',
    '.dv-btn-gho{background:var(--tint2);color:inherit;}',
    '.dv-btn-gho:hover:not(:disabled){background:var(--hov);}',
    '.dv-btn-danger{background:rgba(239,68,68,.12);color:#ef4444;}',
    '.dv-btn-danger:hover:not(:disabled){background:rgba(239,68,68,.2);}',
    '.dv-tbl-wrap{flex:1;min-height:180px;overflow:auto;scrollbar-width:thin;}',
    '.dv-tbl-wrap::-webkit-scrollbar{width:8px;height:8px;}',
    '.dv-tbl-wrap::-webkit-scrollbar-thumb{background:var(--scroll);border-radius:4px;}',
    '.dv-tbl{width:100%;border-collapse:collapse;font-size:13px;table-layout:fixed;}',
    '.dv-tbl thead th{position:sticky;top:0;z-index:2;text-align:left;font-weight:550;font-size:10px;text-transform:uppercase;letter-spacing:.05em;opacity:.55;padding:8px 12px;background:var(--bg);border-bottom:1px solid var(--line);}',
    '.dv-tbl thead th:first-child{width:36px;padding-left:16px;}',
    '.dv-tbl thead th:nth-child(2){width:auto;}',
    '.dv-tbl thead th:nth-child(3),.dv-tbl thead th:nth-child(4){width:140px;}',
    '.dv-tbl thead th:last-child{width:100px;}',
    '.dv-tbl tbody tr{border-bottom:1px solid var(--line-faint);transition:background .12s;}',
    '.dv-tbl tbody tr:hover{background:var(--hov-soft);}',
    '.dv-tbl tbody td{padding:9px 12px;vertical-align:middle;}',
    '.dv-tbl tbody td:first-child{padding-left:16px;}',
    '.dv-name{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-weight:500;}',
    '.dv-date{font-size:11px;opacity:.65;font-variant-numeric:tabular-nums;}',
    '.dv-cb{appearance:none;-webkit-appearance:none;width:15px;height:15px;margin:0;border:1.5px solid var(--cb-br);border-radius:4px;background:transparent;cursor:pointer;transition:.15s;position:relative;}',
    '.dv-cb:hover{border-color:#6366f1;}',
    '.dv-cb:checked{background:#6366f1;border-color:#6366f1;}',
    '.dv-cb:checked::after{content:"\\2713";color:#fff;position:absolute;top:-2px;left:2px;font-size:11px;font-weight:700;}',
    '.dv-cb:indeterminate{background:#6366f1;border-color:#6366f1;}',
    '.dv-cb:indeterminate::after{content:"\\2212";color:#fff;position:absolute;top:-3px;left:3px;font-size:11px;font-weight:700;}',
    '.dv-st{display:inline-flex;align-items:center;gap:5px;font-size:11px;opacity:.85;}',
    '.dv-st::before{content:"";width:6px;height:6px;border-radius:50%;background:currentColor;}',
    '.dv-st.empty{opacity:.3;}',
    '.dv-st[data-s=exp]{color:#f59e0b;}',
    '.dv-st[data-s=done]{color:#10b981;}',
    '.dv-st[data-s=err]{color:#ef4444;}',
    '.dv-empty{display:flex;flex-direction:column;align-items:center;justify-content:center;gap:10px;padding:48px 20px;text-align:center;opacity:.55;font-size:13px;}',
    '.dv-empty .ico{font-size:32px;}',
    '.dv-warn{padding:9px 20px;background:rgba(245,158,11,.1);color:#d97706;font-size:12px;border-bottom:1px solid rgba(245,158,11,.2);line-height:1.45;}',
    '#dv64-overlay.dv-dark .dv-warn{background:rgba(245,158,11,.12);color:#fbbf24;border-color:rgba(245,158,11,.25);}',
    '.dv-warn.danger{background:rgba(239,68,68,.1);color:#dc2626;border-color:rgba(239,68,68,.2);}',
    '#dv64-overlay.dv-dark .dv-warn.danger{background:rgba(239,68,68,.12);color:#f87171;}',
    '.dv-page{display:flex;align-items:center;justify-content:center;gap:8px;padding:9px;border-top:1px solid var(--line-soft);font-size:12px;}',
    '.dv-page:empty{display:none;}',
    '.dv-page button{width:28px;height:28px;border-radius:6px;border:1px solid var(--input-br);background:transparent;cursor:pointer;font:inherit;color:inherit;}',
    '.dv-page button:disabled{opacity:.3;cursor:not-allowed;}',
    '.dv-foot{display:flex;align-items:center;justify-content:space-between;gap:10px;padding:12px 20px;border-top:1px solid var(--line);background:var(--tint);flex-wrap:wrap;}',
    '.dv-foot-l,.dv-foot-r{display:flex;align-items:center;gap:6px;flex-wrap:wrap;}',
    '.dv-lbl{font-size:11px;opacity:.65;font-weight:550;}',
    '.dv-prog{padding:0 20px 11px;}',
    '.dv-prog-lbl{display:flex;justify-content:space-between;font-size:11px;opacity:.75;margin-bottom:6px;}',
    '.dv-prog-tr{height:4px;background:var(--chip);border-radius:2px;overflow:hidden;}',
    '.dv-prog-fl{height:100%;background:linear-gradient(90deg,#6366f1,#8b5cf6);border-radius:2px;transition:width .3s;}',
    '.dv-logs-bar{display:flex;align-items:center;gap:6px;padding:11px 20px;border-bottom:1px solid var(--line-soft);background:var(--tint);flex-wrap:wrap;}',
    '.dv-logs{flex:1;overflow:auto;font-family:ui-monospace,SFMono-Regular,Menlo,Consolas,monospace;font-size:11.5px;line-height:1.55;scrollbar-width:thin;}',
    '.dv-logs::-webkit-scrollbar{width:8px;}.dv-logs::-webkit-scrollbar-thumb{background:var(--scroll);border-radius:4px;}',
    '.dv-log{display:flex;gap:8px;padding:5px 20px;border-bottom:1px solid var(--line-faint);cursor:text;}',
    '.dv-log:hover{background:var(--hov-soft);}',
    '.dv-log-t{flex-shrink:0;opacity:.5;font-size:10.5px;}',
    '.dv-log-l{flex-shrink:0;font-weight:650;font-size:10px;padding:0 5px;border-radius:3px;text-transform:uppercase;line-height:1.6;}',
    '.dv-log-l.info{background:rgba(99,102,241,.12);color:#6366f1;}',
    '.dv-log-l.warn{background:rgba(245,158,11,.12);color:#f59e0b;}',
    '.dv-log-l.error{background:rgba(239,68,68,.15);color:#ef4444;}',
    '.dv-log-m{flex:1;white-space:pre-wrap;word-break:break-word;}',
    '.dv-set{flex:1;overflow:auto;scrollbar-width:thin;}',
    '.dv-set::-webkit-scrollbar{width:8px;}.dv-set::-webkit-scrollbar-thumb{background:var(--scroll);border-radius:4px;}',
    '.dv-set-sec{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;opacity:.5;padding:16px 20px 6px;}',
    '.dv-set-sec:first-child{padding-top:12px;}',
    '.dv-set-row{display:flex;align-items:center;justify-content:space-between;gap:14px;padding:10px 20px;border-bottom:1px solid var(--line-faint);}',
    '.dv-set-l{min-width:0;}',
    '.dv-set-n{font-size:13px;font-weight:550;}',
    '.dv-set-d{font-size:11px;opacity:.6;margin-top:1px;line-height:1.4;}',
    '.dv-set-row .dv-input,.dv-set-row .dv-select{flex-shrink:0;}',
    '.dv-set-actions{padding:16px 20px 20px;}',
    '#dv64-launcher{position:fixed;right:20px;bottom:20px;width:50px;height:50px;border-radius:50%;border:none;cursor:pointer;z-index:2147483645;background:linear-gradient(135deg,#6366f1,#8b5cf6);color:#fff;font-size:22px;display:flex;align-items:center;justify-content:center;box-shadow:0 4px 16px rgba(99,102,241,.45),0 2px 6px rgba(0,0,0,.2);transition:transform .15s,box-shadow .15s;}',
    '#dv64-launcher:hover{transform:scale(1.08);box-shadow:0 6px 24px rgba(99,102,241,.6);}',
    '#dv64-launcher:active{transform:scale(.94);}',
    '#dv64-launcher:disabled{opacity:.6;cursor:not-allowed;transform:none;}',
    '#dv64-toasts{--tbg:#fff;--tfg:#0f172a;--tbr:rgba(0,0,0,.08);position:fixed;bottom:22px;right:22px;z-index:2147483647;display:flex;flex-direction:column;gap:6px;pointer-events:none;font-family:system-ui,-apple-system,"Segoe UI","Noto Sans","Noto Sans Arabic","Microsoft YaHei",sans-serif;}',
    '#dv64-toasts.dv-dark{--tbg:#1a1a1e;--tfg:#e4e4e7;--tbr:rgba(255,255,255,.1);}',
    '.dv-toast{min-width:220px;max-width:340px;padding:11px 15px;background:var(--tbg);color:var(--tfg);border:1px solid var(--tbr);border-radius:9px;font-size:12.5px;box-shadow:0 8px 24px rgba(0,0,0,.15);pointer-events:auto;animation:dv-toast-in .2s cubic-bezier(.16,1,.3,1);}',
    '.dv-toast[data-t=success]{border-left:3px solid #10b981;}',
    '.dv-toast[data-t=error]{border-left:3px solid #ef4444;}',
    '.dv-toast[data-t=warn]{border-left:3px solid #f59e0b;}',
    '.dv-toast[data-t=info]{border-left:3px solid #6366f1;}',
    '@keyframes dv-toast-in{from{opacity:0;transform:translateX(20px);}to{opacity:1;transform:translateX(0);}}'
  ].join('');

  var styleEl = document.createElement('style');
  styleEl.textContent = CSS;
  (document.head || document.documentElement).appendChild(styleEl);

  function toast(msg, type, dur) {
    type = type || 'info';
    dur = dur || settings.toastDurationMs;
    var c = document.getElementById('dv64-toasts');
    if (!c) { c = document.createElement('div'); c.id = 'dv64-toasts'; document.body.appendChild(c); }
    c.classList.toggle('dv-dark', isDarkMode());
    var node = el('div', { class: 'dv-toast', dataset: { t: type }, text: msg });
    c.appendChild(node);
    setTimeout(function () { node.remove(); }, dur);
  }

  function confirmDlg(title, msg, okText, isDanger) {
    return new Promise(function (resolve) {
      var wrap = el('div', { style: 'position:fixed;inset:0;z-index:2147483648;display:flex;align-items:center;justify-content:center;padding:20px;background:rgba(0,0,0,.55);backdrop-filter:blur(6px);' });
      var dark = isDarkMode();
      var bg = dark ? '#1a1a1e' : '#fff';
      var fg = dark ? '#e4e4e7' : '#0f172a';
      var muted = dark ? '#a1a1a5' : '#6b7280';
      var border = dark ? 'rgba(255,255,255,.1)' : 'rgba(0,0,0,.08)';
      var btnColor = isDanger ? '#ef4444' : '#6366f1';
      wrap.innerHTML = '<div style="max-width:420px;width:100%;background:' + bg + ';color:' + fg + ';border:1px solid ' + border + ';border-radius:14px;padding:22px;box-shadow:0 20px 48px rgba(0,0,0,.45);font-family:system-ui">'
        + '<h3 style="margin:0 0 8px;font-size:15px;font-weight:650">' + esc(title) + '</h3>'
        + '<p style="margin:0 0 18px;font-size:13px;color:' + muted + ';line-height:1.5">' + esc(msg) + '</p>'
        + '<div style="display:flex;gap:8px;justify-content:flex-end">'
        + '<button data-a="c" style="height:34px;padding:0 15px;font-size:13px;border-radius:7px;background:transparent;color:' + fg + ';border:1px solid ' + border + ';cursor:pointer">' + esc(t('cancel')) + '</button>'
        + '<button data-a="o" style="height:34px;padding:0 15px;font-size:13px;border-radius:7px;background:' + btnColor + ';color:#fff;border:none;cursor:pointer;font-weight:550">' + esc(okText || t('confirm')) + '</button>'
        + '</div></div>';
      document.body.appendChild(wrap);
      var done = false;
      var kh = function (e) { if (e.key === 'Escape') close(false); };
      function close(v) {
        if (done) return; done = true;
        document.removeEventListener('keydown', kh);
        wrap.remove(); resolve(v);
      }
      document.addEventListener('keydown', kh);
      wrap.querySelector('[data-a=c]').onclick = function () { close(false); };
      wrap.querySelector('[data-a=o]').onclick = function () { close(true); };
      wrap.addEventListener('click', function (e) { if (e.target === wrap) close(false); });
    });
  }

  /* ═══════ BASE64 / JWT ═══════ */
  function b64ToBytes(s) {
    var e = String(s || '').replace(/-/g, '+').replace(/_/g, '/');
    while (e.length % 4) e += '=';
    var bin = atob(e);
    var arr = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
    return arr;
  }
  function toAB(a) { return Uint8Array.from(a).buffer; }
  function decodeJwt(tok) {
    try {
      var p = String(tok).split('.')[1];
      return p ? JSON.parse(new TextDecoder().decode(b64ToBytes(p))) : null;
    } catch (e) { return null; }
  }

  /* ═══════ AUTH (interception first, MSAL decrypt fallback) ═══════ */
  function getCookie(k) {
    var m = document.cookie.match(new RegExp('(^|;)\\s*' + k + '\\s*=\\s*([^;]+)'));
    return m ? m.pop() : '';
  }

  function tokenClaimsOk(claims) {
    if (!claims || !claims.oid || !claims.tid) return false;
    var now = Math.floor(Date.now() / 1000);
    var exp = Number(claims.exp || 0);
    if (exp && exp < now + REFRESH_MARGIN_SEC) return false;
    var aud = claims.aud;
    var audList = Array.isArray(aud) ? aud : [aud];
    return audList.some(function (a) {
      var s = String(a || '').toLowerCase();
      return s.indexOf('substrate') !== -1 || s.indexOf('office.com') !== -1;
    });
  }

  /** Store a Bearer token observed on the wire if it targets Substrate and is fresh enough. */
  function considerAuthHeader(value) {
    var m = String(value || '').match(/^Bearer\s+(.+)$/i);
    if (!m) return;
    var token = m[1];
    if (token.length < 100) return;
    var claims = decodeJwt(token);
    if (!tokenClaimsOk(claims)) return;
    if (capturedAuth && capturedAuth.claims && Number(capturedAuth.claims.exp || 0) >= Number(claims.exp || 0)) return;
    capturedAuth = { token: token, claims: claims };
    ceLog('info', 'Auth token captured via network interception (tenant=' + String(claims.tid).slice(0, 8) + '\u2026)');
  }

  function extractAuthFromHeaders(h) {
    try {
      if (!h) return null;
      if (typeof Headers !== 'undefined' && h instanceof Headers) return h.get('authorization');
      if (Array.isArray(h)) {
        for (var i = 0; i < h.length; i++) {
          if (h[i] && String(h[i][0]).toLowerCase() === 'authorization') return h[i][1];
        }
        return null;
      }
      if (typeof h === 'object') {
        for (var k in h) {
          if (Object.prototype.hasOwnProperty.call(h, k) && k.toLowerCase() === 'authorization') return h[k];
        }
      }
    } catch (e) { /* malformed headers */ }
    return null;
  }

  async function getEncryptionKey() {
    var raw = decodeURIComponent(getCookie('msal.cache.encryption'));
    if (!raw) throw locErr('errEncCookieMissing');
    var parsed;
    try { parsed = JSON.parse(raw); } catch (e) { throw locErr('errEncCookieBad'); }
    if (!parsed.key) throw locErr('errMsalKey');
    var key = await crypto.subtle.importKey('raw', toAB(b64ToBytes(parsed.key)), 'HKDF', false, ['deriveKey']);
    return { key: key, id: parsed.id };
  }

  async function decryptMsal(baseKey, nonce, data) {
    var dk = await crypto.subtle.deriveKey(
      { name: 'HKDF', salt: toAB(b64ToBytes(nonce)), hash: 'SHA-256', info: new TextEncoder().encode(CLIENT_ID) },
      baseKey, { name: 'AES-GCM', length: 256 }, false, ['decrypt']
    );
    var plain = await crypto.subtle.decrypt({ name: 'AES-GCM', iv: new Uint8Array(12) }, dk, toAB(b64ToBytes(data)));
    return new TextDecoder().decode(plain);
  }

  async function getArchiveAuth() {
    // 1) Prefer a token observed on the wire — no decryption, always current.
    if (capturedAuth && tokenClaimsOk(capturedAuth.claims)) {
      ceLog('info', 'Using intercepted token (tenant=' + String(capturedAuth.claims.tid).slice(0, 8) + '\u2026)');
      return {
        token: capturedAuth.token,
        localAccountId: capturedAuth.claims.oid,
        tenantId: capturedAuth.claims.tid
      };
    }

    // 2) Fallback: decrypt the MSAL cache (may break if Microsoft changes the cache format).
    ceLog('info', 'Acquiring substrate token from MSAL cache\u2026');
    var enc = await getEncryptionKey();
    var keysRaw = localStorage.getItem('msal.3.token.keys.' + CLIENT_ID);
    var scopeKeys = [];
    if (keysRaw) {
      try {
        var tk = JSON.parse(keysRaw);
        if (tk && Array.isArray(tk.accessToken)) {
          scopeKeys = tk.accessToken.filter(function (k) { return /substrate|sydney/i.test(k); });
        }
      } catch (e) { /* fall through */ }
    }
    if (!scopeKeys.length) {
      for (var i = 0; i < localStorage.length; i++) {
        var k = localStorage.key(i);
        if (k && /msal.*token/i.test(k) && /substrate|sydney/i.test(k)) scopeKeys.push(k);
      }
    }

    var candidates = [];
    for (var s = 0; s < scopeKeys.length; s++) {
      try {
        var entry = JSON.parse(localStorage.getItem(scopeKeys[s]) || 'null');
        if (!entry || !entry.nonce || !entry.data) continue;
        var dec = JSON.parse(await decryptMsal(enc.key, entry.nonce, entry.data));
        var token = dec && dec.secret;
        if (typeof token !== 'string' || token.length < 100) continue;
        var claims = decodeJwt(token);
        if (!tokenClaimsOk(claims)) continue;
        var isSydney = /sydney/i.test(scopeKeys[s]) || /sydney/i.test(dec.target || '');
        candidates.push({ token: token, claims: claims, isSydney: isSydney });
      } catch (e) { continue; }
    }
    if (!candidates.length) throw locErr('errNoToken');

    // Prefer sydney-scope tokens
    candidates.sort(function (a, b) { return (b.isSydney ? 1 : 0) - (a.isSydney ? 1 : 0); });
    var chosen = candidates[0];
    ceLog('info', 'Token acquired (sydney=' + chosen.isSydney + ', tenant=' + chosen.claims.tid.slice(0, 8) + '\u2026)');
    return {
      token: chosen.token,
      localAccountId: chosen.claims.oid,
      tenantId: chosen.claims.tid
    };
  }

  /* ═══════ API ═══════ */
  function apiHeaders(auth) {
    return {
      authorization: 'Bearer ' + auth.token,
      'content-type': 'application/json',
      'x-anchormailbox': 'Oid:' + auth.localAccountId + '@' + auth.tenantId,
      'x-clientrequestid': crypto.randomUUID().replace(/-/g, ''),
      'x-routingparameter-sessionkey': auth.localAccountId,
      'x-scenario': 'OfficeWebIncludedCopilot'
    };
  }

  /**
   * authedFetch — merges headers, 429 backoff with Retry-After, 401 refresh
   */
  async function authedFetch(url, authRef, options) {
    options = options || {};
    var signal = options.signal;
    var buildInit = function (auth) {
      var h = Object.assign({}, apiHeaders(auth), options.headers || {});
      return Object.assign({}, options, { headers: h, signal: signal });
    };
    var attempt = function () { return fetch(url, buildInit(authRef.value)); };

    var res = await attempt();

    // Rate limit: 429 or 503 → respect Retry-After
    if (res.status === 429 || res.status === 503) {
      var ra = Number(res.headers.get('retry-after') || 3);
      ra = isFinite(ra) ? Math.min(ra, 30) : 3;
      ceLog('warn', 'Rate limited (' + res.status + '), waiting ' + ra + 's\u2026');
      await sleep(ra * 1000);
      res = await attempt();
    }

    // Auth expired: 401 → drop captured token, try refresh from storage
    if (res.status === 401) {
      ceLog('warn', '401 — attempting token refresh');
      capturedAuth = null;
      try {
        var oldToken = authRef.value.token;
        var newAuth = await getArchiveAuth();
        if (newAuth.token === oldToken) {
          throw locErr('errTokenUnchanged');
        }
        authRef.value = newAuth;
        ceLog('info', 'Token refreshed, retrying\u2026');
        res = await attempt();
      } catch (e) {
        ceLog('error', 'Auth refresh failed:', e.message);
        throw locErr('errAuthExpired');
      }
    }

    return res;
  }

  async function apiGetChatsMerged(authRef, max, signal) {
    var variants = 'feature.EnableLastMessageForGetChats,feature.EnableMRUAgents,feature.EnableHasLoopPages';
    async function fetchType(threadType) {
      var req = { source: 'officeweb', traceId: crypto.randomUUID(), threadType: threadType, MaxReturnedChatsCount: max };
      var url = SUBSTRATE + '/GetChats?request=' + encodeURIComponent(JSON.stringify(req)) + '&variants=' + encodeURIComponent(variants);
      try {
        var res = await authedFetch(url, authRef, { method: 'GET', headers: { 'x-variants': variants }, signal: signal });
        if (!res.ok) return [];
        var data = await res.json();
        return Array.isArray(data && data.chats) ? data.chats : (Array.isArray(data && data.value) ? data.value : []);
      } catch (e) {
        ceLog('warn', threadType + ' fetch failed:', e.message);
        return [];
      }
    }
    var results = await Promise.all([fetchType('webchat'), fetchType('bizchat')]);
    var merged = [];
    var seen = new Set();
    results.forEach(function (list) {
      list.forEach(function (chat) {
        if (chat && chat.conversationId && !seen.has(chat.conversationId)) {
          seen.add(chat.conversationId);
          merged.push(chat);
        }
      });
    });
    ceLog('info', 'Merged chats: ' + merged.length + ' (webchat=' + results[0].length + ', bizchat=' + results[1].length + ')');
    return merged;
  }

  async function apiGetConversation(authRef, conversationId, signal) {
    if (!conversationId || typeof conversationId !== 'string') {
      throw locErr('errBadConvId');
    }
    var req = { conversationId: conversationId, source: 'officeweb', traceId: crypto.randomUUID().replace(/-/g, '') };
    var url = SUBSTRATE + '/GetConversation?request=' + encodeURIComponent(JSON.stringify(req));
    var res = await authedFetch(url, authRef, { method: 'GET', signal: signal });
    if (!res.ok) throw new Error('GetConversation HTTP ' + res.status);
    var raw = await res.text();
    var conv;
    try { conv = JSON.parse(raw); }
    catch (e) { throw locErr('errBadJson'); }
    if (!conv || !conv.conversationId || !Array.isArray(conv.messages)) throw locErr('errConvShape');
    if (conv.conversationId !== conversationId) throw locErr('errConvMismatch');
    rememberConversation(conv);
    return conv;
  }

  /* ═══════ MESSAGE FILTERING ═══════ */
  var INTERNAL_TYPES = {
    'InternalSearchResult': 1, 'Internal': 1, 'InternalCodeGenerationTask': 1,
    'InternalComputerToolsRoomInfo': 1, 'AgentToAgentExecutionInfo': 1,
    'InvokeAction': 1, 'GenerateContentQuery': 1, 'Search': 1,
    'SearchResults': 1, 'Progress': 1, 'Canceled': 1, 'Cancelled': 1,
    'EndOfRequest': 1
  };
  var PROGRESS_RE = /^(searching|working|thinking|loading|processing|analyzing|retrieving|composing|drafting)\.{0,3}$/i;

  function isInternalMessage(m) {
    if (!m) return true;
    if (m.messageType && INTERNAL_TYPES[m.messageType]) return true;
    var text = String(m.text || '').trim();
    if (!text && !(m.adaptiveCards && m.adaptiveCards.length)) return true;
    if (text.length < 30 && PROGRESS_RE.test(text)) return true;
    return false;
  }

  function filterMessages(messages, includeInternal) {
    if (includeInternal) return messages.slice();
    return messages.filter(function (m) { return !isInternalMessage(m); });
  }

  /* ═══════ TEXT EXTRACTION ═══════ */
  var CIT_RE = /【([^】]+)】/g;

  function adaptiveText(m) {
    if (!m || !m.adaptiveCards) return '';
    var parts = [];
    m.adaptiveCards.forEach(function (card) {
      if (!card || !card.body) return;
      function walk(node, d) {
        if (!node || d > 6) return;
        if (Array.isArray(node)) { node.forEach(function (n) { walk(n, d + 1); }); return; }
        if (typeof node === 'object') {
          if (typeof node.text === 'string' && node.text.trim()) {
            var ty = node.type || '';
            if (!ty || ty === 'TextBlock' || ty === 'RichTextBlock') parts.push(node.text);
          }
          ['body', 'items', 'columns', 'facts'].forEach(function (k) { if (node[k]) walk(node[k], d + 1); });
        }
      }
      walk(card.body, 0);
    });
    return parts.join('\n');
  }

  /**
   * pickText — prefers text WITH citations over text without
   */
  function pickText(m) {
    var adText = adaptiveText(m).trim();
    var apiText = String(m.text || '').trim();
    if (!adText) return apiText;
    if (!apiText) return adText;
    var adHasCit = /【[^】]+】/.test(adText);
    var apiHasCit = /【[^】]+】/.test(apiText);
    if (adHasCit && !apiHasCit) return adText;
    if (apiHasCit && !adHasCit) return apiText;
    // Same citation status → prefer longer
    return adText.length > apiText.length ? adText : apiText;
  }

  function extractCitations(m, text) {
    var refs = (m && m.references) || {};
    var used = new Map();
    var counter = 0;
    var cites = [];
    var rewritten = String(text || '').replace(CIT_RE, function (match, key) {
      var r = refs[key];
      if (!r || !r.targetLink || typeof r.targetLink !== 'string' || !r.targetLink.trim()) {
        return match;
      }
      if (used.has(key)) return '[' + used.get(key) + ']';
      counter++;
      var meta = {};
      try { meta = JSON.parse((r.displayData && r.displayData.content) || '{}'); } catch (e) { /* */ }
      var label = meta.label || meta.Title || ('Source ' + counter);
      var title = meta.Title || meta.title || label;
      used.set(key, counter);
      cites.push({ num: counter, label: label, title: title, url: r.targetLink });
      return '[' + counter + ']';
    });
    return { text: rewritten, citations: cites };
  }

  /* ═══════ CONVERSATION NORMALIZATION ═══════ */
  function toEpochSeconds(v) {
    if (v == null || v === '') return null;
    var ms = typeof v === 'number' ? (v < 1e11 ? v * 1000 : v) : Date.parse(v);
    return isFinite(ms) ? ms / 1000 : null;
  }
  function toEpochMs(v) {
    var s = toEpochSeconds(v);
    return s === null ? 0 : s * 1000;
  }
  function toIso(v) {
    var s = toEpochSeconds(v);
    return s === null ? '' : new Date(s * 1000).toISOString();
  }

  function cleanTitle(str) {
    return String(str || '').replace(/\s+/g, ' ').trim();
  }

  function normalizeConversation(src, includeInternal) {
    var msgs = [];
    var usedIds = new Set();
    var list = filterMessages(src.messages || [], includeInternal);
    list.forEach(function (m, i) {
      var id = m.messageId || ('msg-' + i);
      while (usedIds.has(id)) id += '-' + i;
      usedIds.add(id);
      var rawText = pickText(m);
      var cit = extractCitations(m, rawText);
      msgs.push({
        id: id,
        role: m.author === 'user' ? 'user' : 'assistant',
        text: cit.text,
        citations: cit.citations,
        createdAt: m.createdAt || m.timestamp || null
      });
    });
    return {
      conversationId: src.conversationId,
      title: cleanTitle(src.chatName || src.title || 'Copilot conversation'),
      createdAt: toEpochSeconds(src.createTimeUtc),
      updatedAt: toEpochSeconds(src.updateTimeUtc),
      messages: msgs
    };
  }

  /* ═══════ EXPORT FORMATS ═══════ */
  function markdownFromConv(c) {
    var lines = ['# ' + c.title, ''];
    if (c.conversationId) {
      lines.push('- ID: ' + c.conversationId);
      // Use current origin, not hardcoded domain
      lines.push('- Source: ' + location.origin + '/chat/conversation/' + c.conversationId);
    }
    if (c.createdAt) lines.push('- Created: ' + new Date(c.createdAt * 1000).toISOString());
    if (c.updatedAt) lines.push('- Updated: ' + new Date(c.updatedAt * 1000).toISOString());
    lines.push('');
    c.messages.forEach(function (m) {
      var ts = m.createdAt ? ' \u00B7 ' + new Date(Date.parse(m.createdAt)).toLocaleString() : '';
      lines.push('## ' + (m.role === 'user' ? 'You' : 'Copilot') + ts, '', m.text, '');
      if (m.citations.length) {
        lines.push('**Sources:**');
        m.citations.forEach(function (ci) {
          lines.push('[' + ci.num + '] [' + ci.title + '](' + ci.url + ')');
        });
        lines.push('');
      }
      lines.push('---', '');
    });
    return lines.join('\n');
  }

  function gptFromConv(c) {
    var root = 'root-' + Math.random().toString(36).slice(2);
    var mapping = {};
    mapping[root] = { id: root, message: null, parent: null, children: [] };
    var parent = root, current = root;
    c.messages.forEach(function (m) {
      var node = {
        id: m.id, parent: parent, children: [],
        message: {
          id: m.id,
          author: { role: m.role, name: null, metadata: {} },
          create_time: m.createdAt ? Date.parse(m.createdAt) / 1000 : null,
          update_time: null,
          content: { content_type: 'text', parts: [m.text] },
          status: 'finished_successfully',
          end_turn: m.role === 'assistant',
          weight: 1, metadata: {}, recipient: 'all', channel: null
        }
      };
      mapping[m.id] = node;
      mapping[parent].children.push(m.id);
      parent = current = m.id;
    });
    return {
      title: c.title, create_time: c.createdAt, update_time: c.updatedAt,
      mapping: mapping, moderation_results: [], current_node: current,
      id: c.conversationId
    };
  }

  function cleanJsonFromConv(c) {
    return {
      title: c.title,
      conversationId: c.conversationId,
      createdAt: c.createdAt,
      updatedAt: c.updatedAt,
      messages: c.messages.map(function (m) {
        return { role: m.role, text: m.text, citations: m.citations, createdAt: m.createdAt };
      })
    };
  }

  function blobFromConv(c, fmt) {
    if (fmt === 'md') return new Blob([markdownFromConv(c)], { type: 'text/markdown' });
    if (fmt === 'gpt') return new Blob([JSON.stringify([gptFromConv(c)], null, 2)], { type: 'application/json' });
    if (fmt === 'clean') return new Blob([JSON.stringify(cleanJsonFromConv(c), null, 2)], { type: 'application/json' });
    return new Blob([JSON.stringify(c, null, 2)], { type: 'application/json' });
  }

  function safeName(s) {
    return cleanTitle(s || 'chat').replace(/[\\/:*?"<>|\x00-\x1f]/g, '_').slice(0, 100) || 'chat';
  }
  function fname(c, fmt) {
    var ext = fmt === 'md' ? 'md' : 'json';
    var id = String(c.conversationId || '').slice(0, 8);
    return 'copilot-' + safeName(c.title) + (id ? '-' + id : '') + '.' + ext;
  }

  function downloadBlob(blob, name) {
    var u = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = u; a.download = name; a.style.display = 'none';
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(function () { URL.revokeObjectURL(u); }, 30000);
  }
  function downloadText(text, name, mime) { downloadBlob(new Blob([text], { type: mime }), name); }

  /* ═══════ ZIP (deflate when available, STORE fallback, correct CRC32) ═══════ */
  var CRC_TABLE = (function () {
    var tab = new Uint32Array(256);
    for (var i = 0; i < 256; i++) {
      var c = i;
      for (var j = 0; j < 8; j++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      tab[i] = c >>> 0;
    }
    return tab;
  })();

  function crc32(bytes) {
    var c = 0xFFFFFFFF;
    for (var i = 0; i < bytes.length; i++) {
      c = ((c >>> 8) ^ CRC_TABLE[(c ^ bytes[i]) & 0xFF]) >>> 0;
    }
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  function deflateRaw(bytes) {
    return new Promise(function (resolve, reject) {
      try {
        var stream = new Blob([bytes]).stream().pipeThrough(new CompressionStream('deflate-raw'));
        new Response(stream).arrayBuffer().then(
          function (buf) { resolve(new Uint8Array(buf)); },
          reject
        );
      } catch (e) { reject(e); }
    });
  }

  async function buildZip(files, useDeflate) {
    var canDeflate = !!useDeflate && typeof CompressionStream !== 'undefined';
    var enc = new TextEncoder();
    var parts = [], central = [], offset = 0;
    var d = new Date();
    var dt = ((d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1)) & 0xFFFF;
    var dd = (((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate()) & 0xFFFF;

    for (var i = 0; i < files.length; i++) {
      var f = files[i];
      var nameBytes = enc.encode(f.name);
      var data = f.data;
      var crc = crc32(data);
      var method = 0, stored = data;
      if (canDeflate && data.length > 64) {
        try {
          var comp = await deflateRaw(data);
          if (comp.length < data.length) { stored = comp; method = 8; }
        } catch (e) { /* keep STORE for this entry */ }
      }
      var version = method === 8 ? 20 : 10;
      var lh = new Uint8Array(30 + nameBytes.length);
      var lv = new DataView(lh.buffer);
      lv.setUint32(0, 0x04034b50, true);
      lv.setUint16(4, version, true);
      lv.setUint16(6, 0x0800, true);
      lv.setUint16(8, method, true);
      lv.setUint16(10, dt, true);
      lv.setUint16(12, dd, true);
      lv.setUint32(14, crc, true);
      lv.setUint32(18, stored.length, true);
      lv.setUint32(22, data.length, true);
      lv.setUint16(26, nameBytes.length, true);
      lv.setUint16(28, 0, true);
      lh.set(nameBytes, 30);
      parts.push(lh, stored);
      central.push({ nameBytes: nameBytes, crc: crc, method: method, csize: stored.length, usize: data.length, version: version, offset: offset });
      offset += lh.length + stored.length;
      if (i % 10 === 0) await new Promise(function (r) { setTimeout(r, 0); });
    }

    var cdStart = offset;
    for (var ci = 0; ci < central.length; ci++) {
      var e = central[ci];
      var cd = new Uint8Array(46 + e.nameBytes.length);
      var cv = new DataView(cd.buffer);
      cv.setUint32(0, 0x02014b50, true);
      cv.setUint16(4, 20, true);
      cv.setUint16(6, e.version, true);
      cv.setUint16(8, 0x0800, true);
      cv.setUint16(10, e.method, true);
      cv.setUint16(12, dt, true);
      cv.setUint16(14, dd, true);
      cv.setUint32(16, e.crc, true);
      cv.setUint32(20, e.csize, true);
      cv.setUint32(24, e.usize, true);
      cv.setUint16(28, e.nameBytes.length, true);
      cv.setUint32(42, e.offset, true);
      cd.set(e.nameBytes, 46);
      parts.push(cd);
      offset += cd.length;
    }

    var eocd = new Uint8Array(22);
    var ev = new DataView(eocd.buffer);
    ev.setUint32(0, 0x06054b50, true);
    ev.setUint16(8, central.length, true);
    ev.setUint16(10, central.length, true);
    ev.setUint32(12, offset - cdStart, true);
    ev.setUint32(16, cdStart, true);
    parts.push(eocd);

    return new Blob(parts, { type: 'application/zip' });
  }

  /* ═══════ DOM EXTRACTION (with partial detection) ═══════ */
  function findButton(scope, re) {
    if (!scope) return null;
    var btns = scope.querySelectorAll('button,[role="button"]');
    for (var i = 0; i < btns.length; i++) {
      var b = btns[i];
      var lbl = ((b.getAttribute('aria-label') || '') + ' ' + (b.innerText || b.textContent || '')).trim();
      if (re.test(lbl)) return b;
    }
    return null;
  }

  function sanitizeContent(node) {
    var clone = node.cloneNode(true);
    var removeSel = 'button,[role="button"],script,style,iframe,svg,canvas,[aria-hidden="true"],[data-testid="CopyButtonTestId"],[data-testid*="feedback"],[data-testid*="Feedback"]';
    clone.querySelectorAll(removeSel).forEach(function (n) { n.remove(); });
    var allowed = { 'A':1,'B':1,'STRONG':1,'I':1,'EM':1,'S':1,'DEL':1,'U':1,'P':1,'BR':1,'HR':1,
      'H1':1,'H2':1,'H3':1,'H4':1,'H5':1,'H6':1,'DIV':1,'SPAN':1,'SECTION':1,'ARTICLE':1,
      'UL':1,'OL':1,'LI':1,'PRE':1,'CODE':1,'BLOCKQUOTE':1,'TABLE':1,'THEAD':1,'TBODY':1,
      'TR':1,'TH':1,'TD':1,'SUB':1,'SUP':1,'IMG':1 };
    function clean(nodeEl) {
      if (!allowed[nodeEl.tagName]) {
        var p = nodeEl.parentNode;
        var kids = Array.prototype.slice.call(nodeEl.children);
        while (nodeEl.firstChild) p.insertBefore(nodeEl.firstChild, nodeEl);
        nodeEl.remove();
        kids.forEach(clean);
        return;
      }
      // Preserve language class for code blocks
      var preservedClass = null;
      if (nodeEl.tagName === 'PRE' || nodeEl.tagName === 'CODE') {
        var cls = nodeEl.getAttribute('class');
        if (cls) {
          var m = cls.match(/language-[\w+-]+/);
          if (m) preservedClass = m[0];
        }
      }
      var href = nodeEl.tagName === 'A' ? nodeEl.getAttribute('href') : null;
      Array.prototype.slice.call(nodeEl.attributes).forEach(function (a) { nodeEl.removeAttribute(a.name); });
      if (preservedClass) nodeEl.setAttribute('class', preservedClass);
      if (nodeEl.tagName === 'A' && href) {
        try {
          var parsed = new URL(href, location.href);
          if (['http:', 'https:', 'mailto:'].indexOf(parsed.protocol) !== -1) {
            nodeEl.setAttribute('href', parsed.href);
            nodeEl.setAttribute('rel', 'noopener noreferrer');
          }
        } catch (e) { /* */ }
      }
      if (nodeEl.tagName === 'IMG') {
        var src = nodeEl.getAttribute('src') || '';
        var alt = nodeEl.getAttribute('alt') || '';
        nodeEl.removeAttribute('src');
        if (alt) nodeEl.setAttribute('alt', alt);
        nodeEl.setAttribute('data-had-src', src ? '1' : '0');
      }
      Array.prototype.slice.call(nodeEl.children).forEach(clean);
    }
    Array.prototype.slice.call(clone.children).forEach(clean);
    return clone;
  }

  function domToMarkdown(root) {
    function block(txt) { return txt.trim() ? '\n\n' + txt.trim() + '\n\n' : ''; }
    function render(node) {
      if (node.nodeType === Node.TEXT_NODE) return (node.nodeValue || '').replace(/\r/g, '').replace(/[\t ]+/g, ' ');
      if (node.nodeType !== Node.ELEMENT_NODE) return '';
      var tag = node.tagName;
      if (['SCRIPT','STYLE','BUTTON','SVG','IFRAME'].indexOf(tag) !== -1) return '';
      if (tag === 'BR') return '\n';
      if (tag === 'HR') return '\n\n---\n\n';
      if (tag === 'IMG') {
        var alt = node.getAttribute('alt');
        return alt ? '[Image: ' + alt + ']' : '[Image]';
      }
      if (tag === 'PRE') {
        var code = (node.innerText || node.textContent || '').replace(/\n+$/, '');
        var cls = node.getAttribute('class') || '';
        var m = cls.match(/language-([\w+-]+)/);
        var lang = m ? m[1] : '';
        var fence = '```';
        while (code.indexOf(fence) !== -1) fence += '`';
        return block(fence + lang + '\n' + code + '\n' + fence);
      }
      if (tag === 'CODE') {
        var v = (node.textContent || '').replace(/`/g, '\\`');
        return '`' + v + '`';
      }
      if (/^H[1-6]$/.test(tag)) {
        var lvl = Number(tag.slice(1));
        var hc = '';
        Array.prototype.forEach.call(node.childNodes, function (ch) { hc += render(ch); });
        return block(Array(lvl + 1).join('#') + ' ' + hc.trim());
      }
      if (tag === 'A') {
        var label = '';
        Array.prototype.forEach.call(node.childNodes, function (ch) { label += render(ch); });
        var href = node.getAttribute('href');
        return href ? '[' + (label.trim() || href) + '](' + href + ')' : label;
      }
      if (tag === 'STRONG' || tag === 'B') {
        var s = '';
        Array.prototype.forEach.call(node.childNodes, function (ch) { s += render(ch); });
        return '**' + s + '**';
      }
      if (tag === 'EM' || tag === 'I') {
        var e2 = '';
        Array.prototype.forEach.call(node.childNodes, function (ch) { e2 += render(ch); });
        return '*' + e2 + '*';
      }
      if (tag === 'DEL' || tag === 'S') {
        var del = '';
        Array.prototype.forEach.call(node.childNodes, function (ch) { del += render(ch); });
        return '~~' + del + '~~';
      }
      if (tag === 'BLOCKQUOTE') {
        var bq = '';
        Array.prototype.forEach.call(node.childNodes, function (ch) { bq += render(ch); });
        return block(bq.split('\n').map(function (l) { return '> ' + l; }).join('\n'));
      }
      if (tag === 'LI') {
        var parent = node.parentElement;
        var sibs = parent ? Array.prototype.filter.call(parent.children, function (x) { return x.tagName === 'LI'; }) : [node];
        var idx = sibs.indexOf(node);
        var prefix = parent && parent.tagName === 'OL' ? (idx + 1) + '. ' : '- ';
        var c2 = '';
        Array.prototype.forEach.call(node.childNodes, function (ch) { c2 += render(ch); });
        return prefix + c2.trim().replace(/\n/g, '\n  ') + '\n';
      }
      if (tag === 'UL' || tag === 'OL') {
        var ulc = '';
        Array.prototype.filter.call(node.children, function (x) { return x.tagName === 'LI'; }).forEach(function (li) { ulc += render(li); });
        return block(ulc);
      }
      if (tag === 'TABLE') {
        var rows = Array.prototype.map.call(node.querySelectorAll('tr'), function (row) {
          var cells = Array.prototype.filter.call(row.children, function (x) { return x.tagName === 'TH' || x.tagName === 'TD'; });
          return '| ' + cells.map(function (cell) {
            var ct = '';
            Array.prototype.forEach.call(cell.childNodes, function (ch) { ct += render(ch); });
            return ct.replace(/\|/g, '\\|').trim();
          }).join(' | ') + ' |';
        });
        if (rows.length > 1) {
          var cols = rows[0].split('|').length - 2;
          var sep = '|';
          for (var k = 0; k < cols; k++) sep += ' --- |';
          rows.splice(1, 0, sep);
        }
        return block(rows.join('\n'));
      }
      var content = '';
      Array.prototype.forEach.call(node.childNodes, function (ch) { content += render(ch); });
      if (['P','DIV','SECTION','ARTICLE','TR'].indexOf(tag) !== -1) return block(content);
      return content;
    }
    var out = '';
    Array.prototype.forEach.call(root.childNodes, function (n) { out += render(n); });
    return out.replace(/[\t ]+\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim();
  }

  function extractMessageTimestamp(wrapper) {
    var timeEl = wrapper.querySelector('time[datetime]');
    if (timeEl) {
      var dt = timeEl.getAttribute('datetime');
      if (dt && !isNaN(Date.parse(dt))) return dt;
    }
    return null;
  }

  async function extractCurrentConversation() {
    var wrappers = document.querySelectorAll(CHAT_WRAPPER);
    if (!wrappers.length) throw locErr('errNoPageMsgs');
    var messages = [];
    for (var i = 0; i < wrappers.length; i++) {
      var w = wrappers[i];
      var ts = extractMessageTimestamp(w);
      var pScope = w.querySelector('[data-testid="chatQuestion"]') || w;
      var pNode = pScope.querySelector('[data-testid="chatOutput"]') || pScope;
      var aScope = w.querySelector('[data-testid="copilot-message-div"]') || w.querySelector('[data-testid="copilot-message-reply-div"]') || w;
      var aNode = aScope.querySelector('[data-testid="markdown-reply"]') || aScope.querySelector('[data-testid="copilot-message-reply-div"]');

      if (pNode) {
        var more = findButton(pNode, /^(see more|show more|المزيد|查看更多)$/i);
        if (more) { try { more.click(); await sleep(90); } catch (e) { /* */ } }
        var clean = sanitizeContent(pNode);
        messages.push({ role: 'user', text: (pNode.innerText || '').trim(), html: clean.innerHTML, markdown: domToMarkdown(clean), timestamp: ts });
        var less = findButton(pNode, /^(see less|show less|显示更少|收起)$/i);
        if (less) { try { less.click(); await sleep(30); } catch (e) { /* */ } }
      }
      if (aNode) {
        var aMore = findButton(aNode, /^(see more|show more|المزيد|查看更多)$/i);
        if (aMore) { try { aMore.click(); await sleep(90); } catch (e) { /* */ } }
        var aClean = sanitizeContent(aNode);
        messages.push({ role: 'assistant', text: (aNode.innerText || '').trim(), html: aClean.innerHTML, markdown: domToMarkdown(aClean), timestamp: ts });
        var aLess = findButton(aNode, /^(see less|show less|显示更少|收起)$/i);
        if (aLess) { try { aLess.click(); await sleep(30); } catch (e) { /* */ } }
      }
    }
    // Drop only CONSECUTIVE duplicates — repeated-but-separated messages are legitimate.
    messages = messages.filter(function (m, i, arr) {
      if (i > 0 && arr[i - 1].role === m.role && arr[i - 1].text === m.text) return false;
      return true;
    });

    // Detect partial (virtualized) view
    var reportedSize = null;
    var indices = [];
    for (var wi = 0; wi < wrappers.length; wi++) {
      var sz = Number(wrappers[wi].getAttribute('aria-setsize'));
      if (Number.isFinite(sz) && sz > 0 && (reportedSize === null || sz > reportedSize)) reportedSize = sz;
      var idx = Number(wrappers[wi].getAttribute('data-message-index'));
      if (Number.isFinite(idx)) indices.push(idx);
    }
    var firstIndex = indices.length ? Math.min.apply(null, indices) : 0;
    var lastIndex = indices.length ? Math.max.apply(null, indices) : (wrappers.length - 1);
    var partial = firstIndex > 0 || (reportedSize !== null && reportedSize > wrappers.length);

    return {
      title: safeTitle(),
      conversationId: currentConversationId(),
      url: location.href,
      extractedAt: new Date().toISOString(),
      messages: messages,
      partial: partial,
      reportedCount: reportedSize,
      visibleCount: wrappers.length,
      firstIndex: firstIndex,
      lastIndex: lastIndex
    };
  }

  function safeTitle() {
    var h1 = document.querySelector('main h1, [role="main"] h1, h1');
    var fromH = h1 && h1.innerText && h1.innerText.trim();
    var fromT = document.title.replace(/\s*[|–-]\s*(Microsoft 365 Copilot|Microsoft Copilot|Copilot).*$/i, '').trim();
    return cleanTitle((fromH && fromH.length < 180 ? fromH : fromT) || 'Copilot conversation');
  }
  function currentConversationId() {
    try {
      var u = new URL(location.href);
      var m = u.pathname.match(/(?:^|\/)conversation\/([0-9a-f-]{20,})(?:$|\/)/i);
      if (m) return m[1];
      var uuid = u.href.match(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i);
      return uuid ? uuid[0] : (lastConversationId || null);
    } catch (e) { return null; }
  }

  /* ═══════ NETWORK CAPTURE (conversation bodies + auth headers) ═══════ */
  function isMsApiUrl(u) {
    try {
      var p = new URL(String(u || ''), location.href);
      var h = p.hostname.toLowerCase();
      var allowed = h === 'substrate.office.com' || h === 'sydney.bing.com' || h === location.hostname.toLowerCase();
      return allowed && /substrate|sydney|m365copilot/i.test(p.pathname);
    } catch (e) { return false; }
  }
  function isConvUrl(u) {
    try {
      var p = new URL(String(u || ''), location.href);
      var h = p.hostname.toLowerCase();
      var allowed = h === 'substrate.office.com' || h === location.hostname.toLowerCase();
      var path = p.pathname.replace(/\/{2,}/g, '/');
      return allowed && /\/(?:api\/proxy\/(?:substrate|sydney)\/)?m365Copilot\/GetConversation\/?$/i.test(path);
    } catch (e) { return false; }
  }
  function rememberConversation(conv) {
    if (!conv || !conv.conversationId || !Array.isArray(conv.messages)) return;
    lastConversation = conv;
    lastConversationId = conv.conversationId;
    ceLog('info', 'Captured conversation (' + conv.messages.length + ' messages)');
  }
  function installPassiveCapture() {
    try {
      var orig = window.fetch;
      if (orig && !orig.__dv64Wrapped) {
        var wrapped = function () {
          var args = arguments;
          var url = typeof args[0] === 'string' ? args[0] : (args[0] && args[0].url) || String(args[0]);
          if (isMsApiUrl(url)) {
            try { considerAuthHeader(extractAuthFromHeaders((args[1] && args[1].headers) || (args[0] && args[0].headers))); } catch (e) { /* */ }
          }
          var res = Reflect.apply(orig, this, args);
          if (!isConvUrl(url)) return res;
          return Promise.resolve(res).then(function (r) {
            try {
              r.clone().text().then(function (body) {
                if (body && body.length > 20) {
                  try { rememberConversation(JSON.parse(body)); } catch (e) { /* */ }
                }
              }).catch(function () { /* */ });
            } catch (e) { /* */ }
            return r;
          });
        };
        Object.defineProperty(wrapped, '__dv64Wrapped', { value: true });
        try { Object.defineProperty(wrapped, 'name', { value: orig.name }); } catch (e) { /* */ }
        try { Object.defineProperty(wrapped, 'length', { value: orig.length }); } catch (e) { /* */ }
        window.fetch = wrapped;
      }
    } catch (e) { ceLog('warn', 'fetch capture failed:', e.message); }

    try {
      var proto = window.XMLHttpRequest && window.XMLHttpRequest.prototype;
      if (proto && !proto.open.__dv64Wrapped) {
        var oOpen = proto.open, oSend = proto.send, oSetHeader = proto.setRequestHeader;
        var wOpen = function (method, url) {
          try { this[SYMBOL_URL] = url; this[SYMBOL_HDR] = null; } catch (e) { /* */ }
          return Reflect.apply(oOpen, this, arguments);
        };
        var wSetHeader = function (name, value) {
          try {
            if (String(name).toLowerCase() === 'authorization') {
              this[SYMBOL_HDR] = this[SYMBOL_HDR] || {};
              this[SYMBOL_HDR].authorization = value;
            }
          } catch (e) { /* */ }
          return Reflect.apply(oSetHeader, this, arguments);
        };
        var wSend = function () {
          var url, hdrs;
          try { url = this[SYMBOL_URL]; hdrs = this[SYMBOL_HDR]; } catch (e) { /* */ }
          if (isMsApiUrl(url) && hdrs) {
            try { considerAuthHeader(hdrs.authorization); } catch (e) { /* */ }
          }
          if (isConvUrl(url)) {
            var self = this;
            this.addEventListener('load', function () {
              try {
                if (self.responseType && self.responseType !== 'text') return;
                rememberConversation(JSON.parse(self.responseText));
              } catch (e) { /* */ }
            }, { once: true });
          }
          return Reflect.apply(oSend, this, arguments);
        };
        Object.defineProperty(wOpen, '__dv64Wrapped', { value: true });
        proto.open = wOpen;
        proto.send = wSend;
        proto.setRequestHeader = wSetHeader;
      }
    } catch (e) { ceLog('warn', 'XHR capture failed:', e.message); }
  }

  /* ═══════ UI STATE ═══════ */
  var ui = {
    selected: new Set(),
    page: 0,
    logsFilter: 'all',
    search: '',
    sort: 'updated'
  };

  /* ═══════ MODAL ═══════ */
  function fmtLabel(v) {
    if (v === 'zip') return t('fmtZip');
    if (v === 'md') return t('fmtMd');
    if (v === 'clean') return t('fmtClean');
    if (v === 'gpt') return t('fmtGpt');
    return t('fmtRaw');
  }
  function fmtOptions(sel, list) {
    return list.map(function (v) {
      return '<option value="' + v + '"' + (v === sel ? ' selected' : '') + '>' + esc(fmtLabel(v)) + '</option>';
    }).join('');
  }

  function modalHtml() {
    return ''
      + '<div id="dv64-modal">'
      + '  <header class="dv-h">'
      + '    <div class="dv-h-l">'
      + '      <div class="dv-h-ico">\u2726</div>'
      + '      <div>'
      + '        <h2>CopilotKeep</h2>'
      + '        <p>' + esc(t('appTagline', { v: VERSION })) + '</p>'
      + '      </div>'
      + '    </div>'
      + '    <button class="dv-x" id="dv64-close" title="' + esc(t('close')) + '" aria-label="' + esc(t('close')) + '">\u00D7</button>'
      + '  </header>'
      + '  <nav class="dv-tabs">'
      + '    <button class="dv-tab active" data-tab="current">' + esc(t('tabCurrent')) + '</button>'
      + '    <button class="dv-tab" data-tab="archive">' + esc(t('tabArchive')) + '</button>'
      + '    <button class="dv-tab" data-tab="logs">' + esc(t('tabLogs')) + ' <span class="dv-badge" data-n="0" id="dv64-badge"></span></button>'
      + '    <button class="dv-tab" data-tab="settings">' + esc(t('tabSettings')) + '</button>'
      + '  </nav>'
      + '  <section class="dv-panel active" data-panel="current">'
      + '    <div class="dv-warn" id="dv64-current-info">' + esc(t('loading')) + '</div>'
      + '    <div class="dv-bar">'
      + '      <div class="dv-bar-l">'
      + '        <label class="dv-chk"><input type="checkbox" class="dv-cb" id="dv64-inc-user"' + (settings.includeUser ? ' checked' : '') + '> ' + esc(t('userMessages')) + '</label>'
      + '        <label class="dv-chk"><input type="checkbox" class="dv-cb" id="dv64-inc-asst"' + (settings.includeAssistant ? ' checked' : '') + '> ' + esc(t('copilotReplies')) + '</label>'
      + '        <label class="dv-chk"><input type="checkbox" class="dv-cb" id="dv64-inc-internal"' + (settings.includeInternal ? ' checked' : '') + '> ' + esc(t('includeInternal')) + '</label>'
      + '      </div>'
      + '    </div>'
      + '    <div style="padding:14px 20px;">'
      + '      <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:12px;">'
      + '        <button class="dv-btn dv-btn-gho" data-action="copy-page">' + esc(t('copyPageBtn')) + '</button>'
      + '        <button class="dv-btn dv-btn-gho" data-action="dl-page-md">' + esc(t('dlPageBtn', { fmt: 'MD' })) + '</button>'
      + '        <button class="dv-btn dv-btn-gho" data-action="dl-page-html">' + esc(t('dlPageBtn', { fmt: 'HTML' })) + '</button>'
      + '        <button class="dv-btn dv-btn-gho" data-action="dl-page-json">' + esc(t('dlPageBtn', { fmt: 'JSON' })) + '</button>'
      + '      </div>'
      + '      <div style="border-top:1px solid rgba(0,0,0,.08);padding-top:12px;margin-top:4px;">'
      + '        <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;">'
      + '          <label class="dv-lbl">' + esc(t('fullExportLabel')) + '</label>'
      + '          <select class="dv-select" id="dv64-full-fmt">' + fmtOptions(settings.defaultFullFmt, FULL_FMTS) + '</select>'
      + '          <button class="dv-btn dv-btn-pri" data-action="full-export">' + esc(t('exportFromMs')) + '</button>'
      + '        </div>'
      + '        <p style="font-size:11px;opacity:.65;margin:10px 0 0;line-height:1.5;">' + esc(t('currentHelp')) + '</p>'
      + '      </div>'
      + '    </div>'
      + '    <div id="dv64-progress-current"></div>'
      + '  </section>'
      + '  <section class="dv-panel" data-panel="archive">'
      + '    <div class="dv-bar">'
      + '      <div class="dv-bar-l">'
      + '        <label class="dv-lbl">' + esc(t('fetchLatest')) + '</label>'
      + '        <input type="number" class="dv-input" id="dv64-max" value="' + settings.maxChats + '" min="1" max="500">'
      + '        <button class="dv-btn dv-btn-pri" data-action="load-archive">' + esc(t('load')) + '</button>'
      + '      </div>'
      + '      <div class="dv-bar-r">'
      + '        <input type="search" class="dv-input" id="dv64-search" placeholder="' + esc(t('searchPh')) + '">'
      + '        <select class="dv-select" id="dv64-sort">'
      + '          <option value="updated">' + esc(t('sortUpdated')) + '</option>'
      + '          <option value="created">' + esc(t('sortCreated')) + '</option>'
      + '          <option value="name">' + esc(t('sortName')) + '</option>'
      + '        </select>'
      + '        <span class="dv-count" id="dv64-count">0 / 0</span>'
      + '      </div>'
      + '    </div>'
      + '    <div class="dv-tbl-wrap">'
      + '      <table class="dv-tbl">'
      + '        <thead><tr><th></th><th>' + esc(t('colName')) + '</th><th>' + esc(t('colCreated')) + '</th><th>' + esc(t('colUpdated')) + '</th><th>' + esc(t('colStatus')) + '</th></tr></thead>'
      + '        <tbody id="dv64-tbody"></tbody>'
      + '      </table>'
      + '    </div>'
      + '    <div class="dv-page" id="dv64-page"></div>'
      + '    <div id="dv64-progress-archive"></div>'
      + '    <div class="dv-foot">'
      + '      <div class="dv-foot-l">'
      + '        <label class="dv-chk"><input type="checkbox" class="dv-cb" id="dv64-all"> ' + esc(t('selectAll')) + '</label>'
      + '        <span class="dv-lbl">' + esc(t('format')) + '</span>'
      + '        <select class="dv-select" id="dv64-arc-fmt">' + fmtOptions(settings.defaultArchiveFmt, ARC_FMTS) + '</select>'
      + '      </div>'
      + '      <div class="dv-foot-r">'
      + '        <button class="dv-btn dv-btn-pri" data-action="export-archive" disabled>' + esc(t('exportSelected')) + '</button>'
      + '      </div>'
      + '    </div>'
      + '    <div class="dv-warn" style="font-size:11px;padding:9px 20px;border-top:1px solid rgba(0,0,0,.05);border-bottom:none;background:rgba(0,0,0,.02);color:inherit;opacity:.75;">'
      + '      ' + esc(t('archiveNote'))
      + '    </div>'
      + '  </section>'
      + '  <section class="dv-panel" data-panel="logs">'
      + '    <div class="dv-logs-bar">'
      + '      <button class="dv-btn dv-btn-gho active" data-filter="all">' + esc(t('logAll')) + '</button>'
      + '      <button class="dv-btn dv-btn-gho" data-filter="info">' + esc(t('logInfo')) + '</button>'
      + '      <button class="dv-btn dv-btn-gho" data-filter="warn">' + esc(t('logWarn')) + '</button>'
      + '      <button class="dv-btn dv-btn-gho" data-filter="error">' + esc(t('logError')) + '</button>'
      + '      <span style="flex:1"></span>'
      + '      <button class="dv-btn dv-btn-gho" id="dv64-copy-logs">' + esc(t('copy')) + '</button>'
      + '      <button class="dv-btn dv-btn-gho" id="dv64-clear-logs">' + esc(t('clear')) + '</button>'
      + '    </div>'
      + '    <div class="dv-logs" id="dv64-logs"></div>'
      + '  </section>'
      + '  <section class="dv-panel" data-panel="settings">'
      + '    <div class="dv-set">'
      + '      <div class="dv-set-sec">' + esc(t('secGeneral')) + '</div>'
      + '      <div class="dv-set-row">'
      + '        <div class="dv-set-l"><div class="dv-set-n">' + esc(t('rowLanguage')) + '</div><div class="dv-set-d">' + esc(t('rowLanguageDesc')) + '</div></div>'
      + '        <select class="dv-select" id="dv64-set-lang">'
      + '          <option value="en"' + (settings.lang === 'en' ? ' selected' : '') + '>' + esc(t('langEn')) + '</option>'
      + '          <option value="ar"' + (settings.lang === 'ar' ? ' selected' : '') + '>' + esc(t('langAr')) + '</option>'
      + '          <option value="zh"' + (settings.lang === 'zh' ? ' selected' : '') + '>' + esc(t('langZh')) + '</option>'
      + '        </select>'
      + '      </div>'
      + '      <div class="dv-set-row">'
      + '        <div class="dv-set-l"><div class="dv-set-n">' + esc(t('rowTheme')) + '</div><div class="dv-set-d">' + esc(t('rowThemeDesc')) + '</div></div>'
      + '        <select class="dv-select" id="dv64-set-theme">'
      + '          <option value="auto"' + (settings.theme === 'auto' ? ' selected' : '') + '>' + esc(t('themeAuto')) + '</option>'
      + '          <option value="light"' + (settings.theme === 'light' ? ' selected' : '') + '>' + esc(t('themeLight')) + '</option>'
      + '          <option value="dark"' + (settings.theme === 'dark' ? ' selected' : '') + '>' + esc(t('themeDark')) + '</option>'
      + '        </select>'
      + '      </div>'
      + '      <div class="dv-set-row">'
      + '        <div class="dv-set-l"><div class="dv-set-n">' + esc(t('rowToastDur')) + '</div><div class="dv-set-d">' + esc(t('rowToastDurDesc')) + '</div></div>'
      + '        <input type="number" class="dv-input" id="dv64-set-toastdur" min="1000" max="15000" step="500" value="' + settings.toastDurationMs + '">'
      + '      </div>'
      + '      <div class="dv-set-sec">' + esc(t('secDefaults')) + '</div>'
      + '      <div class="dv-set-row">'
      + '        <div class="dv-set-l"><div class="dv-set-n">' + esc(t('rowFullFmt')) + '</div><div class="dv-set-d">' + esc(t('rowFullFmtDesc')) + '</div></div>'
      + '        <select class="dv-select" id="dv64-set-fullfmt">' + fmtOptions(settings.defaultFullFmt, FULL_FMTS) + '</select>'
      + '      </div>'
      + '      <div class="dv-set-row">'
      + '        <div class="dv-set-l"><div class="dv-set-n">' + esc(t('rowArcFmt')) + '</div><div class="dv-set-d">' + esc(t('rowArcFmtDesc')) + '</div></div>'
      + '        <select class="dv-select" id="dv64-set-arcfmt">' + fmtOptions(settings.defaultArchiveFmt, ARC_FMTS) + '</select>'
      + '      </div>'
      + '      <div class="dv-set-row">'
      + '        <div class="dv-set-l"><div class="dv-set-n">' + esc(t('rowZip')) + '</div><div class="dv-set-d">' + esc(t('rowZipDesc')) + '</div></div>'
      + '        <label class="dv-chk"><input type="checkbox" class="dv-cb" id="dv64-set-zip"' + (settings.zipCompress ? ' checked' : '') + '></label>'
      + '      </div>'
      + '      <div class="dv-set-row">'
      + '        <div class="dv-set-l"><div class="dv-set-n">' + esc(t('rowInternal')) + '</div><div class="dv-set-d">' + esc(t('rowInternalDesc')) + '</div></div>'
      + '        <label class="dv-chk"><input type="checkbox" class="dv-cb" id="dv64-set-internal"' + (settings.includeInternal ? ' checked' : '') + '></label>'
      + '      </div>'
      + '      <div class="dv-set-sec">' + esc(t('secPerformance')) + '</div>'
      + '      <div class="dv-set-row">'
      + '        <div class="dv-set-l"><div class="dv-set-n">' + esc(t('rowFetchDelay')) + '</div><div class="dv-set-d">' + esc(t('rowFetchDelayDesc')) + '</div></div>'
      + '        <input type="number" class="dv-input" id="dv64-set-delay" min="100" max="5000" step="100" value="' + settings.fetchDelayMs + '">'
      + '      </div>'
      + '      <div class="dv-set-row">'
      + '        <div class="dv-set-l"><div class="dv-set-n">' + esc(t('rowPageSize')) + '</div><div class="dv-set-d">' + esc(t('rowPageSizeDesc')) + '</div></div>'
      + '        <input type="number" class="dv-input" id="dv64-set-pagesize" min="5" max="100" step="5" value="' + settings.pageSize + '">'
      + '      </div>'
      + '      <div class="dv-set-row">'
      + '        <div class="dv-set-l"><div class="dv-set-n">' + esc(t('rowMaxChats')) + '</div><div class="dv-set-d">' + esc(t('rowMaxChatsDesc')) + '</div></div>'
      + '        <input type="number" class="dv-input" id="dv64-set-maxchats" min="1" max="500" step="10" value="' + settings.maxChats + '">'
      + '      </div>'
      + '      <div class="dv-set-sec">' + esc(t('secData')) + '</div>'
      + '      <div class="dv-set-actions">'
      + '        <button class="dv-btn dv-btn-danger" id="dv64-set-reset">' + esc(t('resetBtn')) + '</button>'
      + '      </div>'
      + '    </div>'
      + '  </section>'
      + '</div>';
  }

  function openModal() {
    if (activeModal) return;
    var overlay = el('div', { id: 'dv64-overlay' });
    overlay.classList.toggle('dv-dark', isDarkMode());
    overlay.classList.toggle('rtl', settings.lang === 'ar');
    overlay.innerHTML = modalHtml();

    document.body.appendChild(overlay);
    activeModal = overlay;

    document.getElementById('dv64-close').onclick = function () {
      if (isBusy) { toast(t('waitOrCancel'), 'warn'); return; }
      closeModal();
    };
    overlay.addEventListener('click', function (e) {
      if (e.target === overlay && !isBusy) closeModal();
    });
    document.addEventListener('keydown', escHandler);

    overlay.querySelectorAll('.dv-tab').forEach(function (tabBtn) {
      tabBtn.onclick = function () { switchTab(tabBtn.dataset.tab); };
    });

    overlay.querySelector('[data-action="copy-page"]').onclick = function () { copyPage(); };
    overlay.querySelector('[data-action="dl-page-md"]').onclick = function () { downloadPage('md'); };
    overlay.querySelector('[data-action="dl-page-html"]').onclick = function () { downloadPage('html'); };
    overlay.querySelector('[data-action="dl-page-json"]').onclick = function () { downloadPage('json'); };
    overlay.querySelector('[data-action="full-export"]').onclick = fullExport;
    overlay.querySelector('[data-action="load-archive"]').onclick = loadArchive;
    overlay.querySelector('[data-action="export-archive"]').onclick = exportArchive;

    // Current-tab controls persist to settings
    bindCheckboxSetting('dv64-inc-user', 'includeUser');
    bindCheckboxSetting('dv64-inc-asst', 'includeAssistant');
    bindCheckboxSetting('dv64-inc-internal', 'includeInternal');
    bindSelectSetting('dv64-full-fmt', 'defaultFullFmt', FULL_FMTS, function (v) {
      var sel = document.getElementById('dv64-set-fullfmt');
      if (sel) sel.value = v;
    });
    bindSelectSetting('dv64-arc-fmt', 'defaultArchiveFmt', ARC_FMTS, function (v) {
      var sel = document.getElementById('dv64-set-arcfmt');
      if (sel) sel.value = v;
    });
    var maxInput = document.getElementById('dv64-max');
    maxInput.onchange = function () {
      settings.maxChats = clampNum(maxInput.value, 1, 500, DEFAULT_SETTINGS.maxChats);
      maxInput.value = settings.maxChats;
      saveSettings();
      var synced = document.getElementById('dv64-set-maxchats');
      if (synced) synced.value = settings.maxChats;
    };

    var searchInput = document.getElementById('dv64-search');
    searchInput.value = ui.search;
    searchInput.oninput = function () {
      ui.search = searchInput.value;
      ui.page = 0;
      renderArchiveTable();
    };
    var sortSel = document.getElementById('dv64-sort');
    sortSel.value = ui.sort;
    sortSel.onchange = function () {
      ui.sort = sortSel.value;
      ui.page = 0;
      renderArchiveTable();
    };

    var allCb = document.getElementById('dv64-all');
    allCb.onchange = function () {
      var cbs = document.querySelectorAll('#dv64-tbody input.dv-cb');
      cbs.forEach(function (cb) {
        cb.checked = allCb.checked;
        if (allCb.checked) ui.selected.add(cb.dataset.id);
        else ui.selected.delete(cb.dataset.id);
      });
      updateCount();
    };

    overlay.querySelectorAll('[data-filter]').forEach(function (b) {
      b.classList.toggle('active', b.dataset.filter === ui.logsFilter);
      b.onclick = function () {
        ui.logsFilter = b.dataset.filter;
        overlay.querySelectorAll('[data-filter]').forEach(function (x) { x.classList.toggle('active', x === b); });
        renderLogs();
      };
    });
    document.getElementById('dv64-copy-logs').onclick = function () {
      var txt = ceLogs.map(function (l) { return '[' + new Date(l.t).toISOString() + '] [' + l.level.toUpperCase() + '] ' + l.msg; }).join('\n');
      navigator.clipboard.writeText(txt).then(function () { toast(t('copied'), 'success'); }).catch(function () { toast(t('copyFailed'), 'error'); });
    };
    document.getElementById('dv64-clear-logs').onclick = function () { ceLogs.length = 0; renderLogs(); };

    wireSettingsPanel();

    updateCurrentInfo();
    renderArchiveTable();
    renderLogs();
  }

  function bindCheckboxSetting(id, key) {
    var cb = document.getElementById(id);
    if (!cb) return;
    cb.onchange = function () {
      settings[key] = cb.checked;
      saveSettings();
    };
  }
  function bindSelectSetting(id, key, valid, after) {
    var sel = document.getElementById(id);
    if (!sel) return;
    sel.onchange = function () {
      if (valid.indexOf(sel.value) === -1) return;
      settings[key] = sel.value;
      saveSettings();
      if (after) after(sel.value);
    };
  }

  function wireSettingsPanel() {
    var langSel = document.getElementById('dv64-set-lang');
    if (langSel) langSel.onchange = function () {
      settings.lang = LANGS.indexOf(langSel.value) !== -1 ? langSel.value : 'en';
      saveSettings();
      // Rebuild the whole modal so every label switches language.
      closeModal();
      openModal();
      switchTab('settings');
      toast(t('settingsSaved'), 'success', 1800);
    };

    var themeSel = document.getElementById('dv64-set-theme');
    if (themeSel) themeSel.onchange = function () {
      settings.theme = ['auto', 'light', 'dark'].indexOf(themeSel.value) !== -1 ? themeSel.value : 'auto';
      saveSettings();
      applyModalTheme();
      toast(t('settingsSaved'), 'success', 1800);
    };

    var toastDur = document.getElementById('dv64-set-toastdur');
    if (toastDur) toastDur.onchange = function () {
      settings.toastDurationMs = clampNum(toastDur.value, 1000, 15000, DEFAULT_SETTINGS.toastDurationMs);
      toastDur.value = settings.toastDurationMs;
      saveSettings();
      toast(t('settingsSaved'), 'success', 1800);
    };

    bindSelectSetting('dv64-set-fullfmt', 'defaultFullFmt', FULL_FMTS, function (v) {
      var sel = document.getElementById('dv64-full-fmt');
      if (sel) sel.value = v;
    });
    bindSelectSetting('dv64-set-arcfmt', 'defaultArchiveFmt', ARC_FMTS, function (v) {
      var sel = document.getElementById('dv64-arc-fmt');
      if (sel) sel.value = v;
    });

    var zipCb = document.getElementById('dv64-set-zip');
    if (zipCb) zipCb.onchange = function () {
      settings.zipCompress = zipCb.checked;
      saveSettings();
      toast(t('settingsSaved'), 'success', 1800);
    };

    var internalCb = document.getElementById('dv64-set-internal');
    if (internalCb) internalCb.onchange = function () {
      settings.includeInternal = internalCb.checked;
      saveSettings();
      var main = document.getElementById('dv64-inc-internal');
      if (main) main.checked = internalCb.checked;
      toast(t('settingsSaved'), 'success', 1800);
    };

    var delayInput = document.getElementById('dv64-set-delay');
    if (delayInput) delayInput.onchange = function () {
      settings.fetchDelayMs = clampNum(delayInput.value, 100, 5000, DEFAULT_SETTINGS.fetchDelayMs);
      delayInput.value = settings.fetchDelayMs;
      saveSettings();
      toast(t('settingsSaved'), 'success', 1800);
    };

    var pageSizeInput = document.getElementById('dv64-set-pagesize');
    if (pageSizeInput) pageSizeInput.onchange = function () {
      settings.pageSize = clampNum(pageSizeInput.value, 5, 100, DEFAULT_SETTINGS.pageSize);
      pageSizeInput.value = settings.pageSize;
      saveSettings();
      renderArchiveTable();
      toast(t('settingsSaved'), 'success', 1800);
    };

    var maxChatsInput = document.getElementById('dv64-set-maxchats');
    if (maxChatsInput) maxChatsInput.onchange = function () {
      settings.maxChats = clampNum(maxChatsInput.value, 1, 500, DEFAULT_SETTINGS.maxChats);
      maxChatsInput.value = settings.maxChats;
      saveSettings();
      var main = document.getElementById('dv64-max');
      if (main) main.value = settings.maxChats;
      toast(t('settingsSaved'), 'success', 1800);
    };

    var resetBtn = document.getElementById('dv64-set-reset');
    if (resetBtn) resetBtn.onclick = async function () {
      var ok = await confirmDlg(t('resetConfirmT'), t('resetConfirmB'), t('resetBtn'), true);
      if (!ok) return;
      settings = mergeSettings(null);
      saveSettings();
      closeModal();
      openModal();
      switchTab('settings');
      toast(t('settingsReset'), 'success');
    };
  }

  function escHandler(e) {
    if (e.key === 'Escape' && !isBusy) closeModal();
  }
  function closeModal() {
    document.removeEventListener('keydown', escHandler);
    if (activeModal) { activeModal.remove(); activeModal = null; }
  }
  function switchTab(name) {
    document.querySelectorAll('.dv-tab').forEach(function (tabBtn) { tabBtn.classList.toggle('active', tabBtn.dataset.tab === name); });
    document.querySelectorAll('.dv-panel').forEach(function (p) { p.classList.toggle('active', p.dataset.panel === name); });
    if (name === 'logs') renderLogs();
    if (name === 'current') updateCurrentInfo();
  }

  function updateCurrentInfo() {
    var info = document.getElementById('dv64-current-info');
    if (!info) return;
    var wrappers = document.querySelectorAll(CHAT_WRAPPER);
    var id = currentConversationId();
    var captured = lastConversation && lastConversation.conversationId === id
      ? ' \u00B7 ' + t('cachedHistory', { n: lastConversation.messages.length })
      : ' \u00B7 ' + t('notCapturedYet');
    // Detect partial via aria-setsize
    var reportedSize = null;
    wrappers.forEach(function (w) {
      var sz = Number(w.getAttribute('aria-setsize'));
      if (Number.isFinite(sz) && sz > 0 && (reportedSize === null || sz > reportedSize)) reportedSize = sz;
    });
    var partialWarn = (reportedSize && reportedSize > wrappers.length)
      ? ' \u00B7 ' + t('partialVisible', { a: wrappers.length, b: reportedSize })
      : '';
    info.textContent = wrappers.length
      ? safeTitle() + ' \u00B7 ' + t('visibleMsgs', { n: wrappers.length }) + captured + partialWarn + (id ? ' \u00B7 ' + t('idShort', { id: id.slice(0, 8) + '\u2026' }) : '')
      : t('openChatFirst');
    if (reportedSize && reportedSize > wrappers.length) info.classList.add('danger');
    else info.classList.remove('danger');
  }

  function renderLogs() {
    var list = document.getElementById('dv64-logs');
    if (!list) return;
    var arr = ui.logsFilter === 'all' ? ceLogs : ceLogs.filter(function (l) { return l.level === ui.logsFilter; });
    var errCount = ceLogs.filter(function (l) { return l.level === 'error'; }).length;
    var badge = document.getElementById('dv64-badge');
    if (badge) { badge.dataset.n = String(errCount); badge.textContent = errCount > 0 ? String(errCount) : ''; }
    if (!arr.length) {
      list.innerHTML = '';
      var emptyMsg = ui.logsFilter === 'all' ? t('noLogsYet') : t('noLogsLevel', { level: levelLabel(ui.logsFilter) });
      list.appendChild(el('div', { class: 'dv-empty' }, el('div', { class: 'ico', text: '\u2139' }), el('div', { text: emptyMsg })));
      return;
    }
    list.innerHTML = '';
    arr.slice(-300).forEach(function (l) {
      list.appendChild(el('div', { class: 'dv-log' },
        el('span', { class: 'dv-log-t', text: new Date(l.t).toLocaleTimeString() }),
        el('span', { class: 'dv-log-l ' + l.level, text: l.level }),
        el('span', { class: 'dv-log-m', text: l.msg })
      ));
    });
    list.scrollTop = list.scrollHeight;
  }
  function levelLabel(level) {
    return level === 'warn' ? t('logWarn') : level === 'error' ? t('logError') : t('logInfo');
  }

  /* ═══════ PAGE EXPORT ═══════ */
  async function copyPage() {
    if (isBusy) return;
    setBusy(true, t('extracting'));
    try {
      var conv = await extractCurrentConversation();
      var incU = document.getElementById('dv64-inc-user').checked;
      var incA = document.getElementById('dv64-inc-asst').checked;
      var sel = conv.messages.filter(function (m) {
        return (m.role === 'user' && incU) || (m.role === 'assistant' && incA);
      });
      var plain = conv.title + '\n\n' + sel.map(function (m) {
        return (m.role === 'user' ? 'You:' : 'Copilot:') + '\n' + m.text + '\n';
      }).join('\n');
      await navigator.clipboard.writeText(plain);
      var suffix = conv.partial ? ' (' + t('partialNote', { a: sel.length, b: conv.reportedCount || '?' }) + ')' : '';
      toast(t('copiedN', { n: sel.length }) + suffix, conv.partial ? 'warn' : 'success', conv.partial ? 5000 : 0);
    } catch (e) { toast(locMsg(e), 'error'); }
    finally { setBusy(false); }
  }

  async function downloadPage(fmt) {
    if (isBusy) return;
    setBusy(true, t('extracting'));
    try {
      var conv = await extractCurrentConversation();
      var incU = document.getElementById('dv64-inc-user').checked;
      var incA = document.getElementById('dv64-inc-asst').checked;
      conv.messages = conv.messages.filter(function (m) {
        return (m.role === 'user' && incU) || (m.role === 'assistant' && incA);
      });
      var title = safeName(conv.title);
      var stamp = new Date().toISOString().replace(/[:.]/g, '-');
      if (fmt === 'md') {
        var md = '# ' + conv.title + '\n\n';
        if (conv.partial) md += '> \u26A0 ' + t('partialLong', { a: conv.messages.length, b: conv.reportedCount || '?' }) + '\n\n';
        md += conv.messages.map(function (m) {
          return '## ' + (m.role === 'user' ? 'You' : 'Copilot') + '\n\n' + m.markdown + '\n';
        }).join('\n---\n\n');
        downloadText(md, 'copilot-page-' + title + '-' + stamp + '.md', 'text/markdown;charset=utf-8');
      } else if (fmt === 'html') {
        var html = '<!doctype html><html><head><meta charset="utf-8"><title>' + esc(conv.title) + '</title>'
          + '<style>body{max-width:900px;margin:32px auto;padding:0 20px;font-family:system-ui;line-height:1.6}h1{line-height:1.2}.msg{margin:18px 0;padding:14px;border-radius:10px;border:1px solid #ccc}.msg.user{border-left:4px solid #64748b}.msg.assistant{border-left:4px solid #2563eb}.role{font-weight:700;margin-bottom:8px}.warn{padding:10px;background:#fff3cd;border:1px solid #f59e0b;border-radius:8px;margin:12px 0}</style>'
          + '</head><body><h1>' + esc(conv.title) + '</h1>';
        if (conv.partial) html += '<div class="warn">\u26A0 ' + esc(t('partialLong', { a: conv.messages.length, b: conv.reportedCount || '?' })) + '</div>';
        html += conv.messages.map(function (m) {
          return '<div class="msg ' + m.role + '"><div class="role">' + (m.role === 'user' ? 'You' : 'Copilot') + '</div>' + (m.html || '<pre>' + esc(m.text) + '</pre>') + '</div>';
        }).join('');
        html += '</body></html>';
        downloadText(html, 'copilot-page-' + title + '-' + stamp + '.html', 'text/html;charset=utf-8');
      } else {
        downloadText(JSON.stringify(conv, null, 2), 'copilot-page-' + title + '-' + stamp + '.json', 'application/json;charset=utf-8');
      }
      toast(t('downloadedN', { n: conv.messages.length }) + (conv.partial ? ' (' + t('partialNote', { a: conv.messages.length, b: conv.reportedCount || '?' }) + ')' : ''), conv.partial ? 'warn' : 'success', conv.partial ? 5000 : 0);
    } catch (e) { toast(locMsg(e), 'error'); }
    finally { setBusy(false); }
  }

  /* ═══════ FULL EXPORT (always fresh) ═══════ */
  async function fullExport() {
    if (isBusy) return;
    var cid = currentConversationId() || (lastConversation && lastConversation.conversationId);
    if (!cid) { toast(t('openChatFirstShort'), 'warn'); return; }
    var fmt = document.getElementById('dv64-full-fmt').value;
    if (fmt === 'raw') {
      var ok = await confirmDlg(t('rawConfirmT'), t('rawConfirmBodyFull'), t('exportRaw'), true);
      if (!ok) return;
    }
    activeAbort = new AbortController();
    setBusy(true, t('preparingFull'));
    try {
      // Always fetch fresh
      var authRef = { value: await getArchiveAuth() };
      var conv = null;
      var usedCache = false;
      try {
        setBusy(true, t('fetchingMs'));
        conv = await apiGetConversation(authRef, cid, activeAbort.signal);
      } catch (fetchErr) {
        if (lastConversation && lastConversation.conversationId === cid) {
          ceLog('warn', 'Fresh fetch failed, using cache:', fetchErr.message);
          conv = lastConversation;
          usedCache = true;
          toast('\u26A0 ' + t('usingCache'), 'warn', 5000);
        } else {
          throw fetchErr;
        }
      }
      var includeInternal = document.getElementById('dv64-inc-internal').checked;
      var norm = normalizeConversation(conv, includeInternal);
      var title = safeName(norm.title);
      var stamp = new Date().toISOString().replace(/[:.]/g, '-');
      var blob = blobFromConv(norm, fmt);
      downloadBlob(blob, 'copilot-full-' + title + '-' + stamp + '.' + (fmt === 'md' ? 'md' : 'json'));
      toast(t('exportedNMsgs', { n: norm.messages.length }) + (usedCache ? ' ' + t('cachedNote') : ''), 'success');
    } catch (e) {
      if (e && e.name === 'AbortError') toast(t('cancelled'), 'info');
      else toast(locMsg(e), 'error');
    } finally {
      activeAbort = null;
      setBusy(false);
    }
  }

  /* ═══════ ARCHIVE ═══════ */
  function archiveView() {
    var q = ui.search.trim().toLowerCase();
    var list = chatCache.filter(function (c) {
      return !q || String(c.chatName || '').toLowerCase().indexOf(q) !== -1;
    });
    var sorters = {
      updated: function (a, b) { return toEpochMs(b.updateTimeUtc) - toEpochMs(a.updateTimeUtc); },
      created: function (a, b) { return toEpochMs(b.createTimeUtc) - toEpochMs(a.createTimeUtc); },
      name: function (a, b) { return String(a.chatName || '').localeCompare(String(b.chatName || '')); }
    };
    return list.slice().sort(sorters[ui.sort] || sorters.updated);
  }

  async function loadArchive() {
    if (isBusy) return;
    setBusy(true, t('loadingArchive'));
    try {
      var maxInput = document.getElementById('dv64-max');
      var max = clampNum(maxInput.value, 1, 500, DEFAULT_SETTINGS.maxChats);
      maxInput.value = max;
      settings.maxChats = max;
      saveSettings();
      var authRef = { value: await getArchiveAuth() };
      chatCache = await apiGetChatsMerged(authRef, max);
      ui.page = 0;
      ui.selected.clear();
      renderArchiveTable();
      toast(t('loadedN', { n: chatCache.length }), 'success');
    } catch (e) { toast(locMsg(e), 'error'); }
    finally { setBusy(false); }
  }

  function renderArchiveTable() {
    var tbody = document.getElementById('dv64-tbody');
    if (!tbody) return;
    var view = archiveView();
    if (!chatCache.length) {
      tbody.innerHTML = '';
      tbody.appendChild(el('tr', {}, el('td', { colspan: '5' },
        el('div', { class: 'dv-empty' }, el('div', { class: 'ico', text: '\u2139' }), el('div', { text: t('noConvs') }))
      )));
      updateCount();
      var pg0 = document.getElementById('dv64-page');
      if (pg0) pg0.innerHTML = '';
      return;
    }
    if (!view.length) {
      tbody.innerHTML = '';
      tbody.appendChild(el('tr', {}, el('td', { colspan: '5' },
        el('div', { class: 'dv-empty' }, el('div', { class: 'ico', text: '\u2315' }), el('div', { text: t('noMatch') }))
      )));
      updateCount();
      var pg1 = document.getElementById('dv64-page');
      if (pg1) pg1.innerHTML = '';
      return;
    }
    var pageSize = settings.pageSize;
    var pages = Math.ceil(view.length / pageSize);
    if (ui.page >= pages) ui.page = pages - 1;
    var start = ui.page * pageSize;
    var slice = view.slice(start, start + pageSize);
    tbody.innerHTML = '';
    slice.forEach(function (c) {
      var checked = ui.selected.has(c.conversationId);
      var d1 = c.createTimeUtc ? new Date(c.createTimeUtc).toLocaleDateString() : '\u2014';
      var d2 = c.updateTimeUtc ? new Date(c.updateTimeUtc).toLocaleDateString() : '\u2014';
      var tr = el('tr', { dataset: { id: c.conversationId } },
        el('td', {}, el('input', { type: 'checkbox', class: 'dv-cb', dataset: { id: c.conversationId }, checked: checked })),
        el('td', { class: 'dv-name', title: cleanTitle(c.chatName), text: cleanTitle(c.chatName) || t('untitled') }),
        el('td', { class: 'dv-date', text: d1 }),
        el('td', { class: 'dv-date', text: d2 }),
        el('td', {}, el('span', { class: 'dv-st empty', text: '\u2014' }))
      );
      tr.querySelector('input').onchange = function (e) {
        if (e.target.checked) ui.selected.add(c.conversationId);
        else ui.selected.delete(c.conversationId);
        updateCount();
        updateSelectAllState();
      };
      tbody.appendChild(tr);
    });
    var pg = document.getElementById('dv64-page');
    pg.innerHTML = '';
    if (pages > 1) {
      var prev = el('button', { text: '\u2190', disabled: ui.page === 0 });
      var next = el('button', { text: '\u2192', disabled: ui.page >= pages - 1 });
      prev.onclick = function () { ui.page--; renderArchiveTable(); };
      next.onclick = function () { ui.page++; renderArchiveTable(); };
      pg.appendChild(prev);
      pg.appendChild(el('span', { text: (ui.page + 1) + ' / ' + pages }));
      pg.appendChild(next);
    }
    updateCount();
    updateSelectAllState();
  }

  function updateCount() {
    var countEl = document.getElementById('dv64-count');
    if (countEl) countEl.textContent = ui.selected.size + ' / ' + chatCache.length;
    var btn = document.querySelector('[data-action="export-archive"]');
    if (btn) btn.disabled = ui.selected.size === 0 || isBusy;
  }

  // Sync "select all" with individual selections (indeterminate state)
  function updateSelectAllState() {
    var allCb = document.getElementById('dv64-all');
    if (!allCb) return;
    var visible = document.querySelectorAll('#dv64-tbody input.dv-cb');
    var visibleChecked = 0;
    visible.forEach(function (cb) { if (cb.checked) visibleChecked++; });
    if (visible.length === 0) {
      allCb.checked = false; allCb.indeterminate = false;
    } else if (visibleChecked === visible.length && ui.selected.size === chatCache.length) {
      allCb.checked = true; allCb.indeterminate = false;
    } else if (visibleChecked === 0 && ui.selected.size === 0) {
      allCb.checked = false; allCb.indeterminate = false;
    } else {
      allCb.checked = false; allCb.indeterminate = true;
    }
  }

  function setRowStatus(cid, s, msg) {
    var row = document.querySelector('#dv64-tbody tr[data-id="' + cid + '"]');
    if (!row) return;
    var cell = row.querySelector('.dv-st');
    if (!cell) return;
    var labels = { exp: t('stExporting'), done: t('stDone'), err: t('stError') };
    cell.className = 'dv-st' + (s ? '' : ' empty');
    cell.dataset.s = s || '';
    cell.title = msg || '';
    cell.textContent = labels[s] || '\u2014';
  }

  async function exportArchive() {
    if (isBusy || !ui.selected.size) return;
    var fmt = document.getElementById('dv64-arc-fmt').value;
    if (fmt === 'raw') {
      var ok = await confirmDlg(t('rawConfirmT'), t('rawConfirmBodyArc'), t('exportRaw'), true);
      if (!ok) return;
    }
    var items = chatCache.filter(function (c) { return ui.selected.has(c.conversationId); });
    var includeInternal = document.getElementById('dv64-inc-internal').checked;

    activeAbort = new AbortController();
    setBusy(true, t('startingExport'));
    var prog = showProgress('dv64-progress-archive', items.length, t('starting'));
    var results = { ok: [], err: [] };
    var files = [];
    var combined = [];

    try {
      var authRef = { value: await getArchiveAuth() };
      for (var i = 0; i < items.length; i++) {
        var c = items[i];
        setRowStatus(c.conversationId, 'exp');
        prog.update(i, cleanTitle(c.chatName) || t('untitled'));
        try {
          var conv = await apiGetConversation(authRef, c.conversationId, activeAbort.signal);
          var norm = normalizeConversation(conv, includeInternal);
          if (fmt === 'zip') {
            var blob = blobFromConv(norm, 'md');
            files.push({ name: fname(norm, 'md'), data: new Uint8Array(await blob.arrayBuffer()) });
          } else if (fmt === 'md') {
            combined.push(markdownFromConv(norm));
          } else if (fmt === 'clean') {
            combined.push(cleanJsonFromConv(norm));
          } else if (fmt === 'gpt') {
            combined.push(gptFromConv(norm));
          } else {
            combined.push(conv);
          }
          results.ok.push(c.conversationId);
          setRowStatus(c.conversationId, 'done');
        } catch (e) {
          if (e && e.name === 'AbortError') throw e;
          ceLog('error', 'Failed: ' + c.chatName, e);
          results.err.push({ id: c.conversationId, msg: e.message });
          setRowStatus(c.conversationId, 'err', e.message);
        }
        prog.update(i + 1);
        if (i < items.length - 1) await sleep(settings.fetchDelayMs);
      }
      prog.done();
      if (!results.ok.length) throw locErr('allFailed');

      var stamp = new Date().toISOString().replace(/[:.]/g, '-');
      if (fmt === 'zip') {
        var totalBytes = files.reduce(function (a, f) { return a + f.data.length; }, 0);
        if (totalBytes > RAM_WARN_MB * 1024 * 1024) {
          toast('\u26A0 ' + t('buildingZip', { mb: (totalBytes / 1024 / 1024).toFixed(1) }), 'warn', 5000);
        }
        var zipBlob = await buildZip(files, settings.zipCompress);
        downloadBlob(zipBlob, 'copilot-archive-' + stamp + '.zip');
      } else if (fmt === 'md') {
        downloadText(combined.join('\n\n---\n\n'), 'copilot-archive-' + stamp + '.md', 'text/markdown;charset=utf-8');
      } else if (fmt === 'clean' || fmt === 'gpt') {
        downloadText(JSON.stringify(combined, null, 2), 'copilot-archive-' + stamp + '.json', 'application/json;charset=utf-8');
      } else {
        downloadText(JSON.stringify(combined, null, 2), 'copilot-archive-raw-' + stamp + '.json', 'application/json;charset=utf-8');
      }

      if (results.err.length) {
        toast(t('exportedPart', { a: results.ok.length, b: items.length, c: results.err.length }), 'warn', 5000);
        ceLog('warn', 'Failed chats:', results.err);
      } else {
        toast(t('exportedNConvs', { n: results.ok.length }), 'success');
      }
    } catch (e) {
      if (e && e.name === 'AbortError') toast(t('cancelled'), 'info');
      else toast(locMsg(e), 'error');
    } finally {
      activeAbort = null;
      setBusy(false);
    }
  }

  function showProgress(elId, total, label) {
    var slot = document.getElementById(elId);
    if (!slot) return { update: function () {}, done: function () {} };
    slot.innerHTML = '';
    var cur = el('span', { text: label });
    var cnt = el('span', { text: '0 / ' + total });
    var fl = el('div', { class: 'dv-prog-fl', style: 'width:0%' });
    slot.appendChild(el('div', { class: 'dv-prog' },
      el('div', { class: 'dv-prog-lbl' }, cur, cnt),
      el('div', { class: 'dv-prog-tr' }, fl)
    ));
    return {
      update: function (done, lbl) {
        if (lbl) cur.textContent = String(lbl).slice(0, 60);
        cnt.textContent = done + ' / ' + total;
        fl.style.width = (done / total * 100) + '%';
      },
      done: function () { setTimeout(function () { slot.innerHTML = ''; }, 2000); }
    };
  }

  function setBusy(busy, msg) {
    isBusy = busy;
    ['copy-page', 'dl-page-md', 'dl-page-html', 'dl-page-json', 'full-export', 'load-archive', 'export-archive'].forEach(function (action) {
      var b = document.querySelector('[data-action="' + action + '"]');
      if (b) b.disabled = busy;
    });
    var closeBtn = document.getElementById('dv64-close');
    if (closeBtn) closeBtn.disabled = busy;
    var launcher = document.getElementById('dv64-launcher');
    if (launcher) launcher.disabled = busy;
    if (busy && msg) toast(msg, 'info', 2000);
    updateCount();
  }

  /* ═══════ INIT ═══════ */
  function injectLauncher() {
    if (document.getElementById('dv64-launcher')) return;
    var btn = el('button', { id: 'dv64-launcher', type: 'button', title: 'Open CopilotKeep', text: '\u2193' });
    btn.onclick = openModal;
    document.body.appendChild(btn);
  }

  ceLog('info', 'CopilotKeep v' + VERSION + ' initialized (lang=' + settings.lang + ', theme=' + settings.theme + ')');
  installPassiveCapture();

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', injectLauncher);
  } else {
    injectLauncher();
  }

  var reinjT = null;
  new MutationObserver(function () {
    clearTimeout(reinjT);
    reinjT = setTimeout(function () {
      if (!document.getElementById('dv64-launcher') && document.body) injectLauncher();
    }, 300);
  }).observe(document.documentElement, { childList: true, subtree: true });

  /* ═══════ TEST HOOKS (pure functions, for offline verification) ═══════ */
  window.__dv64Exporter = {
    VERSION: VERSION,
    LANGS: LANGS,
    STR: STR,
    DEFAULT_SETTINGS: DEFAULT_SETTINGS,
    mergeSettings: mergeSettings,
    t: t,
    locErr: locErr,
    locMsg: locMsg,
    isInternalMessage: isInternalMessage,
    filterMessages: filterMessages,
    adaptiveText: adaptiveText,
    pickText: pickText,
    extractCitations: extractCitations,
    normalizeConversation: normalizeConversation,
    markdownFromConv: markdownFromConv,
    gptFromConv: gptFromConv,
    cleanJsonFromConv: cleanJsonFromConv,
    toEpochSeconds: toEpochSeconds,
    toIso: toIso,
    cleanTitle: cleanTitle,
    safeName: safeName,
    fname: fname,
    crc32: crc32,
    buildZip: buildZip,
    fmtLabel: fmtLabel
  };
})();

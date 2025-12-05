/**
 * ============================================================================
 * INTERNATIONALIZATION (I18N) - Multi-Language Support
 * ============================================================================
 *
 * Provides translation support for Spanish and other languages.
 * Enables the dashboard to serve Spanish-speaking union members.
 *
 * @module I18n
 * @version 2.1.0
 * ============================================================================
 */

/**
 * Supported languages
 * @const {Object}
 */
const SUPPORTED_LANGUAGES = {
  EN: 'en',  // English
  ES: 'es'   // Spanish (Español)
};

/**
 * Translation strings
 * @const {Object}
 */
const TRANSLATIONS = {
  en: {
    // Dashboard
    DASHBOARD_TITLE: '📊 LOCAL 509 DASHBOARD',
    MEMBER_METRICS: '👥 MEMBER METRICS',
    GRIEVANCE_METRICS: '📋 GRIEVANCE METRICS',
    LAST_UPDATED: 'Last Updated',

    // Member metrics
    TOTAL_MEMBERS: 'Total Members',
    ACTIVE_STEWARDS: 'Active Stewards',
    AVG_OPEN_RATE: 'Avg Open Rate',
    YTD_VOL_HOURS: 'YTD Vol. Hours',

    // Grievance metrics
    OPEN_GRIEVANCES: 'Open Grievances',
    PENDING_INFO: 'Pending Info',
    SETTLED_THIS_MONTH: 'Settled This Month',
    AVG_DAYS_OPEN: 'Avg Days Open',
    UPCOMING_DEADLINES: 'Upcoming Deadlines',

    // Buttons
    START_GRIEVANCE: 'Start Grievance',
    CANCEL: 'Cancel',
    SAVE: 'Save',
    DELETE: 'Delete',
    EXPORT: 'Export',
    REFRESH: 'Refresh',
    SEARCH: 'Search',
    SUBMIT: 'Submit',

    // Status
    OPEN: 'Open',
    CLOSED: 'Closed',
    PENDING: 'Pending Info',
    SETTLED: 'Settled',
    WITHDRAWN: 'Withdrawn',
    APPEALED: 'Appealed',

    // Forms
    MEMBER_ID: 'Member ID',
    FIRST_NAME: 'First Name',
    LAST_NAME: 'Last Name',
    EMAIL: 'Email Address',
    PHONE: 'Phone Number',
    JOB_TITLE: 'Job Title',
    LOCATION: 'Work Location',
    UNIT: 'Unit',

    // Messages
    SUCCESS: 'Success',
    ERROR: 'Error',
    WARNING: 'Warning',
    INFO: 'Information',
    LOADING: 'Loading...',
    PROCESSING: 'Processing...',
    PLEASE_WAIT: 'Please wait...',

    // Common phrases
    YES: 'Yes',
    NO: 'No',
    SELECT: 'Select',
    PLEASE_SELECT: '-- Please Select --',
    REQUIRED: 'Required',
    OPTIONAL: 'Optional',
    ALL: 'All',
    NONE: 'None'
  },

  es: {
    // Dashboard
    DASHBOARD_TITLE: '📊 PANEL LOCAL 509',
    MEMBER_METRICS: '👥 MÉTRICAS DE MIEMBROS',
    GRIEVANCE_METRICS: '📋 MÉTRICAS DE QUEJAS',
    LAST_UPDATED: 'Última Actualización',

    // Member metrics
    TOTAL_MEMBERS: 'Total de Miembros',
    ACTIVE_STEWARDS: 'Delegados Activos',
    AVG_OPEN_RATE: 'Tasa de Apertura Promedio',
    YTD_VOL_HOURS: 'Horas de Voluntariado del Año',

    // Grievance metrics
    OPEN_GRIEVANCES: 'Quejas Abiertas',
    PENDING_INFO: 'Información Pendiente',
    SETTLED_THIS_MONTH: 'Resueltas Este Mes',
    AVG_DAYS_OPEN: 'Días Abiertos Promedio',
    UPCOMING_DEADLINES: 'Fechas Límite Próximas',

    // Buttons
    START_GRIEVANCE: 'Iniciar Queja',
    CANCEL: 'Cancelar',
    SAVE: 'Guardar',
    DELETE: 'Eliminar',
    EXPORT: 'Exportar',
    REFRESH: 'Actualizar',
    SEARCH: 'Buscar',
    SUBMIT: 'Enviar',

    // Status
    OPEN: 'Abierta',
    CLOSED: 'Cerrada',
    PENDING: 'Información Pendiente',
    SETTLED: 'Resuelta',
    WITHDRAWN: 'Retirada',
    APPEALED: 'Apelada',

    // Forms
    MEMBER_ID: 'ID del Miembro',
    FIRST_NAME: 'Nombre',
    LAST_NAME: 'Apellido',
    EMAIL: 'Correo Electrónico',
    PHONE: 'Número de Teléfono',
    JOB_TITLE: 'Título del Trabajo',
    LOCATION: 'Ubicación de Trabajo',
    UNIT: 'Unidad',

    // Messages
    SUCCESS: 'Éxito',
    ERROR: 'Error',
    WARNING: 'Advertencia',
    INFO: 'Información',
    LOADING: 'Cargando...',
    PROCESSING: 'Procesando...',
    PLEASE_WAIT: 'Por favor espere...',

    // Common phrases
    YES: 'Sí',
    NO: 'No',
    SELECT: 'Seleccionar',
    PLEASE_SELECT: '-- Por Favor Seleccione --',
    REQUIRED: 'Requerido',
    OPTIONAL: 'Opcional',
    ALL: 'Todos',
    NONE: 'Ninguno'
  }
};

/**
 * Gets the current user's preferred language
 * Checks User Settings sheet, falls back to English
 *
 * @returns {string} Language code (en, es, etc.)
 */
function getUserLanguage() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userSettings = ss.getSheetByName(SHEETS.USER_SETTINGS);

    if (!userSettings) {
      return SUPPORTED_LANGUAGES.EN; // Default to English
    }

    const userEmail = getCurrentUserEmail();
    const data = userSettings.getDataRange().getValues();

    // Find user's language preference
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === userEmail) {
        const language = data[i][1]; // Assuming column B is language
        return language || SUPPORTED_LANGUAGES.EN;
      }
    }

    return SUPPORTED_LANGUAGES.EN;

  } catch (error) {
    Logger.log('Error getting user language: ' + error.message);
    return SUPPORTED_LANGUAGES.EN;
  }
}

/**
 * Sets the user's preferred language
 *
 * @param {string} language - Language code (en, es)
 */
function setUserLanguage(language) {
  if (!Object.values(SUPPORTED_LANGUAGES).includes(language)) {
    throw new Error(`Unsupported language: ${language}`);
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let userSettings = ss.getSheetByName(SHEETS.USER_SETTINGS);

    if (!userSettings) {
      userSettings = ss.insertSheet(SHEETS.USER_SETTINGS);
      userSettings.appendRow(['User Email', 'Language', 'Last Updated']);
    }

    const userEmail = getCurrentUserEmail();
    const data = userSettings.getDataRange().getValues();
    let rowFound = false;

    // Update existing row
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === userEmail) {
        userSettings.getRange(i + 1, 2).setValue(language);
        userSettings.getRange(i + 1, 3).setValue(new Date());
        rowFound = true;
        break;
      }
    }

    // Add new row if not found
    if (!rowFound) {
      userSettings.appendRow([userEmail, language, new Date()]);
    }

    SpreadsheetApp.flush();

    SpreadsheetApp.getActive().toast(
      language === 'es' ? 'Idioma cambiado a Español' : 'Language changed to English',
      'Settings',
      NUMERIC_CONSTANTS.TOAST_DURATION_SHORT
    );

  } catch (error) {
    Logger.log('Error setting user language: ' + error.message);
    throw error;
  }
}

/**
 * Translates a key to the current user's language
 *
 * @param {string} key - Translation key
 * @param {string} locale - Optional language override
 * @returns {string} Translated text
 *
 * @example
 * t('DASHBOARD_TITLE')  // Returns "📊 LOCAL 509 DASHBOARD" or "📊 PANEL LOCAL 509"
 * t('START_GRIEVANCE', 'es')  // Force Spanish
 */
function t(key, locale = null) {
  const language = locale || getUserLanguage();

  if (!TRANSLATIONS[language]) {
    Logger.log(`Language not found: ${language}, using English`);
    return TRANSLATIONS[SUPPORTED_LANGUAGES.EN][key] || key;
  }

  return TRANSLATIONS[language][key] || TRANSLATIONS[SUPPORTED_LANGUAGES.EN][key] || key;
}

/**
 * Shows language selection dialog
 */
function showLanguageSelector() {
  const currentLanguage = getUserLanguage();

  const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        ${getCommonDialogStyles()}
        .language-option {
          padding: 15px;
          margin: 10px 0;
          border: 2px solid #ddd;
          border-radius: 8px;
          cursor: pointer;
          transition: all 0.3s;
        }
        .language-option:hover {
          border-color: #1a73e8;
          background: #e8f0fe;
        }
        .language-option.selected {
          border-color: #1a73e8;
          background: #e8f0fe;
          font-weight: bold;
        }
        .language-flag {
          font-size: 32px;
          margin-right: 15px;
        }
      </style>
    </head>
    <body>
      <div class="container">
        <h2>🌍 Select Language / Seleccionar Idioma</h2>

        <div class="language-option ${currentLanguage === 'en' ? 'selected' : ''}"
             onclick="selectLanguage('en')">
          <span class="language-flag">🇺🇸</span>
          <strong>English</strong><br>
          <small>English (United States)</small>
        </div>

        <div class="language-option ${currentLanguage === 'es' ? 'selected' : ''}"
             onclick="selectLanguage('es')">
          <span class="language-flag">🇪🇸</span>
          <strong>Español</strong><br>
          <small>Spanish (Español)</small>
        </div>

        ${createInfoBox(
          'Note / Nota',
          'Your language preference will be saved and applied across the dashboard. / ' +
          'Su preferencia de idioma se guardará y aplicará en todo el panel.',
          'info'
        )}
      </div>

      <script>
        function selectLanguage(lang) {
          google.script.run
            .withSuccessHandler(onLanguageChanged)
            .withFailureHandler(onError)
            .setUserLanguage(lang);
        }

        function onLanguageChanged() {
          alert('✅ Language updated! Please refresh the page to see changes.\\n\\n' +
                '✅ ¡Idioma actualizado! Actualice la página para ver los cambios.');
          google.script.host.close();
        }

        function onError(error) {
          alert('❌ Error: ' + error.message);
        }
      </script>
    </body>
    </html>
  `;

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(500).setHeight(400),
    'Language / Idioma'
  );
}

/**
 * Gets all available languages
 *
 * @returns {Array<Object>} Array of {code, name, nativeName}
 */
function getAvailableLanguages() {
  return [
    {
      code: SUPPORTED_LANGUAGES.EN,
      name: 'English',
      nativeName: 'English',
      flag: '🇺🇸'
    },
    {
      code: SUPPORTED_LANGUAGES.ES,
      name: 'Spanish',
      nativeName: 'Español',
      flag: '🇪🇸'
    }
  ];
}

/**
 * Translates an entire object's values
 *
 * @param {Object} obj - Object with keys to translate
 * @param {string} locale - Language code
 * @returns {Object} Object with translated values
 */
function translateObject(obj, locale = null) {
  const translated = {};

  for (const key in obj) {
    if (obj.hasOwnProperty(key)) {
      translated[key] = t(obj[key], locale);
    }
  }

  return translated;
}

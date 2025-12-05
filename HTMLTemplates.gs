/**
 * ============================================================================
 * HTML TEMPLATES - Reusable HTML Generation Helpers
 * ============================================================================
 *
 * Provides functions to generate HTML components consistently across the app.
 * Refactors long HTML string generation into smaller, testable functions.
 *
 * @module HTMLTemplates
 * @version 2.1.0
 * ============================================================================
 */

/**
 * Creates a complete HTML page with consistent structure
 *
 * @param {string} title - Page title
 * @param {string} styles - CSS styles (without <style> tags)
 * @param {string} body - HTML body content
 * @param {string} scripts - JavaScript code (without <script> tags)
 * @returns {string} Complete HTML page
 */
function createHTMLPage(title, styles, body, scripts = '') {
  return `<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${sanitizeHTML(title)}</title>
  <style>${styles}</style>
</head>
<body>
  ${body}
  ${scripts ? `<script>${scripts}</script>` : ''}
</body>
</html>`;
}

/**
 * Common CSS styles for dialogs
 * @returns {string} CSS styles
 */
function getCommonDialogStyles() {
  return `
    body {
      font-family: 'Roboto', Arial, sans-serif;
      padding: 20px;
      margin: 0;
      background: #f5f5f5;
    }
    .container {
      background: white;
      padding: 25px;
      border-radius: 8px;
      box-shadow: 0 2px 8px rgba(0,0,0,0.1);
      max-width: 100%;
    }
    h2 {
      color: #1a73e8;
      margin-top: 0;
      border-bottom: 3px solid #1a73e8;
      padding-bottom: 10px;
    }
    .form-group {
      margin-bottom: 20px;
    }
    label {
      display: block;
      font-weight: bold;
      margin-bottom: 8px;
      color: #333;
    }
    select, input, textarea {
      width: 100%;
      padding: 10px;
      border: 2px solid #ddd;
      border-radius: 4px;
      font-size: 14px;
      box-sizing: border-box;
    }
    select:focus, input:focus, textarea:focus {
      outline: none;
      border-color: #1a73e8;
    }
    .info-box {
      background: #e8f0fe;
      padding: 15px;
      border-radius: 4px;
      margin: 15px 0;
      border-left: 4px solid #1a73e8;
    }
    .warning-box {
      background: #fff3cd;
      padding: 15px;
      border-radius: 4px;
      margin: 15px 0;
      border-left: 4px solid #ffc107;
    }
    .error-box {
      background: #f8d7da;
      padding: 15px;
      border-radius: 4px;
      margin: 15px 0;
      border-left: 4px solid #dc3545;
    }
    .success-box {
      background: #d4edda;
      padding: 15px;
      border-radius: 4px;
      margin: 15px 0;
      border-left: 4px solid #28a745;
    }
    .button-group {
      display: flex;
      gap: 10px;
      margin-top: 25px;
    }
    button {
      flex: 1;
      padding: 12px 24px;
      border: none;
      border-radius: 4px;
      font-size: 16px;
      font-weight: bold;
      cursor: pointer;
      transition: all 0.3s;
    }
    .btn-primary {
      background: #1a73e8;
      color: white;
    }
    .btn-primary:hover {
      background: #1557b0;
      transform: translateY(-2px);
      box-shadow: 0 4px 12px rgba(26,115,232,0.4);
    }
    .btn-primary:disabled {
      background: #ccc;
      cursor: not-allowed;
      transform: none;
    }
    .btn-secondary {
      background: #f1f3f4;
      color: #333;
    }
    .btn-secondary:hover {
      background: #e8eaed;
    }
    .btn-danger {
      background: #dc3545;
      color: white;
    }
    .btn-danger:hover {
      background: #c82333;
    }
    .loading {
      display: none;
      text-align: center;
      padding: 20px;
      color: #1a73e8;
    }
    .spinner {
      border: 4px solid #f3f3f3;
      border-top: 4px solid #1a73e8;
      border-radius: 50%;
      width: 40px;
      height: 40px;
      animation: spin 1s linear infinite;
      margin: 0 auto;
    }
    @keyframes spin {
      0% { transform: rotate(0deg); }
      100% { transform: rotate(360deg); }
    }

    /* Mobile responsive */
    @media (max-width: 768px) {
      body {
        padding: 10px;
      }
      .container {
        padding: 15px;
      }
      .button-group {
        flex-direction: column;
      }
      button {
        width: 100%;
      }
    }

    /* Touch-friendly on mobile */
    @media (hover: none) and (pointer: coarse) {
      button, select, input {
        min-height: 44px;
        font-size: 16px;
      }
    }
  `;
}

/**
 * Creates an info box component
 *
 * @param {string} title - Box title
 * @param {string} content - Box content
 * @param {string} type - Box type: 'info', 'warning', 'error', 'success'
 * @returns {string} HTML for info box
 */
function createInfoBox(title, content, type = 'info') {
  const icons = {
    info: 'ℹ️',
    warning: '⚠️',
    error: '❌',
    success: '✅'
  };

  const icon = icons[type] || icons.info;

  return `
    <div class="${type}-box">
      <strong>${icon} ${sanitizeHTML(title)}</strong><br>
      ${sanitizeHTML(content)}
    </div>
  `;
}

/**
 * Creates a loading spinner component
 *
 * @param {string} message - Loading message
 * @param {string} id - Element ID (default: 'loading')
 * @returns {string} HTML for loading spinner
 */
function createLoadingSpinner(message = 'Loading...', id = 'loading') {
  return `
    <div class="loading" id="${sanitizeHTML(id)}">
      <div class="spinner"></div>
      <p>${sanitizeHTML(message)}</p>
    </div>
  `;
}

/**
 * Creates button HTML
 *
 * @param {string} text - Button text
 * @param {string} onClick - JavaScript onClick handler
 * @param {string} type - Button type: 'primary', 'secondary', 'danger'
 * @param {boolean} disabled - Whether button is disabled
 * @returns {string} HTML for button
 */
function createButton(text, onClick, type = 'primary', disabled = false) {
  return `
    <button
      class="btn-${type}"
      onclick="${sanitizeHTML(onClick)}"
      ${disabled ? 'disabled' : ''}
    >
      ${sanitizeHTML(text)}
    </button>
  `;
}

/**
 * Creates a form group (label + input/select)
 *
 * @param {string} label - Form label
 * @param {string} inputHTML - Input element HTML
 * @param {string} helpText - Optional help text
 * @returns {string} HTML for form group
 */
function createFormGroup(label, inputHTML, helpText = '') {
  return `
    <div class="form-group">
      <label>${sanitizeHTML(label)}</label>
      ${inputHTML}
      ${helpText ? `<small style="color: #666;">${sanitizeHTML(helpText)}</small>` : ''}
    </div>
  `;
}

/**
 * Creates a select dropdown
 *
 * @param {string} id - Select ID
 * @param {Array<Object>} options - Array of {value, text} objects
 * @param {string} onChange - onChange handler (optional)
 * @returns {string} HTML for select element
 */
function createSelect(id, options, onChange = '') {
  const optionsHTML = options.map(opt =>
    `<option value="${sanitizeHTML(opt.value)}">${sanitizeHTML(opt.text)}</option>`
  ).join('');

  return `
    <select id="${sanitizeHTML(id)}" ${onChange ? `onchange="${sanitizeHTML(onChange)}"` : ''}>
      ${optionsHTML}
    </select>
  `;
}

/**
 * Creates a text input
 *
 * @param {string} id - Input ID
 * @param {string} placeholder - Placeholder text
 * @param {string} type - Input type (text, email, phone, etc.)
 * @returns {string} HTML for input element
 */
function createInput(id, placeholder = '', type = 'text') {
  return `
    <input
      type="${sanitizeHTML(type)}"
      id="${sanitizeHTML(id)}"
      placeholder="${sanitizeHTML(placeholder)}"
    >
  `;
}

/**
 * Creates a textarea
 *
 * @param {string} id - Textarea ID
 * @param {string} placeholder - Placeholder text
 * @param {number} rows - Number of rows
 * @returns {string} HTML for textarea element
 */
function createTextarea(id, placeholder = '', rows = 4) {
  return `
    <textarea
      id="${sanitizeHTML(id)}"
      placeholder="${sanitizeHTML(placeholder)}"
      rows="${rows}"
    ></textarea>
  `;
}

/**
 * Creates a table
 *
 * @param {Array<string>} headers - Table headers
 * @param {Array<Array<string>>} rows - Table rows
 * @param {string} className - Optional CSS class
 * @returns {string} HTML for table
 */
function createTable(headers, rows, className = '') {
  const headerHTML = headers.map(h => `<th>${sanitizeHTML(h)}</th>`).join('');

  const rowsHTML = rows.map(row => {
    const cells = row.map(cell => `<td>${sanitizeHTML(cell)}</td>`).join('');
    return `<tr>${cells}</tr>`;
  }).join('');

  return `
    <table ${className ? `class="${sanitizeHTML(className)}"` : ''}>
      <thead><tr>${headerHTML}</tr></thead>
      <tbody>${rowsHTML}</tbody>
    </table>
  `;
}

/**
 * Creates a modal dialog
 *
 * @param {string} id - Modal ID
 * @param {string} title - Modal title
 * @param {string} content - Modal content
 * @returns {string} HTML for modal
 */
function createModal(id, title, content) {
  return `
    <div id="${sanitizeHTML(id)}" class="modal" style="display:none;">
      <div class="modal-content">
        <div class="modal-header">
          <h3>${sanitizeHTML(title)}</h3>
          <span class="modal-close" onclick="document.getElementById('${sanitizeHTML(id)}').style.display='none'">&times;</span>
        </div>
        <div class="modal-body">
          ${content}
        </div>
      </div>
    </div>
    <style>
      .modal {
        position: fixed;
        z-index: 1000;
        left: 0;
        top: 0;
        width: 100%;
        height: 100%;
        background-color: rgba(0,0,0,0.4);
      }
      .modal-content {
        background-color: white;
        margin: 5% auto;
        padding: 20px;
        border-radius: 8px;
        width: 80%;
        max-width: 600px;
        box-shadow: 0 4px 20px rgba(0,0,0,0.3);
      }
      .modal-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 20px;
        border-bottom: 2px solid #1a73e8;
        padding-bottom: 10px;
      }
      .modal-close {
        color: #aaa;
        font-size: 28px;
        font-weight: bold;
        cursor: pointer;
      }
      .modal-close:hover {
        color: #000;
      }
    </style>
  `;
}

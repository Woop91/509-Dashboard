/**
 * ============================================================================
 * NOTIFICATION CENTER
 * ============================================================================
 *
 * Centralized notification management system
 * Features:
 * - In-app notifications
 * - Notification history
 * - Priority levels
 * - Mark as read/unread
 * - Notification preferences
 * - Push notifications (toast)
 * - Notification badges
 */

/**
 * Notification types and priorities
 */
const NOTIFICATION_CONFIG = {
  TYPES: {
    INFO: 'info',
    SUCCESS: 'success',
    WARNING: 'warning',
    ERROR: 'error',
    SYSTEM: 'system'
  },
  PRIORITIES: {
    LOW: 1,
    MEDIUM: 2,
    HIGH: 3,
    URGENT: 4
  },
  MAX_NOTIFICATIONS: 100
};

/**
 * Shows notification center
 */
function showNotificationCenter() {
  const html = createNotificationCenterHTML();
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(800)
    .setHeight(700);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '🔔 Notifications');
}

/**
 * Creates HTML for notification center
 */
function createNotificationCenterHTML() {
  const notifications = getNotifications();
  const unreadCount = notifications.filter(n => !n.read).length;

  let notificationRows = '';
  if (notifications.length === 0) {
    notificationRows = '<div class="no-notifications">📭 No notifications yet</div>';
  } else {
    notifications.forEach((notif, index) => {
      const icon = getNotificationIcon(notif.type);
      const priorityClass = `priority-${notif.priority || 1}`;

      notificationRows += `
        <div class="notification-card ${notif.read ? 'read' : 'unread'} ${priorityClass}" onclick="toggleRead(${index})">
          <div class="notification-header">
            <span class="notification-icon">${icon}</span>
            <span class="notification-title">${notif.title}</span>
            <span class="notification-time">${getTimeAgo(notif.timestamp)}</span>
          </div>
          <div class="notification-body">${notif.message}</div>
          ${!notif.read ? '<div class="unread-indicator"></div>' : ''}
        </div>
      `;
    });
  }

  return `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { font-family: Arial, sans-serif; padding: 20px; margin: 0; background: #f5f5f5; }
    .container { max-width: 750px; margin: 0 auto; background: white; padding: 25px; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); max-height: 650px; overflow-y: auto; }
    h2 { color: #1a73e8; margin-top: 0; display: flex; align-items: center; justify-content: space-between; }
    .badge { background: #f44336; color: white; padding: 4px 12px; border-radius: 12px; font-size: 14px; font-weight: bold; }
    .controls { margin: 20px 0; display: flex; gap: 10px; flex-wrap: wrap; }
    button { background: #1a73e8; color: white; border: none; padding: 10px 20px; font-size: 13px; border-radius: 6px; cursor: pointer; }
    button:hover { background: #1557b0; }
    button.secondary { background: #6c757d; }
    .notification-card { background: #f8f9fa; padding: 15px; margin: 12px 0; border-radius: 8px; border-left: 4px solid #ddd; cursor: pointer; transition: all 0.2s; position: relative; }
    .notification-card:hover { transform: translateX(5px); box-shadow: 0 2px 8px rgba(0,0,0,0.1); }
    .notification-card.unread { background: #e8f0fe; border-left-color: #1a73e8; font-weight: 500; }
    .notification-card.priority-4 { border-left-color: #f44336; }
    .notification-card.priority-3 { border-left-color: #ff9800; }
    .notification-header { display: flex; align-items: center; gap: 10px; margin-bottom: 8px; }
    .notification-icon { font-size: 20px; }
    .notification-title { flex: 1; font-weight: 600; color: #333; }
    .notification-time { font-size: 12px; color: #999; }
    .notification-body { color: #666; font-size: 14px; line-height: 1.6; }
    .unread-indicator { position: absolute; top: 20px; right: 15px; width: 10px; height: 10px; background: #1a73e8; border-radius: 50%; }
    .no-notifications { text-align: center; padding: 60px; font-size: 18px; color: #999; }
  </style>
</head>
<body>
  <div class="container">
    <h2>
      <span>🔔 Notifications</span>
      ${unreadCount > 0 ? `<span class="badge">${unreadCount} new</span>` : ''}
    </h2>

    <div class="controls">
      <button onclick="markAllRead()">✅ Mark All Read</button>
      <button onclick="clearAll()" class="secondary">🗑️ Clear All</button>
      <button onclick="refreshNotifications()" class="secondary">🔄 Refresh</button>
    </div>

    <div id="notificationList">
      ${notificationRows}
    </div>
  </div>

  <script>
    function toggleRead(index) {
      google.script.run
        .withSuccessHandler(() => location.reload())
        .toggleNotificationRead(index);
    }

    function markAllRead() {
      google.script.run
        .withSuccessHandler(() => location.reload())
        .markAllNotificationsRead();
    }

    function clearAll() {
      if (confirm('Clear all notifications?')) {
        google.script.run
          .withSuccessHandler(() => location.reload())
          .clearAllNotifications();
      }
    }

    function refreshNotifications() {
      location.reload();
    }
  </script>
</body>
</html>
  `;
}

/**
 * Gets notification icon based on type
 * @param {string} type - Notification type
 * @returns {string} Icon
 */
function getNotificationIcon(type) {
  switch (type) {
    case NOTIFICATION_CONFIG.TYPES.INFO: return 'ℹ️';
    case NOTIFICATION_CONFIG.TYPES.SUCCESS: return '✅';
    case NOTIFICATION_CONFIG.TYPES.WARNING: return '⚠️';
    case NOTIFICATION_CONFIG.TYPES.ERROR: return '❌';
    case NOTIFICATION_CONFIG.TYPES.SYSTEM: return '⚙️';
    default: return '🔔';
  }
}

/**
 * Gets time ago string
 * @param {string} timestamp - ISO timestamp
 * @returns {string} Time ago
 */
function getTimeAgo(timestamp) {
  const now = new Date();
  const then = new Date(timestamp);
  const seconds = Math.floor((now - then) / 1000);

  if (seconds < 60) return 'Just now';
  if (seconds < 3600) return Math.floor(seconds / 60) + 'm ago';
  if (seconds < 86400) return Math.floor(seconds / 3600) + 'h ago';
  if (seconds < 604800) return Math.floor(seconds / 86400) + 'd ago';
  return then.toLocaleDateString();
}

/**
 * Gets all notifications for current user
 * @returns {Array} Notifications
 */
function getNotifications() {
  const props = PropertiesService.getUserProperties();
  const notificationsJSON = props.getProperty('notifications');

  if (!notificationsJSON) {
    return [];
  }

  const notifications = JSON.parse(notificationsJSON);
  return notifications.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
}

/**
 * Adds a notification
 * @param {string} title - Notification title
 * @param {string} message - Notification message
 * @param {string} type - Notification type
 * @param {number} priority - Priority level (1-4)
 */
function addNotification(title, message, type = NOTIFICATION_CONFIG.TYPES.INFO, priority = NOTIFICATION_CONFIG.PRIORITIES.MEDIUM) {
  const props = PropertiesService.getUserProperties();
  let notifications = getNotifications();

  notifications.unshift({
    title: title,
    message: message,
    type: type,
    priority: priority,
    timestamp: new Date().toISOString(),
    read: false
  });

  // Limit notifications
  if (notifications.length > NOTIFICATION_CONFIG.MAX_NOTIFICATIONS) {
    notifications = notifications.slice(0, NOTIFICATION_CONFIG.MAX_NOTIFICATIONS);
  }

  props.setProperty('notifications', JSON.stringify(notifications));

  // Show toast for high/urgent priority
  if (priority >= NOTIFICATION_CONFIG.PRIORITIES.HIGH) {
    const icon = getNotificationIcon(type);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      message,
      `${icon} ${title}`,
      5
    );
  }
}

/**
 * Toggles notification read status
 * @param {number} index - Notification index
 */
function toggleNotificationRead(index) {
  const props = PropertiesService.getUserProperties();
  const notifications = getNotifications();

  if (index >= 0 && index < notifications.length) {
    notifications[index].read = !notifications[index].read;
    props.setProperty('notifications', JSON.stringify(notifications));
  }
}

/**
 * Marks all notifications as read
 */
function markAllNotificationsRead() {
  const props = PropertiesService.getUserProperties();
  const notifications = getNotifications();

  notifications.forEach(n => n.read = true);
  props.setProperty('notifications', JSON.stringify(notifications));

  SpreadsheetApp.getActiveSpreadsheet().toast(
    '✅ All notifications marked as read',
    'Notifications',
    3
  );
}

/**
 * Clears all notifications
 */
function clearAllNotifications() {
  const props = PropertiesService.getUserProperties();
  props.deleteProperty('notifications');

  SpreadsheetApp.getActiveSpreadsheet().toast(
    '✅ All notifications cleared',
    'Notifications',
    3
  );
}

/**
 * Gets unread notification count
 * @returns {number} Count
 */
function getUnreadCount() {
  const notifications = getNotifications();
  return notifications.filter(n => !n.read).length;
}

/**
 * Sends welcome notification (for new users)
 */
function sendWelcomeNotification() {
  addNotification(
    'Welcome to 509 Dashboard!',
    'Explore the features using the menu. Check out the Help Center to get started.',
    NOTIFICATION_CONFIG.TYPES.INFO,
    NOTIFICATION_CONFIG.PRIORITIES.LOW
  );
}

/**
 * Sends system notification
 * @param {string} message - Notification message
 */
function sendSystemNotification(message) {
  addNotification(
    'System Notification',
    message,
    NOTIFICATION_CONFIG.TYPES.SYSTEM,
    NOTIFICATION_CONFIG.PRIORITIES.MEDIUM
  );
}

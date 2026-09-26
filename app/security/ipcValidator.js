/**
 * IPC Security Validation Module
 * 
 * Provides security validation for IPC channels as a compensating control
 * for disabled contextIsolation and sandbox features.
 */

// Allowlist of legitimate IPC channels used by Outlook for Linux
const allowedChannels = new Set([
  // Core application channels
  'config-file-changed',
  'config-changed',
  'get-config',
  'get-app-version',

  // Zoom and display controls
  'get-zoom-level',
  'save-zoom-level',

  // Notifications and user interaction
  'play-notification-sound',
  'show-notification',
  // main -> renderer only (webContents.send to the renderer that created the
  // notification). Not gated by this validator, which wraps ipcMain
  // handle/on/once only; listed so the allowlist stays authoritative. No
  // ipcMain handler exists for either, so these entries grant a renderer nothing.
  'notification-closed',
  'notification-clicked',
  'notification-show-toast',
  'notification-toast-click',
  'set-badge-count',
  'tray-update',

  // Authentication and forms
  'submitForm',

  // Connection management
  'offline-retry',

  // Navigation controls
  'navigate-back',
  'navigate-forward',
  'get-navigation-state',
  'navigation-state-changed',

  // Renderer-side error forwarding (registered in app/browser/preload.js)
  'unhandled-rejection',
  'window-error',

  // WebAuthn / FIDO2 security key support
  'webauthn:create',
  'webauthn:get',
  'webauthn:pin-submit',
  'webauthn:pin-cancel',
  'webauthn:touch-cancel',

  // Shared secure-prompt dialog (smartcard / PKCS#11 client-certificate PIN)
  'secure-prompt:submit',
  'secure-prompt:cancel',
]);

const DANGEROUS_PROPS = new Set(['__proto__', 'constructor', 'prototype']);
const MAX_SANITIZE_DEPTH = 10;

/**
 * Recursively sanitizes an object to remove prototype pollution vectors.
 * @param {any} obj - The object to sanitize
 * @param {number} depth - Current recursion depth (prevents stack overflow on circular refs)
 */
function sanitizePayload(obj, depth = 0) {
  if (!obj || typeof obj !== 'object' || depth > MAX_SANITIZE_DEPTH) {
    return;
  }

  for (const prop of DANGEROUS_PROPS) {
    if (Object.hasOwn(obj, prop)) {
      delete obj[prop];
    }
  }

  for (const key of Object.keys(obj)) {
    if (obj[key] && typeof obj[key] === 'object') {
      sanitizePayload(obj[key], depth + 1);
    }
  }
}

/**
 * Validates an IPC channel request
 * @param {string} channel - The IPC channel name
 * @param {any} payload - The payload being sent
 * @returns {boolean} - True if request is valid, false if blocked
 */
function validateIpcChannel(channel, payload = null) {
  // Check channel allowlist
  if (!allowedChannels.has(channel)) {
    console.warn(`[IPC Security] Blocked unauthorized channel: ${channel}`);
    return false;
  }

  // Recursive payload sanitization to prevent prototype pollution
  sanitizePayload(payload);

  return true;
}

module.exports = { validateIpcChannel, allowedChannels };
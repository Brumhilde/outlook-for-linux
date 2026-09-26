# Notification Service Module

Handles native desktop notifications and notification sounds for Outlook mail and calendar reminders.

## NotificationService Class

Manages OS notifications and sound playback based on configuration.

**Dependencies:**
- `soundPlayer` - Audio player instance (NodeSound)
- `config` - Application configuration
- `mainWindow` - Main application window

**IPC Channels:**
- `show-notification` - Display native notification
- `play-notification-sound` - Play notification sound

**Usage:**
```javascript
const notificationService = new NotificationService(
  player,
  config,
  mainWindow
);
notificationService.initialize();
```

**Sound Playback:**
Sound is controlled by `config.disableNotificationSound` (global disable).

**Notification Types:**
- `new-message` - Plays new_message.wav

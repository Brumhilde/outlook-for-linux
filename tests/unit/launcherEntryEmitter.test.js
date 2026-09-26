'use strict';

const { describe, it } = require('node:test');
const assert = require('node:assert');

const { computeDesktopUri } = require('../../app/downloadManager/launcherEntryEmitter');

// Receivers (Ubuntu Dock, Dash-to-Dock) match LauncherEntry signals against
// the desktop file as exported on the host, so the URI has to follow the
// packaging. A wrong URI is dropped silently: no badge, no progress.
describe('launcherEntryEmitter.computeDesktopUri', () => {
	it('uses the plain app name for deb/rpm/AppImage installs', () => {
		assert.strictEqual(
			computeDesktopUri('outlook-for-linux', {}),
			'application://outlook-for-linux.desktop',
		);
	});

	it('prefixes the snap instance name the way snapd exports desktop files', () => {
		assert.strictEqual(
			computeDesktopUri('outlook-for-linux', {
				SNAP_NAME: 'outlook-for-linux',
				SNAP_INSTANCE_NAME: 'outlook-for-linux',
			}),
			'application://outlook-for-linux_outlook-for-linux.desktop',
		);
	});

	it('keeps parallel snap instances distinct', () => {
		assert.strictEqual(
			computeDesktopUri('outlook-for-linux', {
				SNAP_NAME: 'outlook-for-linux',
				SNAP_INSTANCE_NAME: 'outlook-for-linux_work',
			}),
			'application://outlook-for-linux_work_outlook-for-linux.desktop',
		);
	});

	it('ignores a custom app name under snap, where the desktop basename is fixed at build time', () => {
		assert.strictEqual(
			computeDesktopUri('my-teams', {
				CHROME_DESKTOP: 'my-teams.desktop',
				SNAP_NAME: 'outlook-for-linux',
				SNAP_INSTANCE_NAME: 'outlook-for-linux',
			}),
			'application://outlook-for-linux_outlook-for-linux.desktop',
		);
	});

	it('falls back to the app name when SNAP_NAME is missing', () => {
		assert.strictEqual(
			computeDesktopUri('outlook-for-linux', { SNAP_INSTANCE_NAME: 'outlook-for-linux' }),
			'application://outlook-for-linux_outlook-for-linux.desktop',
		);
	});

	it('uses the Flatpak app id as the desktop file name', () => {
		assert.strictEqual(
			computeDesktopUri('outlook-for-linux', {
				FLATPAK_ID: 'io.github.outlook_for_linux',
			}),
			'application://io.github.outlook_for_linux.desktop',
		);
	});

	it('lets the snap prefix win when both snap and flatpak variables are present', () => {
		assert.strictEqual(
			computeDesktopUri('outlook-for-linux', {
				SNAP_NAME: 'outlook-for-linux',
				SNAP_INSTANCE_NAME: 'outlook-for-linux',
				FLATPAK_ID: 'com.example.ignored',
			}),
			'application://outlook-for-linux_outlook-for-linux.desktop',
		);
	});

	it('follows the desktop entry Electron was given, not the display name', () => {
		// app.name is "Outlook for Linux" on Linux so notifications carry that
		// header; the desktop file is still outlook-for-linux.desktop.
		assert.strictEqual(
			computeDesktopUri('Outlook for Linux', { CHROME_DESKTOP: 'outlook-for-linux.desktop' }),
			'application://outlook-for-linux.desktop',
		);
	});

	it('follows a custom app name (config `class`)', () => {
		assert.strictEqual(
			computeDesktopUri('my-teams', {}),
			'application://my-teams.desktop',
		);
	});
});

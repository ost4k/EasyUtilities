# EasyUtilities 1.4.3.0 Release Notes

## GitHub

### Fixed
- Fixed "Start with Windows" behavior for packaged (MSIX) builds by switching to `StartupTask` API.
- Kept fallback startup registration via `HKCU\Software\Microsoft\Windows\CurrentVersion\Run` for unpackaged/debug runs.
- Added startup activation handling so app minimizes to tray when started by Windows startup task.
- Added automatic cleanup of legacy `Run` registry entry in packaged mode to prevent stale startup records.

### Technical changes
- Added `windows.startupTask` extension to `Package.appxmanifest`.
- Refactored startup service methods to async and updated UI call sites accordingly.

## Microsoft Store (EN)

What is new in 1.4.3.0:
- Fixed autostart reliability for the Microsoft Store/MSIX version.
- Improved startup behavior when Windows launches the app at sign-in.
- Removed legacy startup compatibility issues from older versions.
- Internal stability and startup-service improvements.

## Microsoft Store (UK)

Що нового в 1.4.3.0:
- Виправлено надійність автозапуску для версії Microsoft Store/MSIX.
- Покращено поведінку старту застосунку при запуску Windows після входу в систему.
- Усунено проблеми сумісності зі старими механізмами автозапуску з попередніх версій.
- Внутрішні покращення стабільності та сервісу автозапуску.

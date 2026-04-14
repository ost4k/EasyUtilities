# EasyUtilities

EasyUtilities - це десктопний застосунок для Windows (WinUI 3 + .NET 8), який об'єднує набір щоденних системних утиліт: автозапуск, трей-режим, глобальні гарячі клавіші, OCR з виділеної області екрана, експорт списку файлів з Explorer, керування desktop-ярликами та backup/restore налаштувань.

## Що вміє застосунок

- Автозапуск з Windows: для MSIX через `StartupTask`, для unpackaged-запуску через `HKCU\Software\Microsoft\Windows\CurrentVersion\Run` (із запуском у згорнутому режимі).
- Згортання в трей замість закриття вікна.
- Глобальні гарячі клавіші:
  - `Ctrl+Alt+T` - увімкнути/вимкнути `Always on top` для активного вікна.
  - `Ctrl+Alt+O` - OCR з виділеної області екрана (результат у буфер обміну).
  - `Ctrl+Alt+M` - показати/сховати головне вікно застосунку.
  - `Ctrl+Alt+R` - перезапустити `explorer.exe`.
  - `Ctrl+Alt+F` - експорт списку файлів з активної папки Explorer.
- Middle-click по верхній частині вікон: згортання більшості програм середньою кнопкою миші.
- Приховування підписів desktop-ярликів (візуально, без обов'язкового перейменування).
- Приховування стрілок для internet-ярликів (`.url`) через shell-налаштування для поточного користувача.
- Експорт списку файлів з активного Explorer:
  - у буфер обміну;
  - у CSV-файл у тій самій папці.
- Save/Restore розкладки desktop-іконок у JSON (знімок відповідних ключів реєстру).
- Backup/Restore налаштувань застосунку у JSON.
- Локалізація інтерфейсу: `Українська`, `English`, `System`.
- Тема: `System`, `Light`, `Dark`.

## Технології

- .NET 8
- WinUI 3 (`Microsoft.WindowsAppSDK`)
- Windows Forms (для tray icon і file dialogs)
- Windows OCR API (`Windows.Media.Ocr`)
- Win32 API (hotkeys, hooks, window ops)
- MSIX packaging (x86/x64/arm64)

## Структура репозиторію

- `EasyUtilities/` - основний WinUI-застосунок.
- `EasyUtilities.TrayHost/` - helper-бібліотека для tray icon (`NotifyIcon`).
- `docs/PRIVACY_POLICY.md` - політика приватності.
- `docs/STORE_PUBLISH.md` - нотатки по публікації в Microsoft Store.
- `EasyUtilities.slnx` - solution.

Ключові файли:

- `EasyUtilities/MainWindow.xaml` - UI всіх карток з фічами.
- `EasyUtilities/MainWindow.xaml.cs` - логіка UI, гарячі клавіші, трей, backup/restore.
- `EasyUtilities/Services/*.cs` - доменні сервіси (startup, shell tweaks, OCR, export, settings, localization тощо).
- `EasyUtilities/NativeMethods.cs` - P/Invoke до Win32.

## Вимоги

- Windows 10/11 (мінімальна версія платформи: `10.0.17763.0`)
- .NET SDK 8.x
- Більшість shell-змін застосовуються без прав адміністратора; приховування стрілок internet-ярликів працює на рівні поточного користувача

## Локальний запуск

З кореня репозиторію:

```powershell
dotnet restore EasyUtilities\EasyUtilities.csproj
dotnet build EasyUtilities\EasyUtilities.csproj -c Debug
dotnet run --project EasyUtilities\EasyUtilities.csproj
```

## Аргументи запуску

- `--start-minimized` - запуск і згортання в трей.
- `--set-shortcut-arrows=on` - приховати стрілки internet-ярликів (`.url`).
- `--set-shortcut-arrows=off` - повернути стрілки internet-ярликів (`.url`).

## Де зберігаються дані

Локально в профілі користувача:

- Налаштування: `%LOCALAPPDATA%\EasyUtilities\settings.json`
- Журнал фатальних винятків: `%LOCALAPPDATA%\EasyUtilities\fatal.log`

Експортовані файли (`backup`, `desktop icons`, `csv`) зберігаються у шлях, який обирає користувач.

## Приватність

Застосунок працює локально, без вбудованої аналітики та трекерів. Детально:

- `docs/PRIVACY_POLICY.md`

## Публікація в Microsoft Store

Інструкція та приклад `dotnet publish`:

- `docs/STORE_PUBLISH.md`
- Важливо: у перших 2 рядках Store-опису вкажіть залежність `Microsoft .NET 8.0 Desktop Runtime (x64)`.

## Відомі особливості

- Для застосування частини shell-змін може знадобитися перезапуск `Explorer`.
- Якщо стрілки раніше змінювались старою версією застосунку через `HKLM`, при першому вимкненні може знадобитися одноразове підтвердження UAC для очищення legacy-параметрів.
- Гаряча клавіша OCR залежить від доступності системних OCR-компонентів/мовних пакетів.
- Функція приховування стрілок стосується саме internet-ярликів (`.url`), а не всіх ярликів `.lnk`.

## Ліцензія

Ліцензія в репозиторії наразі не вказана.

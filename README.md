# VerifyHash

**Check a downloaded file against its published hash, from the Windows
right-click menu.** One PowerShell file, no dependency, no administrator
rights.

*Read this in [Français](README.fr.md).*

![Platform](https://img.shields.io/badge/platform-Windows%2010%20%7C%2011-0078D6)
![PowerShell](https://img.shields.io/badge/PowerShell-5.1-5391FE)
![License](https://img.shields.io/badge/license-MIT-green)

<img src="docs/screenshot.png" alt="The verification window, green, showing a computed and an expected SHA-256 that match" width="716">

## The problem

A project publishes a checksum next to its download. Comparing 64 hexadecimal
characters by eye is slow and unreliable — and the eye is worst exactly where
it matters, in the middle of the string.

`Get-FileHash` gives you the value, but you still have to compare it yourself.
VerifyHash does the whole thing: right-click the file, paste the published
hash, and the window turns green or red.

## Install

### Option 1 — Download (recommended)

1. Click **Code → Download ZIP** at the top of this page.
2. Extract it somewhere it can stay — the menu entry will point at that
   location.
3. Double-click **`VerifyHash.bat`**, then click **Install**.

That's it. The launcher runs the script with an execution policy bypass, so
there is nothing to unblock and no policy to change.

### Option 2 — Clone

```powershell
git clone https://github.com/Defacedz/VerifyHash.git
cd VerifyHash
.\VerifyHash.ps1
```

### Running the .ps1 directly

If you would rather use **Right-click → Run with PowerShell** on
`VerifyHash.ps1`, unblock it first. Windows tags files that came from the
internet, and PowerShell then refuses to load them — the console shows a red
error and closes instantly:

```powershell
Unblock-File .\VerifyHash.ps1
```

Or right-click the file → **Properties** → tick **Unblock** → OK.

## Usage

Right-click any file → **Verify hash**.

> On Windows 11 the entry may sit behind **Show more options**, or `Shift` +
> right-click.

Paste the value published by the vendor into the second field. The comparison
runs as you type; there is no button to confirm.

- **MD5, SHA-1, SHA-256, SHA-512.**
- **The algorithm follows the hash you paste.** A 32-character value switches
  to MD5, 40 to SHA-1, 64 to SHA-256, 128 to SHA-512. You rarely need to press
  the algorithm buttons yourself.
- **Pasted text is cleaned up.** Spaces, line breaks and prefixes such as
  `SHA256:` are dropped, and the case is normalised — copying a whole line off
  a download page works.
- **Differences are highlighted** in both fields, with their positions, so you
  can tell a truncated paste from a genuinely different file.
- **The window appears instantly** and a progress bar moves while the file is
  read, even on several gigabytes. Switching algorithm mid-hash cancels and
  restarts immediately.
- **Drop another file on the window** to check that one next.
- A band across the bottom turns green when the hashes match, red when they do
  not — readable from across the room.
- `Esc` closes the window.

The interface is in English, and in French on a French Windows. Force one with
`-Language en` or `-Language fr`.

## Privacy

The script never writes to the clipboard — the **Paste** button only reads it.
No network access, no telemetry, no file created anywhere. The only thing
written to the machine is one registry key under `HKEY_CURRENT_USER`.

## Uninstall

Double-click `VerifyHash.bat` again and click **Uninstall**. The registry key
is removed and the folder can then be deleted.

## How it works

- The menu entry is a single key under
  `HKEY_CURRENT_USER\Software\Classes\*\shell\VerifyHash`. `*` means "every
  file type". Being under `HKCU` and not `HKLM`, it applies to your account
  only and needs no elevation.
- That key is written through the .NET `Microsoft.Win32.Registry` API rather
  than `New-Item`. The path contains a `*`, which the PowerShell Registry
  provider would treat as a wildcard.
- Hashing reads the file in 4 MB chunks through `TransformBlock`, pumping the
  message loop between chunks. That is what keeps the window responsive
  without a second thread — and what makes cancelling a 10 GB hash instant.
- The file is encoded as **UTF-8 with a BOM**. Windows PowerShell 5.1 does not
  detect UTF-8 without one and would render the accented French strings as
  garbage. Keep the BOM if you edit it.

## Requirements

Windows 10 or 11, Windows PowerShell 5.1 (ships with Windows). No
dependencies, no .NET install, no administrator rights.

## License

MIT — see [LICENSE](LICENSE).

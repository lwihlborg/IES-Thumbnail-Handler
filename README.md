# IES Photometry Thumbnail Handler for Windows

A Windows Shell Extension that displays thumbnail previews of IES photometry files directly in Windows Explorer. The thumbnails show polar plots of both vertical and horizontal candela distributions.

<img width="814" height="447" alt="image" src="https://github.com/user-attachments/assets/a2c80bef-01f3-47d7-84be-aaaa2198500a" />

## Features

- **Native Windows Integration** - Thumbnails appear automatically in Explorer, no separate application needed
- **Dual Distribution Display** - Shows both vertical (V) and horizontal (H) candela distributions
- **Automatic Symmetry Expansion** - Correctly handles quadrant, bilateral, and full photometric data
- **Clean Visualization** - Blue and red polar plots on white background for clear viewing
- **Supports All IES Formats** - Works with IESNA LM-63-1986, LM-63-1991, LM-63-1995, and LM-63-2002

## System Requirements

- Windows 10 or Windows 11 (64-bit recommended)
- .NET Framework 4.8 (included with Windows 10 1903+)

## Quick Installation

### Pre-built Binary

1. Download the latest release from the [Releases](../../releases) page
2. Extract the ZIP file
3. Run `install.bat`
4. Done! Navigate to a folder with IES files to see thumbnails

### Build from Source

1. Install [.NET SDK 6.0+](https://dotnet.microsoft.com/download) or Visual Studio 2022
2. Open a terminal in the project directory
3. Build the project:
   ```cmd
   dotnet build -c Release
   ```
4. Run `install.bat`

## How It Works

The handler registers itself as a Windows Shell Thumbnail Provider for `.ies` files. When Explorer needs to display a thumbnail:

1. Windows loads the CLR via mscoree.dll, which instantiates our handler and passes the IES file stream
2. The handler parses the photometric data (angles and candela values)
3. A combined polar plot is generated with both distributions overlaid:
   - **Blue**: Vertical candela distribution (0° nadir at bottom, 180° zenith at top)
   - **Red**: Horizontal candela distribution at ~45° vertical angle
4. The bitmap is returned to Explorer for display

### Thumbnail Visualization

```
         180° (zenith)
              │
              │    ╭─── Red = Horizontal
         ┌────┴────┐
        ╱    ┌─┐    ╲
       │    │   │    │
  270°─┼────┼───┼────┼─90°
       │    │   │    │
        ╲   └───┘   ╱
         └────┬────┘
              │    └─── Blue = Vertical
              │
          0° (nadir)
```

- **Blue (Vertical)**: Shows how light is distributed from nadir (down) through horizontal to zenith (up)
- **Red (Horizontal)**: Shows the beam pattern around the luminaire at 45° from nadir

## Troubleshooting

### Thumbnails not appearing

1. **Check Explorer View Settings**
   - Open File Explorer → View → Options → View tab
   - Ensure "Always show icons, never thumbnails" is **unchecked**

2. **Clear Thumbnail Cache**
   - Press `Win + R`, type `cleanmgr`, press Enter
   - Select your drive and click OK
   - Check "Thumbnails" and click OK

3. **Restart Explorer**
   - Open Task Manager (Ctrl+Shift+Esc)
   - Find "Windows Explorer" in the processes
   - Right-click and select "Restart"

4. **Verify Installation**
   - Open Registry Editor (regedit)
   - Navigate to: `HKCU\Software\Classes\.ies\ShellEx\{e357fccd-a995-4576-b01f-234630154e96}`
   - Should contain: `{7B3FC2A1-E8D4-4F5B-9A2C-8E1D3F6B5C4A}`

### Thumbnails appear blank or incorrect

- Ensure the IES file is valid (has proper TILT= line and numeric data)
- Some very old or non-standard IES files may not parse correctly
- Try opening the IES file in a text editor to verify its format

## Uninstallation

1. Run `uninstall.bat`
2. The handler will be unregistered and files removed

## Technical Details

### Registry Keys Created

```
HKCU\Software\Classes\.ies\ShellEx\{e357fccd-a995-4576-b01f-234630154e96}
HKCU\Software\Classes\.IES\ShellEx\{e357fccd-a995-4576-b01f-234630154e96}
HKCU\Software\Classes\CLSID\{7B3FC2A1-E8D4-4F5B-9A2C-8E1D3F6B5C4A}
HKCU\Software\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\{7B3FC2A1-E8D4-4F5B-9A2C-8E1D3F6B5C4A}
```

### Files Installed

```
%LocalAppData%\IESThumbnailHandler\IESThumbnailHandler.dll
```

### COM Interfaces Implemented

- `IThumbnailProvider` - Provides the thumbnail bitmap
- `IInitializeWithStream` - Receives the file stream from Explorer

## Building for Different Architectures

```cmd
# 64-bit (recommended)
dotnet build -c Release -p:Platform=x64

# 32-bit
dotnet build -c Release -p:Platform=x86
```

## License

MIT License - Feel free to modify and distribute.

## Acknowledgments

- IES file format specification: IESNA LM-63
- Windows Shell Extension documentation: Microsoft Docs

---

*Created for lighting designers who browse lots of IES files*

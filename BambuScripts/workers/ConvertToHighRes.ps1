param(
    [string]$InputFile
)
$ErrorActionPreference = 'Stop'

# ════════════════════════════════════════════════════════════════════════════════
#  ConvertToHighRes.ps1
#  Copies *Final.3mf files to a High-tag folder structure with renamed files,
#  then clears VLH (removes layer_heights_profile.txt) and sets layer_height=0.09.
#
#  Input:  X1C_CalicoCat_Farm_Final.3mf
#  Output: {base}\X1C_High_Farm\X1C_HighCalicoCat_Farm\X1C_HighCalicoCat_Farm_Final.3mf
# ════════════════════════════════════════════════════════════════════════════════

Add-Type -AssemblyName 'System.IO.Compression.FileSystem'

# ── 1. Collect all Final.3mf files from dropped paths ──
$inputs = [System.IO.File]::ReadAllLines($InputFile) | Where-Object { $_ -ne '' }

$files = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
foreach ($input in $inputs) {
    if (Test-Path $input -PathType Container) {
        Get-ChildItem -Path $input -Recurse -Filter '*Final.3mf' | ForEach-Object { $files.Add($_) }
    } elseif ((Test-Path $input -PathType Leaf) -and $input -imatch 'Final\.3mf$') {
        $files.Add((Get-Item $input))
    }
}

if ($files.Count -eq 0) {
    Write-Host "No Final.3mf files found in the dropped items."
    exit 1
}

Write-Host "Found $($files.Count) Final.3mf file(s):"
foreach ($f in $files) { Write-Host "  $($f.Name)" }
Write-Host ""

# ── 2. Prompt for destination base folder (modern IFileOpenDialog picker) ──
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public class HighResFolderPicker {
    [ComImport, Guid("DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7")]
    private class FileOpenDialogImpl {}

    [ComImport, Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IShellItem {
        void BindToHandler([In] IntPtr pbc, [In] ref Guid bhid, [In] ref Guid riid, out IntPtr ppv);
        void GetParent([MarshalAs(UnmanagedType.Interface)] out IShellItem ppsi);
        void GetDisplayName([In] uint sigdnName, [MarshalAs(UnmanagedType.LPWStr)] out string ppszName);
        void GetAttributes([In] uint sfgaoMask, out uint psfgaoAttribs);
        void Compare([In, MarshalAs(UnmanagedType.Interface)] IShellItem psi, [In] uint hint, out int piOrder);
    }

    [ComImport, Guid("d57c7288-d4ad-4768-be02-9d969532d960"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IFileOpenDialog {
        [PreserveSig] int Show([In] IntPtr parent);
        void SetFileTypes([In] uint cFileTypes, [In] IntPtr rgFilterSpec);
        void SetFileTypeIndex([In] uint iFileType);
        void GetFileTypeIndex(out uint piFileType);
        void Advise([In, MarshalAs(UnmanagedType.Interface)] IntPtr pfde, out uint pdwCookie);
        void Unadvise([In] uint dwCookie);
        void SetOptions([In] uint fos);
        void GetOptions(out uint pfos);
        void SetDefaultFolder([In, MarshalAs(UnmanagedType.Interface)] IntPtr psi);
        void SetFolder([In, MarshalAs(UnmanagedType.Interface)] IntPtr psi);
        void GetFolder([MarshalAs(UnmanagedType.Interface)] out IntPtr ppsi);
        void GetCurrentSelection([MarshalAs(UnmanagedType.Interface)] out IntPtr ppsi);
        void SetFileName([In, MarshalAs(UnmanagedType.LPWStr)] string pszName);
        void GetFileName([MarshalAs(UnmanagedType.LPWStr)] out string pszName);
        void SetTitle([In, MarshalAs(UnmanagedType.LPWStr)] string pszTitle);
        void SetOkButtonLabel([In, MarshalAs(UnmanagedType.LPWStr)] string pszText);
        void SetFileNameLabel([In, MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
        void GetResult([MarshalAs(UnmanagedType.Interface)] out IShellItem ppsi);
        void AddPlace([In, MarshalAs(UnmanagedType.Interface)] IntPtr psi, int fdap);
        void SetDefaultExtension([In, MarshalAs(UnmanagedType.LPWStr)] string pszDefaultExtension);
        void Close([MarshalAs(UnmanagedType.Error)] int hr);
        void SetClientGuid([In] ref Guid guid);
        void ClearClientData();
        void SetFilter([MarshalAs(UnmanagedType.Interface)] IntPtr pFilter);
        void GetResults([MarshalAs(UnmanagedType.Interface)] out IntPtr ppenum);
        void GetSelectedItems([MarshalAs(UnmanagedType.Interface)] out IntPtr ppsai);
    }

    public static string Pick(string title) {
        try {
            IFileOpenDialog dialog = (IFileOpenDialog)new FileOpenDialogImpl();
            uint options;
            dialog.GetOptions(out options);
            // FOS_PICKFOLDERS (0x20) | FOS_FORCEFILESYSTEM (0x40)
            dialog.SetOptions(options | 0x00000020 | 0x00000040);
            if (!string.IsNullOrEmpty(title)) dialog.SetTitle(title);
            int hr = dialog.Show(IntPtr.Zero);
            if (hr != 0) return null;
            IShellItem item;
            dialog.GetResult(out item);
            string path;
            item.GetDisplayName(0x80058000, out path); // SIGDN_FILESYSPATH
            return path;
        } catch { return null; }
    }
}
"@

$baseDir = [HighResFolderPicker]::Pick("Select the base destination folder for High tag files")
if ([string]::IsNullOrEmpty($baseDir)) {
    Write-Host "Cancelled."
    exit 0
}

# ── 3. Process each file ──
$errors = 0
foreach ($f in $files) {
    # Stem without extension: e.g. "X1C_CalicoCat_Farm_Final"
    $stem        = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
    $nameNoFinal = $stem -replace '(?i)_Final$', ''
    $parts       = $nameNoFinal -split '_'

    if ($parts.Count -lt 3) {
        Write-Warning "Skipping '$($f.Name)' - expected format Printer_Name_Theme_Final.3mf (got $($parts.Count) segment(s))"
        $errors++
        continue
    }

    $printer  = $parts[0]
    $theme    = $parts[-1]
    $charName = if ($parts.Count -gt 3) { ($parts[1..($parts.Count - 2)]) -join '_' } else { $parts[1] }

    $newBaseName = "${printer}_High${charName}_${theme}"
    $newFileName = "${newBaseName}_Final.3mf"
    $tagFolder   = "${printer}_High_${theme}"
    $destDir     = Join-Path (Join-Path $baseDir $tagFolder) $newBaseName
    $destFile    = Join-Path $destDir $newFileName

    try {
        New-Item -ItemType Directory -Path $destDir -Force | Out-Null
        Copy-Item -Path $f.FullName -Destination $destFile -Force

        # ── Patch the copied 3MF ──
        $zip = [System.IO.Compression.ZipFile]::Open($destFile, 'Update')
        try {
            # Clear VLH: remove layer_heights_profile.txt if present
            $vlhEntry = $zip.GetEntry('Metadata/layer_heights_profile.txt')
            if ($null -ne $vlhEntry) { $vlhEntry.Delete() }

            # Set base layer height to 0.09 in project_settings.config
            $cfgEntry = $zip.GetEntry('Metadata/project_settings.config')
            if ($null -ne $cfgEntry) {
                $reader  = New-Object System.IO.StreamReader($cfgEntry.Open())
                $content = $reader.ReadToEnd()
                $reader.Close()

                $patched = $content -replace '"layer_height"\s*:\s*"[^"]*"', '"layer_height": "0.09"'

                if ($patched -ne $content) {
                    $cfgEntry.Delete()
                    $newEntry = $zip.CreateEntry('Metadata/project_settings.config')
                    $writer   = New-Object System.IO.StreamWriter($newEntry.Open())
                    $writer.Write($patched)
                    $writer.Close()
                }
            }
        } finally {
            $zip.Dispose()
        }

        Write-Host "OK  $($f.Name)"
        Write-Host "    -> $tagFolder\$newBaseName\$newFileName"
    } catch {
        Write-Warning "Failed on '$($f.Name)': $_"
        $errors++
    }
}

Write-Host ""
if ($errors -eq 0) {
    Write-Host "Done! $($files.Count) file(s) copied."
} else {
    Write-Host "Done with $errors error(s). $($files.Count - $errors) of $($files.Count) file(s) copied."
}

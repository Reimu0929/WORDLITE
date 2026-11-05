# Capture-Word2026UI-FailingParts.ps1
# Word 2026 の各ダイアログを Win32 API (FindWindow + BitBlt) で直接キャプチャ
# 出力: C:\Users\%USERNAME%\source\repos\Word2026_Captures

param(
    [string]$OutputDir = "C:\Users\$env:USERNAME\source\repos\Word2026_Captures"
)

$ErrorActionPreference = "Stop"
if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

# ───────────────────────────────────────────────────────────────
# 1. Win32 API 読み込み (C# 2.0 互換の簡潔な構文)
# ───────────────────────────────────────────────────────────────
try {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Text;

public static class W32Api {
    [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int nCmdShow);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);

    public struct RECT { public int Left, Top, Right, Bottom; }
}
"@ -ErrorAction Stop
} catch {
    # 既存の型は無視
}

[W32Api]::SetProcessDPIAware() | Out-Null

try {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Text;

public class WinEnum {
    delegate bool Callback(IntPtr h, IntPtr l);
    [DllImport("user32.dll")] static extern bool EnumWindows(Callback cb, IntPtr l);
    [DllImport("user32.dll")] static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll")] static extern int GetWindowTextLength(IntPtr h);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);

    public static List<Window> All() {
        List<Window> list = new List<Window>();
        EnumWindows(delegate(IntPtr h, IntPtr l) {
            Window w = new Window();
            w.Handle = h;
            w.Visible = IsWindowVisible(h);
            int len = GetWindowTextLength(h);
            if (len > 0) {
                StringBuilder sb = new StringBuilder(len + 1);
                GetWindowText(h, sb, sb.Capacity);
                w.Title = sb.ToString();
            }
            list.Add(w);
            return true;
        }, IntPtr.Zero);
        return list;
    }

    public static IntPtr FindByTitle(string part) {
        foreach (Window w in All()) {
            if (w.Visible && w.Title != null && w.Title.Contains(part)) {
                return w.Handle;
            }
        }
        return IntPtr.Zero;
    }

    public class Window {
        public IntPtr Handle;
        public string Title;
        public bool Visible;
    }
}
"@ -ErrorAction Stop
} catch {
    # 既存の型は無視
}

Add-Type -AssemblyName System.Drawing -ErrorAction Stop

# ───────────────────────────────────────────────────────────────
# 2. 変数・ユーティリティ
# ───────────────────────────────────────────────────────────────
$script:WordHwnd = [IntPtr]::Zero
$script:WordRect = New-Object W32Api+RECT
$script:Shell = New-Object -ComObject wscript.shell

function Write-Log($level, $msg) {
    Write-Host "  [$level] $msg"
}

function Activate-Word() {
    if ($script:WordHwnd -ne [IntPtr]::Zero) {
        # 9 = SW_RESTORE: 最小化されていたら元に戻す
        [W32Api]::ShowWindow($script:WordHwnd, 9) | Out-Null
        Start-Sleep -Milliseconds 400
        [W32Api]::SetForegroundWindow($script:WordHwnd) | Out-Null
        Start-Sleep -Milliseconds 600
    }
}

function Find-WordWindow() {
    $proc = Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($proc -and $proc.MainWindowHandle -ne [IntPtr]::Zero) {
        $script:WordHwnd = $proc.MainWindowHandle
        [W32Api]::GetWindowRect($script:WordHwnd, [ref]$script:WordRect) | Out-Null
        return $script:WordHwnd
    }
    Write-Log "Warn" "Word window not found"
    return [IntPtr]::Zero
}

function List-VisibleWindows() {
    foreach ($w in [WinEnum]::All()) {
        if ($w.Visible -and $w.Title) {
            Write-Log "dlg" "$($w.Title)  (H=$($w.Handle.ToInt64()))"
        }
    }
}

function Capture-Window($hwnd, $filename) {
    if ($hwnd -eq [IntPtr]::Zero) { return $false }
    [W32Api]::SetForegroundWindow($hwnd) | Out-Null
    Start-Sleep -Milliseconds 400

    $r = New-Object W32Api+RECT
    [W32Api]::GetWindowRect($hwnd, [ref]$r) | Out-Null
    $w = $r.Right - $r.Left
    $h = $r.Bottom - $r.Top
    if ($w -le 0 -or $h -le 0) { return $false }

    $bmp = New-Object System.Drawing.Bitmap($w, $h)
    try {
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        try {
            $size = New-Object System.Drawing.Size($w, $h)
            $g.CopyFromScreen($r.Left, $r.Top, 0, 0, $size)
        } finally {
            $g.Dispose()
        }
    } finally {
        $out = Join-Path $OutputDir "$filename.png"
        $bmp.Save($out, [System.Drawing.Imaging.ImageFormat]::Png)
        $bmp.Dispose()
        Write-Log "Saved" $out
    }
    return $true
}

function Find-Dialog($titlePart) {
    $h = [WinEnum]::FindByTitle($titlePart)
    if ($h -ne [IntPtr]::Zero) {
        Write-Log "Found" "$titlePart -> H=$($h.ToInt64())"
    }
    return $h
}

function Send-Keys($sendKeys, $waitAfter) {
    # 通常の同時押しキー (Ctrl+H など)
    Activate-Word
    $script:Shell.SendKeys($sendKeys)
    Start-Sleep -Milliseconds $waitAfter
}

function Send-Ribbon($sequence, $finalWait) {
    # リボン操作: Alt を1回押して離し、指定文字を1つずつ送信
    # 例: "hbb" -> Alt, h, b, b
    Activate-Word
    $script:Shell.SendKeys("%")
    Start-Sleep -Milliseconds 700
    foreach ($c in $sequence.ToCharArray()) {
        $script:Shell.SendKeys([string]$c)
        Start-Sleep -Milliseconds 600
    }
    Start-Sleep -Milliseconds $finalWait
}

function Wait-Manual($prompt) {
    Write-Host ""
    Write-Host "  [MANUAL] $prompt"
    Write-Host "  開いたら Enter キーを押してください..."
    Read-Host | Out-Null
}

function Close-Dialog() {
    $script:Shell.SendKeys("{ESC}")
    Start-Sleep -Milliseconds 400
}

# ───────────────────────────────────────────────────────────────
# 3. 各ダイアログのキャプチャ処理
# ───────────────────────────────────────────────────────────────

function Capture-PageSetup() {
    Write-Host "`n[Phase 1] Page Setup"
    Send-Ribbon "p" 1500

    $h = Find-Dialog "ページ設定"
    if ($h -eq [IntPtr]::Zero) {
        Write-Log "Warn" "Page Setup not found"
        List-VisibleWindows
        Wait-Manual "ページ設定 ダイアログを手動で開いてください (例: [レイアウト]タブ → 右下の小さな矢印)"
        $h = Find-Dialog "ページ設定"
        if ($h -eq [IntPtr]::Zero) {
            Write-Log "Error" "Still not found. Skipping."
            return
        }
    }

    Capture-Window $h "04_dialog_page_setup"

    # タブ切り替え: Ctrl+Tab (無効な場合は手動)
    $tabs = @("余白", "用紙", "レイアウト")
    foreach ($t in $tabs) {
        $script:Shell.SendKeys("^{TAB}")
        Start-Sleep -Milliseconds 700
        Capture-Window $h "04_dialog_page_setup_$t"
    }

    Close-Dialog
}

function Capture-BordersShading() {
    Write-Host "`n[Phase 2] Borders and Shading"
    Send-Ribbon "hbb" 1500

    $h = Find-Dialog "枠線と網かけ"
    if ($h -eq [IntPtr]::Zero) {
        Write-Log "Warn" "Borders and Shading not found"
        List-VisibleWindows
        Wait-Manual "枠線と網かけ ダイアログを手動で開いてください (例: [ホーム]タブ → 枠線▼ → 枠線と網かけ)"
        $h = Find-Dialog "枠線と網かけ"
        if ($h -eq [IntPtr]::Zero) {
            Write-Log "Error" "Still not found. Skipping."
            return
        }
    }

    Capture-Window $h "04_dialog_borders_shading"
    $tabs = @("枠線", "網かけ", "ページ罫線")
    foreach ($t in $tabs) {
        $script:Shell.SendKeys("^{TAB}")
        Start-Sleep -Milliseconds 700
        Capture-Window $h "04_dialog_borders_shading_$t"
    }

    Close-Dialog
}

function Capture-Columns() {
    Write-Host "`n[Phase 3] Columns"
    Send-Ribbon "pw" 1500

    $h = Find-Dialog "段組"
    if ($h -eq [IntPtr]::Zero) {
        Write-Log "Warn" "Columns not found"
        List-VisibleWindows
        Wait-Manual "段組 ダイアログを手動で開いてください (例: [レイアウト]タブ → 段組)"
        $h = Find-Dialog "段組"
        if ($h -eq [IntPtr]::Zero) {
            Write-Log "Error" "Still not found. Skipping."
            return
        }
    }

    Capture-Window $h "04_dialog_columns"
    Close-Dialog
}

function Capture-Options() {
    Write-Host "`n[Phase 4] Options"
    Send-Ribbon "ft" 2500

    $h = Find-Dialog "オプション"
    if ($h -eq [IntPtr]::Zero) {
        Write-Log "Warn" "Options not found"
        List-VisibleWindows
        Wait-Manual "オプション ダイアログを手動で開いてください (例: [ファイル] → オプション)"
        $h = Find-Dialog "オプション"
        if ($h -eq [IntPtr]::Zero) {
            Write-Log "Error" "Still not found. Skipping."
            return
        }
    }

    Capture-Window $h "04_dialog_options"
    $tabs = @("一般", "表示", "校正", "保存", "言語", "詳細設定", "カスタマイズ", "リボンのユーザー設定")
    foreach ($t in $tabs) {
        $script:Shell.SendKeys("^{TAB}")
        Start-Sleep -Milliseconds 800
        Capture-Window $h "04_dialog_options_$t"
    }

    Close-Dialog
}

function Capture-Hyperlink() {
    Write-Host "`n[Phase 5] Hyperlink"
    Send-Keys "^k" 1500

    $h = Find-Dialog "ハイパーリンク"
    if ($h -eq [IntPtr]::Zero) {
        Write-Log "Warn" "Hyperlink not found"
        List-VisibleWindows
        Wait-Manual "ハイパーリンク ダイアログを手動で開いてください (例: Ctrl+K)"
        $h = Find-Dialog "ハイパーリンク"
        if ($h -eq [IntPtr]::Zero) {
            Write-Log "Error" "Still not found. Skipping."
            return
        }
    }

    Capture-Window $h "04_dialog_hyperlink"
    Close-Dialog
}

function Capture-Footnote() {
    Write-Host "`n[Phase 6] Footnote"
    Send-Ribbon "snd" 1500

    $h = Find-Dialog "脚注"
    if ($h -eq [IntPtr]::Zero) {
        Write-Log "Warn" "Footnote not found"
        List-VisibleWindows
        Wait-Manual "脚注 ダイアログを手動で開いてください (例: [参照]タブ → 脚注 → 小さな矢印)"
        $h = Find-Dialog "脚注"
        if ($h -eq [IntPtr]::Zero) {
            Write-Log "Error" "Still not found. Skipping."
            return
        }
    }

    Capture-Window $h "04_dialog_footnote"
    Close-Dialog
}

function Capture-StyleManager() {
    Write-Host "`n[Phase 7] Style Manager"
    Send-Ribbon "hoa" 1500

    $h = Find-Dialog "スタイルの管理"
    if ($h -eq [IntPtr]::Zero) {
        Write-Log "Warn" "Style Manager not found"
        List-VisibleWindows
        Wait-Manual "スタイルの管理 ダイアログを手動で開いてください (例: Alt+Ctrl+Shift+S → 管理)"
        $h = Find-Dialog "スタイルの管理"
        if ($h -eq [IntPtr]::Zero) {
            Write-Log "Error" "Still not found. Skipping."
            return
        }
    }

    Capture-Window $h "04_dialog_style_manager"
    Close-Dialog
}

function Capture-Editor() {
    Write-Host "`n[Phase 8] Editor / Accessibility"
    Send-Ribbon "ra" 2500

    $h = Find-Dialog "アクセシビリティ"
    if ($h -eq [IntPtr]::Zero) {
        Write-Log "Warn" "Accessibility not found"
        List-VisibleWindows
        Close-Dialog
        return
    }

    Capture-Window $h "05_scenario_editor_score_reviewtab"

    $ed = Find-Dialog "エディター"
    if ($ed -ne [IntPtr]::Zero) {
        Capture-Window $ed "05_scenario_editor_score_pane"
    }

    Close-Dialog
}

function Capture-Replace() {
    Write-Host "`n[Phase 9] Find and Replace"
    Send-Keys "^h" 1500

    $h = Find-Dialog "検索と置換"
    if ($h -eq [IntPtr]::Zero) {
        Write-Log "Warn" "Find and Replace not found"
        List-VisibleWindows
        Close-Dialog
        return
    }

    Capture-Window $h "05_scenario_replace_dialog"

    # 詳細オプションを開く (ボタンをスペースで押す想定)
    $script:Shell.SendKeys("{TAB}")
    Start-Sleep -Milliseconds 500
    $script:Shell.SendKeys(" ")
    Start-Sleep -Milliseconds 800
    Capture-Window $h "05_scenario_replace_dialog_options"

    Close-Dialog
}

# ───────────────────────────────────────────────────────────────
# 4. メイン処理
# ───────────────────────────────────────────────────────────────
Write-Host "========================================"
Write-Host "Word 2026 Dialog Capture (Win32 API)"
Write-Host "Output: $OutputDir"
Write-Host "========================================"

Find-WordWindow | Out-Null
if ($script:WordHwnd -eq [IntPtr]::Zero) {
    Write-Log "Error" "Word is not running. Please start Word 2026 first."
    exit 1
}

Capture-PageSetup
Capture-BordersShading
Capture-Columns
Capture-Options
Capture-Hyperlink
Capture-Footnote
Capture-StyleManager
Capture-Editor
Capture-Replace

Write-Host "`n========================================"
Write-Host "Capture complete."
Write-Host "Output: $OutputDir"
Write-Host "========================================"

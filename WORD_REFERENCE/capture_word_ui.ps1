#Requires -Version 5.1
# Capture-Word2026UI.ps1
# Word 2026 (WINWORD.EXE) のリボン UI とバックステージを自動でキャプチャする
# 実行: powershell.exe -ExecutionPolicy Bypass -File .\WORD_REFERENCE\capture_word_ui.ps1

$ErrorActionPreference = "Stop"
$OutputDir = "C:\Users\$env:USERNAME\source\repos\Word2026_Captures"
if (-not (Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }

# Win32 API / .NET の読み込み
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class WordCaptureApi {
  [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int c);
  [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
  [DllImport("user32.dll")] public static extern void mouse_event(uint f, uint x, uint y, uint d, IntPtr e);
  [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint flags);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  [DllImport("dwmapi.dll")] public static extern int DwmGetWindowAttribute(IntPtr hwnd, int attr, out RECT rect, int size);
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
}
"@

[WordCaptureApi]::SetProcessDPIAware() | Out-Null

# 定数
$AE = [System.Windows.Automation.AutomationElement]
$TS = [System.Windows.Automation.TreeScope]
$CT = [System.Windows.Automation.ControlType]

# タブ名の期待リスト（Word 2026 標準）
$ExpectedRibbonTabs = @("ホーム", "挿入", "デザイン", "レイアウト", "参考資料", "差し込み文書", "校閲", "表示", "ヘルプ")

# バックステージ左メニュー順序（Word 2026 スクリーンショット確認済: 02_backstage_home.png）
# 順序: ホーム→新規→開く→共有→情報→上書き保存→名前を付けて保存→印刷→エクスポート→閉じる→アカウント→その他...
$BackstageLeftItems = @("新規", "開く", "共有", "情報", "上書き保存", "名前を付けて保存", "印刷", "エクスポート", "閉じる", "アカウント", "その他...")

# 校閲タブのスプリットボタン展開ファイル名リスト（順序保証）
# 画像で確認済みの順序:
#   1: スペル チェックと文章校正
#   2: (元に戻せません)
#   3: アクセシビリティ チェック
#   4: 削除
#   5: コメントの表示
#   6: [変更履歴] ウィンドウ
#   7: インクを非表示にする
$ReviewTabSplits = @("スペル チェックと文章校正", "アクセシビリティ チェック", "削除", "コメントの表示", "すべての変更履歴/コメント", "変更履歴とコメントの表示", "[変更履歴] ウィンドウ", "変更履歴", "インクを非表示にする")

# 必須キャプチャ（順序付き）
$RequiredCaptures = @{
    "00_initial"                          = "初期画面（文書編集領域表示中）"
    "00_start_screen"                     = "スタート画面（ようこそ画面）"
    "02_backstage_home"                   = "バックステージ ホーム"
    "02_backstage_new"                    = "バックステージ 新規（テンプレート ギャラリー）"
    "02_backstage_new_scroll"             = "バックステージ 新規（スクロール後）"
    "02_backstage_closed"                 = "バックステージ閉じた後"
    "04_dialog_paragraph"                 = "段落 ダイアログ（基本タブ）"
    "04_dialog_paragraph_行とページ間の改行" = "段落 ダイアログ（行とページ間の改行）"
    "04_dialog_font"                      = "フォント ダイアログ（基本設定）"
    "04_dialog_font_詳細設定"             = "フォント ダイアログ（詳細設定）"
    "04_dialog_bullets_numbering"         = "箇条書きと段落番号 ダイアログ"
    "04_dialog_bullets_numbering_イメージ付き箇条書き" = "箇条書きと段落番号（イメージ付き箇条）"
    "04_dialog_page_setup"                = "ページ設定 ダイアログ（余白）"
    "04_dialog_page_setup_余白"           = "ページ設定 ダイアログ（余白）"
    "04_dialog_page_setup_用紙"           = "ページ設定 ダイアログ（用紙）"
    "04_dialog_page_setup_レイアウト"     = "ページ設定 ダイアログ（レイアウト）"
    "04_dialog_borders_shading"           = "枠線と網かけ ダイアログ（枠線）"
    "04_dialog_borders_shading_網かけ"    = "枠線と網かけ ダイアログ（網かけ）"
    "04_dialog_borders_shading_ページ罫線" = "枠線と網かけ ダイアログ（ページ罫線）"
    "04_dialog_columns"                   = "段組 ダイアログ"
    "04_dialog_style_pane"                = "スタイル ウィンドウ"
    "04_dialog_style_manager"             = "スタイルの管理 ダイアログ"
    "04_dialog_options"                   = "オプション ダイアログ（全般）"
    "04_dialog_options_表示"              = "オプション ダイアログ（表示）"
    "04_dialog_options_校正"              = "オプション ダイアログ（校正）"
    "04_dialog_options_保存"              = "オプション ダイアログ（保存）"
    "04_dialog_options_言語"              = "オプション ダイアログ（言語）"
    "04_dialog_options_詳細設定"          = "オプション ダイアログ（詳細設定）"
    "04_dialog_options_customization"     = "オプション ダイアログ（カスタマイズ）"
    "04_dialog_hyperlink"                 = "ハイパーリンク ダイアログ"
    "04_dialog_footnote"                  = "脚注/注釈 ダイアログ"
    "05_scenario_editor_score_pane"       = "エディター スコア ペイン"
    "05_scenario_search_pane"             = "検索パネル"
    "05_scenario_replace_dialog"          = "検索と置換 ダイアログ"
    "05_scenario_replace_dialog_opts"     = "検索と置換（オプション展開）"
}

# グローバル状態
$script:WordWindow = $null
$script:WinRect = New-Object WordCaptureApi+RECT

# ============================================================
# ユーティリティ関数
# ============================================================

function Open-BlankDocumentIfNeeded() {
    # Word 2026 では Ctrl+N は「新規」タブ（バックステージ）を開くだけで、
    # 自動的には空文書が作成されない。「白紙の文書」テンプレートを明示的にクリックする。
    $doc = $script:WordWindow.FindFirst($TS::Descendants,
        (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $CT::Document)))
    if (-not $doc) {
        Write-Host "  [Info] 文書編集領域が見つからないため、「白紙の文書」を開きます。"
        Open-ToBlankDocument
        Start-Sleep -Milliseconds 2000
    }
}

# Word 2026 の新規作成フロー: バックステージ → 新規 → 白紙の文書
# Ctrl+N は「新規」タブを表示するだけで、白紙の文書は自動で作成されない
function Open-ToBlankDocument() {
    $wshell = New-Object -ComObject wscript.shell
    [WordCaptureApi]::SetForegroundWindow($script:WordWindow.Current.NativeWindowHandle) | Out-Null
    Start-Sleep -Milliseconds 300

    # バックステージを開く (Alt+F)
    $wshell.SendKeys("%f")
    Start-Sleep -Milliseconds 1500

    # 新規テンプレートをクリック
    $newBtn = Find-ButtonByName "新規"
    if (-not $newBtn) { $newBtn = Find-HyperlinkByName "新規" }
    if ($newBtn) {
        Invoke-SafeClick $newBtn
        Start-Sleep -Milliseconds 1500

        # 「白紙の文書」テンプレートをクリック
        $blankDoc = Find-ButtonByName "白紙の文書"
        if (-not $blankDoc) { $blankDoc = Find-HyperlinkByName "白紙の文書" }
        if ($blankDoc) {
            Invoke-SafeClick $blankDoc
            Start-Sleep -Milliseconds 1500
            Write-Host "  [OK] 白紙の文書を開きました。"
            return
        } else {
            Write-Host "  [Warn] 「白紙の文書」ボタンが見つかりません。Alt+N を試します。"
        }
    }

    # フォールバック: Alt+N (新規作成ショートカット)
    Write-Host "  [Info] フォールバック: Alt+N を送信"
    Close-Backstage
    Start-Sleep -Milliseconds 300
    [WordCaptureApi]::SetForegroundWindow($script:WordWindow.Current.NativeWindowHandle) | Out-Null
    Start-Sleep -Milliseconds 300
    $wshell.SendKeys("%n")
    Start-Sleep -Milliseconds 2500
}

function Get-WordMainWindow() {
    $proc = Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $proc) { throw "WINWORD プロセスが見つかりません。Word 2026 を起動してください。" }
    [WordCaptureApi]::ShowWindow($proc.MainWindowHandle, 9) | Out-Null
    Start-Sleep -Milliseconds 400
    [WordCaptureApi]::SetWindowPos($proc.MainWindowHandle, [IntPtr](-1), 60, 20, 1784, 972, 0x0040) | Out-Null
    [WordCaptureApi]::SetForegroundWindow($proc.MainWindowHandle) | Out-Null
    Start-Sleep -Milliseconds 1000

    $r = New-Object WordCaptureApi+RECT
    [WordCaptureApi]::GetWindowRect($proc.MainWindowHandle, [ref]$r) | Out-Null
    $f = New-Object WordCaptureApi+RECT
    if ([WordCaptureApi]::DwmGetWindowAttribute($proc.MainWindowHandle, 9, [ref]$f, 16) -eq 0) { $r = $f }
    $script:WinRect = $r

    $root = $AE::RootElement
    $cond = New-Object System.Windows.Automation.PropertyCondition($AE::ProcessIdProperty, $proc.Id)
    $script:WordWindow = $root.FindFirst($TS::Children, $cond)
    if (-not $script:WordWindow) { throw "UI Automation で Word ウィンドウが見つかりません。" }
    Write-Host "Word ウィンドウを取得: PID=$($proc.Id), Rect=($($r.Left),$($r.Top),$($r.Right),$($r.Bottom))"
}

function Invoke-SafeClick($element, [switch]$Double) {
    if (-not $element) { return }
    try {
        $pt = $element.Current.BoundingRectangle
        if ($pt.Width -le 0 -or $pt.Height -le 0) { return }
        $x = [int]($pt.X + $pt.Width / 2)
        $y = [int]($pt.Y + $pt.Height / 2)
        [WordCaptureApi]::SetCursorPos($x, $y) | Out-Null
        Start-Sleep -Milliseconds 300
        [WordCaptureApi]::mouse_event(0x0002, 0, 0, 0, [IntPtr]::Zero)
        Start-Sleep -Milliseconds 120
        [WordCaptureApi]::mouse_event(0x0004, 0, 0, 0, [IntPtr]::Zero)
        if ($Double) {
            Start-Sleep -Milliseconds 100
            [WordCaptureApi]::mouse_event(0x0002, 0, 0, 0, [IntPtr]::Zero)
            Start-Sleep -Milliseconds 120
            [WordCaptureApi]::mouse_event(0x0004, 0, 0, 0, [IntPtr]::Zero)
        }
        Start-Sleep -Milliseconds 200
    } catch {}
}

function Take-Screenshot($name) {
    # カーソルを画面外に移動してから撮影
    [WordCaptureApi]::SetCursorPos($script:WinRect.Right + 200, $script:WinRect.Bottom + 200) | Out-Null
    Start-Sleep -Milliseconds 1200

    $w = $script:WinRect.Right - $script:WinRect.Left
    $h = $script:WinRect.Bottom - $script:WinRect.Top
    $bmp = New-Object System.Drawing.Bitmap $w, $h
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($script:WinRect.Left, $script:WinRect.Top, 0, 0, $bmp.Size)

    $path = Join-Path $OutputDir "$name.png"
    $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose(); $bmp.Dispose()
    Write-Host "  [Saved] $path"
}

function Find-VisibleChildren($parent, $controlType, $maxDepth = 10) {
    $results = @()
    if (-not $parent) { return $results }
    try {
        $cond = New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $controlType)
        $found = $parent.FindAll($TS::Descendants, $cond)
        foreach ($e in $found) {
            try {
                $rect = $e.Current.BoundingRectangle
                if (-not $e.Current.IsOffscreen -and $rect.Width -gt 0 -and $rect.Height -gt 0) {
                    $results += $e
                }
            } catch {}
        }
    } catch {}
    return $results | Sort-Object { $_.Current.BoundingRectangle.X }
}

function Find-ButtonByName($name) {
    $all = Find-VisibleChildren $script:WordWindow $CT::Button
    foreach ($e in $all) {
        try {
            if ($e.Current.Name -eq $name) { return $e }
        } catch {}
    }
    return $null
}

function Find-HyperlinkByName($name) {
    $all = Find-VisibleChildren $script:WordWindow $CT::Hyperlink
    foreach ($e in $all) {
        try {
            if ($e.Current.Name -eq $name) { return $e }
        } catch {}
    }
    return $null
}

function Find-ToggleButtonByName($name) {
    $all = Find-VisibleChildren $script:WordWindow $CT::ToggleButton
    foreach ($e in $all) {
        try {
            if ($e.Current.Name -eq $name) { return $e }
        } catch {}
    }
    return $null
}

function Find-TabItemByNameGeneral($name) {
    $all = Find-VisibleChildren $script:WordWindow $CT::TabItem
    foreach ($e in $all) {
        try {
            if ($e.Current.Name -eq $name) { return $e }
        } catch {}
    }
    return $null
}

function Find-AnyByName($name) {
    # Button → ToggleButton → Hyperlink → TabItem の順で探索
    $btn = Find-ButtonByName $name
    if ($btn) { return $btn }
    $tbtn = Find-ToggleButtonByName $name
    if ($tbtn) { return $tbtn }
    $hl = Find-HyperlinkByName $name
    if ($hl) { return $hl }
    $ti = Find-TabItemByNameGeneral $name
    if ($ti) { return $ti }
    # ControlType 非依存の最終手段: 名前部分一致で全要素探索
    try {
        $cond = New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, $name)
        $all = $script:WordWindow.FindAll($TS::Descendants, $cond)
        foreach ($e in $all) {
            try {
                $rect = $e.Current.BoundingRectangle
                if (-not $e.Current.IsOffscreen -and $rect.Width -gt 0) {
                    return $e
                }
            } catch {}
        }
    } catch {}
    return $null
}

# ダイアログ ウィンドウを検出し、その中で要素を探す関数
function Find-DialogElement($dialogTitle, $controlType, $name) {
    try {
        $root = $AE::RootElement
        
        # まず $script:WordWindow 内を探索（Word 2026 のダイアログはここにある）
        $dlg = $script:WordWindow.FindFirst($TS::Descendants,
            (New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, $dialogTitle)))
        if (-not $dlg) {
            # 名前一致の Window 要素を探す
            $allWindows = $script:WordWindow.FindAll($TS::Descendants,
                (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $CT::Window)))
            foreach ($w in $allWindows) {
                try {
                    if ($w.Current.Name -like "*$dialogTitle*") { $dlg = $w; break }
                } catch {}
            }
        }
        # 見つからなければ RootElement のトップレベルも探索
        if (-not $dlg) {
            $dlg = $root.FindFirst($TS::Children,
                (New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, $dialogTitle)))
        }
        if (-not $dlg) {
            $allDlgs = $root.FindAll($TS::Children,
                (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $CT::Window)))
            foreach ($d in $allDlgs) {
                try {
                    if ($d.Current.Name -like "*$dialogTitle*") { $dlg = $d; break }
                } catch {}
            }
        }
        if (-not $dlg) { return $null }

        # ControlType で探索 → ループ内で Name でフィルタ
        $c = New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $controlType)
        $all = $dlg.FindAll($TS::Descendants, $c)
        foreach ($e in $all) {
            try {
                $rect = $e.Current.BoundingRectangle
                if (-not $e.Current.IsOffscreen -and $rect.Width -gt 0 -and $e.Current.Name -eq $name) {
                    return $e
                }
            } catch {}
        }
    } catch {}
    return $null
}

# ダイアログ内のタブを探す - $script:WordWindow 内も探索
function Find-DialogTabByName($dialogTitle, $tabName) {
    try {
        $root = $AE::RootElement
        $dlg = $null
        
        # まず $script:WordWindow 内を検索（Word 2026 のダイアログはここ）
        $dlg = $script:WordWindow.FindFirst($TS::Descendants,
            (New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, $dialogTitle)))
        if (-not $dlg) {
            $allWins = $script:WordWindow.FindAll($TS::Descendants,
                (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $CT::Window)))
            foreach ($w in $allWins) {
                try { if ($w.Current.Name -like "*$dialogTitle*") { $dlg = $w; break } } catch {}
            }
        }
        # 次に RootElement のトップレベルも探索
        if (-not $dlg) {
            $dlg = $root.FindFirst($TS::Children,
                (New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, $dialogTitle)))
        }
        if (-not $dlg) {
            $allDlgs = $root.FindAll($TS::Children,
                (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $CT::Window)))
            foreach ($d in $allDlgs) {
                try {
                    if ($d.Current.Name -like "*$dialogTitle*") { $dlg = $d; break }
                } catch {}
            }
        }
        if (-not $dlg) {
            Write-Host "  [Debug] ダイアログ '$dialogTitle' が見つかりません"
            $allDlgs2 = $root.FindAll($TS::Children,
                (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $CT::Window)))
            $cnt = 0
            foreach ($d2 in $allDlgs2) {
                try { if ($d2.Current.Name -and $cnt -lt 20) { Write-Host "    [dlg] '$($d2.Current.Name)'"; $cnt++ } } catch {}
            }
            # WordWindow 内の Window 要素も列挙
            $wWins = $script:WordWindow.FindAll($TS::Descendants,
                (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $CT::Window)))
            if ($wWins.Count -gt 0) {
                Write-Host "  [Debug] WordWindow 内の Window($($wWins.Count)):"
                $cnt2 = 0
                foreach ($ww in $wWins) {
                    try { if ($ww.Current.Name -and $cnt2 -lt 30) { Write-Host "    [win] '$($ww.Current.Name)'"; $cnt2++ } } catch {}
                }
            }
            return $null
        }
        
        # タブを複数の ControlType で探索
        foreach ($ctName in @($CT::Button, $CT::ToggleButton, $CT::Hyperlink, $CT::ListItem)) {
            $c = New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $ctName)
            $all = $dlg.FindAll($TS::Descendants, $c)
            foreach ($e in $all) {
                try {
                    $rect = $e.Current.BoundingRectangle
                    if (-not $e.Current.IsOffscreen -and $rect.Width -gt 0 -and $e.Current.Name -eq $tabName) {
                        return $e
                    }
                } catch {}
            }
        }
        # それでもなければ名前部分一致で探索
        $cond = New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, $tabName)
        $all = $dlg.FindAll($TS::Descendants, $cond)
        foreach ($e in $all) {
            try {
                $rect = $e.Current.BoundingRectangle
                if (-not $e.Current.IsOffscreen -and $rect.Width -gt 0) { return $e }
            } catch {}
        }
    } catch { Write-Host "  [Warn] Find-DialogTabByName 失敗: $_" }
    return $null
}

# 現在の Word ウィンドウを前面に出し、UI 状態をリセット
function Reset-WordFocus() {
    $wshell = New-Object -ComObject wscript.shell
    try {
        # ダイアログ/メニューを閉じる
        $wshell.SendKeys("{ESC}")
        Start-Sleep -Milliseconds 300
        $wshell.SendKeys("{ESC}")
        Start-Sleep -Milliseconds 300
        # 前面ウィンドウを再取得して Word を前面に
        try {
            $proc = Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($proc) {
                [WordCaptureApi]::SetForegroundWindow($proc.MainWindowHandle) | Out-Null
            }
        } catch {}
        Start-Sleep -Milliseconds 500
        # UIA ツリーを再取得
        try {
            $proc2 = Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($proc2) {
                $root2 = $AE::RootElement
                $cond2 = New-Object System.Windows.Automation.PropertyCondition($AE::ProcessIdProperty, $proc2.Id)
                $script:WordWindow = $root2.FindFirst($TS::Children, $cond2)
                Write-Host "  [Info] UIA ツリーを再取得 (PID=$($proc2.Id))"
            }
        } catch { Write-Host "  [Warn] UIA ツリー再取得失敗: $_" }
    } catch { Write-Host "  [Warn] Reset-WordFocus 失敗: $_" }
    # ホーム タブを表示して状態確認
    try {
        $homeTab = Find-TabByName "ホーム"
        if ($homeTab) {
            Invoke-SafeClick $homeTab
            Start-Sleep -Milliseconds 500
        }
    } catch {}
}

function Find-TabByName($name) {
    # ControlType で探索 → ループ内で Name でフィルタ
    $all = Find-VisibleChildren $script:WordWindow $CT::TabItem
    foreach ($e in $all) {
        try {
            if ($e.Current.Name -eq $name) { return $e }
        } catch {}
    }

    # Button としてフォールバック
    $btn = Find-ButtonByName $name
    if ($btn) { return $btn }

    return $null
}

function Find-TabItemByName($name) {
    $all = Find-VisibleChildren $script:WordWindow $CT::TabItem
    foreach ($e in $all) {
        try {
            if ($e.Current.Name -eq $name) { return $e }
        } catch {}
    }
    return $null
}

function Close-Backstage() {
    # エスケープキーでバックステージを閉じる試み
    $wshell = New-Object -ComObject wscript.shell
    $wshell.SendKeys("{ESC}")
    Start-Sleep -Milliseconds 500
}

function Close-Dialog() {
    # ESC または Cancel/X でダイアログを閉じる
    $wshell = New-Object -ComObject wscript.shell
    $wshell.SendKeys("{ESC}")
    Start-Sleep -Milliseconds 500
}

function Invoke-RightClick($x, $y) {
    [WordCaptureApi]::SetCursorPos($x, $y) | Out-Null
    Start-Sleep -Milliseconds 300
    [WordCaptureApi]::mouse_event(0x0002, 0, 0, 0, [IntPtr]::Zero)
    Start-Sleep -Milliseconds 120
    [WordCaptureApi]::mouse_event(0x0008, 0, 0, 0, [IntPtr]::Zero)
    Start-Sleep -Milliseconds 120
    [WordCaptureApi]::mouse_event(0x0010, 0, 0, 0, [IntPtr]::Zero)
    Start-Sleep -Milliseconds 1000
}

function Close-AllMenus() {
    $wshell = New-Object -ComObject wscript.shell
    $wshell.SendKeys("{ESC}")
    Start-Sleep -Milliseconds 400
}

# ============================================================
# Phase 0: スタート画面 / ようこそ画面のキャプチャ
# ============================================================
function Capture-StartScreen() {
    Write-Host "`n[Phase 0] スタート画面のキャプチャを開始"

    # WinWORD を起動して、ようこそ画面（バックステージ）を表示
    # 既定では新しい文書が表示されるため、Alt+F でバックステージ ホームを表示
    $wshell = New-Object -ComObject wscript.shell
    [WordCaptureApi]::SetForegroundWindow($script:WordWindow.Current.NativeWindowHandle) | Out-Null
    Start-Sleep -Milliseconds 300
    $wshell.SendKeys("%f")
    Start-Sleep -Milliseconds 1500

    # バックステージ ホーム（最近使ったファイル）
    Take-Screenshot "00_start_backstage"

    # 新規テンプレート
    $newBtn = Find-ButtonByName "新規"
    if (-not $newBtn) { $newBtn = Find-HyperlinkByName "新規" }
    if ($newBtn) {
        Invoke-SafeClick $newBtn
        Start-Sleep -Milliseconds 1500
        Take-Screenshot "00_start_templates"

        # バックステージの「新規」タブ: テンプレート一覧を撮影
        Take-Screenshot "02_backstage_new"

        # 「白紙の文書」をクリックして実際に空文書を開く
        $blankDoc = Find-ButtonByName "白紙の文書"
        if (-not $blankDoc) { $blankDoc = Find-HyperlinkByName "白紙の文書" }
        if ($blankDoc) {
            Write-Host "  [Info] 白紙の文書テンプレートをクリック"
            Invoke-SafeClick $blankDoc
            Start-Sleep -Milliseconds 1500
        }
    }

    # バックステージを閉じて、すでに空文書が開いていれば初期状態を撮影
    Close-Backstage
    Start-Sleep -Milliseconds 800

    # 文書編集領域を確認
    $doc = $script:WordWindow.FindFirst($TS::Descendants,
        (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $CT::Document)))
    if (-not $doc) {
        Write-Host "  [Info] まだ文書が開いていないため、Open-ToBlankDocument を実行"
        Open-ToBlankDocument
        Start-Sleep -Milliseconds 1500
    }
    Take-Screenshot "00_initial"
}

# ============================================================
# Phase 1: リボン タブのキャプチャ（各タブ + 主要スプリットボタン）
# ============================================================
function Capture-RibbonTabs() {
    Write-Host "`n[Phase 1] リボン タブのキャプチャを開始"

    foreach ($tabName in $ExpectedRibbonTabs) {
        Write-Host "  [Tab] $tabName"
        $tab = Find-TabByName $tabName
        if (-not $tab) {
            Write-Host "  [Skip] タブが見つかりません: $tabName"
            continue
        }
        Invoke-SafeClick $tab
        Start-Sleep -Milliseconds 800
        Take-Screenshot "01_ribbon_$($tabName)"

        # === 特定タブの特殊キャプチャ ===
        if ($tabName -eq "挿入") {
            # 3D モデル → オンライン3D モデル ダイアログ
            $modelBtn = Find-ButtonByName "3D モデル"
            if ($modelBtn) {
                Invoke-SafeClick $modelBtn
                Start-Sleep -Milliseconds 1000
                Take-Screenshot "01_ribbon_挿入_split_3D_models"
                Close-AllMenus
            }

            # ハイパーリンク ダイアログ
            $linkBtn = Find-ButtonByName "リンク"
            if ($linkBtn) {
                Invoke-SafeClick $linkBtn
                Start-Sleep -Milliseconds 1000
                Take-Screenshot "01_ribbon_挿入_split_hyperlink"
                Close-AllMenus
            }

            # オブジェクトの挿入 ダイアログ
            $objBtn = Find-ButtonByName "オブジェクト"
            if ($objBtn) {
                Invoke-SafeClick $objBtn
                Start-Sleep -Milliseconds 1000
                Take-Screenshot "01_ribbon_挿入_split_object"
                Close-AllMenus
            }

            # 署名欄 ダイアログ
            $sigBtn = Find-ButtonByName "署名欄"
            if ($sigBtn) {
                Invoke-SafeClick $sigBtn
                Start-Sleep -Milliseconds 1000
                Take-Screenshot "01_ribbon_挿入_split_signature"
                Close-AllMenus
            }
        }

        if ($tabName -eq "校閲") {
            # 校閲タブ固有のスプリットボタンを順序どおりにキャプチャ
            # 1. スペル チェックと文章校正
            $spellBtn = Find-ButtonByName "スぺル チェックと文章校正"
            if ($spellBtn) {
                Invoke-SafeClick $spellBtn
                Start-Sleep -Milliseconds 1000
                Take-Screenshot "01_ribbon_校閲_split_1_スペル チェックと文章校正"
                Close-AllMenus
            }

            # 2. アクセシビリティ チェック ドロップダウン
            $a11yBtn = Find-ButtonByName "アクセシビリティ チェック"
            if ($a11yBtn) {
                Invoke-SafeClick $a11yBtn
                Start-Sleep -Milliseconds 1000
                Take-Screenshot "01_ribbon_校閲_split_アクセシビリティ チェック"
                Close-AllMenus
            }

            # 3. コメントの表示 ドロップダウン
            $commentBtn = Find-ButtonByName "コメントの表示"
            if ($commentBtn) {
                Invoke-SafeClick $commentBtn
                Start-Sleep -Milliseconds 1000
                Take-Screenshot "01_ribbon_校閲_split_コメントの表示"
                Close-AllMenus
            }

            # 4. [変更履歴] ウィンドウ ドロップダウン → 変更履歴ペイン表示
            $revHistBtn = Find-ButtonByName "[変更履歴] ウィンドウ"
            if ($revHistBtn) {
                Invoke-SafeClick $revHistBtn
                Start-Sleep -Milliseconds 1000
                Take-Screenshot "01_ribbon_校閲_split_変更履歴ウィンドウ"
            }

            # 5. インクを非表示にする ドロップダウン
            $inkBtn = Find-ButtonByName "インクを非表示にする"
            if ($inkBtn) {
                Invoke-SafeClick $inkBtn
                Start-Sleep -Milliseconds 1000
                Take-Screenshot "01_ribbon_校閲_split_インクを非表示にする"
                Close-AllMenus
            }
        }

        if ($tabName -eq "表示") {
            # マクロ ドロップダウン
            $macroBtn = Find-ButtonByName "マクロ"
            if ($macroBtn) {
                Invoke-SafeClick $macroBtn
                Start-Sleep -Milliseconds 1000
                Take-Screenshot "01_ribbon_表示_split_マクロ"
                Close-AllMenus
            }

            # イマーシブ リーダー ボタン（存在する場合）
            $immersiveBtn = Find-ButtonByName "イマーシブ リーダー"
            if ($immersiveBtn) {
                # 開かない（別のウィンドウが開く可能性がある）のでスナップショットのみ
                Take-Screenshot "01_ribbon_表示_イマーシブリーダーステータス"
            }
        }

        if ($tabName -eq "ホーム") {
            # スタイル ドロップダウン（右端の▼）
            $styleBtn = Find-ButtonByName "スタイル"
            if (-not $styleBtn) {
                # 「標準」「見出し 1」などのスタイルプレビューから探す
                $allButtons = Find-VisibleChildren $script:WordWindow $CT::Button
                foreach ($b in $allButtons) {
                    try {
                        if ($b.Current.Name -match "スタイル|標準|見出し") {
                            $rect = $b.Current.BoundingRectangle
                            if ($rect.X -gt 1200) { # リボンの右側
                                Invoke-SafeClick $b
                                Start-Sleep -Milliseconds 1000
                                Take-Screenshot "01_ribbon_ホーム_split_スタイルギャラリー"
                                Close-AllMenus
                                break
                            }
                        }
                    } catch {}
                }
            }
        }

        if ($tabName -eq "参考資料") {
            # 脚注 ドロップダウン
            $footnoteBtn = Find-ButtonByName "脚注の挿入"
            if ($footnoteBtn) {
                Invoke-SafeClick $footnoteBtn
                Start-Sleep -Milliseconds 1000
                Take-Screenshot "01_ribbon_参考資料_split_脚注"
                Close-AllMenus
            }

            # 資料文献の管理
            $citeBtn = Find-ButtonByName "資料文献の管理"
            if ($citeBtn) {
                Invoke-SafeClick $citeBtn
                Start-Sleep -Milliseconds 1000
                Take-Screenshot "01_ribbon_参考資料_split_資料文献管理"
                Close-AllMenus
            }
        }

        if ($tabName -eq "デザイン") {
            # 数式插入 → 数式タブ表示
            $eqBtn = Find-ButtonByName "数式"
            if ($eqBtn) {
                Invoke-SafeClick $eqBtn
                Start-Sleep -Milliseconds 800
                # 数式タブが追加表示される
                Take-Screenshot "01_ribbon_デザイン_split_数式挿入"
                # 数式を選択状態にしてデザインタブのスタイルギャラリー表示
                Take-Screenshot "01_ribbon_デザイン_split_スタイルギャラリー"
            }
        }

        if ($tabName -eq "差し込み文書") {
            # 差し込みフィールドの挿入 ドロップダウン
            $fieldBtn = Find-ButtonByName "差し込みフィールドの挿入"
            if ($fieldBtn) {
                Invoke-SafeClick $fieldBtn
                Start-Sleep -Milliseconds 1000
                Take-Screenshot "01_ribbon_差し込み文書_split_差し込みフィールドの挿入"
                Close-AllMenus
            }
        }

        # 一般的な SplitButton を列挙
        $ribbon = Find-RibbonContainer
        $splitButtons = Find-VisibleChildren $ribbon $CT::SplitButton
        $sIdx = 0
        foreach ($sb in $splitButtons) {
            $sIdx++
            try {
                $n = $sb.Current.Name
                if ([string]::IsNullOrWhiteSpace($n)) { continue }
                # ディクテーション は明示的にスキップ
                if ($n -eq "ディクテーション" -or $n -like "*ディクテーション*") { continue }
                # すでにキャプチャ済みかどうかをファイル名でチェック
                $safeName = ($n -replace '[\\/:*?"<>|]', '_')
                # 校閲タブの既知キャプチャと重複しないかチェック
                $knownCaptures = @(
                    "01_ribbon_校閲_split_1_スペル チェックと文章校正",
                    "01_ribbon_校閲_split_アクセシビリティ チェック",
                    "01_ribbon_校閲_split_コメントの表示",
                    "01_ribbon_校閲_split_変更履歴ウィンドウ",
                    "01_ribbon_校閲_split_インクを非表示にする",
                    "01_ribbon_表示_split_マクロ",
                    "01_ribbon_ホーム_split_10_ディクテーション"
                )
                $isKnown = $false
                foreach ($k in $knownCaptures) {
                    if ($sb.Current.Name -match ($k -replace '_', ' ')) { $isKnown = $true; break }
                }

                # すでに特定キャプチャが存在するかチェック
                $expectedNames = @(
                    "01_ribbon_ホーム_split_1_貼り付け",
                    "01_ribbon_ホーム_split_2_下線",
                    "01_ribbon_ホーム_split_3_元に戻す 下線のスタイル",
                    "01_ribbon_ホーム_split_4_蛍光ペンの色",
                    "01_ribbon_ホーム_split_5_フォントの色",
                    "01_ribbon_ホーム_split_6_箇条書き",
                    "01_ribbon_ホーム_split_7_段落番号",
                    "01_ribbon_ホーム_split_8_塗りつぶし",
                    "01_ribbon_ホーム_split_9_罫線",
                    "01_ribbon_挿入_split_3D モデル",
                    "01_ribbon_挿入_split_4_リンク",
                    "01_ribbon_挿入_split_5_オブジェクト...",
                    "01_ribbon_挿入_split_6_署名欄",
                    "01_ribbon_参考資料_split_1_元に戻す 文字の塗りつぶし",
                    "01_ribbon_参考資料_split_2_次の脚注",
                    "01_ribbon_レイアウト_split_1_元に戻す 既定の箇条書きの設定",
                    "01_ribbon_レイアウト_split_2_前面へ移動",
                    "01_ribbon_レイアウト_split_3_背面へ移動",
                    "01_ribbon_校閲_split_1_スペル チェックと文章校正",
                    "01_ribbon_校閲_split_2_元に戻せません",
                    "01_ribbon_校閲_split_3_アクセシビリティ チェック",
                    "01_ribbon_校閲_split_4_削除",
                    "01_ribbon_校閲_split_5_コメントの表示",
                    "01_ribbon_校閲_split_6_[変更履歴] ウィンドウ",
                    "01_ribbon_校閲_split_7_インクを非表示にする",
                    "01_ribbon_ヘルプ_split_1_元に戻せません",
                    "01_ribbon_差し込み文書_split_1_元に戻せません",
                    "01_ribbon_差し込み文書_split_2_差し込みフィールドの挿入",
                    "01_ribbon_表示_split_1_元に戻せません",
                    "01_ribbon_表示_split_2_マクロ"
                )
                $filePrefix = "01_ribbon_$tabName_split_"
                $allFiles = Get-ChildItem $OutputDir -Filter "$filePrefix*" -ErrorAction SilentlyContinue | ForEach-Object { $_.Name }
                if ($allFiles) {
                    foreach ($f in $allFiles) {
                        if ($f -match "($n)") { $isKnown = $true; break }
                    }
                }

                if ($isKnown) { continue }

                if ($safeName -match '_') { continue } # 名前に_があるものは既知の可能性

                $safeName = ($n -replace '[\\/:*?"<>|]', '_')
                Invoke-SafeClick $sb
                Start-Sleep -Milliseconds 700
                Take-Screenshot "01_ribbon_$($tabName)_split_$($sIdx)_$safeName"
                Close-AllMenus
            } catch {}
        }

        # 通常ボタンのうち、ドロップダウンを開く可能性のあるもの
        $buttons = Find-VisibleChildren $ribbon $CT::Button | Where-Object {
            try { $_.Current.Name -match '[▼v↓]|More|追加|その他' } catch { $false }
        }
        $bidx = 0
        foreach ($btn in $buttons) {
            $bidx++
            try {
                $n = $btn.Current.Name
                if ([string]::IsNullOrWhiteSpace($n)) { continue }
                $safeName = ($n -replace '[\\/:*?"<>|]', '_')
                # 既知のキャプチャと重複チェック
                $expectedNames = @("01_ribbon_ホーム_split_10_ディクテーション", "01_ribbon_表示_split_マクロ")
                $isKnown = $false
                foreach ($en in $expectedNames) {
                    if ($safeName -match ($en.Replace('_', '\_'))) { $isKnown = $true; break }
                }
                if ($isKnown) { continue }

                Invoke-SafeClick $btn
                Start-Sleep -Milliseconds 700
                Take-Screenshot "01_ribbon_$($tabName)_btn_$($bidx)_$safeName"
                Close-AllMenus
            } catch {}
        }
    }
}

function Find-RibbonContainer() {
    $homeTab = Find-TabByName "ホーム"
    if (-not $homeTab) { return $script:WordWindow }
    $current = $homeTab
    for ($i = 0; $i -lt 10; $i++) {
        try { $parent = [System.Windows.Automation.TreeWalker]::RawViewParent.GetParent($current) } catch { break }
        if (-not $parent) { break }
        try {
            $ct = $parent.Current.ControlType.ProgrammaticName
            if ($ct -eq "ControlType.ToolBar" -or $ct -eq "ControlType.Tab") {
                return $parent
            }
        } catch {}
        $current = $parent
    }
    return $script:WordWindow
}

# ============================================================
# Phase 1b: ホームタブの主要スプリットボタン追加キャプチャ
# ============================================================
function Capture-HomeSplitButtons() {
    Write-Host "`n[Phase 1b] ホーム タブの主要スプリットボタンを追加キャプチャ"

    $homeTab = Find-TabByName "ホーム"
    if ($homeTab) {
        Invoke-SafeClick $homeTab
        Start-Sleep -Milliseconds 800
    }

    # 貼り付け
    $pasteBtn = Find-ButtonByName "貼り付け"
    if ($pasteBtn) {
        Invoke-SafeClick $pasteBtn
        Start-Sleep -Milliseconds 1000
        Take-Screenshot "01_ribbon_ホーム_split_貼り付け"
        Close-AllMenus
    }

    # 下線
    $underlineBtn = Find-ButtonByName "下線"
    if ($underlineBtn) {
        Invoke-SafeClick $underlineBtn
        Start-Sleep -Milliseconds 1000
        Take-Screenshot "01_ribbon_ホーム_split_下線"
        Close-AllMenus
    }

    # 蛍光ペン
    $hilightBtn = Find-ButtonByName "蛍光ペン"
    if ($hilightBtn) {
        Invoke-SafeClick $hilightBtn
        Start-Sleep -Milliseconds 1000
        Take-Screenshot "01_ribbon_ホーム_split_蛍光ペンの色"
        Close-AllMenus
    }

    # フォントの色
    $fontColorBtn = Find-ButtonByName "フォントの色"
    if ($fontColorBtn) {
        Invoke-SafeClick $fontColorBtn
        Start-Sleep -Milliseconds 1000
        Take-Screenshot "01_ribbon_ホーム_split_フォントの色"
        Close-AllMenus
    }

    # 箇条書き
    $bulletBtn = Find-ButtonByName "箇条書き"
    if ($bulletBtn) {
        Invoke-SafeClick $bulletBtn
        Start-Sleep -Milliseconds 1000
        Take-Screenshot "01_ribbon_ホーム_split_箇条書き"
        Close-AllMenus
    }

    # 段落番号
    $numberBtn = Find-ButtonByName "段落番号"
    if ($numberBtn) {
        Invoke-SafeClick $numberBtn
        Start-Sleep -Milliseconds 1000
        Take-Screenshot "01_ribbon_ホーム_split_段落番号"
        Close-AllMenus
    }

    # 塗りつぶし
    $fillBtn = Find-ButtonByName "塗りつぶし"
    if ($fillBtn) {
        Invoke-SafeClick $fillBtn
        Start-Sleep -Milliseconds 1000
        Take-Screenshot "01_ribbon_ホーム_split_塗りつぶし"
        Close-AllMenus
    }

    # 罫線
    $borderBtn = Find-ButtonByName "罫線"
    if ($borderBtn) {
        Invoke-SafeClick $borderBtn
        Start-Sleep -Milliseconds 1000
        Take-Screenshot "01_ribbon_ホーム_split_罫線"
        Close-AllMenus
    }

    # 行間（段落グループ）
    $lineSpacingBtn = Find-ButtonByName "行間"
    if ($lineSpacingBtn) {
        Invoke-SafeClick $lineSpacingBtn
        Start-Sleep -Milliseconds 1000
        Take-Screenshot "04_dialog_paragraph"
        Close-AllMenus
    }
}

# ============================================================
# Phase 2: バックステージ（ファイルメニュー）のキャプチャ
# ============================================================
function Capture-Backstage() {
    Write-Host "`n[Phase 2] バックステージのキャプチャを開始"

    # ファイルタブは TabItem ではないため、Alt+F でバックステージを開く
    $wshell = New-Object -ComObject wscript.shell
    [WordCaptureApi]::SetForegroundWindow($script:WordWindow.Current.NativeWindowHandle) | Out-Null
    Start-Sleep -Milliseconds 300
    $wshell.SendKeys("%f")
    Start-Sleep -Milliseconds 1500

    # ホーム（最近使ったファイル）は既に開いている状態で撮影済み
    Take-Screenshot "02_backstage_home"

    foreach ($itemName in $BackstageLeftItems) {
        Write-Host "  [Item] $itemName"

        # バックステージを再表示 → ホームに戻して UIA ツリーをリフレッシュ
        $wshell.SendKeys("%f")
        Start-Sleep -Milliseconds 1000
        # 必ずホームに戻す
        try {
            $homeHome = Find-TabItemByNameGeneral "ホーム"
            if ($homeHome) { Invoke-SafeClick $homeHome; Start-Sleep -Milliseconds 500 }
        } catch {}

        $btn = Find-AnyByName $itemName

        # 「その他...」の場合は特別な処理が必要（オプションなどを開く）
        if ($itemName -eq "その他...") {
            # 「その他...」をクリック後、オプションなどを探す
            if ($btn) {
                Invoke-SafeClick $btn
                Start-Sleep -Milliseconds 1000
                Take-Screenshot "02_backstage_other"
            }
            Close-AllMenus
            continue
        }

        if (-not $btn) {
            Write-Host "  [Skip] バックステージ項目が見つかりません: $itemName"
            continue
        }

        Invoke-SafeClick $btn
        Start-Sleep -Milliseconds 1200

        # テンプレート ギャラリーが表示されている場合はスクロール
        if ($itemName -eq "新規") {
            # テンプレート一覧を撮影
            Take-Screenshot "02_backstage_新規"
            # テンプレートを右にスクロール
            $wshell.SendKeys("{RIGHT}")
            Start-Sleep -Milliseconds 500
            Take-Screenshot "02_backstage_新規_scroll"
            # 左に戻す
            $wshell.SendKeys("{LEFT}")
            Start-Sleep -Milliseconds 500

            # ★ 白紙の文書をクリックして空文書を開く（Word 2026 では自動では開かない）
            $blankDoc = Find-ButtonByName "白紙の文書"
            if (-not $blankDoc) { $blankDoc = Find-HyperlinkByName "白紙の文書" }
            if ($blankDoc) {
                Write-Host "  [Info] バックステージ: 白紙の文書をクリックして空文書を開く"
                Invoke-SafeClick $blankDoc
                Start-Sleep -Milliseconds 1500
                # 文書が開いたら撮影（後で Capture-DynamicUIScenarios で撮影する可能性がある）
            }
            # 必ず「ホーム」に戻す
            $homeHome = Find-TabItemByNameGeneral "ホーム"
            if ($homeHome) { Invoke-SafeClick $homeHome; Start-Sleep -Milliseconds 800 }
        }
        # 印刷ページの場合は印刷プレビューが展開されている
        elseif ($itemName -eq "印刷") {
            Take-Screenshot "02_backstage_印刷"
        }
        # 閉じるの場合は保存ダイアログが出現する可能性がある
        elseif ($itemName -eq "閉じる") {
            # 保存ダイアログが開くことを期待
            Start-Sleep -Milliseconds 2000
            # ダイアログがあるか確認
            $saveDlgs = $script:WordWindow.FindFirst($TS::Descendants,
                (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $CT::Window)))
            if ($saveDlgs) {
                Start-Sleep -Milliseconds 1000
                Take-Screenshot "02_backstage_閉じる"
            }
        }
        else {
            Take-Screenshot "02_backstage_$itemName"
        }

        # メニューを閉じる
        Close-AllMenus
    }

    # 「その他...」がキャプチャできなかった場合（オプションなど）
    # 別方法でオプション ダイアログを開く
    $optBtn = Find-ButtonByName "オプション"
    if ($optBtn) {
        Invoke-SafeClick $optBtn
        Start-Sleep -Milliseconds 1000
        Take-Screenshot "02_backstage_options"
    }

    # バックステージを閉じて通常表示に戻す
    Close-Backstage
    Start-Sleep -Milliseconds 800
    Take-Screenshot "02_backstage_closed"
}

# ============================================================
# Phase 3: エディタ内のコンテキストメニュー（テキスト選択時）
# ============================================================
function Capture-ContextMenus() {
    Write-Host "`n[Phase 3] コンテキストメニューのキャプチャを開始"

    # ホームタブに戻す
    $homeTab = Find-TabByName "ホーム"
    if ($homeTab) { Invoke-SafeClick $homeTab; Start-Sleep -Milliseconds 800 }

    # エディタ領域をクリックしてフォーカス
    $editArea = $script:WordWindow.FindFirst($TS::Descendants,
        (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $CT::Document)))
    if (-not $editArea) {
        $editArea = $script:WordWindow.FindFirst($TS::Descendants,
            (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $CT::Pane)))
    }
    if ($editArea) { Invoke-SafeClick $editArea }
    Start-Sleep -Milliseconds 500

    # テキストを入力してから選択 → 右クリック
    $wshell = New-Object -ComObject wscript.shell
    $wshell.SendKeys("これはテスト文です。WordPad Lite キャプチャ用。")
    Start-Sleep -Milliseconds 300

    # テキストをドラッグ選択
    try {
        $pt = $editArea.Current.BoundingRectangle
        $x1 = [int]($pt.X + 100)
        $y1 = [int]($pt.Y + 100)
        $x2 = [int]($pt.X + 300)
        $y2 = [int]($pt.Y + 100)
        [WordCaptureApi]::SetCursorPos($x1, $y1) | Out-Null
        [WordCaptureApi]::mouse_event(0x0002, 0, 0, 0, [IntPtr]::Zero)
        Start-Sleep -Milliseconds 100
        [WordCaptureApi]::SetCursorPos($x2, $y2) | Out-Null
        Start-Sleep -Milliseconds 300
        [WordCaptureApi]::mouse_event(0x0004, 0, 0, 0, [IntPtr]::Zero)
        Start-Sleep -Milliseconds 300

        # 選択範囲の中央で右クリック
        $mx = [int](($x1 + $x2) / 2)
        Invoke-RightClick $mx ($pt.Y + 100)
        Take-Screenshot "03_context_text"

        # 閉じる
        $wshell.SendKeys("{ESC}")
        Start-Sleep -Milliseconds 400
    } catch {
        Write-Host "  [Skip] テキスト右クリックメニューのキャプチャに失敗: $_"
    }

    # 段落番号の続行ドロップダウンのコンテキストメニュー
    # （箇条書きを適用 → 右クリック → 段落番号の続行が表示される状態）
    $bulletBtn = Find-ButtonByName "箇条書き"
    if ($bulletBtn) {
        Invoke-SafeClick $bulletBtn
        Start-Sleep -Milliseconds 500
        # 箇条書きの記号をクリック
        $bullets = Find-VisibleChildren $script:WordWindow $CT::ListItem
        if ($bullets.Count -gt 0) {
            Invoke-SafeClick $bullets[0]
        } else {
            $wshell.SendKeys("{ENTER}")
        }
        Start-Sleep -Milliseconds 1000
        Close-AllMenus

        # テキストを選択して右クリック（箇条書き/段落番号の継続が表示される状態）
        try {
            $pt = $editArea.Current.BoundingRectangle
            $mx = [int]($pt.X + 120)
            $my = [int]($pt.Y + 150)
            [WordCaptureApi]::SetCursorPos($mx, $my) | Out-Null
            [WordCaptureApi]::mouse_event(0x0002, 0, 0, 0, [IntPtr]::Zero)
            Start-Sleep -Milliseconds 100
            [WordCaptureApi]::SetCursorPos([int]($mx + 200), $my) | Out-Null
            Start-Sleep -Milliseconds 300
            [WordCaptureApi]::mouse_event(0x0004, 0, 0, 0, [IntPtr]::Zero)
            Start-Sleep -Milliseconds 300
            Invoke-RightClick $mx $my
            Take-Screenshot "03_context_text_numbering"
            $wshell.SendKeys("{ESC}")
            Start-Sleep -Milliseconds 400
        } catch {
            Write-Host "  [Skip] 段落番号の続行メニューのキャプチャに失敗: $_"
        }
    }
}

# ============================================================
# Phase 4: ダイアログのキャプチャ
# ============================================================
function Capture-Dialogs() {
    Write-Host "`n[Phase 4] ダイアログのキャプチャを開始"

    # ホームタブを表示
    $homeTab = Find-TabByName "ホーム"
    if ($homeTab) { Invoke-SafeClick $homeTab; Start-Sleep -Milliseconds 800 }

    $wshell = New-Object -ComObject wscript.shell

    # =====================================================
    # 段落ダイアログ - 全タブ
    # =====================================================
    Write-Host "  [段落] 段落ダイアログ（全タブ）"
    # ホームタブの段落グループのダイアログ ランチャーを探す
    $homeTab = Find-TabByName "ホーム"
    if ($homeTab) { Invoke-SafeClick $homeTab; Start-Sleep -Milliseconds 800 }
    # ダイアログ ランチャーは通常、GroupItem の中に小さな Button がある
    $homePane = $script:WordWindow.FindFirst($TS::Descendants,
        (New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, "段落")))
    if ($homePane) {
        try {
            $ctlType = $homePane.Current.ControlType.ProgrammaticName
            if ($ctlType -eq "ControlType.GroupItem") {
                # GroupItem の中からダイアログ ランチャーを探す
                $launcherBtns = Find-VisibleChildren $homePane $CT::Button
                if ($launcherBtns.Count -gt 0) {
                    # 右下に位置するボタン（ダイアログ ランチャー）をクリック
                    $bestBtn = $null
                    $maxXPlusY = 0
                    foreach ($lb in $launcherBtns) {
                        try {
                            $rect = $lb.Current.BoundingRectangle
                            $sum = $rect.X + $rect.Y
                            if ($sum -gt $maxXPlusY) { $maxXPlusY = $sum; $bestBtn = $lb }
                        } catch {}
                    }
                    if ($bestBtn) { Invoke-SafeClick $bestBtn; Start-Sleep -Milliseconds 800 }
                }
            }
        } catch { Write-Host "  [Warn] 段落ダイアログ ランチャー探索失敗: $_" }
    }
    Take-Screenshot "04_dialog_paragraph"

    # 段落ダイアログ内のタブ: 行とページ間の改行
    # Win32 ダイアログ: TAB でタブコントロールにフォーカス → RIGHT で次タブ
    $wshell.SendKeys("{TAB}")
    Start-Sleep -Milliseconds 600
    $wshell.SendKeys("{RIGHT}")
    Start-Sleep -Milliseconds 600
    Take-Screenshot "04_dialog_paragraph_行とページ間の改行"
    Reset-WordFocus

    # =====================================================
    # フォントダイアログ - 全タブ
    # =====================================================
    Write-Host "  [フォント] フォントダイアログ（全タブ）"
    $wshell.SendKeys("^d")
    Start-Sleep -Milliseconds 1200
    Take-Screenshot "04_dialog_font"

    # フォント ダイアログのタブ: 詳細設定
    # Win32 ダイアログ: TAB でタブコントロールにフォーカス → RIGHT で次タブ
    $wshell.SendKeys("{TAB}")
    Start-Sleep -Milliseconds 600
    $wshell.SendKeys("{RIGHT}")
    Start-Sleep -Milliseconds 600
    Take-Screenshot "04_dialog_font_詳細設定"
    Reset-WordFocus

    # =====================================================
    # 箇条書きと段落番号 ダイアログ（Ctrl+L）
    # =====================================================
    Write-Host "  [箇条書き/段落番号] ダイアログ"
    $wshell.SendKeys("^l")
    Start-Sleep -Milliseconds 1200
    Take-Screenshot "04_dialog_bullets_numbering"

    # 箇条書きと段落番号 ダイアログ内のタブ（イメージ付き箇条書き）
    # Win32 ダイアログ: TAB でタブコントロールにフォーカス → RIGHT で次タブ
    $wshell.SendKeys("{TAB}")
    Start-Sleep -Milliseconds 600
    $wshell.SendKeys("{RIGHT}")
    Start-Sleep -Milliseconds 600
    Take-Screenshot "04_dialog_bullets_numbering_イメージ付き箇条書き"
    Reset-WordFocus

    # =====================================================
    # ページ設定 ダイアログ - 全タブ
    # =====================================================
    Write-Host "  [ページ設定] ダイアログ（全タブ）"
    $layoutTab = Find-TabByName "レイアウト"
    if ($layoutTab) {
        Invoke-SafeClick $layoutTab
        Start-Sleep -Milliseconds 800
    }
    # ダイアログ ランチャーを探す: ページ設定グループ内の右下のボタン
    $pageSetupGroup = $script:WordWindow.FindFirst($TS::Descendants,
        (New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, "ページ設定")))
    $clicked = $false
    if ($pageSetupGroup) {
        try {
            $ctlType = $pageSetupGroup.Current.ControlType.ProgrammaticName
            Write-Host "  [Info] ページ設定グループ ControlType: $ctlType"
            if ($ctlType -eq "ControlType.GroupItem") {
                # GroupItem 右下のダイアログ ランチャーを探す（複数 ControlType 対応）
                $launcherCandidates = @()
                foreach ($ct in @($CT::Button, $CT::ToggleButton, $CT::Document, $CT::Text, $CT::Hyperlink)) {
                    $children = Find-VisibleChildren $pageSetupGroup $ct
                    if ($children.Count -gt 0) {
                        foreach ($c in $children) {
                            try {
                                if ($c.Current.BoundingRectangle.Width -gt 0) {
                                    $launcherCandidates += $c
                                }
                            } catch {}
                        }
                    }
                }
                if ($launcherCandidates.Count -gt 0) {
                    Write-Host "  [Info] ページ設定ランチャー候補($($launcherCandidates.Count)): $($launcherCandidates[0].Current.Name)"
                    # 右下（最大 X+Y）のボタンをクリック
                    $bestBtn = $null
                    $maxSum = 0
                    foreach ($db in $launcherCandidates) {
                        try {
                            $rect = $db.Current.BoundingRectangle
                            if ($rect.Width -gt 0) {
                                $sum = $rect.X + $rect.Y
                                if ($sum -gt $maxSum) { $maxSum = $sum; $bestBtn = $db }
                            }
                        } catch {}
                    }
                    if ($bestBtn) {
                        Invoke-SafeClick $bestBtn
                        Start-Sleep -Milliseconds 1000
                        $clicked = $true
                    }
                }
            } else {
                # GroupItem 以外の場合、直接名前一致でダイアログランチャーを探す
                $dlgLaunchers = $script:WordWindow.FindAll($TS::Descendants,
                    (New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, "ページ設定")))
                foreach ($dl in $dlgLaunchers) {
                    try {
                        if (-not $dl.Current.IsOffscreen -and $dl.Current.BoundingRectangle.Width -gt 0) {
                            Invoke-SafeClick $dl
                            Start-Sleep -Milliseconds 1000
                            $clicked = $true; break
                        }
                    } catch {}
                }
            }
        } catch { Write-Host "  [Warn] ページ設定ランチャー探索失敗: $_" }
    }
    # ダイアログが開いたか確認（$script:WordWindow 内 + RootElement 両方を検索）
    if ($clicked -or ($pageSetupGroup -ne $null)) {
        $foundDlg = $false
        # $script:WordWindow 内を先に探索（Word 2026 のダイアログはここ）
        $allInner = $script:WordWindow.FindAll($TS::Descendants,
            (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $CT::Window)))
        foreach ($d in $allInner) {
            try {
                if ($d.Current.Name -like "*ページ設定*" -or $d.Current.Name -like "*Page Setup*") {
                    $foundDlg = $true; break
                }
            } catch {}
        }
        # 見つからなければ RootElement のトップレベルも探索
        if (-not $foundDlg) {
            $dlgRoot = $AE::RootElement
            $allDlgs = $dlgRoot.FindAll($TS::Children,
                (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $CT::Window)))
            foreach ($d in $allDlgs) {
                try {
                    if ($d.Current.Name -like "*ページ設定*" -or $d.Current.Name -like "*Page Setup*") {
                        $foundDlg = $true; break
                    }
                } catch {}
            }
        }
        if ($foundDlg) {
            Write-Host "  [Info] ページ設定ダイアログを検出"
            Take-Screenshot "04_dialog_page_setup"

            # ページ設定ダイアログの各タブ: 余白→用紙→レイアウト
            # Win32 ダイアログ: TAB でタブコントロールにフォーカス → 矢印キーでタブ切替
            $wshell.SendKeys("{TAB}")
            Start-Sleep -Milliseconds 600
            Take-Screenshot "04_dialog_page_setup_余白"
            # 用紙タブに切り替え
            $wshell.SendKeys("{RIGHT}")
            Start-Sleep -Milliseconds 600
            Take-Screenshot "04_dialog_page_setup_用紙"
            # レイアウトタブに切り替え
            $wshell.SendKeys("{RIGHT}")
            Start-Sleep -Milliseconds 600
            Take-Screenshot "04_dialog_page_setup_レイアウト"
            Reset-WordFocus
        } else {
            Write-Host "  [Warn] ページ設定ダイアログが見つかりません"
        }
    }
}

# ============================================================
# Phase 4b: 詳細ウィンドウ（ダイアログ内タブ）の追加キャプチャ
# ============================================================
function Capture-DetailedDialogs() {
    Write-Host "`n[Phase 4b] 詳細ウィンドウのキャプチャを開始"
    $wshell = New-Object -ComObject wscript.shell

    # =====================================================
    # 枠線と網かけ ダイアログ
    # =====================================================
    Write-Host "  [枠線と網かけ] ダイアログ（全タブ）"
    $homeTab = Find-TabByName "ホーム"
    if ($homeTab) { Invoke-SafeClick $homeTab; Start-Sleep -Milliseconds 800 }

    # 罫線スプリットボタンから枠線と網かけを起動
    $borderBtn = Find-ButtonByName "罫線"
    if (-not $borderBtn) { $borderBtn = Find-ButtonByName "罫線と網かけ" }
    if ($borderBtn) {
        Invoke-SafeClick $borderBtn
        Start-Sleep -Milliseconds 800
        # ドロップダウンから「枠線と網かけ」をクリック（複数 ControlType 対応）
        $borderShading = Find-ButtonByName "枠線と網かけ"
        if (-not $borderShading) { $borderShading = Find-ButtonByName "枠線と網かけ..." }
        # MenuItem としても探索
        if (-not $borderShading) {
            $allMI = $script:WordWindow.FindAll($TS::Descendants,
                (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $CT::MenuItem)))
            foreach ($mi in $allMI) {
                try {
                    if ($mi.Current.Name -eq "枠線と網かけ" -or $mi.Current.Name -eq "枠線と網かけ...") {
                        $borderShading = $mi; break
                    }
                } catch {}
            }
        }
        # 名前部分一致でも探索
        if (-not $borderShading) {
            $allNames = $script:WordWindow.FindAll($TS::Descendants,
                (New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, "枠線")))
            foreach ($n in $allNames) {
                try {
                    if ($n.Current.Name -like "*枠線と網かけ*" -and -not $n.Current.IsOffscreen -and $n.Current.BoundingRectangle.Width -gt 0) {
                        $borderShading = $n; break
                    }
                } catch {}
            }
        }
        if ($borderShading) {
            try {
                $pattern = $borderShading.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::RuntimeId)
            } catch {
                Invoke-SafeClick $borderShading
                Start-Sleep -Milliseconds 1200
                Take-Screenshot "04_dialog_borders_shading"

                # 枠線と網かけダイアログの各タブ（Win32 なので SendKeys TAB → 矢印キー）
                # タブにフォーカス → 右矢印で次タブ
                $wshell.SendKeys("{TAB}")
                Start-Sleep -Milliseconds 600
                Take-Screenshot "04_dialog_borders_shading_網かけ"
                $wshell.SendKeys("{RIGHT}")
                Start-Sleep -Milliseconds 600
                Take-Screenshot "04_dialog_borders_shading_ページ罫線"
                Reset-WordFocus
            }
        } else {
            Write-Host "  [Warn] 枠線と網かけが見つかりません"
        }
        Close-AllMenus
    }

    # =====================================================
    # 段組 ダイアログ
    # =====================================================
    Write-Host "  [段組] ダイアログ"
    $homeTab = Find-TabByName "ホーム"
    if ($homeTab) { Invoke-SafeClick $homeTab; Start-Sleep -Milliseconds 500 }
    $layoutTab = Find-TabByName "レイアウト"
    if ($layoutTab) {
        Invoke-SafeClick $layoutTab
        Start-Sleep -Milliseconds 800
    }
    # 段組ダイアログを開く: ドロップダウンから「段組...」を探す
    $columnBtn = Find-ButtonByName "段組"
    if ($columnBtn) {
        Invoke-SafeClick $columnBtn
        Start-Sleep -Milliseconds 600
        # ドロップダウン内のメニューアイテムを複数 ControlType で探索
        $columnDialog = Find-ButtonByName "段組..."
        if (-not $columnDialog) { $columnDialog = Find-ButtonByName "段組" }
        if (-not $columnDialog) {
            # MenuItem として探索
            $allMI = $script:WordWindow.FindAll($TS::Descendants,
                (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $CT::MenuItem)))
            foreach ($mi in $allMI) {
                try {
                    if ($mi.Current.Name -eq "段組" -or $mi.Current.Name -eq "段組...") {
                        $columnDialog = $mi; break
                    }
                } catch {}
            }
        }
        # 名前部分一致でも探索
        if (-not $columnDialog) {
            $allNames = $script:WordWindow.FindAll($TS::Descendants,
                (New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, "段")))
            foreach ($n in $allNames) {
                try {
                    if ($n.Current.Name -like "*段組*" -and -not $n.Current.IsOffscreen -and $n.Current.BoundingRectangle.Width -gt 0) {
                        $columnDialog = $n; break
                    }
                } catch {}
            }
        }
        if ($columnDialog) {
            Invoke-SafeClick $columnDialog
            Start-Sleep -Milliseconds 1200
            Take-Screenshot "04_dialog_columns"
            # 段組ダイアログのタブ切替（SendKeys TAB → 矢印キー）
            $wshell.SendKeys("{TAB}")
            Start-Sleep -Milliseconds 600
            Take-Screenshot "04_dialog_columns"
            Reset-WordFocus
        }
    }

    # =====================================================
    # スタイル ダイアログ（スタイル ウィンドウ）
    # =====================================================
    Write-Host "  [スタイル] スタイル ウィンドウ"
    $homeTab = Find-TabByName "ホーム"
    if ($homeTab) { Invoke-SafeClick $homeTab; Start-Sleep -Milliseconds 800 }

    # スタイル ウィンドウはスタイル ギャラリーの右下ランチャーから
    $styleLauncher = $script:WordWindow.FindFirst($TS::Descendants,
        (New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, "スタイル")))
    if (-not $styleLauncher) {
        # ホーム タブ内を探索
        $homePane = Find-TabByName "ホーム"
        if ($homePane) {
            $styleLauncher = $homePane.FindFirst($TS::Descendants,
                (New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, "スタイル")))
        }
    }
    if (-not $styleLauncher) {
        # Alt+Ctrl+Shift+S (スタイル ウィンドウ ショートカット)
        Write-Host "  [Info] スタイル ウィンドウを Alt+Ctrl+Shift+S で開く"
        $wshell.SendKeys("%+{S}")
        Start-Sleep -Milliseconds 1500
    }
    # スタイル ウィンドウが開いていれば撮影
    $styleWin = $script:WordWindow.FindFirst($TS::Descendants,
        (New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, "スタイル")))
    if ($styleWin) {
        Take-Screenshot "04_dialog_style_pane"
        # スタイルの管理ボタン
        $manageBtn = Find-ButtonByName "スタイルの管理"
        if ($manageBtn) {
            Invoke-SafeClick $manageBtn
            Start-Sleep -Milliseconds 1000
            Take-Screenshot "04_dialog_style_manager"
            Reset-WordFocus
        }
    }

    # =====================================================
    # オプション ダイアログ（ファイル → オプション）
    # =====================================================
    Write-Host "  [オプション] ダイアログ（主要カテゴリ）"
    # Ctrl+F1 でオプションダイアログを直接開く（Word 2026 互換）
    $wshell.SendKeys("^+{F1}")
    Start-Sleep -Milliseconds 2000
    $optBtn = Find-ButtonByName "オプション"
    if (-not $optBtn) {
        # フォールバック: Alt+O（通常オプションダイアログ）
        $wshell.SendKeys("%o")
        Start-Sleep -Milliseconds 1500
        $optBtn = Find-ButtonByName "オプション"
    }
    if ($optBtn) {
        Invoke-SafeClick $optBtn
        Start-Sleep -Milliseconds 1500
        Take-Screenshot "04_dialog_options"

        # オプション ダイアログの主要カテゴリ（Win32 ダイアログ: TAB → 矢印キーでタブ切替）
        $optTabs = @("表示", "校正", "保存", "言語", "詳細設定", "カスタマイズ")
        foreach ($t in $optTabs) {
            $wshell.SendKeys("{RIGHT}")
            Start-Sleep -Milliseconds 800
            Take-Screenshot "04_dialog_options_$($t)"
        }
        Reset-WordFocus
    }
    Close-Backstage
    Start-Sleep -Milliseconds 500

    # =====================================================
    # ハイパーリンク ダイアログの詳細
    # =====================================================
    Write-Host "  [ハイパーリンク] ダイアログ"
    Reset-WordFocus
    try { $wshell.SendKeys("%i") } catch {}  # Alt+I で挿入
    Start-Sleep -Milliseconds 1000
    $linkBtn = Find-ButtonByName "リンク"
    if (-not $linkBtn) {
        # MenuItem として探索
        $allMI = $script:WordWindow.FindAll($TS::Descendants,
            (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $CT::MenuItem)))
        foreach ($mi in $allMI) {
            try {
                if ($mi.Current.Name -eq "リンク") { $linkBtn = $mi; break }
            } catch {}
        }
    }
    if ($linkBtn) {
        Invoke-SafeClick $linkBtn
        Start-Sleep -Milliseconds 1200
        Take-Screenshot "04_dialog_hyperlink"
        Reset-WordFocus
    }

    # =====================================================
    # 脚注ダイアログ（参考資料タブ）
    # =====================================================
    Write-Host "  [脚注/注釈] ダイアログ"
    Reset-WordFocus
    # 参考資料タブに遷移
    try { $wshell.SendKeys("%r") } catch {}  # Alt+R で参考資料
    Start-Sleep -Milliseconds 1000
    # 脚注グループのダイアログ ランチャーを探す（複数 ControlType 対応）
    $footnoteLauncher = $script:WordWindow.FindFirst($TS::Descendants,
        (New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, "脚注")))
    if ($footnoteLauncher) {
        $foundDlgBtn = $false
        try {
            $ctlType = $footnoteLauncher.Current.ControlType.ProgrammaticName
            Write-Host "  [Info] 脚注グループ ControlType: $ctlType"
            if ($ctlType -eq "ControlType.GroupItem") {
                # GroupItem 内から複数 ControlType でランチャー探索
                $candidates = @()
                foreach ($ct in @($CT::Button, $CT::Document, $CT::Text, $CT::Hyperlink, $CT::ToggleButton)) {
                    $ch = Find-VisibleChildren $footnoteLauncher $ct
                    if ($ch.Count -gt 0) { foreach ($c in $ch) { $candidates += $c } }
                }
                if ($candidates.Count -gt 0) {
                    # 右下（最大 X+Y）のボタンをクリック
                    $bestBtn = $null
                    $maxSum = 0
                    foreach ($db in $candidates) {
                        try {
                            $rect = $db.Current.BoundingRectangle
                            if ($rect.Width -gt 0) {
                                $sum = $rect.X + $rect.Y
                                if ($sum -gt $maxSum) { $maxSum = $sum; $bestBtn = $db }
                            }
                        } catch {}
                    }
                    if ($bestBtn) {
                        Invoke-SafeClick $bestBtn
                        Start-Sleep -Milliseconds 1200
                        Take-Screenshot "04_dialog_footnote"
                        Reset-WordFocus
                        $foundDlgBtn = $true
                    }
                }
            }
        } catch { Write-Host "  [Warn] 脚注ダイアログ探索失敗: $_" }
        # GroupItem 以外の場合、直接名前一致でダイアログランチャーを探す
        if (-not $foundDlgBtn) {
            $dlgLaunchers = $script:WordWindow.FindAll($TS::Descendants,
                (New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, "脚注")))
            foreach ($dl in $dlgLaunchers) {
                try {
                    if (-not $dl.Current.IsOffscreen -and $dl.Current.BoundingRectangle.Width -gt 0) {
                        Invoke-SafeClick $dl
                        Start-Sleep -Milliseconds 1200
                        Take-Screenshot "04_dialog_footnote"
                        Reset-WordFocus
                        $foundDlgBtn = $true; break
                    }
                } catch {}
            }
        }
    }
    if (-not $foundDlgBtn) {
        Write-Host "  [Warn] 脚注ダイアログ ランチャーが見つかりません"
    }
}

# ============================================================
# Phase 5: 動的 / 条件付き UI のキャプチャ
# ============================================================
function Capture-DynamicUIScenarios() {
    Write-Host "`n[Phase 5] 動的 UI シナリオのキャプチャを開始"
    $wshell = New-Object -ComObject wscript.shell

    # Phase 4b 終了後の UI 状態をリセット
    Reset-WordFocus

    # シナリオ 1: 100 文字以上入力してエディター スコアを確認
    $homeTab = Find-TabByName "ホーム"
    if ($homeTab) { Invoke-SafeClick $homeTab; Start-Sleep -Milliseconds 500 }

    $text = "これは WordPad Lite の UI キャプチャ用のテスト文章です。校正機能やエディター スコアが表示されることを確認するために、十分な長さの文章を入力しておきます。日本語でも英語でも文字数はカウントされます。This is a sample sentence for editor score. "
    try {
        Set-Clipboard -Value $text
        Start-Sleep -Milliseconds 300
        $wshell.SendKeys("^v")
        Start-Sleep -Milliseconds 800
    } catch {
        Write-Host "  [Warn] クリップボード経由のテキスト貼り付けに失敗: $_"
    }

    $reviewTab = Find-TabByName "校閲"
    if ($reviewTab) {
        Invoke-SafeClick $reviewTab
        Start-Sleep -Milliseconds 1000
        Take-Screenshot "05_scenario_editor_score_reviewtab"

        # エディター / 文章校正 ボタンを探す
        $candidates = @("エディター", "スぺル チェックと文章校正", "文章校正", "校正", "Editor")
        foreach ($c in $candidates) {
            $editorBtn = Find-ButtonByName $c
            if ($editorBtn) {
                Invoke-SafeClick $editorBtn
                Start-Sleep -Milliseconds 1500
                Take-Screenshot "05_scenario_editor_score_pane"
                # 閉じる（Esc）
                $wshell.SendKeys("{ESC}")
                Start-Sleep -Milliseconds 500
                break
            }
        }

        # 変更履歴ウィンドウが開いている場合
        $revHistWindow = Find-ButtonByName "[変更履歴] ウィンドウ"
        if ($revHistWindow) {
            Invoke-SafeClick $revHistWindow
            Start-Sleep -Milliseconds 1000
            Take-Screenshot "05_scenario_change_history_pane"
        }
    }

    # シナリオ 2: 検索パネル (Ctrl+F)
    Reset-WordFocus
    $wshell.SendKeys("^f")
    Start-Sleep -Milliseconds 1200
    Take-Screenshot "05_scenario_search_pane"
    $wshell.SendKeys("{ESC}")
    Start-Sleep -Milliseconds 500
    Reset-WordFocus

    # シナリオ 3: 置換ダイアログ (Ctrl+H)
    $wshell.SendKeys("^h")
    Start-Sleep -Milliseconds 1200
    Take-Screenshot "05_scenario_replace_dialog"

    # シナリオ 3b: 置換ダイアログのオプション展開
    $optBtn = Find-ButtonByName "オプション"
    if ($optBtn) {
        Invoke-SafeClick $optBtn
        Start-Sleep -Milliseconds 1000
        Take-Screenshot "05_scenario_replace_dialog_opts"
    }
    Reset-WordFocus

    # シナリオ 4: ナビゲーション ウィンドウ
    $viewTab = Find-TabByName "表示"
    if ($viewTab) {
        Invoke-SafeClick $viewTab
        Start-Sleep -Milliseconds 800
        $navBtn = Find-ButtonByName "ナビゲーション ウィンドウ"
        if ($navBtn) {
            Invoke-SafeClick $navBtn
            Start-Sleep -Milliseconds 1200
            Take-Screenshot "05_scenario_navigation_pane"
            Invoke-SafeClick $navBtn
            Start-Sleep -Milliseconds 500
        }

        # ルーラー表示切替
        $rulerCheck = Find-ButtonByName "ルーラー"
        if ($rulerCheck) {
            Invoke-SafeClick $rulerCheck
            Start-Sleep -Milliseconds 500
            Take-Screenshot "05_scenario_ruler_visible"
            Invoke-SafeClick $rulerCheck
            Start-Sleep -Milliseconds 500
        }
    }

    # シナリオ 5: コマンド パレット (Ctrl+K)
    $wshell.SendKeys("^k")
    Start-Sleep -Milliseconds 1500
    Take-Screenshot "05_scenario_command_palette"
    $wshell.SendKeys("{ESC}")
    Start-Sleep -Milliseconds 500

    # シナリオ 6: テーマ選択（デザインタブ）
    $designTab = Find-TabByName "デザイン"
    if ($designTab) {
        Invoke-SafeClick $designTab
        Start-Sleep -Milliseconds 800
        # テーマ ドロップダウン
        $themeBtn = Find-ButtonByName "テーマ"
        if ($themeBtn) {
            Invoke-SafeClick $themeBtn
            Start-Sleep -Milliseconds 1000
            Take-Screenshot "01_ribbon_デザイン_split_テーマ"
            Close-AllMenus
        }
        # 配色 ドロップダウン
        $colorBtn = Find-ButtonByName "配色"
        if ($colorBtn) {
            Invoke-SafeClick $colorBtn
            Start-Sleep -Milliseconds 1000
            Take-Screenshot "01_ribbon_デザイン_split_配色"
            Close-AllMenus
        }
    }
}

# ============================================================
# Phase 6: 補足キャプチャ（画像で確認した但未実装のもの）
# ============================================================
function Capture-AdditionalUI() {
    Write-Host "`n[Phase 6] 補足 UI のキャプチャを開始"
    Reset-WordFocus

    # 1. スタイル ギャラリー（ホームタブ）
    $homeTab = Find-TabByName "ホーム"
    if ($homeTab) {
        Invoke-SafeClick $homeTab
        Start-Sleep -Milliseconds 800

        # スタイル プレビューからドロップダウンを開く
        $stylePreview = $script:WordWindow.FindFirst($TS::Descendants,
            (New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, "見出し 1")))
        if ($stylePreview) {
            # スタイル プレビューの右側の▼を探す
            $parent = $stylePreview
            for ($i = 0; $i -lt 5; $i++) {
                try {
                    $parent = [System.Windows.Automation.TreeWalker]::RawViewParent.GetParent($parent)
                    $dropBtns = Find-VisibleChildren $parent $CT::Button
                    foreach ($db in $dropBtns) {
                        try {
                            if ($db.Current.Name -match '[▼v↓]') {
                                Invoke-SafeClick $db
                                Start-Sleep -Milliseconds 1000
                                Take-Screenshot "01_ribbon_ホーム_split_スタイルギャラリー"
                                Close-AllMenus
                                break
                            }
                        } catch {}
                    }
                } catch { break }
            }
        }
    }

    # 2. マクロ ドロップダウン（表示タブ）
    $viewTab = Find-TabByName "表示"
    if ($viewTab) {
        Invoke-SafeClick $viewTab
        Start-Sleep -Milliseconds 800
        $macroBtn = Find-ButtonByName "マクロ"
        if ($macroBtn) {
            Invoke-SafeClick $macroBtn
            Start-Sleep -Milliseconds 1000
            Take-Screenshot "01_ribbon_表示_split_マクロ"
            Close-AllMenus
        }
    }

    # 3. ヘルプタブの「新機能ガイド」
    $helpTab = Find-TabByName "ヘルプ"
    if ($helpTab) {
        Invoke-SafeClick $helpTab
        Start-Sleep -Milliseconds 800
        Take-Screenshot "01_ribbon_ヘルプ"

        $newFeatBtn = Find-ButtonByName "新機能ガイド"
        if ($newFeatBtn) {
            Invoke-SafeClick $newFeatBtn
            Start-Sleep -Milliseconds 1000
            Take-Screenshot "01_ribbon_ヘルプ_split_新機能ガイド"
        }
    }

    # 4. 差し込み文書タブのドロップダウン追加
    $mailMergeTab = Find-TabByName "差し込み文書"
    if ($mailMergeTab) {
        Invoke-SafeClick $mailMergeTab
        Start-Sleep -Milliseconds 800

        # 差し込み印刷の開始
        $mmBtn = Find-ButtonByName "差し込み印刷の開始"
        if ($mmBtn) {
            Invoke-SafeClick $mmBtn
            Start-Sleep -Milliseconds 1000
            Take-Screenshot "01_ribbon_差し込み文書_split_差し込み印刷の開始"
            Close-AllMenus
        }

        # 宛先の選択
        $dstBtn = Find-ButtonByName "宛先の選択"
        if ($dstBtn) {
            Invoke-SafeClick $dstBtn
            Start-Sleep -Milliseconds 1000
            Take-Screenshot "01_ribbon_差し込み文書_split_宛先の選択"
            Close-AllMenus
        }
    }

    # 5. 参考資料タブの「文字の塗りつぶし」など
    $refTab = Find-TabByName "参考資料"
    if ($refTab) {
        Invoke-SafeClick $refTab
        Start-Sleep -Milliseconds 800

        # 脚注の挿入 ドロップダウン
        $footnoteBtn = Find-ButtonByName "脚注の挿入"
        if ($footnoteBtn) {
            Invoke-SafeClick $footnoteBtn
            Start-Sleep -Milliseconds 1000
            Take-Screenshot "01_ribbon_参考資料_split_脚注の挿入"
            Close-AllMenus
        }

        # 次の脚注
        $nextFootBtn = Find-ButtonByName "次の脚注"
        if ($nextFootBtn) {
            Invoke-SafeClick $nextFootBtn
            Start-Sleep -Milliseconds 1000
            Take-Screenshot "01_ribbon_参考資料_split_次の脚注"
        }
    }

    # 6. デザインタブの数式ツール（数式挿入時）
    $insertTab = Find-TabByName "挿入"
    if ($insertTab) {
        Invoke-SafeClick $insertTab
        Start-Sleep -Milliseconds 800
        $eqBtn = Find-ButtonByName "数式"
        if ($eqBtn) {
            Invoke-SafeClick $eqBtn
            Start-Sleep -Milliseconds 1000
            Take-Screenshot "01_ribbon_数式"
            # 数式ツール（数式タブ）が表示される
            $eqTools = Find-VisibleChildren $script:WordWindow $CT::Button
            foreach ($et in $eqTools) {
                try {
                    if ($et.Current.Name -match "分数|積分|累乗根|大型演算子") {
                        Invoke-SafeClick $et
                        Start-Sleep -Milliseconds 1000
                        Take-Screenshot "01_ribbon_数式_split_構造"
                        Close-AllMenus
                        break
                    }
                } catch {}
            }
        }
    }

    # 7. レイアウトタブのドロップダウン追加
    $layoutTab = Find-TabByName "レイアウト"
    if ($layoutTab) {
        Invoke-SafeClick $layoutTab
        Start-Sleep -Milliseconds 800

        # 区切り
        $breakBtn = Find-ButtonByName "区切り"
        if ($breakBtn) {
            Invoke-SafeClick $breakBtn
            Start-Sleep -Milliseconds 1000
            Take-Screenshot "01_ribbon_レイアウト_split_区切り"
            Close-AllMenus
        }

        # 行番号
        $lineNumBtn = Find-ButtonByName "行番号"
        if ($lineNumBtn) {
            Invoke-SafeClick $lineNumBtn
            Start-Sleep -Milliseconds 1000
            Take-Screenshot "01_ribbon_レイアウト_split_行番号"
            Close-AllMenus
        }

        # 配置
        $posBtn = Find-ButtonByName "配置"
        if ($posBtn) {
            Invoke-SafeClick $posBtn
            Start-Sleep -Milliseconds 1000
            Take-Screenshot "01_ribbon_レイアウト_split_配置"
            Close-AllMenus
        }

        # グループ化
        $groupBtn = Find-ButtonByName "グループ化"
        if ($groupBtn) {
            Invoke-SafeClick $groupBtn
            Start-Sleep -Milliseconds 1000
            Take-Screenshot "01_ribbon_レイアウト_split_グループ化"
            Close-AllMenus
        }

        # 回転
        $rotateBtn = Find-ButtonByName "回転"
        if ($rotateBtn) {
            Invoke-SafeClick $rotateBtn
            Start-Sleep -Milliseconds 1000
            Take-Screenshot "01_ribbon_レイアウト_split_回転"
            Close-AllMenus
        }
    }

    # 8. 校閲タブの削除と変更履歴ドロップダウン追加
    $reviewTab = Find-TabByName "校閲"
    if ($reviewTab) {
        Invoke-SafeClick $reviewTab
        Start-Sleep -Milliseconds 800

        # 削除 ドロップダウン
        $delBtn = Find-ButtonByName "削除"
        if ($delBtn) {
            Invoke-SafeClick $delBtn
            Start-Sleep -Milliseconds 1000
            Take-Screenshot "01_ribbon_校閲_split_削除"
            Close-AllMenus
        }

        # 変更履歴 ドロップダウン
        $revBtn = Find-ButtonByName "変更履歴"
        if ($revBtn) {
            Invoke-SafeClick $revBtn
            Start-Sleep -Milliseconds 1000
            Take-Screenshot "01_ribbon_校閲_split_変更履歴"
            Close-AllMenus
        }

        # 比較 ドロップダウン
        $compareBtn = Find-ButtonByName "比較"
        if ($compareBtn) {
            Invoke-SafeClick $compareBtn
            Start-Sleep -Milliseconds 1000
            Take-Screenshot "01_ribbon_校閲_split_比較"
            Close-AllMenus
        }

        # 保護 ドロップダウン
        $protectBtn = Find-ButtonByName "保護"
        if ($protectBtn) {
            Invoke-SafeClick $protectBtn
            Start-Sleep -Milliseconds 1000
            Take-Screenshot "01_ribbon_校閲_split_保護"
            Close-AllMenus
        }
    }

    # 9. 挿入タブのページグループ追加
    $insertTab = Find-TabByName "挿入"
    if ($insertTab) {
        Invoke-SafeClick $insertTab
        Start-Sleep -Milliseconds 800

        # 表紙
        $coverBtn = Find-ButtonByName "表紙"
        if ($coverBtn) {
            Invoke-SafeClick $coverBtn
            Start-Sleep -Milliseconds 1000
            Take-Screenshot "01_ribbon_挿入_split_表紙"
            Close-AllMenus
        }

        # ヘッダー
        $headerBtn = Find-ButtonByName "ヘッダー"
        if ($headerBtn) {
            Invoke-SafeClick $headerBtn
            Start-Sleep -Milliseconds 1000
            Take-Screenshot "01_ribbon_挿入_split_ヘッダー"
            Close-AllMenus
        }

        # フッター
        $footerBtn = Find-ButtonByName "フッター"
        if ($footerBtn) {
            Invoke-SafeClick $footerBtn
            Start-Sleep -Milliseconds 1000
            Take-Screenshot "01_ribbon_挿入_split_フッター"
            Close-AllMenus
        }

        # ページ番号
        $pageNumBtn = Find-ButtonByName "ページ番号"
        if ($pageNumBtn) {
            Invoke-SafeClick $pageNumBtn
            Start-Sleep -Milliseconds 1000
            Take-Screenshot "01_ribbon_挿入_split_ページ番号"
            Close-AllMenus
        }

        # テキスト ボックス
        $textBoxBtn = Find-ButtonByName "テキストボックス"
        if ($textBoxBtn) {
            Invoke-SafeClick $textBoxBtn
            Start-Sleep -Milliseconds 1000
            Take-Screenshot "01_ribbon_挿入_split_テキストボックス"
            Close-AllMenus
        }

        # ワードアート
        $wordArtBtn = Find-ButtonByName "ワードアート"
        if ($wordArtBtn) {
            Invoke-SafeClick $wordArtBtn
            Start-Sleep -Milliseconds 1000
            Take-Screenshot "01_ribbon_挿入_split_ワードアート"
            Close-AllMenus
        }

        # ドロップ キャップ
        $dropCapBtn = Find-ButtonByName "ドロップ キャップ"
        if ($dropCapBtn) {
            Invoke-SafeClick $dropCapBtn
            Start-Sleep -Milliseconds 1000
            Take-Screenshot "01_ribbon_挿入_split_ドロップキャップ"
            Close-AllMenus
        }

        # 記号と特殊文字
        $symbolBtn = Find-ButtonByName "記号と特殊文字"
        if ($symbolBtn) {
            Invoke-SafeClick $symbolBtn
            Start-Sleep -Milliseconds 1000
            Take-Screenshot "01_ribbon_挿入_split_記号と特殊文字"
            Close-AllMenus
        }

        # eSignature フィールド
        $esigBtn = Find-ButtonByName "eSignature フィールド"
        if ($esigBtn) {
            Invoke-SafeClick $esigBtn
            Start-Sleep -Milliseconds 1000
            Take-Screenshot "01_ribbon_挿入_split_eSignature"
        }
    }
}

# ============================================================
# メイン処理
# ============================================================
Get-WordMainWindow

# 「ようこそ画面」またはバックステージが開いている場合は新規文書に
# Word 2026 は初期状態ですでに文書を開いている場合が多い
# Open-BlankDocumentIfNeeded は Phase 0 で適切に処理する

# Phase 0: スタート画面（ようこそ画面）
# 既存のキャプチャでは 00_initial.png が空文書表示だったため、
# スタート画面のキャプチャは任意（Word 2026 の設定による）
# 既に文書が開いている場合はスキップ
$doc = $script:WordWindow.FindFirst($TS::Descendants,
    (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $CT::Document)))
if (-not $doc) {
    Capture-StartScreen
} else {
    Write-Host "`n[Info] 既に文書が開いているため、スタート画面キャプチャをスキップします。"
    Write-Host "        スクリーンショット 00_initial を撮影します。"
    Take-Screenshot "00_initial"
}

Capture-RibbonTabs

# Phase 1b: ホームタブの主要スプリットボタン追加
Capture-HomeSplitButtons

Capture-Backstage
Capture-ContextMenus
Capture-Dialogs
Capture-DetailedDialogs
Capture-DynamicUIScenarios
Capture-AdditionalUI

Write-Host "`n========================================"
Write-Host "キャプチャ完了。出力先: $OutputDir"
Write-Host "========================================"
Write-Host "`n生成予定ファイル一覧:"
foreach ($key in $RequiredCaptures.Keys) {
    $path = Join-Path $OutputDir "$key.png"
    $status = if (Test-Path $path) { "✓ 作成済み" } else { "○ 未作成" }
    Write-Host "  $key : $status - $($RequiredCaptures[$key])"
}
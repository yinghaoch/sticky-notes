# --- 自動權限設定與繞過機制 ---
try {
    $currentPolicy = Get-ExecutionPolicy -Scope CurrentUser
    if ($currentPolicy -eq "Restricted" -or $null -eq $currentPolicy) {
        Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force -ErrorAction Stop
    }
} catch {
    if ($MyInvocation.MyCommand.Path -and $args -notcontains "-SelfBypassed") {
        Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -NoProfile -File `"$($MyInvocation.MyCommand.Path)`" -SelfBypassed"
        exit
    }
}
# ------------------------------

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# 設定資料儲存路徑（與 .ps1 同資料夾）
$ScriptDir = $PSScriptRoot
if ([string]::IsNullOrEmpty($ScriptDir)) { $ScriptDir = $PWD.Path }
$DataFile = Join-Path $ScriptDir "sticky_notes_data.json"

# 初始化資料結構
$global:notesData = [PSCustomObject]@{
    tasks   = @()
    alpha   = 0.9
    topmost = $true
}

function Load-Data {
    if (Test-Path $DataFile) {
        try {
            $json = Get-Content $DataFile -Raw -Encoding UTF8
            $data = $json | ConvertFrom-Json
            if ($data.tasks) { $global:notesData.tasks = @($data.tasks) }
            if ($data.alpha) { $global:notesData.alpha = $data.alpha }
            if ($null -ne $data.topmost) { $global:notesData.topmost = $data.topmost }
        } catch {}
    }
}

function Save-Data {
    try {
        $global:notesData | ConvertTo-Json -Depth 5 | Set-Content $DataFile -Encoding UTF8
    } catch {}
}

Load-Data

# 時間格式化輔助 (支援 0900 自動轉為 09:00)
function Format-TimeInput($val) {
    $val = $val.Trim()
    if ($val -notmatch ":" -and $val.Length -eq 4 -and $val -match '^\d+$') {
        $val = $val.Substring(0, 2) + ":" + $val.Substring(2, 2)
    }
    return $val
}

# 建立主視窗
$form = New-Object System.Windows.Forms.Form
$form.Text = "PowerShell 輕量便箋"
$form.Size = New-Object System.Drawing.Size(340, 390)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::FromArgb(255, 255, 242)
$form.TopMost = $global:notesData.topmost
$form.Opacity = $global:notesData.alpha

# 1. 上方輸入區
$panelTop = New-Object System.Windows.Forms.Panel
$panelTop.Dock = "Top"
$panelTop.Height = 45
$panelTop.Padding = New-Object System.Windows.Forms.Padding(8)

$txtInput = New-Object System.Windows.Forms.TextBox
$txtInput.Location = New-Object System.Drawing.Point(8, 10)
$txtInput.Size = New-Object System.Drawing.Size(210, 25)
$txtInput.Font = New-Object System.Drawing.Font("微軟正黑體", 10)

$btnAdd = New-Object System.Windows.Forms.Button
$btnAdd.Location = New-Object System.Drawing.Point(225, 9)
$btnAdd.Size = New-Object System.Drawing.Size(85, 27)
$btnAdd.Text = "新增代辦"
$btnAdd.Font = New-Object System.Drawing.Font("微軟正黑體", 9)
$btnAdd.BackColor = [System.Drawing.Color]::FromArgb(232, 213, 139)

$panelTop.Controls.Add($txtInput)
$panelTop.Controls.Add($btnAdd)
$form.Controls.Add($panelTop)

# 2. 中間清單區
$checklist = New-Object System.Windows.Forms.CheckedListBox
$checklist.Location = New-Object System.Drawing.Point(8, 50)
$checklist.Size = New-Object System.Drawing.Size(306, 230)
$checklist.Font = New-Object System.Drawing.Font("微軟正黑體", 10)
$checklist.BackColor = [System.Drawing.Color]::FromArgb(255, 253, 240)
$checklist.CheckOnClick = $true
$form.Controls.Add($checklist)

function Refresh-Listbox {
    $checklist.Items.Clear()
    foreach ($t in $global:notesData.tasks) {
        $tag = ""
        $r = if ($t.reminder_type) { $t.reminder_type } else { "none" }
        if ($r -eq "once") { $tag = " (單次: $($t.date) $($t.time))" }
        elseif ($r -eq "weekday") { $tag = " (平日 $($t.time))" }
        elseif ($r -eq "weekend") { $tag = " (假日 $($t.time))" }
        elseif ($r -eq "monthly") { 
            $daysStr = if ($t.days_of_month) { $t.days_of_month -join "," } else { "" }
            $tag = " (每月 ${daysStr}日 $($t.time))" 
        }
        $checklist.Items.Add("$($t.content)$tag", $t.checked)
    }
}
Refresh-Listbox

# 原生 C# 輸入框替代方案（指定 Owner 避免被置頂主視窗擋住）
function Show-InputDialog($ownerForm, $title, $prompt, $default) {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = $title
    $dlg.Size = New-Object System.Drawing.Size(320, 160)
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = "FixedDialog"
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.Owner = $ownerForm # 指定擁有者，讓子視窗正確浮在主視窗之上

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $prompt
    $lbl.Location = New-Object System.Drawing.Point(12, 15)
    $lbl.Size = New-Object System.Drawing.Size(280, 25)
    $lbl.Font = New-Object System.Drawing.Font("微軟正黑體", 9)
    $dlg.Controls.Add($lbl)

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Text = $default
    $txt.Location = New-Object System.Drawing.Point(12, 45)
    $txt.Size = New-Object System.Drawing.Size(280, 25)
    $txt.Font = New-Object System.Drawing.Font("微軟正黑體", 10)
    $dlg.Controls.Add($txt)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "確定"
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $btnOk.Location = New-Object System.Drawing.Point(136, 85)
    $btnOk.Size = New-Object System.Drawing.Size(75, 28)
    $dlg.AcceptButton = $btnOk
    $dlg.Controls.Add($btnOk)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "取消"
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $btnCancel.Location = New-Object System.Drawing.Point(217, 85)
    $btnCancel.Size = New-Object System.Drawing.Size(75, 28)
    $dlg.CancelButton = $btnCancel
    $dlg.Controls.Add($btnCancel)

    if ($dlg.ShowDialog($ownerForm) -eq [System.Windows.Forms.DialogResult]::OK) {
        return $txt.Text
    }
    return $null
}

# 右鍵選單（設定提醒）
$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$menuOnce = $contextMenu.Items.Add("📅 設定單次提醒")
$menuWeekday = $contextMenu.Items.Add("🔄 設定平日循環 (週一~週五)")
$menuWeekend = $contextMenu.Items.Add("🌴 設定假日循環 (週六~週日)")
$menuMonthly = $contextMenu.Items.Add("🗓️ 設定每月特定日期循環")
[void]$contextMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
$menuClear = $contextMenu.Items.Add("❌ 清除此項提醒")

$checklist.ContextMenuStrip = $contextMenu

# 3. 下方控制區
$panelBottom = New-Object System.Windows.Forms.Panel
$panelBottom.Dock = "Bottom"
$panelBottom.Height = 45
$panelBottom.BackColor = [System.Drawing.Color]::FromArgb(252, 235, 166)

$btnDel = New-Object System.Windows.Forms.Button
$btnDel.Location = New-Object System.Drawing.Point(8, 8)
$btnDel.Size = New-Object System.Drawing.Size(65, 28)
$btnDel.Text = "刪除"
$btnDel.Font = New-Object System.Drawing.Font("微軟正黑體", 9)
$btnDel.BackColor = [System.Drawing.Color]::FromArgb(232, 213, 139)

$btnSimple = New-Object System.Windows.Forms.Button
$btnSimple.Location = New-Object System.Drawing.Point(76, 8)
$btnSimple.Size = New-Object System.Drawing.Size(65, 28)
$btnSimple.Text = "簡易畫面"
$btnSimple.Font = New-Object System.Drawing.Font("微軟正黑體", 9)
$btnSimple.BackColor = [System.Drawing.Color]::FromArgb(232, 213, 139)

$chkTopmost = New-Object System.Windows.Forms.CheckBox
$chkTopmost.Location = New-Object System.Drawing.Point(145, 11)
$chkTopmost.Size = New-Object System.Drawing.Size(55, 24)
$chkTopmost.Text = "置頂"
$chkTopmost.Font = New-Object System.Drawing.Font("微軟正黑體", 9)
$chkTopmost.Checked = $global:notesData.topmost

$trackAlpha = New-Object System.Windows.Forms.TrackBar
$trackAlpha.Location = New-Object System.Drawing.Point(200, 8)
$trackAlpha.Size = New-Object System.Drawing.Size(115, 30)
$trackAlpha.Minimum = 3
$trackAlpha.Maximum = 10
$trackAlpha.Value = [int]($global:notesData.alpha * 10)
$trackAlpha.TickStyle = "None"

$panelBottom.Controls.Add($btnDel)
$panelBottom.Controls.Add($btnSimple)
$panelBottom.Controls.Add($chkTopmost)
$panelBottom.Controls.Add($trackAlpha)
$form.Controls.Add($panelBottom)

# 4. 簡易畫面容器
$simplePanel = New-Object System.Windows.Forms.Panel
$simplePanel.Dock = "Fill"
$simplePanel.BackColor = [System.Drawing.Color]::FromArgb(255, 248, 214)
$simplePanel.Visible = $false
$form.Controls.Add($simplePanel)

$lblSimpleText = New-Object System.Windows.Forms.Label
$lblSimpleText.Location = New-Object System.Drawing.Point(12, 12)
$lblSimpleText.Size = New-Object System.Drawing.Size(300, 300)
$lblSimpleText.Font = New-Object System.Drawing.Font("微軟正黑體", 11)
$lblSimpleText.ForeColor = [System.Drawing.Color]::FromArgb(44, 44, 44)
$simplePanel.Controls.Add($lblSimpleText)

# 支援滑鼠拖曳移動視窗
$script:isDragging = $false
$script:mouseCursor = $null

$SimpleDragStart = {
    if ($_.Button -eq "Left") {
        $script:isDragging = $true
        $script:mouseCursor = $_.Location
    }
}
$SimpleDragMove = {
    if ($script:isDragging) {
        $currentPos = [System.Windows.Forms.Cursor]::Position
        $form.Location = New-Object System.Drawing.Point(($currentPos.X - [int]$script:mouseCursor.X), ($currentPos.Y - [int]$script:mouseCursor.Y))
    }
}
$SimpleDragEnd = { $script:isDragging = $false }

$simplePanel.Add_MouseDown($SimpleDragStart)
$simplePanel.Add_MouseMove($SimpleDragMove)
$simplePanel.Add_MouseUp($SimpleDragEnd)

$lblSimpleText.Add_MouseDown({
    if ($_.Button -eq "Left") {
        $script:isDragging = $true
        $script:mouseCursor = New-Object System.Drawing.Point(([int]$_.X + [int]$lblSimpleText.Left), ([int]$_.Y + [int]$lblSimpleText.Top))
    }
})
$lblSimpleText.Add_MouseMove($SimpleDragMove)
$lblSimpleText.Add_MouseUp($SimpleDragEnd)

# 雙擊切換回主畫面
$simplePanel.Add_DoubleClick({ Switch-To-Main })
$lblSimpleText.Add_DoubleClick({ Switch-To-Main })

function Switch-To-Simple {
    $checkedItems = @()
    for ($i = 0; $i -lt $checklist.Items.Count; $i++) {
        if ($checklist.GetItemChecked($i)) {
            $checkedItems += "• " + $global:notesData.tasks[$i].content
        }
    }
    if ($checkedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("目前沒有打勾的代辦事項可顯示在簡易畫面！", "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    $lblSimpleText.Text = $checkedItems -join "`n"
    $form.FormBorderStyle = "None"
    $form.Size = New-Object System.Drawing.Size(300, [Math]::Max(70, $checkedItems.Count * 30 + 32))
    $panelTop.Visible = $false
    $checklist.Visible = $false
    $panelBottom.Visible = $false
    $simplePanel.Visible = $true
}

function Switch-To-Main {
    $form.FormBorderStyle = "Sizable"
    $form.Size = New-Object System.Drawing.Size(340, 390)
    $simplePanel.Visible = $false
    $panelTop.Visible = $true
    $checklist.Visible = $true
    $panelBottom.Visible = $true
    Refresh-Listbox
}

$btnSimple.Add_Click({ Switch-To-Simple })

# --- 事件處理 ---

$AddAction = {
    $val = $txtInput.Text.Trim()
    if (-not [string]::IsNullOrEmpty($val)) {
        $newItem = [PSCustomObject]@{ 
            content = $val; 
            checked = $false; 
            reminder_type = "none";
            time = (Get-Date).ToString("HH:mm");
            date = (Get-Date).ToString("yyyy-MM-dd");
            days_of_month = @();
            notified = $false
        }
        $global:notesData.tasks += $newItem
        Refresh-Listbox
        Save-Data
        $txtInput.Text = ""
    }
}
$btnAdd.Add_Click($AddAction)
$txtInput.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        &$AddAction
        $_.SuppressKeyPress = $true
    }
})

$btnDel.Add_Click({
    $selectedIndices = @($checklist.CheckedIndices) + @($checklist.SelectedIndex) | Select-Object -Unique | Sort-Object -Descending
    foreach ($idx in $selectedIndices) {
        if ($idx -ge 0 -and $idx -lt $global:notesData.tasks.Count) {
            $global:notesData.tasks = @($global:notesData.tasks | Where-Object { $_ -ne $global:notesData.tasks[$idx] })
        }
    }
    Refresh-Listbox
    Save-Data
})

$checklist.Add_ItemCheck({
    param($sender, $e)
    $tempTasks = @()
    for ($i = 0; $i -lt $checklist.Items.Count; $i++) {
        $isChecked = if ($i -eq $e.Index) { ($e.NewValue -eq 'Checked') } else { $checklist.GetItemChecked($i) }
        $t = $global:notesData.tasks[$i]
        $tempTasks += [PSCustomObject]@{
            content       = $t.content
            checked       = $isChecked
            reminder_type = $t.reminder_type
            time          = $t.time
            date          = $t.date
            days_of_month = $t.days_of_month
            notified      = $t.notified
        }
    }
    $global:notesData.tasks = $tempTasks
    Save-Data
})

# 右鍵選單功能
$menuOnce.Add_Click({
    $idx = $checklist.SelectedIndex
    if ($idx -lt 0) { return }
    $t = $global:notesData.tasks[$idx]
    
    $dInput = Show-InputDialog $form "設定單次提醒" "請輸入提醒日期 (YYYY-MM-DD):" $t.date
    if ([string]::IsNullOrEmpty($dInput)) { return }
    $tInput = Show-InputDialog $form "設定單次提醒" "請輸入提醒時間 (格式 HH:MM 或 0900):" $t.time
    if ([string]::IsNullOrEmpty($tInput)) { return }

    $formattedTime = Format-TimeInput $tInput
    $t.reminder_type = "once"
    $t.date = $dInput.Trim()
    $t.time = $formattedTime
    $t.notified = $false
    Refresh-Listbox
    Save-Data
})

$menuWeekday.Add_Click({
    $idx = $checklist.SelectedIndex
    if ($idx -lt 0) { return }
    $t = $global:notesData.tasks[$idx]
    $tInput = Show-InputDialog $form "設定平日循環" "請輸入平日提醒時間 (格式 HH:MM 或 0900):" $t.time
    if ([string]::IsNullOrEmpty($tInput)) { return }

    $t.reminder_type = "weekday"
    $t.time = Format-TimeInput $tInput
    $t.notified = $false
    Refresh-Listbox
    Save-Data
})

$menuWeekend.Add_Click({
    $idx = $checklist.SelectedIndex
    if ($idx -lt 0) { return }
    $t = $global:notesData.tasks[$idx]
    $tInput = Show-InputDialog $form "設定假日循環" "請輸入假日提醒時間 (格式 HH:MM 或 0900):" $t.time
    if ([string]::IsNullOrEmpty($tInput)) { return }

    $t.reminder_type = "weekend"
    $t.time = Format-TimeInput $tInput
    $t.notified = $false
    Refresh-Listbox
    Save-Data
})

$menuMonthly.Add_Click({
    $idx = $checklist.SelectedIndex
    if ($idx -lt 0) { return }
    $t = $global:notesData.tasks[$idx]
    
    $defaultDays = if ($t.days_of_month) { $t.days_of_month -join "," } else { "1" }
    $dInput = Show-InputDialog $form "每月特定日期循環" "請輸入每月哪幾日提醒（以逗號分隔，例如：1,15,25）:" $defaultDays
    if ([string]::IsNullOrEmpty($dInput)) { return }

    $tInput = Show-InputDialog $form "每月特定日期循環" "請輸入提醒時間 (格式 HH:MM 或 0900):" $t.time
    if ([string]::IsNullOrEmpty($tInput)) { return }

    $daysArr = $dInput -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^\d+$' } | ForEach-Object { [int]$_ }
    $t.reminder_type = "monthly"
    $t.days_of_month = @($daysArr)
    $t.time = Format-TimeInput $tInput
    $t.notified = $false
    Refresh-Listbox
    Save-Data
})

$menuClear.Add_Click({
    $idx = $checklist.SelectedIndex
    if ($idx -lt 0) { return }
    $t = $global:notesData.tasks[$idx]
    $t.reminder_type = "none"
    $t.notified = $false
    Refresh-Listbox
    Save-Data
})

$chkTopmost.Add_CheckedChanged({
    $form.TopMost = $chkTopmost.Checked
    $global:notesData.topmost = $chkTopmost.Checked
    Save-Data
})

$trackAlpha.Add_Scroll({
    $alphaVal = $trackAlpha.Value / 10.0
    $form.Opacity = $alphaVal
    $global:notesData.alpha = $alphaVal
    Save-Data
})

# 背景鬧鐘計時器（每 10 秒檢查一次）
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 10000
$timer.Add_Tick({
    $now = Get-Date
    $nowDateStr = $now.ToString("yyyy-MM-dd")
    $nowTimeStr = $now.ToString("HH:mm")
    $isWeekday = ($now.DayOfWeek -ne [DayOfWeek]::Saturday -and $now.DayOfWeek -ne [DayOfWeek]::Sunday)
    $isWeekend = ($now.DayOfWeek -eq [DayOfWeek]::Saturday -or $now.DayOfWeek -eq [DayOfWeek]::Sunday)
    $currentDay = $now.Day

    foreach ($t in $global:notesData.tasks) {
        $r = if ($t.reminder_type) { $t.reminder_type } else { "none" }
        if ($r -eq "none") { continue }

        $trigger = $false
        if ($r -eq "once" -and -not $t.notified) {
            if ($t.date -eq $nowDateStr -and $t.time -eq $nowTimeStr) {
                $trigger = $true
                $t.notified = $true
            }
        }
        elseif ($r -eq "weekday" -and $isWeekday) {
            if ($t.time -eq $nowTimeStr) { $trigger = $true }
        }
        elseif ($r -eq "weekend" -and $isWeekend) {
            if ($t.time -eq $nowTimeStr) { $trigger = $true }
        }
        elseif ($r -eq "monthly") {
            if ($t.days_of_month -contains $currentDay -and $t.time -eq $nowTimeStr) {
                $trigger = $true
            }
        }

        if ($trigger) {
            Save-Data
            Switch-To-Main
            $form.Show()
            $form.WindowState = "Normal"
            $form.Opacity = 1.0
            $trackAlpha.Value = 10
            $form.TopMost = $true
            $chkTopmost.Checked = $true
            [System.Windows.Forms.MessageBox]::Show($form, "提醒您：`n`n$($t.content)", "時間到！", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    }
})
$timer.Start()

[void]$form.ShowDialog()
#Requires -Version 3.0
# ================================================================
#  Outlook Email Notifier
#  Author  : Pravin Kadam
#  Version : 1.0
#  Info    : Monitors Outlook inbox and pops a desktop alert
#            whenever a new email lands — fully offline, no
#            third-party tools required.
# ================================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---- Settings (edit if needed) ---------------------------------
$POLL_SECONDS = 30      # how often to check inbox
$TOAST_MS     = 8000    # how long the popup stays on screen
$TOAST_W      = 375
$TOAST_H      = 108

# ---- Colours ---------------------------------------------------
$C_BG     = [System.Drawing.Color]::FromArgb(26, 26, 26)
$C_ACCENT = [System.Drawing.Color]::FromArgb(0, 120, 215)
$C_WHITE  = [System.Drawing.Color]::White
$C_GREY   = [System.Drawing.Color]::FromArgb(158, 158, 158)
$C_DIMX   = [System.Drawing.Color]::FromArgb(90, 90, 90)

# ================================================================
#  Connect to a running Outlook instance (or start one)
# ================================================================
function Connect-Outlook {
    try   { return [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application") }
    catch {
        try   { return New-Object -ComObject Outlook.Application }
        catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Outlook does not appear to be running.`n`nPlease open Outlook first, then start the notifier again.",
                "Email Notifier", "OK", "Warning")
            exit
        }
    }
}

# ================================================================
#  Build and show the toast popup
# ================================================================
function Show-Toast {
    param(
        [string]$User,
        [string]$Sender,
        [string]$Subject,
        [string]$Preview
    )

    # Reset elapsed counter for this toast
    $script:_ms = 0

    # ---- Form ----
    $toast = New-Object System.Windows.Forms.Form
    $toast.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
    $toast.TopMost         = $true
    $toast.ShowInTaskbar   = $false
    $toast.Size            = New-Object System.Drawing.Size($TOAST_W, $TOAST_H)
    $toast.BackColor       = $C_BG
    $toast.Opacity         = 0.96

    # Start just below screen edge; slide up to $endY
    $wa   = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $endY = $wa.Bottom - $TOAST_H - 14
    $toast.Location = New-Object System.Drawing.Point(($wa.Right - $TOAST_W - 14), ($wa.Bottom + 4))

    # ---- Left accent stripe ----
    $stripe = New-Object System.Windows.Forms.Panel
    $stripe.BackColor = $C_ACCENT
    $stripe.Bounds    = '0,0,4,{0}' -f $TOAST_H | ForEach-Object {
        [System.Drawing.Rectangle]::new(0, 0, 4, $TOAST_H)
    }
    $toast.Controls.Add($stripe)

    # ---- Envelope icon ----
    $ico = New-Object System.Windows.Forms.Label
    $ico.Text      = [char]0x2709        # ✉
    $ico.Font      = New-Object System.Drawing.Font("Segoe UI Symbol", 17)
    $ico.ForeColor = $C_ACCENT
    $ico.Bounds    = [System.Drawing.Rectangle]::new(10, 12, 34, 34)
    $ico.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $toast.Controls.Add($ico)

    # ---- Line 1: greeting ----
    $l1 = New-Object System.Windows.Forms.Label
    $l1.Text           = "Hi $User!   New email received"
    $l1.Font           = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $l1.ForeColor      = $C_WHITE
    $l1.Bounds         = [System.Drawing.Rectangle]::new(52, 10, 305, 18)
    $l1.AutoEllipsis   = $true
    $toast.Controls.Add($l1)

    # ---- Line 2: sender ----
    $l2 = New-Object System.Windows.Forms.Label
    $l2.Text           = "From :  $Sender"
    $l2.Font           = New-Object System.Drawing.Font("Segoe UI", 8.5)
    $l2.ForeColor      = $C_ACCENT
    $l2.Bounds         = [System.Drawing.Rectangle]::new(52, 31, 305, 17)
    $l2.AutoEllipsis   = $true
    $toast.Controls.Add($l2)

    # ---- Line 3: subject ----
    $l3 = New-Object System.Windows.Forms.Label
    $l3.Text           = $Subject
    $l3.Font           = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $l3.ForeColor      = $C_GREY
    $l3.Bounds         = [System.Drawing.Rectangle]::new(52, 51, 305, 16)
    $l3.AutoEllipsis   = $true
    $toast.Controls.Add($l3)

    # ---- Line 4: preview ----
    $short = if ($Preview.Length -gt 58) { $Preview.Substring(0, 55) + [char]0x2026 } else { $Preview }
    $l4 = New-Object System.Windows.Forms.Label
    $l4.Text           = $short
    $l4.Font           = New-Object System.Drawing.Font("Segoe UI", 7.5)
    $l4.ForeColor      = $C_GREY
    $l4.Bounds         = [System.Drawing.Rectangle]::new(52, 70, 305, 15)
    $l4.AutoEllipsis   = $true
    $toast.Controls.Add($l4)

    # ---- Progress bar (shrinks as timer counts down) ----
    $bar = New-Object System.Windows.Forms.Panel
    $bar.BackColor = $C_ACCENT
    $bar.Bounds    = [System.Drawing.Rectangle]::new(0, $TOAST_H - 3, $TOAST_W, 3)
    $toast.Controls.Add($bar)

    # ---- Close (×) button ----
    $xBtn = New-Object System.Windows.Forms.Label
    $xBtn.Text      = [char]0x00D7   # ×
    $xBtn.Font      = New-Object System.Drawing.Font("Segoe UI", 10)
    $xBtn.ForeColor = $C_DIMX
    $xBtn.Bounds    = [System.Drawing.Rectangle]::new($TOAST_W - 22, 4, 19, 19)
    $xBtn.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $xBtn.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $xBtn.Add_MouseEnter({ $xBtn.ForeColor = $C_WHITE })
    $xBtn.Add_MouseLeave({ $xBtn.ForeColor = $C_DIMX  })
    $xBtn.Add_Click({ $toast.Close() })
    $toast.Controls.Add($xBtn)

    # ---- Click anywhere to dismiss ----
    foreach ($ctrl in @($l1, $l2, $l3, $l4, $ico)) {
        $ctrl.Add_Click({ $toast.Close() })
        $ctrl.Cursor = [System.Windows.Forms.Cursors]::Hand
    }

    # ---- Slide-up animation ----
    $slide = New-Object System.Windows.Forms.Timer
    $slide.Interval = 8
    $slide.Add_Tick({
        if ($toast.Top -gt $endY) { $toast.Top -= 6 }
        else { $toast.Top = $endY; $slide.Stop() }
    })

    # ---- Auto-close + progress shrink ----
    $auto = New-Object System.Windows.Forms.Timer
    $auto.Interval = 40
    $auto.Add_Tick({
        $script:_ms += 40
        $ratio = [Math]::Max(0.0, 1.0 - ($script:_ms / $TOAST_MS))
        $bar.Width = [int]($TOAST_W * $ratio)
        if ($script:_ms -ge $TOAST_MS) { $auto.Stop(); $toast.Close() }
    })

    $toast.Add_Shown({ $slide.Start(); $auto.Start() })

    $toast.ShowDialog() | Out-Null

    $slide.Dispose()
    $auto.Dispose()
    $toast.Dispose()
}

# ================================================================
#  Main
# ================================================================

$ol   = Connect-Outlook
$ns   = $ol.GetNamespace("MAPI")

# Get inbox and current user
$inbox     = $ns.GetDefaultFolder(6)   # 6 = olFolderInbox
$fullName  = $ns.CurrentUser.Name
$firstName = ($fullName -split "[\s,]+" | Where-Object { $_ } | Select-Object -First 1)
if (-not $firstName) { $firstName = "User" }

# Baseline: don't alert about emails that arrived before we started
$inbox.Items.Sort("[ReceivedTime]", $true)
$seed          = $inbox.Items.GetFirst()
$script:latest = if ($seed) { $seed.ReceivedTime } else { [datetime]::Now }

Write-Host ""
Write-Host "  Email Notifier is running for: $fullName"
Write-Host "  Inbox check every $POLL_SECONDS seconds."
Write-Host "  Right-click the system tray icon to exit."
Write-Host ""

# ---- System tray icon ----
$tray                   = New-Object System.Windows.Forms.NotifyIcon
$tray.Icon              = [System.Drawing.SystemIcons]::Mail
$tray.Visible           = $true
$tray.Text              = "Email Notifier  —  $firstName"
$tray.BalloonTipTitle   = "Email Notifier started"
$tray.BalloonTipText    = "Watching inbox for $fullName"
$tray.BalloonTipIcon    = "Info"
$tray.ShowBalloonTip(3500)

$ctx      = New-Object System.Windows.Forms.ContextMenuStrip
$miExit   = $ctx.Items.Add("Exit Email Notifier")
$miExit.Add_Click({
    $tray.Visible = $false
    $poll.Stop()
    [System.Windows.Forms.Application]::Exit()
})
$tray.ContextMenuStrip = $ctx

# ---- Poll timer ----
$poll          = New-Object System.Windows.Forms.Timer
$poll.Interval = $POLL_SECONDS * 1000
$poll.Add_Tick({
    try {
        # Re-bind in case Outlook was restarted
        if (-not $script:ol -or [Runtime.InteropServices.Marshal]::IsComObject($script:ol) -eq $false) {
            $script:ol    = Connect-Outlook
            $script:ns    = $script:ol.GetNamespace("MAPI")
            $script:inbox = $script:ns.GetDefaultFolder(6)
        }

        $script:inbox.Items.Sort("[ReceivedTime]", $true)
        $item = $script:inbox.Items.GetFirst()

        if ($item -and ($item.ReceivedTime -gt $script:latest)) {
            $script:latest = $item.ReceivedTime

            $sender  = if ($item.SenderName)  { $item.SenderName }  else { "Unknown" }
            $subject = if ($item.Subject)      { $item.Subject }     else { "(no subject)" }
            $body    = ($item.Body -replace "[\r\n\t]+"," ").Trim()

            Show-Toast -User    $firstName `
                       -Sender  $sender    `
                       -Subject $subject   `
                       -Preview $body
        }
    }
    catch {
        Write-Warning "Poll error: $_"
    }
})

# Expose to closure scope
$script:ol    = $ol
$script:ns    = $ns
$script:inbox = $inbox

$poll.Start()

# Run the WinForms message loop (keeps script alive)
[System.Windows.Forms.Application]::Run()

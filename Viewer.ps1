Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ===========================================================================
# FIX: Force IE11 rendering engine for WebBrowser control
# ===========================================================================
$regPath = "HKCU:\Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION"
$exeName = [System.IO.Path]::GetFileName([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)
if (-not (Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }
Set-ItemProperty -Path $regPath -Name $exeName -Value 11001 -Type DWord -Force

# ===========================================================================
# C# HELPER
# ===========================================================================
$csharpCode = @"
using System;
using System.IO;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections.Generic;

public class MsgHelper {

    public static bool IsOutlookInstalled() {
        return Type.GetTypeFromProgID("Outlook.Application") != null;
    }

    public static string GetStrings(byte[] data) {
        var sb = new StringBuilder();
        foreach (byte b in data) {
            if (b >= 32 && b <= 126) sb.Append((char)b);
            else if (b == 10 || b == 13) sb.Append((char)b);
        }
        return sb.ToString();
    }

    public static string[] ExtractLinks(string html, string plain) {
        var links = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        string combined = (html ?? "") + "\n" + (plain ?? "");
        foreach (Match m in Regex.Matches(combined,
            @"(?:href|src|url)\s*=\s*[""']?(https?://[^\s""'<>]+)",
            RegexOptions.IgnoreCase))
            links.Add(m.Groups[1].Value.TrimEnd('.', ',', ';', ')', '>', ' '));
        foreach (Match m in Regex.Matches(combined,
            @"https?://[^\s""'<>\[\]{}|\\^]+",
            RegexOptions.IgnoreCase))
            links.Add(m.Groups[0].Value.TrimEnd('.', ',', ';', ')', '>', ' '));
        var list = new List<string>(links);
        list.Sort();
        return list.ToArray();
    }

    public static string UploadToVirusTotal(string apiKey, string fileName, byte[] fileBytes) {
        string boundary = "----VTBoundary" + DateTime.Now.Ticks.ToString("x");
        var req = (HttpWebRequest)WebRequest.Create("https://www.virustotal.com/api/v3/files");
        req.Method = "POST";
        req.Headers.Add("x-apikey", apiKey);
        req.ContentType = "multipart/form-data; boundary=" + boundary;
        using (var ms = new MemoryStream()) {
            byte[] nl = Encoding.ASCII.GetBytes("\r\n");
            string ph = "--" + boundary + "\r\n"
                + "Content-Disposition: form-data; name=\"file\"; filename=\"" + fileName + "\"\r\n"
                + "Content-Type: application/octet-stream\r\n\r\n";
            byte[] phb = Encoding.UTF8.GetBytes(ph);
            ms.Write(phb, 0, phb.Length);
            ms.Write(fileBytes, 0, fileBytes.Length);
            ms.Write(nl, 0, nl.Length);
            byte[] ft = Encoding.ASCII.GetBytes("--" + boundary + "--\r\n");
            ms.Write(ft, 0, ft.Length);
            byte[] body = ms.ToArray();
            req.ContentLength = body.Length;
            using (var s = req.GetRequestStream()) s.Write(body, 0, body.Length);
        }
        try {
            using (var resp = (HttpWebResponse)req.GetResponse())
            using (var sr = new StreamReader(resp.GetResponseStream()))
                return sr.ReadToEnd();
        } catch (WebException ex) {
            using (var sr = new StreamReader(ex.Response.GetResponseStream()))
                return "ERROR: " + sr.ReadToEnd();
        }
    }

    public static string GetAnalysis(string apiKey, string analysisId) {
        var req = (HttpWebRequest)WebRequest.Create(
            "https://www.virustotal.com/api/v3/analyses/" + analysisId);
        req.Method = "GET";
        req.Headers.Add("x-apikey", apiKey);
        try {
            using (var resp = (HttpWebResponse)req.GetResponse())
            using (var sr = new StreamReader(resp.GetResponseStream()))
                return sr.ReadToEnd();
        } catch (WebException ex) {
            using (var sr = new StreamReader(ex.Response.GetResponseStream()))
                return "ERROR: " + sr.ReadToEnd();
        }
    }
}
"@
Add-Type -TypeDefinition $csharpCode -ErrorAction SilentlyContinue

# ===========================================================================
# CONFIG
# ===========================================================================
$VT_API_KEY = "PUTYOURAPI KEY HERE"

# ===========================================================================
# FONTS
# ===========================================================================
$FONT_UI    = New-Object System.Drawing.Font("Segoe UI", 9)
$FONT_BOLD  = New-Object System.Drawing.Font("Segoe UI", 9,  [System.Drawing.FontStyle]::Bold)
$FONT_MONO  = New-Object System.Drawing.Font("Consolas", 9)

# ===========================================================================
# COLORS  (standard Windows light palette)
# ===========================================================================
$CLR_BG       = [System.Drawing.SystemColors]::Control
$CLR_WHITE    = [System.Drawing.Color]::White
$CLR_LBLFG    = [System.Drawing.Color]::FromArgb(30, 30, 30)
$CLR_INPUTBG  = [System.Drawing.Color]::White
$CLR_BTNGREEN = [System.Drawing.Color]::FromArgb(0, 120, 60)
$CLR_BTNBLUE  = [System.Drawing.Color]::FromArgb(0, 90, 180)
$CLR_BTNCOPY  = [System.Drawing.Color]::FromArgb(230, 230, 230)
$CLR_BTNCOPYFG= [System.Drawing.Color]::FromArgb(30, 30, 30)
$CLR_FLASHBG  = [System.Drawing.Color]::FromArgb(200, 240, 210)
$CLR_FLASHFG  = [System.Drawing.Color]::FromArgb(0, 100, 40)
$CLR_VTBG     = [System.Drawing.Color]::FromArgb(245, 250, 245)
$CLR_VTFG     = [System.Drawing.Color]::FromArgb(0, 80, 20)

# ===========================================================================
# STATE
# ===========================================================================
$script:lastOpenDir     = [Environment]::GetFolderPath("Desktop")
$script:attachmentData  = @{}
$script:attachmentBytes = @{}
$script:outlook         = $null
$script:currentMsg      = $null
$script:currentFileExt  = ""
$script:previewTempFile = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "mailviewer_preview.html")

# ===========================================================================
# LAYOUT  (sequential Y - no overlaps)
# ===========================================================================
$M      = 12    # outer margin
$IW     = 930   # inner content width
$LH     = 18    # label height
$SH     = 24    # single-line input height
$BH     = 28    # button height
$GAP    = 4     # label-to-control gap
$SEC    = 12    # section spacing
$COPY_W = 66    # copy button width
$BOX_W  = $IW - $COPY_W - 6   # input width when copy button present

$y = 30   # start below menu bar

# Section 1 - Message-ID
$Y_LBL_MSGID = $y ;  $y += $LH + $GAP
$Y_MSGID     = $y ;  $y += $SH + $SEC

# Section 2 - Headers
$Y_LBL_HDR   = $y ;  $y += $LH + $GAP
$Y_HDR       = $y ;  $HDR_H = 78 ; $y += $HDR_H + $SEC

# Section 3 - Body tabs
$Y_LBL_BODY  = $y ;  $y += $LH + $GAP
$Y_TABS      = $y ;  $TAB_H = 290 ; $y += $TAB_H + $SEC

# Section 4 - Links
$Y_LBL_LINKS = $y ;  $y += $LH + $GAP
$Y_LINKS     = $y ;  $LINK_H = 88 ; $y += $LINK_H + $SEC

# Section 5 - Attachments
$Y_LBL_ATT   = $y ;  $y += $LH + $GAP
$Y_ATT       = $y ;  $ATT_H = 76 ; $y += $ATT_H + $SEC

# Section 6 - Action buttons
$Y_BTNS      = $y ;  $y += $BH + $SEC

# Section 7 - VT Results
$Y_LBL_VT    = $y ;  $y += $LH + $GAP
$Y_VT        = $y ;  $VT_H = 140 ; $y += $VT_H + $M

$FORM_W = $IW + 2 * $M + 16
$FORM_H = $y + 42

# ===========================================================================
# CONTROL FACTORY FUNCTIONS
# ===========================================================================
function New-SectionLabel($text, $x, $y) {
    $l = New-Object System.Windows.Forms.Label
    $l.Text      = $text
    $l.Location  = [System.Drawing.Point]::new($x, $y)
    $l.AutoSize  = $true
    $l.Font      = $FONT_BOLD
    $l.ForeColor = $CLR_LBLFG
    return $l
}

function New-ReadBox($x, $y, $w, $h, $mono) {
    $t = New-Object System.Windows.Forms.TextBox
    $t.Location    = [System.Drawing.Point]::new($x, $y)
    $t.Size        = [System.Drawing.Size]::new($w, $h)
    $t.ReadOnly    = $true
    $t.BackColor   = $CLR_INPUTBG
    $t.BorderStyle = "FixedSingle"
    $t.Font        = if ($mono) { $FONT_MONO } else { $FONT_UI }
    if ($h -gt 28) { $t.Multiline = $true; $t.ScrollBars = "Vertical" }
    return $t
}

function New-CopyBtn($x, $y, $label) {
    $b = New-Object System.Windows.Forms.Button
    $b.Text      = $label
    $b.Location  = [System.Drawing.Point]::new($x, $y)
    $b.Size      = [System.Drawing.Size]::new($COPY_W, $BH)
    $b.FlatStyle = "Flat"
    $b.BackColor = $CLR_BTNCOPY
    $b.ForeColor = $CLR_BTNCOPYFG
    $b.Font      = $FONT_BOLD
    $b.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(180, 180, 180)
    $b.FlatAppearance.BorderSize  = 1
    $b.Cursor    = [System.Windows.Forms.Cursors]::Hand
    return $b
}

function New-ActionBtn($text, $x, $y, $w, $bgColor) {
    $b = New-Object System.Windows.Forms.Button
    $b.Text      = $text
    $b.Location  = [System.Drawing.Point]::new($x, $y)
    $b.Size      = [System.Drawing.Size]::new($w, $BH)
    $b.FlatStyle = "Flat"
    $b.BackColor = $bgColor
    $b.ForeColor = $CLR_WHITE
    $b.Font      = $FONT_BOLD
    $b.FlatAppearance.BorderSize = 0
    $b.Enabled   = $false
    $b.Cursor    = [System.Windows.Forms.Cursors]::Hand
    return $b
}

function Flash-Btn($btn, $resetLabel) {
    $btn.Text      = "Copied!"
    $btn.BackColor = $CLR_FLASHBG
    $btn.ForeColor = $CLR_FLASHFG
    $t = New-Object System.Windows.Forms.Timer
    $t.Interval = 1400
    $captBtn = $btn ; $captReset = $resetLabel ; $captTimer = $t
    $t.Add_Tick({
        $captBtn.Text      = $captReset
        $captBtn.BackColor = $CLR_BTNCOPY
        $captBtn.ForeColor = $CLR_BTNCOPYFG
        $captTimer.Stop() ; $captTimer.Dispose()
    })
    $t.Start()
}

# ===========================================================================
# HELPER: Safely render HTML in WebBrowser (IE11 mode via temp file)
# ===========================================================================
function Set-WebContent($html) {
    if ([string]::IsNullOrWhiteSpace($html)) {
        $fallback = "<html><head><meta charset='utf-8'></head><body style='font-family:Segoe UI;color:gray;padding:16px'>No HTML body available.</body></html>"
        [System.IO.File]::WriteAllText($script:previewTempFile, $fallback, [System.Text.Encoding]::UTF8)
    } else {
        # Wrap bare HTML fragments in a proper shell
        if ($html -notmatch '(?i)<html') {
            $html = "<html><head><meta charset='utf-8'><meta http-equiv='X-UA-Compatible' content='IE=11'></head><body>$html</body></html>"
        } elseif ($html -notmatch '(?i)X-UA-Compatible') {
            # Inject compatibility meta if missing
            $html = $html -replace '(?i)(<head[^>]*>)', '$1<meta http-equiv="X-UA-Compatible" content="IE=11">'
        }
        [System.IO.File]::WriteAllText($script:previewTempFile, $html, [System.Text.Encoding]::UTF8)
    }
    $webBody.Navigate("file:///$($script:previewTempFile.Replace('\','/'))")
}

# ===========================================================================
# FORM
# ===========================================================================
$form = New-Object System.Windows.Forms.Form
$form.Text          = "Mail Security Viewer - ColdFire.at"
$form.ClientSize    = [System.Drawing.Size]::new($FORM_W, $FORM_H)
$form.MinimumSize   = [System.Drawing.Size]::new(700, 700)
$form.StartPosition = "CenterScreen"
$form.BackColor     = $CLR_BG
$form.Font          = $FONT_UI

# Menu
$mainMenu = New-Object System.Windows.Forms.MenuStrip
$fileMenu = New-Object System.Windows.Forms.ToolStripMenuItem("File")
$openItem = New-Object System.Windows.Forms.ToolStripMenuItem("Open Email File...")
$openItem.ShortcutKeys = "Control, O"
$exitItem = New-Object System.Windows.Forms.ToolStripMenuItem("Exit")
$fileMenu.DropDownItems.AddRange(@($openItem, $exitItem))
$mainMenu.Items.Add($fileMenu) | Out-Null
$form.MainMenuStrip = $mainMenu
$form.Controls.Add($mainMenu)

# ===========================================================================
# SECTION 1 - Message-ID
# ===========================================================================
$lblMsgID     = New-SectionLabel "Message-ID" $M $Y_LBL_MSGID
$txtMsgID     = New-ReadBox $M $Y_MSGID $BOX_W $SH $false
$txtMsgID.Anchor = "Top, Left, Right"
$btnCopyMsgID = New-CopyBtn ($M + $BOX_W + 6) $Y_MSGID "Copy"
$btnCopyMsgID.Anchor = "Top, Right"

# ===========================================================================
# SECTION 2 - Mail Headers
# ===========================================================================
$lblHeader     = New-SectionLabel "Mail Headers" $M $Y_LBL_HDR
$txtHeader     = New-ReadBox $M $Y_HDR $BOX_W $HDR_H $false
$txtHeader.Anchor = "Top, Left, Right"
$btnCopyHeader = New-CopyBtn ($M + $BOX_W + 6) $Y_HDR "Copy"
$btnCopyHeader.Anchor = "Top, Right"

# ===========================================================================
# SECTION 3 - Mail Body tabs
# ===========================================================================
$lblBody = New-SectionLabel "Mail Body" $M $Y_LBL_BODY

$tabControl          = New-Object System.Windows.Forms.TabControl
$tabControl.Location = [System.Drawing.Point]::new($M, $Y_TABS)
$tabControl.Size     = [System.Drawing.Size]::new($IW, $TAB_H)
$tabControl.Anchor   = "Top, Left, Right"
$tabControl.Font     = $FONT_UI

$tabRender = New-Object System.Windows.Forms.TabPage ; $tabRender.Text = "HTML Rendered"
$tabText   = New-Object System.Windows.Forms.TabPage ; $tabText.Text   = "Plain Text"
$tabSource = New-Object System.Windows.Forms.TabPage ; $tabSource.Text = "HTML Source"
$tabControl.Controls.AddRange(@($tabRender, $tabText, $tabSource))

$webBody = New-Object System.Windows.Forms.WebBrowser
$webBody.Dock = "Fill"
$webBody.ScriptErrorsSuppressed = $true
$webBody.IsWebBrowserContextMenuEnabled = $false
$tabRender.Controls.Add($webBody)

$txtBody = New-Object System.Windows.Forms.TextBox
$txtBody.Dock = "Fill" ; $txtBody.Multiline = $true ; $txtBody.ScrollBars = "Vertical"
$txtBody.ReadOnly = $true ; $txtBody.Font = $FONT_UI ; $txtBody.BorderStyle = "None"
$tabText.Controls.Add($txtBody)

$txtSource = New-Object System.Windows.Forms.TextBox
$txtSource.Dock = "Fill" ; $txtSource.Multiline = $true ; $txtSource.ScrollBars = "Both"
$txtSource.ReadOnly = $true ; $txtSource.Font = $FONT_MONO ; $txtSource.BorderStyle = "None"
$tabSource.Controls.Add($txtSource)

# ===========================================================================
# SECTION 4 - Links found in mail
# ===========================================================================
$lblLinks = New-SectionLabel "Links found in mail (0)" $M $Y_LBL_LINKS

$listLinks           = New-Object System.Windows.Forms.ListBox
$listLinks.Location  = [System.Drawing.Point]::new($M, $Y_LINKS)
$listLinks.Size      = [System.Drawing.Size]::new($BOX_W, $LINK_H)
$listLinks.Anchor    = "Bottom, Left, Right"
$listLinks.Font      = $FONT_MONO
$listLinks.BorderStyle = "FixedSingle"
$listLinks.SelectionMode = "MultiExtended"
$listLinks.HorizontalScrollbar = $true

$btnCopyLinks = New-CopyBtn ($M + $BOX_W + 6) $Y_LINKS "Copy All"
$btnCopyLinks.Anchor = "Bottom, Right"

# ===========================================================================
# SECTION 5 - Attachments
# ===========================================================================
$lblAttach = New-SectionLabel "Attachments" $M $Y_LBL_ATT

$listAttach          = New-Object System.Windows.Forms.ListBox
$listAttach.Location = [System.Drawing.Point]::new($M, $Y_ATT)
$listAttach.Size     = [System.Drawing.Size]::new($IW, $ATT_H)
$listAttach.Anchor   = "Bottom, Left, Right"
$listAttach.Font     = $FONT_UI
$listAttach.BorderStyle = "FixedSingle"
$listAttach.SelectionMode = "MultiExtended"

# ===========================================================================
# SECTION 6 - Action buttons
# ===========================================================================
$chkOpenDir          = New-Object System.Windows.Forms.CheckBox
$chkOpenDir.Text     = "Open folder after extraction"
$chkOpenDir.Location = [System.Drawing.Point]::new($M, $Y_BTNS)
$chkOpenDir.Size     = [System.Drawing.Size]::new(230, $BH)
$chkOpenDir.Checked  = $true
$chkOpenDir.Font     = $FONT_UI
$chkOpenDir.Anchor   = "Bottom, Left"

$btnExtract = New-ActionBtn "Extract Selected" 250 $Y_BTNS 180 $CLR_BTNGREEN
$btnExtract.Anchor = "Bottom, Left"

$btnUpload  = New-ActionBtn "Upload to VirusTotal" 440 $Y_BTNS 200 $CLR_BTNBLUE
$btnUpload.Anchor  = "Bottom, Left"

# ===========================================================================
# SECTION 7 - VirusTotal results
# ===========================================================================
$lblVTStatus = New-SectionLabel "VirusTotal Results" $M $Y_LBL_VT

$txtVTResult           = New-Object System.Windows.Forms.TextBox
$txtVTResult.Location  = [System.Drawing.Point]::new($M, $Y_VT)
$txtVTResult.Size      = [System.Drawing.Size]::new($BOX_W, $VT_H)
$txtVTResult.Multiline = $true
$txtVTResult.ScrollBars= "Vertical"
$txtVTResult.ReadOnly  = $true
$txtVTResult.BackColor = $CLR_VTBG
$txtVTResult.ForeColor = $CLR_VTFG
$txtVTResult.Font      = $FONT_MONO
$txtVTResult.BorderStyle = "FixedSingle"
$txtVTResult.Anchor    = "Bottom, Left, Right"

$btnCopyVT = New-CopyBtn ($M + $BOX_W + 6) $Y_VT "Copy"
$btnCopyVT.Anchor = "Bottom, Right"

# ===========================================================================
# ADD ALL CONTROLS
# ===========================================================================
$form.Controls.AddRange(@(
    $lblMsgID,    $txtMsgID,    $btnCopyMsgID,
    $lblHeader,   $txtHeader,   $btnCopyHeader,
    $lblBody,     $tabControl,
    $lblLinks,    $listLinks,   $btnCopyLinks,
    $lblAttach,   $listAttach,
    $chkOpenDir,  $btnExtract,  $btnUpload,
    $lblVTStatus, $txtVTResult, $btnCopyVT
))

# ===========================================================================
# HELPER FUNCTIONS
# ===========================================================================
function Get-JsonValue($json, $key) {
    if ($json -match """$key""\s*:\s*""([^""]+)""") { return $Matches[1] }
    if ($json -match """$key""\s*:\s*([0-9]+)")      { return $Matches[1] }
    return $null
}

function Format-VTResult($json) {
    $sb = [System.Text.StringBuilder]::new()
    $status = Get-JsonValue $json "status"
    [void]$sb.AppendLine("STATUS               : $status")
    if ($json -match '"stats"\s*:\s*\{([^}]+)\}') {
        $blk = $Matches[1]
        foreach ($metric in @("malicious","suspicious","undetected","harmless","timeout","failure","type-unsupported")) {
            if ($blk -match """$metric""\s*:\s*([0-9]+)") {
                [void]$sb.AppendLine(("{0,-22} : {1}" -f $metric.ToUpper(), $Matches[1]))
            }
        }
    }
    $id = Get-JsonValue $json "id"
    if ($id) {
        [void]$sb.AppendLine("")
        [void]$sb.AppendLine("Analysis URL:")
        [void]$sb.AppendLine("https://www.virustotal.com/gui/file-analysis/$id")
    }
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("--- Raw JSON (first 2000 chars) ---")
    [void]$sb.AppendLine($json.Substring(0, [Math]::Min($json.Length, 2000)))
    return $sb.ToString()
}

function Populate-Links($html, $plain) {
    $listLinks.Items.Clear()
    $links = [MsgHelper]::ExtractLinks($html, $plain)
    foreach ($u in $links) { $listLinks.Items.Add($u) | Out-Null }
    $lblLinks.Text = "Links found in mail ($($links.Count))"
}

function Reset-All {
    $listAttach.Items.Clear()
    $listLinks.Items.Clear()
    $script:attachmentData  = @{}
    $script:attachmentBytes = @{}
    $txtMsgID.Text    = ""
    $txtHeader.Text   = ""
    $txtBody.Text     = ""
    $txtSource.Text   = ""
    $txtVTResult.Text = ""
    $lblLinks.Text    = "Links found in mail (0)"
    $btnExtract.Enabled = $false
    $btnUpload.Enabled  = $false
    Set-WebContent ""
}

# ===========================================================================
# LOAD FUNCTIONS
# ===========================================================================
function Load-EML($path) {
    try {
        $stream = New-Object -ComObject ADODB.Stream
        $stream.Open()
        $stream.LoadFromFile($path)
        $script:currentMsg = New-Object -ComObject CDO.Message
        $script:currentMsg.DataSource.OpenObject($stream, "_Stream")
        $script:currentFileExt = "eml"

        $txtMsgID.Text  = try { $script:currentMsg.Fields.Item("urn:schemas:mailheader:message-id").Value } catch { "N/A" }
        $txtHeader.Text = ($script:currentMsg.Fields | ForEach-Object { "$($_.Name): $($_.Value)`r`n" }) -join ""
        $txtBody.Text   = $script:currentMsg.TextBody
        $txtSource.Text = $script:currentMsg.HTMLBody
        Set-WebContent $script:currentMsg.HTMLBody
        Populate-Links $script:currentMsg.HTMLBody $script:currentMsg.TextBody

        foreach ($at in $script:currentMsg.Attachments) {
            $listAttach.Items.Add($at.FileName) | Out-Null
            $script:attachmentData[$at.FileName] = $at
            $tmp = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), $at.FileName)
            try {
                $at.SaveToFile($tmp)
                $script:attachmentBytes[$at.FileName] = [System.IO.File]::ReadAllBytes($tmp)
                Remove-Item $tmp -ErrorAction SilentlyContinue
            } catch {}
        }
        $stream.Close()
    } catch {
        [System.Windows.Forms.MessageBox]::Show("EML Load Error: " + $_.Exception.Message, "Error", "OK", "Error")
    }
}

function Load-MSG($path) {
    if (-not [MsgHelper]::IsOutlookInstalled()) {
        $txtMsgID.Text  = "OUTLOOK NOT INSTALLED"
        $txtHeader.Text = "Fallback mode - reading printable strings from binary file."
        $bytes = [System.IO.File]::ReadAllBytes($path)
        $txtBody.Text = [MsgHelper]::GetStrings($bytes)
        Set-WebContent ""
        Populate-Links "" $txtBody.Text
        return
    }
    try {
        if ($null -eq $script:outlook) { $script:outlook = New-Object -ComObject Outlook.Application }
        $fullPath = (Resolve-Path $path).Path
        $script:currentMsg = $script:outlook.CreateItemFromTemplate($fullPath)
        $script:currentFileExt = "msg"

        $txtBody.Text   = $script:currentMsg.Body
        $txtSource.Text = $script:currentMsg.HTMLBody
        Set-WebContent $script:currentMsg.HTMLBody

        try { $txtMsgID.Text  = $script:currentMsg.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x1035001E") } catch { $txtMsgID.Text = "N/A" }
        try { $txtHeader.Text = $script:currentMsg.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E") } catch { $txtHeader.Text = "Headers restricted by Outlook" }

        Populate-Links $script:currentMsg.HTMLBody $script:currentMsg.Body

        for ($i = 1; $i -le $script:currentMsg.Attachments.Count; $i++) {
            $at = $script:currentMsg.Attachments.Item($i)
            $listAttach.Items.Add($at.FileName) | Out-Null
            $script:attachmentData[$at.FileName] = $i
            $tmp = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), $at.FileName)
            try {
                $at.SaveAsFile($tmp)
                $script:attachmentBytes[$at.FileName] = [System.IO.File]::ReadAllBytes($tmp)
                Remove-Item $tmp -ErrorAction SilentlyContinue
            } catch {}
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("MSG Error: " + $_.Exception.Message, "Error", "OK", "Error")
    }
}

# ===========================================================================
# EVENTS
# ===========================================================================

$openItem.Add_Click({
    $fd = New-Object System.Windows.Forms.OpenFileDialog
    $fd.Filter = "Email Files|*.eml;*.msg|All Files|*.*"
    $fd.InitialDirectory = $script:lastOpenDir
    $fd.Title = "Open Email File"
    if ($fd.ShowDialog() -eq "OK") {
        $script:lastOpenDir = [System.IO.Path]::GetDirectoryName($fd.FileName)
        Reset-All
        $ext = [System.IO.Path]::GetExtension($fd.FileName).ToLower()
        if ($ext -eq ".msg") { Load-MSG $fd.FileName } else { Load-EML $fd.FileName }
    }
})

$btnCopyMsgID.Add_Click({
    if ($txtMsgID.Text -ne "") {
        [System.Windows.Forms.Clipboard]::SetText($txtMsgID.Text)
        Flash-Btn $btnCopyMsgID "Copy"
    }
})

$btnCopyHeader.Add_Click({
    if ($txtHeader.Text -ne "") {
        [System.Windows.Forms.Clipboard]::SetText($txtHeader.Text)
        Flash-Btn $btnCopyHeader "Copy"
    }
})

$btnCopyLinks.Add_Click({
    if ($listLinks.Items.Count -gt 0) {
        $all = ($listLinks.Items | ForEach-Object { $_.ToString() }) -join "`r`n"
        [System.Windows.Forms.Clipboard]::SetText($all)
        Flash-Btn $btnCopyLinks "Copy All"
    }
})

$btnCopyVT.Add_Click({
    if ($txtVTResult.Text -ne "") {
        [System.Windows.Forms.Clipboard]::SetText($txtVTResult.Text)
        Flash-Btn $btnCopyVT "Copy"
    }
})

$listLinks.Add_DoubleClick({
    if ($listLinks.SelectedItem) {
        $url = $listLinks.SelectedItem.ToString()
        $r = [System.Windows.Forms.MessageBox]::Show(
            "Open this URL in your default browser?`n`n$url",
            "Open External Link", "YesNo", "Warning")
        if ($r -eq "Yes") { Start-Process $url }
    }
})

$listAttach.Add_SelectedIndexChanged({
    $on = ($listAttach.SelectedItems.Count -gt 0)
    $btnExtract.Enabled = $on
    $btnUpload.Enabled  = $on
})

$btnExtract.Add_Click({
    if ($listAttach.SelectedItems.Count -eq 0) { return }
    $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
    $fbd.Description = "Choose destination folder for extracted attachments"
    if ($fbd.ShowDialog() -eq "OK") {
        $errors = @()
        foreach ($name in $listAttach.SelectedItems) {
            $dest = [System.IO.Path]::Combine($fbd.SelectedPath, $name)
            try {
                if ($script:currentFileExt -eq "msg") {
                    $script:currentMsg.Attachments.Item($script:attachmentData[$name]).SaveAsFile($dest)
                } else {
                    $script:attachmentData[$name].SaveToFile($dest)
                }
            } catch { $errors += "$name : " + $_.Exception.Message }
        }
        if ($errors.Count -gt 0) {
            [System.Windows.Forms.MessageBox]::Show(($errors -join "`n"), "Save Errors", "OK", "Warning")
        }
        if ($chkOpenDir.Checked) { Start-Process explorer.exe $fbd.SelectedPath }
    }
})

$btnUpload.Add_Click({
    if ($listAttach.SelectedItems.Count -eq 0) { return }
    $txtVTResult.Text = "Uploading to VirusTotal - please wait..."
    $form.Refresh()
    $sb = [System.Text.StringBuilder]::new()

    foreach ($name in $listAttach.SelectedItems) {
        [void]$sb.AppendLine("============================================================")
        [void]$sb.AppendLine("FILE: $name")
        [void]$sb.AppendLine("============================================================")

        $bytes = $script:attachmentBytes[$name]
        if ($null -eq $bytes -or $bytes.Length -eq 0) {
            [void]$sb.AppendLine("ERROR: No bytes available for upload.")
            [void]$sb.AppendLine()
            continue
        }

        try {
            [void]$sb.AppendLine("Uploading $([Math]::Round($bytes.Length / 1KB, 1)) KB...")
            $txtVTResult.Text = $sb.ToString()
            $form.Refresh()

            $uploadResp = [MsgHelper]::UploadToVirusTotal($VT_API_KEY, $name, $bytes)
            if ($uploadResp -match "ERROR:") {
                [void]$sb.AppendLine("Upload failed: $uploadResp")
                [void]$sb.AppendLine()
                continue
            }

            $aid = $null
            if ($uploadResp -match '"id"\s*:\s*"([^"]+)"') { $aid = $Matches[1] }
            if (-not $aid) {
                [void]$sb.AppendLine("Could not extract analysis ID from response.")
                [void]$sb.AppendLine($uploadResp)
                [void]$sb.AppendLine()
                continue
            }

            [void]$sb.AppendLine("Analysis ID : $aid")
            [void]$sb.AppendLine("Polling for results (up to 60 seconds)...")
            $txtVTResult.Text = $sb.ToString()
            $form.Refresh()

            $result = "" ; $status = ""
            for ($p = 1; $p -le 12; $p++) {
                Start-Sleep -Seconds 5
                $result = [MsgHelper]::GetAnalysis($VT_API_KEY, $aid)
                $status = Get-JsonValue $result "status"
                [void]$sb.AppendLine("  Poll $p/12 - Status: $status")
                $txtVTResult.Text = $sb.ToString()
                $form.Refresh()
                if ($status -eq "completed") { break }
            }

            [void]$sb.AppendLine()
            [void]$sb.AppendLine((Format-VTResult $result))
            [void]$sb.AppendLine()

        } catch {
            [void]$sb.AppendLine("Exception: " + $_.Exception.Message)
            [void]$sb.AppendLine()
        }
    }
    $txtVTResult.Text = $sb.ToString()
})

# ===========================================================================
# CLEANUP
# ===========================================================================
$form.Add_FormClosing({
    if ($null -ne $script:currentMsg) {
        try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($script:currentMsg) | Out-Null } catch {}
    }
    if ($null -ne $script:outlook) {
        try { $script:outlook.Quit() ; [System.Runtime.Interopservices.Marshal]::ReleaseComObject($script:outlook) | Out-Null } catch {}
    }
    # Clean up temp preview file
    if (Test-Path $script:previewTempFile) {
        Remove-Item $script:previewTempFile -ErrorAction SilentlyContinue
    }
    $script:attachmentBytes = @{}
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
})

$exitItem.Add_Click({ $form.Close() })
$form.ShowDialog() | Out-Null
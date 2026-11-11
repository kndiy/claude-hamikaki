param([string]$InFile, [string]$Out, [string]$Json)
$ErrorActionPreference = 'Stop'
$P = Get-Content -Raw -Encoding UTF8 $Json | ConvertFrom-Json
$M = [System.Reflection.Missing]::Value
$NAVY = 6558747; $TEAL = 10859776; $MINT = 16119784; $YEL = 13431551
$GREEN = 13561798; $RED = 13551615; $BLUE = 16247773; $GRID = 12566463
$ci = [Globalization.CultureInfo]::InvariantCulture
function D($s) { [datetime]::ParseExact($s, 'dd/MM/yyyy', $ci).ToOADate() }
function Cl($sh, [int]$row, [int]$col) { $sh.Range([string][char](64 + $col) + [string]$row) }
function SetV($rg, $v) { [void]$rg.GetType().InvokeMember("Value2", [Reflection.BindingFlags]::SetProperty, $null, $rg, @($v)) }

Copy-Item $InFile $Out -Force
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Open($Out)
  $src = $wb.Worksheets.Item($P.src_sheet)
  $ws = $wb.Worksheets.Add($src); $ws.Move([Type]::Missing, $src); $ws = $wb.Worksheets.Item($src.Index + 1)
  $ws.Name = $P.sheet
  $ws.Tab.Color = $TEAL
  $xl.ActiveWindow.DisplayGridlines = $false
  $ws.Cells.Font.Name = 'Calibri'; $ws.Cells.Font.Size = 11
  $ws.Cells.VerticalAlignment = -4108

  $widths = @(5, 12, 15, 50, 22, 10.5, 10.5, 32, 15, 26, 7, 7, 7, 7, 7)
  for ($i = 0; $i -lt $widths.Count; $i++) { $ws.Columns.Item($i + 1).ColumnWidth = [double]$widths[$i] }

  # letterhead copied from the week sheet
  $src.Range('A1:E5').Copy($ws.Range('A1')) | Out-Null
  for ($r = 1; $r -le 5; $r++) { $ws.Rows.Item($r).RowHeight = $src.Rows.Item($r).RowHeight }
  for ($r = 2; $r -le 5; $r++) { $ws.Range("A${r}:O${r}").UnMerge(); $ws.Range("A${r}:O${r}").Merge() }
  $ws.Range('D1:O1').UnMerge(); $ws.Range('D1:O1').Merge()
  $ws.Range('A1').ClearContents() | Out-Null

  function Band($row, $text, $fill, $size) {
    $rg = $ws.Range("A${row}:O${row}"); $rg.Merge(); $rg.Value2 = $text
    $rg.Interior.Color = $fill; $rg.Font.Color = 16777215; $rg.Font.Bold = $true; $rg.Font.Size = $size
    $rg.HorizontalAlignment = -4131; $rg.IndentLevel = 1
    $ws.Rows.Item($row).RowHeight = [double]22
  }
  function Hdr($row, $cells, $last2) {
    if (-not $last2) { $last2 = $row }
    $rg = $ws.Range("A${row}:O${last2}")
    $rg.Interior.Color = $MINT; $rg.Font.Color = $NAVY; $rg.Font.Bold = $true
    $rg.HorizontalAlignment = -4108; $rg.WrapText = $true
    if ($last2 -eq $row) { $ws.Rows.Item($row).RowHeight = [double]30 }
    foreach ($c in $cells) {
      $one = $ws.Range($c[0]); if ($one.Cells.Count -gt 1) { $one.MergeCells = $true }
      SetV ($ws.Range($c[0]).Cells.Item(1)) ($c[1])
    }
  }
  function Grid($addr) {
    $b = $ws.Range($addr).Borders
    foreach ($k in 7, 8, 9, 10, 11, 12) { $b.Item($k).LineStyle = 1; $b.Item($k).Weight = 2; $b.Item($k).Color = $GRID }
  }

  # title
  $t = $ws.Range('A7:O7'); $t.Merge(); $t.Value2 = $P.title
  $t.Font.Bold = $true; $t.Font.Size = 15; $t.Font.Color = $NAVY; $t.HorizontalAlignment = -4108
  $ws.Rows.Item(7).RowHeight = [double]26
  $t = $ws.Range('A8:O8'); $t.Merge(); $t.Value2 = $P.subtitle
  $t.Font.Italic = $true; $t.Font.Size = 10; $t.WrapText = $true; $t.HorizontalAlignment = -4108
  $ws.Rows.Item(8).RowHeight = [double]30

  # ---- layout rows first (formulas in section 1 need task ranges)
  $gRow = 10; $gHdr = 11; $g0 = 12
  $nG = $P.goals.Count
  $tBand = $g0 + $nG + 1; $tHdr1 = $tBand + 1; $tHdr2 = $tBand + 2
  $r = $tHdr2 + 1; $first = $r
  $taskRows = @(); $checkRows = @()
  foreach ($g in $P.groups) {
    $rg = $ws.Range("A${r}:O${r}"); $rg.Merge(); $rg.Value2 = $g.title
    $rg.Interior.Color = $TEAL; $rg.Font.Color = 16777215; $rg.Font.Bold = $true; $rg.IndentLevel = 1
    $ws.Rows.Item($r).RowHeight = [double]20
    $r++
    $n = 1
    foreach ($tk in $g.tasks) {
      SetV (Cl $ws $r 1) ($n)
      SetV (Cl $ws $r 2) ($g.key)
      SetV (Cl $ws $r 3) ($tk[0])
      SetV (Cl $ws $r 4) ($tk[1])
      SetV (Cl $ws $r 5) ($tk[2])
      SetV (Cl $ws $r 6) ((D $tk[3]))
      SetV (Cl $ws $r 7) ((D $tk[4]))
      SetV (Cl $ws $r 8) ($tk[5])
      SetV (Cl $ws $r 10) ($tk[6])
      if ($tk[7] -eq 1) { $checkRows += $r }
      $taskRows += $r
      $r++; $n++
    }
  }
  $last = $r - 1
  $B = "`$B`$${first}:`$B`$${last}"; $Dd = "`$D`$${first}:`$D`$${last}"
  $G = "`$G`$${first}:`$G`$${last}"; $I = "`$I`$${first}:`$I`$${last}"
  $done = $P.done

  # ---- section 1: goals
  Band $gRow $P.sec_goals $NAVY 12
  $h = $P.goal_hdr
  Hdr $gHdr @(@("A$gHdr", $h[0]), @("B${gHdr}:C$gHdr", $h[1]), @("D$gHdr", $h[2]), @("E$gHdr", $h[3]), @("F$gHdr", $h[4]), @("G$gHdr", $h[5]), @("H$gHdr", $h[6]), @("I$gHdr", $h[7]), @("J${gHdr}:O$gHdr", $h[8]))
  for ($k = 0; $k -lt $nG; $k++) {
    $row = $g0 + $k; $go = $P.goals[$k]
    SetV (Cl $ws $row 1) ($k + 1)
    $ws.Range("B${row}:C$row").Merge(); SetV (Cl $ws $row 2) ($go.key)
    (Cl $ws $row 2).Font.Bold = $true; (Cl $ws $row 2).Font.Color = $NAVY
    SetV (Cl $ws $row 4) ($go.goal)
    SetV (Cl $ws $row 5) ($go.owner)
    SetV (Cl $ws $row 6) ((D $go.due)); (Cl $ws $row 6).Font.Bold = $true
    (Cl $ws $row 7).Formula = "=F$row-TODAY()"
    $cnt = "COUNTIFS($B,`$B$row,$I,`"*$done*`")"; $tot = "COUNTIF($B,`$B$row)"
    (Cl $ws $row 8).Formula = "=$cnt&`"/`"&$tot&`" (`"&TEXT(IFERROR($cnt/$tot,0),`"0%`")&`")`""
    (Cl $ws $row 9).Formula = "=SUMPRODUCT(($B=`$B$row)*($G<TODAY())*ISERROR(SEARCH(`"$done`",$I)))"
    $mn = "AGGREGATE(15,6,$G/(($B=`$B$row)*ISERROR(SEARCH(`"$done`",$I))),1)"
    $ws.Range("J${row}:O$row").Merge()
    (Cl $ws $row 10).Formula = "=IFERROR(TEXT($mn,`"dd/mm`")&`" - `"&INDEX($Dd,MATCH(1,INDEX(($B=`$B$row)*ISERROR(SEARCH(`"$done`",$I))*($G=$mn),0),0)),`"-`")"
    $ws.Rows.Item($row).RowHeight = [double]62
  }
  $gl = $g0 + $nG - 1
  $rg = $ws.Range("A${g0}:O$gl"); $rg.WrapText = $true
  $ws.Range("A${g0}:A$gl").HorizontalAlignment = -4108
  $ws.Range("F${g0}:I$gl").HorizontalAlignment = -4108
  $ws.Range("F${g0}:F$gl").NumberFormat = 'dd/mm/yyyy'
  $ws.Range("G${g0}:G$gl").NumberFormat = '0'
  Grid "A${gHdr}:O$gl"
  foreach ($fr in @($g0..$gl)) { [void]$ws.Rows.Item($fr).AutoFit(); $hh = $ws.Rows.Item($fr).RowHeight; $ws.Rows.Item($fr).RowHeight = [double][Math]::Max(48, $hh + 8) }

  # ---- section 2 header
  Band $tBand $P.sec_tasks $NAVY 12
  $h = $P.task_hdr
  $cells = @()
  $cols = 'A','B','C','D','E','F','G','H','I','J'
  for ($k = 0; $k -lt 10; $k++) { $cells += , @("$($cols[$k])${tHdr1}:$($cols[$k])$tHdr2", $h[$k]) }
  $cells += , @("K${tHdr1}:O$tHdr1", $P.weeks_label)
  Hdr $tHdr1 $cells $tHdr2
  $wc = 'K','L','M','N','O'
  for ($k = 0; $k -lt 5; $k++) { SetV ($ws.Range("$($wc[$k])$tHdr2")) (D $P.weeks[$k]) }
  $ws.Range("K${tHdr2}:O$tHdr2").NumberFormat = 'dd/mm'
  $ws.Rows.Item($tHdr1).RowHeight = [double]18; $ws.Rows.Item($tHdr2).RowHeight = [double]18

  foreach ($row in $taskRows) {
    for ($k = 0; $k -lt 5; $k++) {
      $c = $wc[$k]
      $ws.Range("$c$row").Formula = "=IF(AND(`$F$row<=$c`$$tHdr2+6,`$G$row>=$c`$$tHdr2),1,`"`")"
    }
    $ws.Rows.Item($row).RowHeight = [double]33
  }
  foreach ($row in $checkRows) { $ws.Range("A${row}:J$row").Interior.Color = $YEL }
  $rg = $ws.Range("A${first}:J$last"); $rg.WrapText = $true
  $ws.Range("A${first}:A$last").HorizontalAlignment = -4108
  $ws.Range("F${first}:G$last").HorizontalAlignment = -4108
  $ws.Range("F${first}:G$last").NumberFormat = 'dd/mm'
  $ws.Range("B${first}:C$last").Font.Size = 10
  $ws.Range("J${first}:J$last").Font.Size = 10; $ws.Range("J${first}:J$last").Font.Italic = $true
  $ws.Range("K${first}:O$last").NumberFormat = ';;;'
  Grid "A${tHdr1}:O$last"
  foreach ($fr in $taskRows) { [void]$ws.Rows.Item($fr).AutoFit(); $hh = $ws.Rows.Item($fr).RowHeight; $ws.Rows.Item($fr).RowHeight = [double][Math]::Max(30, $hh + 8) }

  # status dropdown on task rows only
  foreach ($row in $taskRows) {
    $v = $ws.Range("I$row").Validation; $v.Delete(); $v.Add(3, 1, 1, $P.status_list) | Out-Null
  }

  # conditional formats (relative to the first task row)
  $ws.Activate(); $ws.Range("A$first").Select() | Out-Null
  $rI = $ws.Range("I${first}:I$last")
  $fc = $rI.FormatConditions.Add(2, $M, "=ISNUMBER(SEARCH(`"$done`",`$I$first))"); $fc.Interior.Color = $GREEN
  $fc = $rI.FormatConditions.Add(2, $M, "=LEFT(`$I$first,1)=`"$([char]0x26A0)`""); $fc.Interior.Color = $RED
  $fc = $rI.FormatConditions.Add(2, $M, "=AND(`$I$first<>`"`",ISERROR(SEARCH(`"$done`",`$I$first)),LEFT(`$I$first,1)<>`"$([char]0x26A0)`")"); $fc.Interior.Color = $BLUE
  $fc = $ws.Range("G${first}:G$last").FormatConditions.Add(2, $M, "=AND(`$G$first<>`"`",`$G$first<TODAY(),ISERROR(SEARCH(`"$done`",`$I$first)))")
  $fc.Font.Color = 255; $fc.Font.Bold = $true
  $fc = $ws.Range("K${first}:O$last").FormatConditions.Add(2, $M, "=K$first=1"); $fc.Interior.Color = $TEAL
  $ws.Range("A$g0").Select() | Out-Null
  $fc = $ws.Range("I${g0}:I$gl").FormatConditions.Add(2, $M, "=`$I$g0>0"); $fc.Interior.Color = $RED; $fc.Font.Bold = $true
  # today marker on the week header
  $ws.Range("K$tHdr2").Select() | Out-Null
  $fc = $ws.Range("K${tHdr2}:O$tHdr2").FormatConditions.Add(2, $M, "=AND(TODAY()>=K$tHdr2,TODAY()<=K$tHdr2+6)")
  $fc.Interior.Color = $NAVY; $fc.Font.Color = 16777215

  # ---- section 3: risks / decisions
  $r = $last + 2
  Band $r $P.sec_risks $NAVY 12; $r++
  $h = $P.risk_hdr; $rh = $r
  Hdr $r @(@("A$r", $h[0]), @("B${r}:D$r", $h[1]), @("E${r}:H$r", $h[2]), @("I$r", $h[3]), @("J${r}:O$r", $h[4])); $r++
  $n = 1; $r0 = $r
  foreach ($x in $P.risks) {
    SetV (Cl $ws $r 1) ($n)
    $ws.Range("B${r}:D$r").Merge(); SetV (Cl $ws $r 2) ($x[0])
    $ws.Range("E${r}:H$r").Merge(); SetV (Cl $ws $r 5) ($x[1])
    SetV (Cl $ws $r 9) ((D $x[2]))
    $ws.Range("J${r}:O$r").Merge(); SetV (Cl $ws $r 10) ($x[3])
    $ws.Rows.Item($r).RowHeight = [double]33
    $r++; $n++
  }
  $rg = $ws.Range("A${r0}:O$($r-1)"); $rg.WrapText = $true
  $ws.Range("A${r0}:A$($r-1)").HorizontalAlignment = -4108
  $ws.Range("I${r0}:I$($r-1)").HorizontalAlignment = -4108; $ws.Range("I${r0}:I$($r-1)").NumberFormat = 'dd/mm'
  Grid "A${rh}:O$($r-1)"

  # ---- section 4: rules
  $r++
  Band $r $P.sec_rules $NAVY 12; $r++
  foreach ($x in $P.rules) {
    $rg = $ws.Range("A${r}:O$r"); $rg.Merge(); $rg.Value2 = $x; $rg.WrapText = $true; $rg.IndentLevel = 1
    $ws.Rows.Item($r).RowHeight = [double]18; $r++
  }
  $end = $r - 1

  # print setup
  $ps = $ws.PageSetup
  $ps.PrintArea = "`$A`$1:`$O`$$end"; $ps.Orientation = 2; $ps.PaperSize = 9
  $ps.Zoom = $false; $ps.FitToPagesWide = 1; $ps.FitToPagesTall = $false
  $ps.PrintTitleRows = "`$${tHdr1}:`$$tHdr2"
  $ps.LeftMargin = $xl.CentimetersToPoints(1); $ps.RightMargin = $xl.CentimetersToPoints(1)
  $ps.TopMargin = $xl.CentimetersToPoints(1.2); $ps.BottomMargin = $xl.CentimetersToPoints(1.2)
  $ps.CenterFooter = '&P / &N'
  $ws.Range("A$first").Select() | Out-Null
  $xl.ActiveWindow.FreezePanes = $false
  $ws.Range('A1').Select() | Out-Null

  # ---- week sheet: decisions, link, remove the markdown draft table
  $w = $src
  $k = 0
  foreach ($x in $P.decisions_45) {
    $row = 13 + $k
    SetV (Cl $w $row 2) ($x[0])
    SetV (Cl $w $row 4) ((D $x[1])); (Cl $w $row 4).NumberFormat = 'dd/mm/yyyy'
    $k++
  }
  $hl = $w.Range('A18')
  $txt = [string]$hl.Value2
  $w.Hyperlinks.Add($hl, '', "'$($P.sheet)'!A1", '', "$txt   " + [char]0x2192 + "  chi ti" + [char]0x1EBF + "t: sheet '$($P.sheet)'") | Out-Null
  $hl.Font.Color = 16777215; $hl.Font.Underline = -4142; $hl.Font.Bold = $true; $hl.Font.Size = 12; $hl.Font.Name = 'Calibri'
  foreach ($lo in @($w.ListObjects)) { if ($lo.Name -eq 'Table1') { $lo.Delete() } }

  $wb.Worksheets.Item($P.sheet).Activate()
  $wb.Save()
  "OK rows first=$first last=$last end=$end checks=$($checkRows.Count) tasks=$($taskRows.Count)"
  $tmpPdf = [IO.Path]::ChangeExtension($Json, '.pdf')
  $wb.Worksheets.Item($P.sheet).ExportAsFixedFormat(0, $tmpPdf)
  $wb.Worksheets.Item($P.src_sheet).ExportAsFixedFormat(0, [IO.Path]::ChangeExtension($Json, '.w45.pdf'))
  $wb.Close($false)
} finally { $xl.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null }

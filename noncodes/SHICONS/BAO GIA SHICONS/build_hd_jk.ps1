# Build hợp đồng JK from the 2023 JK contract template + final quote 261109 (FN)
$ErrorActionPreference = 'Stop'
$src = "C:\Users\Khiem Nguyen\Documents\02 WORK MR LUAN\Bao gia\Hợp đồng bảo trì hệ thống HVAC và thẩm đinh kho JK.doc"
$dir = "C:\Users\Khiem Nguyen\Documents\projects\noncodes\SHICONS\BAO GIA SHICONS"
$base = "261110 HỢP ĐỒNG SHICONS - JK"
$docx = Join-Path $dir "$base.docx"
$pdf  = Join-Path $dir "$base.pdf"

$w = New-Object -ComObject Word.Application; $w.Visible = $false; $w.DisplayAlerts = 0
$d = $w.Documents.Open($src, $false, $true)
try {

$log = "$env:TEMP\hd_jk_log_{0:HHmmss}.txt" -f (Get-Date); Set-Content $log '' 
function L($m) { Add-Content $log ("{0:HH:mm:ss} {1}" -f (Get-Date), $m) }
function Rep([string]$a, [string]$b) {
  $ranges = @($d.Content)
  foreach ($h in $d.Sections.Item(1).Headers) { $ranges += $h.Range }
  foreach ($h in $d.Sections.Item(1).Footers) { $ranges += $h.Range }
  $hit = $false
  foreach ($r in $ranges) {
    $f = $r.Find; $f.ClearFormatting(); $f.Replacement.ClearFormatting()
    if ($f.Execute($a, $true, $false, $false, $false, $false, $true, 0, $false, $b, 2)) { $hit = $true }
  }
  L ("rep " + $hit + ": " + $a.Substring(0, [Math]::Min(30, $a.Length)))
  if (-not $hit) { Write-Warning "NOT FOUND: $a" }
}
function SetCell($t, $r, $c, [string]$txt) { $t.Cell($r, $c).Range.Text = $txt }

L 'start body'
# ---- Body text ----
Rep 'Số: .....151123.../HĐKT/VIWA-SHICONS' 'Số: ........../HĐKT/JK-SHICONS'
Rep '(Về việc thẩm định kho)' '(Về việc thẩm định kho GSP và bảo trì hệ thống HVAC)'
Rep 'Hôm nay, ngày ...15... tháng ...11... năm 2023' 'Hôm nay, ngày ...... tháng ...... năm 2026'
Rep 'thẩ định' 'thẩm định'
Rep 'thực hiện công việc thẩm định kho  cho Bên A' 'thực hiện công việc thẩm định kho GSP, xe vận chuyển, khu lấy mẫu và vệ sinh máy lạnh cho Bên A'
Rep 'Số 06, Đường số 8, Khu Dân Cư Ven Sông, P. Tân Phong, Quận 7, Hồ Chí  Minh, Việt Nam' '6 đường số 8, Khu dân cư Ven Sông, Phường Tân Hưng, Thành phố Hồ Chí Minh, Việt Nam'
Rep 'dự kiến: từ ngày   20 /11/2023  đến' 'dự kiến: 10 (mười) ngày, từ ngày ....../....../2026 đến'
Rep 'ngày  15/1/2024.' 'ngày ....../....../2026.'
Rep 'Hai mươi ba triệu, sáu trăm năm mươi hai nghìn đồng.' 'Mười chín triệu, bốn trăm bốn mươi nghìn đồng.'
Rep 'Tổng giá trị hợp đồng là 23,652,000 đồng' 'Tổng giá trị hợp đồng là 19,440,000 đồng (đã bao gồm thuế GTGT 8%)'
Rep '04 (năm) bản' '04 (bốn) bản'
Rep 'Hiep Binh Phuoc Ward, Thu Duc City, HCMC' 'Hiep Binh Ward, Ho Chi Minh City'

L 'body done'
# remove the irrelevant QĐ 5522 legal basis paragraph
foreach ($p in $d.Paragraphs) { if ($p.Range.Text -like '*5522/QĐ-UBND*') { $p.Range.Delete() | Out-Null; break } }

L 'qd removed'
# ---- Table 1: parties ----
$t1 = $d.Tables.Item(1)
SetCell $t1 3 1 'Điện thoại'
SetCell $t1 2 2 ': 6 đường số 8, Khu dân cư Ven Sông, Phường Tân Hưng, Thành phố Hồ Chí Minh, Việt Nam'
SetCell $t1 8 2 ': 78H, Đường số 9, Khu phố 5, Phường Hiệp Bình, Thành phố Hồ Chí Minh, Việt Nam'

L 't1 done'
# ---- Table 2: scope/price, mirrors quote 261109 FN ----
$t2 = $d.Tables.Item(2)
$boldG = $t2.Cell(3, 2).Range.Font.Bold
$boldS = $t2.Cell(4, 2).Range.Font.Bold
$t2.Cell(12, 1).Select(); $w.Selection.InsertRowsBelow(3)   # 12 item rows -> 15 (table has vertical merges, so Rows.Item fails)

L 'rows inserted'
$LB = [char]11
$h1 = 'Thẩm định nhiệt độ, độ ẩm kho bảo quản nguyên liệu làm thuốc GDP/GSP và tiến hành các tình huống giả định'
$d1 = $LB + '+ Khu bảo quản + phân phối + biệt trữ: diện tích 76.8 m2, cao 2.6 m' + $LB + '+ Khu chờ xuất nhập: diện tích 27.8 m2, cao 2.6 m' + $LB + '+ Nhiệt độ yêu cầu: 15 - 30 độ C; độ ẩm < 75%'
# row: No, Description, Unit, Qty, UnitPrice, Amount, kind(G=group,S=sub)
$rows = @(
  @('1', ($h1 + $d1), 'Hệ', '1', '10,000,000', '10,000,000', 'G'),
  @('-', 'Nghiên cứu thẩm định trong 7 ngày liên tục', '', '', '', '', 'S'),
  @('-', 'Chi phí thiết bị (24 sensor)', '', '', '', '', 'S'),
  @('-', 'Chi phí lập hồ sơ báo cáo, calib', '', '', '', '', 'S'),
  @('2', 'Thẩm định nhiệt độ, độ ẩm xe vận chuyển', 'Xe', '1', '2,000,000', '2,000,000', 'G'),
  @('-', 'Số lượng: 01 xe, 5 sensor; nhiệt độ yêu cầu 15 - 30 độ C', '', '', '', '', 'S'),
  @('-', 'Thời gian đo 1,5 ngày; hoàn thiện và gửi báo cáo thẩm định', '', '', '', '', 'S'),
  @('3', 'Chi phí vệ sinh, thẩm định AHU-LM', 'Hệ', '1', '4,000,000', '4,000,000', 'G'),
  @('-', 'Vệ sinh lọc thô AHU-LM', '', '', '', '', 'S'),
  @('-', 'Vệ sinh lọc thô tại phòng', '', '', '', '', 'S'),
  @('-', 'Căn chỉnh áp suất, lập hồ sơ thẩm định khu lấy mẫu các chỉ tiêu nhiệt độ, độ ẩm, chênh áp, tiểu phân, số lần trao đổi không khí', '', '', '', '', 'S'),
  @('4', 'Vệ sinh máy lạnh', '', '', '', '', 'G'),
  @('-', 'Vệ sinh máy lạnh FCU', 'Bộ', '2', '1,000,000', '2,000,000', 'S')
)
$r = 3
foreach ($row in $rows) {
  for ($c = 1; $c -le 6; $c++) { SetCell $t2 $r $c $row[$c - 1] }
  SetCell $t2 $r 7 ''
  $b = if ($row[6] -eq 'G') { $boldG } else { $boldS }
  for ($c = 1; $c -le 7; $c++) { $t2.Cell($r, $c).Range.Font.Bold = $b }
  $r++
}
L 'rows filled'
# item 1: only the heading line bold, details normal
$rg = $t2.Cell(3, 2).Range
$rg.Font.Bold = 0
$d.Range($rg.Start, $rg.Start + $h1.Length).Font.Bold = 1
# totals (rows 16 = B, 17..19 merged rows)
SetCell $t2 17 2 '18,000,000'
SetCell $t2 18 2 '1,440,000'
SetCell $t2 19 2 '19,440,000'

L 'totals set'
$d.Convert(); L 'converted'; $tmpDocx = "$env:TEMP\hd_jk_out.docx"; $tmpPdf = "$env:TEMP\hd_jk_out.pdf"; $d.SaveAs2($tmpDocx, 16)
L 'docx saved'
$d.ExportAsFixedFormat($tmpPdf, 17); L 'pdf exported'
"pages=" + $d.ComputeStatistics(2)
$d.SaveAs2("$env:TEMP\hd_jk_check.txt", 7, $false, "", $false, "", $false, $false, $false, $false, $false, 65001)
} finally { $d.Close($false); $w.Quit() }
Copy-Item $tmpDocx $docx -Force; Copy-Item $tmpPdf $pdf -Force; L 'copied'




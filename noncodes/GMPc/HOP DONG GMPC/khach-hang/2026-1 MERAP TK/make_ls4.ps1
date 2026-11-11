# ĐNTT "lần 4" (requests contract installment 5, FPF appraisal; installment 4 is on hold).
# Adds English for the yellow-highlighted Vietnamese and fixes English that does not match.
# Works on a copy in output\ via Word COM; the NAS original is never opened for writing.
$base = Split-Path -Parent $MyInvocation.MyCommand.Path
$out  = Join-Path $base 'output\Ho so thanh toan lan 5'
New-Item -ItemType Directory -Force $out | Out-Null
$src = Get-ChildItem (Join-Path $base 'raw\ls4') -Filter '3. *lan 4.doc' | Select-Object -First 1
$dst = Join-Path $out $src.Name
Copy-Item $src.FullName $dst -Force

# English inserted as a new italic paragraph after the Vietnamese paragraph that starts with the key.
$adds = @(
  @{ key = 'Đối với điều kiện công việc của lần thanh toán thứ 4'; en =
     'With regard to the conditions for the fourth payment under the Contract, GMPc Vietnam will continue to monitor and coordinate the completion of the related procedures, and will request that payment once all conditions under the Contract have been fully met.' },
  @{ key = 'Hiện tại, hạng mục công việc tương ứng'; en =
     'At present, the work item corresponding to the fourth payment under the Contract is awaiting the competent State authority''s review and resolution of the investment registration dossier (the authority responsible for reviewing and approving the amendment of the Investment Registration Certificate has changed to the Provincial People''s Committee, which has caused a delay against the schedule). As a result, the environmental dossier has not yet been completed and is not yet eligible for payment. However, this does not affect the progress or results of the Fire Prevention and Fighting (FPF) design item under the fifth payment. To date, the FPF design item has been appraised by the Fire Prevention, Firefighting and Rescue Police Department of Hung Yen Provincial Police.' }
)
$repl = @(
  # "lần 4" = 4th request sent (advance, 2, 3, this one); it asks for contract installment 5 (user, 2026-11-04)
  @('REQUEST FOR THE FIFTH PAYMENT', 'FOURTH PAYMENT REQUEST'),
  @('From: GMPc Joint Stock Company', 'From: GMPc Vietnam Joint Stock Company'),
  @('to GMPc Vietnam Joint Stock Company, as follows', 'to GMPc Vietnam Joint Stock Company, as follows:')
)

$w = New-Object -ComObject Word.Application; $w.Visible = $false; $w.DisplayAlerts = 0
try {
  $d = $w.Documents.Open($dst)
  foreach ($a in $adds) {
    $p = $d.Paragraphs | Where-Object { $_.Range.Text.TrimStart().StartsWith($a.key) } | Select-Object -First 1
    if (-not $p) { "MISS add: $($a.key)"; continue }
    $p.Range.Font.Italic = 0                     # Vietnamese body text is upright
    $p.Range.InsertParagraphAfter()
    $n = $p.Next().Range; $n.MoveEnd(1, -1) | Out-Null
    $n.Text = $a.en; $n.Font.Italic = 1; $n.Font.Bold = 0
    "OK   add: $($a.key)"
  }
  foreach ($e in $repl) {
    $r = $d.Content; $r.Find.ClearFormatting()
    if ($r.Find.Execute($e[0], $true)) { $r.Text = $e[1]; "OK   repl: $($e[0])" } else { "MISS repl: $($e[0])" }
  }
  # the added English pushes the empty spacer row (row 2, 3.7pt) of the signature table onto a blank page 3
  $sig = $d.Tables($d.Tables.Count)
  if ($sig.Rows.Count -ge 2 -and $sig.Rows(2).Range.Text.Trim([char]13, [char]7, ' ').Length -eq 0) { $sig.Rows(2).Delete(); 'OK   removed empty spacer row in signature table' }
  "pages: $($d.ComputeStatistics(2))"
  $d.Content.HighlightColorIndex = 0   # final copy: drop the review highlights
  $d.Save(); $d.Close(0)
} finally { $w.Quit() }

# Payment dossier #5 (Merap design contract): align the English text with the
# yellow-highlighted Vietnamese, wording taken from contract 1307/HDTV/MERAP-GMPc/2026.
# Works on copies in output\ via Word COM; the NAS originals are never opened for writing.
$base = Split-Path -Parent $MyInvocation.MyCommand.Path
$out  = Join-Path $base 'output\Ho so thanh toan lan 5'
New-Item -ItemType Directory -Force $out | Out-Null

$ffp = 'Fire Fighting & Prevention (FFP) – Lightning Protection and Smoke Extraction Design, Escape Route Review, and FFP Licensing'

$jobs = @(
  @{ src = '1. *.doc'; edits = @() },
  @{ src = '2. *.docx'; edits = @(
      @('Party B has completed the Basic Design, and both Parties have agreed upon it.',
        "Party B has completed the $ffp item, which has been appraised by the competent authority."),
      @('The Parties agree to accept the completed Basic Design work.',
        "The Parties agree to accept the completed $ffp work.")) },
  @{ src = '3. *.doc'; edits = @(
      @('REQUEST FOR THE THIRD PAYMENT', 'REQUEST FOR THE FIFTH PAYMENT'),
      @('The Basic Design has been approved by the Investor (Merap Group Corporation).',
        'The Fire Prevention and Fighting (FPF) design has been appraised by the Fire Prevention, Firefighting and Rescue Police Department of Hung Yen Provincial Police.'),
      @('make the third payment, equivalent to 20% of the service fee',
        'make the fifth payment, equivalent to 10% of the service fee'),
      @('Five hundred eighty-eight million six hundred thirty-four thousand six hundred sixty-eight Vietnamese Dong',
        'Two hundred ninety-four million three hundred seventeen thousand three hundred thirty-four Vietnamese Dong')) }
)

$w = New-Object -ComObject Word.Application; $w.Visible = $false; $w.DisplayAlerts = 0
try {
  foreach ($j in $jobs) {
    $f = Get-ChildItem (Join-Path $base 'raw\ls5') -Filter $j.src | Where-Object { $_.Extension -eq [IO.Path]::GetExtension($j.src) } | Select-Object -First 1
    $dst = Join-Path $out $f.Name
    Copy-Item $f.FullName $dst -Force
    $d = $w.Documents.Open($dst)
    foreach ($e in $j.edits) {
      $r = $d.Content
      $r.Find.ClearFormatting()
      if ($r.Find.Execute($e[0], $true)) { $r.Text = $e[1]; "OK   $($f.Name): $($e[0].Substring(0,[Math]::Min(40,$e[0].Length)))" }
      else { "MISS $($f.Name): $($e[0])" }
    }
    $d.Content.HighlightColorIndex = 0   # final copy: drop the review highlights
    $d.Save(); $d.Close(0)
  }
} finally { $w.Quit() }

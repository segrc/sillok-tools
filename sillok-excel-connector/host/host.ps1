# 크롬 확장 프로그램 -> 열려 있는 엑셀로 데이터를 넣어 주는 연결 프로그램
$ErrorActionPreference = 'Stop'
$stdin  = [Console]::OpenStandardInput()
$stdout = [Console]::OpenStandardOutput()

function Read-Exact($n) {
  $buf = New-Object byte[] $n; $got = 0
  while ($got -lt $n) { $r = $stdin.Read($buf, $got, $n - $got); if ($r -le 0) { return $null }; $got += $r }
  return ,$buf
}
function Send-Msg($obj) {
  $bytes = [Text.Encoding]::UTF8.GetBytes(($obj | ConvertTo-Json -Compress))
  $stdout.Write([BitConverter]::GetBytes([int32]$bytes.Length), 0, 4)
  $stdout.Write($bytes, 0, $bytes.Length); $stdout.Flush()
}

# 화면에 떠 있는 엑셀 창을 직접 찾아서(맨 앞 창부터) 연결.
# GetActiveObject 하나만 쓰면, 엑셀이 여러 개 실행 중이거나 막 켠 직후에 엉뚱한(빈) 엑셀을 잡는 문제가 있음.
$finderSrc = @'
using System; using System.Collections.Generic; using System.Runtime.InteropServices; using System.Text;
public static class SillokXlFinder {
  delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr p, EnumProc cb, IntPtr l);
  [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr h, StringBuilder sb, int n);
  [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
  [DllImport("oleacc.dll")] static extern int AccessibleObjectFromWindow(IntPtr h, uint id, ref Guid iid,
      [MarshalAs(UnmanagedType.IDispatch)] out object o);
  static string Cls(IntPtr h) { StringBuilder sb = new StringBuilder(256); GetClassName(h, sb, 256); return sb.ToString(); }
  public static object[] Windows() {
    List<IntPtr> tops = new List<IntPtr>();
    EnumWindows(delegate(IntPtr h, IntPtr l) { if (Cls(h) == "XLMAIN" && IsWindowVisible(h)) tops.Add(h); return true; }, IntPtr.Zero);
    List<object> res = new List<object>();
    Guid iid = new Guid("00020400-0000-0000-C000-000000000046");
    foreach (IntPtr t in tops) {
      IntPtr found = IntPtr.Zero;
      EnumChildWindows(t, delegate(IntPtr h, IntPtr l) { if (Cls(h) == "EXCEL7") { found = h; return false; } return true; }, IntPtr.Zero);
      if (found == IntPtr.Zero) continue;
      object o;
      if (AccessibleObjectFromWindow(found, 0xFFFFFFF0, ref iid, out o) == 0 && o != null) res.Add(o);
    }
    return res.ToArray();
  }
}
'@
if (-not ('SillokXlFinder' -as [type])) {
  $dll = Join-Path $PSScriptRoot 'SillokXlFinder1.dll'
  try {
    if (-not (Test-Path $dll)) { Add-Type -TypeDefinition $finderSrc -OutputAssembly $dll -OutputType Library }
    Add-Type -Path $dll
  } catch { Add-Type -TypeDefinition $finderSrc }
}

function Find-Sheet {
  $wins = @(); try { $wins = @([SillokXlFinder]::Windows()) } catch {}
  foreach ($w in $wins) {
    try { $s = $w.ActiveSheet; if ($null -ne $s) { return @{ ws = $s; n = $wins.Count } } } catch {}
  }
  try {
    $xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
    $s = $xl.ActiveSheet; if ($null -ne $s) { return @{ ws = $s; n = $wins.Count } }
  } catch {}
  return @{ ws = $null; n = $wins.Count }
}

# 엑셀이 잠깐 바쁠 때(저장·계산·화면 갱신 중)는 "그 동작 하나만" 조금 기다렸다 다시 시도.
# (예전에는 행 전체를 처음부터 다시 써서, 일부만 쓴 행 + 새 행 = 중복 행·링크 누락이 생겼음)
function Invoke-Com([scriptblock]$sb) {
  for ($i = 0; ; $i++) {
    try { return (& $sb) }
    catch {
      if ("$($_.Exception)" -match '80010001|8001010A' -and $i -lt 40) { Start-Sleep -Milliseconds 250; continue }
      throw
    }
  }
}

function Write-Row($msg) {
  $found = Find-Sheet; $ws = $found.ws
  if ($null -eq $ws) {
    if ($found.n -gt 0) { return @{ ok = $false; error = '엑셀 시트를 찾지 못했어요 (보호된 보기면 "편집 사용" 클릭)' } }
    return @{ ok = $false; error = '열려 있는 엑셀을 찾지 못했어요' }
  }
  # 지정한 열들 중 가장 아래 데이터의 바로 다음 행 (한 번만 계산)
  $last = 0
  foreach ($f in $msg.fields) {
    $r = Invoke-Com {
      $c = $ws.Range("$($f.col)1").Column
      $x = $ws.Cells.Item($ws.Rows.Count, $c).End(-4162).Row   # xlUp
      if ($x -eq 1 -and [string]::IsNullOrEmpty($ws.Cells.Item(1, $c).Formula)) { $x = 0 }
      $x
    }
    if ($r -gt $last) { $last = $r }
  }
  $row = $last + 1

  $script:writtenRow = $row
  foreach ($f in $msg.fields) {
    $v = [string]$f.value
    if ($v.Length -gt 32000) { $v = $v.Substring(0, 32000) }   # 엑셀 한 칸 글자 수 제한
    if ($f.link) {
      Invoke-Com {
        $cell = $ws.Range("$($f.col)$row")
        [void]$ws.Hyperlinks.Add($cell, $v)
      } | Out-Null
    } else {
      if ($v -match '^[=+]') { $v = "'" + $v }                 # 수식으로 바뀌지 않게
      Invoke-Com { $ws.Range("$($f.col)$row").Value2 = $v } | Out-Null
      if ($f.wrap -and $v.Contains("`n")) { Invoke-Com { $ws.Range("$($f.col)$row").WrapText = $true } | Out-Null }
    }
  }

  return @{ ok = $true; row = $row; sheet = $ws.Name; book = $ws.Parent.Name }
}

try {
  $lenBytes = Read-Exact 4; if ($null -eq $lenBytes) { exit }
  $body = Read-Exact ([BitConverter]::ToInt32($lenBytes, 0))
  $msg = [Text.Encoding]::UTF8.GetString($body) | ConvertFrom-Json
  if ($msg.ping) { Send-Msg @{ ok = $true; ping = $true }; exit }   # 설치 확인용

  $script:writtenRow = $null
  try { $result = Write-Row $msg }
  catch {
    $e = "$($_.Exception)"
    if ($e -match '800AC472') { $err = '엑셀에서 셀 편집 중이에요 (Esc 후 다시)' }
    elseif ($e -match '80010001|8001010A') { $err = '엑셀이 바빠요 (대화상자가 열려 있는지 확인)' }
    else { $err = "엑셀 오류: $($_.Exception.Message)" }
    if ($script:writtenRow) { $err = "$($script:writtenRow)행 일부만 입력됨 – $err" }
    $result = @{ ok = $false; error = $err }
  }
  Send-Msg $result
}
catch {
  try { Send-Msg @{ ok = $false; error = "연결 프로그램 오류: $($_.Exception.Message)" } } catch {}
}

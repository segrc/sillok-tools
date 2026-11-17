# 실록 도우미 - 엑셀 연결 프로그램 등록 (현재 사용자만, 관리자 권한 필요 없음)

# 이 연결 프로그램을 쓸 수 있는 확장 프로그램 ID
#  - 크롬 웹스토어판 ID가 정해지면 아래 'STORE_ID' 자리에 넣으세요 (영문 소문자 32자)
$extensionIds = @(
  'STORE_ID',                          # 크롬 웹스토어판
  'iabhcmaekhaknmalhibifpfhjdlobolm'     # 직접 로드한 개발용
)

$hostDir  = Join-Path $PSScriptRoot 'host'
$manifest = Join-Path $hostDir 'com.sillok.tools.json'
$origins  = @($extensionIds | Where-Object { $_ -match '^[a-p]{32}$' } | ForEach-Object { "chrome-extension://$_/" })
$json = [ordered]@{
  name = 'com.sillok.tools'
  description = 'Sillok helper - Excel connector'
  path = (Join-Path $hostDir 'host.bat')
  type = 'stdio'
  allowed_origins = $origins
} | ConvertTo-Json
[IO.File]::WriteAllText($manifest, $json, (New-Object Text.UTF8Encoding $false))
foreach ($k in 'HKCU:\Software\Google\Chrome\NativeMessagingHosts\com.sillok.tools',
               'HKCU:\Software\Microsoft\Edge\NativeMessagingHosts\com.sillok.tools') {
  New-Item -Path $k -Force | Out-Null
  Set-ItemProperty -Path $k -Name '(default)' -Value $manifest
}
Write-Host ''
Write-Host '설치 완료!' -ForegroundColor Green
Write-Host '실록 페이지를 새로고침하면 엑셀 버튼이 나타나요.'
Write-Host '(버튼이 안 보이면 실록 도우미 설정 창의 "다시 확인"을 누르세요)'
Write-Host '(이 폴더를 옮기거나 지우면 엑셀 입력이 안 돼요. 옮겼다면 install.bat 을 다시 실행하세요)'

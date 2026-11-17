Remove-Item 'HKCU:\Software\Google\Chrome\NativeMessagingHosts\com.sillok.tools' -ErrorAction SilentlyContinue
Remove-Item 'HKCU:\Software\Microsoft\Edge\NativeMessagingHosts\com.sillok.tools' -ErrorAction SilentlyContinue
Write-Host '연결 프로그램 등록을 해제했어요.'

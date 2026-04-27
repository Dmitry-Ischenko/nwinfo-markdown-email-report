Формирует отчет с помощью NWinfo, парсит его и сохраняет в формате MarkDown файла, и отправляет его на почту.

```powershell
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; $params = @('-nwinfoUrl', 'https://github.com/a1ive/nwinfo/releases/download/v1.6.2/NWinfo.zip', '-smtpServer', 'smtp.yandex.ru', '-smtpPort', '587', '-from', 'from@yandex.ru', '-to', 'to@yandex.ru', '-smtpUser', 'from@yyandex.ru', '-smtpPassword', '!!!Пароль!!!'); iex "& { $([System.Text.Encoding]::UTF8.GetString((Invoke-WebRequest -Uri 'https://github.com/Dmitry-Ischenko/nwinfo-markdown-email-report/releases/download/New-release/nwinfo_md-report-email.ps1' -UseBasicParsing).Content)) } $params"
```

<#
.SYNOPSIS
    Автоматическая генерация отчета о системе с помощью NWinfo и отправка по email
.PARAMETER nwinfoUrl
    URL для скачивания NWinfo.zip (например, https://github.com/a1ive/nwinfo/releases/download/v1.6.2/NWinfo.zip)
.PARAMETER smtpServer
    SMTP сервер для отправки email
.PARAMETER smtpPort
    Порт SMTP сервера
.PARAMETER from
    Email отправителя
.PARAMETER to
    Email получателя
.PARAMETER smtpUser
    Логин для SMTP аутентификации
.PARAMETER smtpPassword
    Пароль для SMTP аутентификации
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$nwinfoUrl,
    
    [Parameter(Mandatory=$true)]
    [string]$smtpServer,
    
    [Parameter(Mandatory=$true)]
    [int]$smtpPort,
    
    [Parameter(Mandatory=$true)]
    [string]$from,
    
    [Parameter(Mandatory=$true)]
    [string]$to,
    
    [Parameter(Mandatory=$true)]
    [string]$smtpUser,
    
    [Parameter(Mandatory=$true)]
    [string]$smtpPassword
)

# Глобальные переменные
$tempFolder = $env:TEMP
$nwinfoZip = Join-Path $tempFolder "NWinfo.zip"
$nwinfoExtractPath = Join-Path $tempFolder "NWinfo"
$nwinfoDir = ""  # Будет установлена после извлечения

try {
    Write-Host "=== Начало работы скрипта ===" -ForegroundColor Cyan
    
    # Шаг 1: Скачивание NWinfo.zip
    Write-Host "`nШаг 1: Скачивание NWinfo..." -ForegroundColor Yellow
    Write-Host "URL: $nwinfoUrl" -ForegroundColor Gray
    
    if (Test-Path $nwinfoZip) {
        Remove-Item $nwinfoZip -Force
        Write-Host "Удален старый архив" -ForegroundColor Gray
    }
    
    Invoke-WebRequest -Uri $nwinfoUrl -OutFile $nwinfoZip -UseBasicParsing
    Write-Host "? Архив успешно скачан: $nwinfoZip" -ForegroundColor Green
    
    # Шаг 2: Извлечение архива
    Write-Host "`nШаг 2: Извлечение архива..." -ForegroundColor Yellow
    
    if (Test-Path $nwinfoExtractPath) {
        Remove-Item $nwinfoExtractPath -Recurse -Force
        Write-Host "Удалена старая папка NWinfo" -ForegroundColor Gray
    }
    
    Expand-Archive -Path $nwinfoZip -DestinationPath $nwinfoExtractPath -Force
    Write-Host "? Архив успешно извлечен в: $nwinfoExtractPath" -ForegroundColor Green
    
    # Установка переменной директории NWinfo
    $nwinfoDir = $nwinfoExtractPath
    Write-Host "? Рабочая директория NWinfo: $nwinfoDir" -ForegroundColor Green
    
    # Шаг 3: Определение исполняемого файла
    Write-Host "`nШаг 3: Определение версии NWinfo..." -ForegroundColor Yellow
    
    $programPath = ""
    
    if ([Environment]::Is64BitOperatingSystem) {
        $possiblePath = Join-Path $nwinfoDir "nwinfo.exe"
        if (Test-Path $possiblePath) {
            $programPath = $possiblePath
            Write-Host "? Обнаружена 64-bit версия: nwinfo.exe" -ForegroundColor Green
        }
        else {
            throw "nwinfo.exe не найден в директории: $nwinfoDir"
        }
    }
    else {
        $possiblePathx86 = Join-Path $nwinfoDir "nwinfox86.exe"
        $possiblePath64 = Join-Path $nwinfoDir "nwinfo.exe"
        
        if (Test-Path $possiblePathx86) {
            $programPath = $possiblePathx86
            Write-Host "? Обнаружена 32-bit версия: nwinfox86.exe" -ForegroundColor Green
        }
        elseif (Test-Path $possiblePath64) {
            $programPath = $possiblePath64
            Write-Host "? Обнаружена 64-bit версия (на 32-bit ОС): nwinfo.exe" -ForegroundColor Green
        }
        else {
            throw "nwinfo.exe или nwinfox86.exe не найдены в директории: $nwinfoDir"
        }
    }
    
    Write-Host "Путь к исполняемому файлу: $programPath" -ForegroundColor Gray
    
    # Шаг 4: Генерация отчета (пример - замените на вашу логику)
    Write-Host "`nШаг 4: Генерация отчета..." -ForegroundColor Yellow
    
    $programArgs = @(
        "--format=json",
        "--cp=utf8",
        "--human",
        "--sys",
        "--cpu",
        "--gpu",
		"--board",
        "--display",
        "--net",
        "--disk=phys",
        "--spd" 
    )
    
    Write-Host "Executing nwinfo..."
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = (Resolve-Path $programPath).Path
    $psi.Arguments = $programArgs -join " "
    $psi.RedirectStandardOutput = $true
    $psi.UseShellExecute = $false
    $psi.StandardOutputEncoding = [System.Text.Encoding]::UTF8

    $proc = [System.Diagnostics.Process]::Start($psi)
    $processOutput = $proc.StandardOutput.ReadToEnd()
    $proc.WaitForExit()

    if ($proc.HasExited -and $proc.ExitCode -ne 0) {
        throw "nwinfo exited with code $($proc.ExitCode)."
    }

    Write-Host "Processing nwinfo output..."
    $utf8Json = $processOutput

    Write-Host "Parsing JSON..."
    $parsedJson = $utf8Json | ConvertFrom-Json -ErrorAction Stop

    # Создание MD-файла
    Write-Host "Generating markdown report..."
    $mdContent = @()
    
    $mdContent += "| Параметр | Значение |"
    $mdContent += "| -------- | -------- |"
    
    # Дата отчета
    $mdContent += "| Дата отчета | $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss') |"
    
    # Имя компьютера
    $computerName = $parsedJson.System.'Computer Name'
    $mdContent += "| Имя компьютера | $computerName |"
    
    # ОС
    $os = $parsedJson.System.OS
    $mdContent += "| Операционная система | $os |"
    
    # Дата установки
    $installDate = $parsedJson.System.'Install Date'
    $mdContent += "| Дата установки | $installDate |"
    $mdContent += "| | |"
    
    # Материнская плата
    $boardName = $parsedJson.Mainboard.'Board Name'
    $mdContent += "| Материнская плата | $boardName |"
    
    # BIOS
    $biosVersion = $parsedJson.Mainboard.'BIOS Version'
    $biosDate = $parsedJson.Mainboard.'BIOS Date'
    $mdContent += "| Версия биос | $biosVersion ($biosDate) |"
    
    # Процессор
    $cpuBrand = $parsedJson.CPUID.'CPU0-G'.Brand
    $mdContent += "| Процессор | $cpuBrand |"
    
    # Температура процессора
    $cpuTemp = $parsedJson.CPUID.'CPU0-G'.'Temperature (C)'
    $mdContent += "| Температура процессора | $cpuTemp |"
    $mdContent += "| | |"
    
    # Память
    $totalMemory = $parsedJson.System.'Physical Memory'.Total
    $mdContent += "| Total Physical Memory | $totalMemory |"
    
    # SPD (RAM модули)
    if ($parsedJson.SPD) {
        for ($i = 0; $i -lt $parsedJson.SPD.Count; $i++) {
            $spd = $parsedJson.SPD[$i]
            $ramInfo = "$($spd.Manufacturer) $($spd.'Part Number') $($spd.Capacity) ($($spd.'Memory Type'))"
            $mdContent += "| Оперативная память bank $($spd.ID) | $ramInfo |"
        }
    }
    $mdContent += "| | |"
    
    # Мониторы
	if ($parsedJson.Display) {
		for ($i = 0; $i -lt $parsedJson.Display.Count; $i++) {
			$display = $parsedJson.Display[$i]
			# Проверяем наличие HWID - если его нет, пропускаем этот элемент
			if ($display.HWID) {
				$displayInfo = "$($display.Manufacturer) $($display.'Display Name')"
				$mdContent += "| Монитор [$i] | $displayInfo |"
				$mdContent += "| Разрешения экрана [$i] | $($display.'Max Resolution') |"
			}
		}
	}
	$mdContent += "| | |"
    
    # Диски
    if ($parsedJson.Disks) {
        for ($i = 0; $i -lt $parsedJson.Disks.Count; $i++) {
            $disk = $parsedJson.Disks[$i]
            $diskInfo = "$($disk.'HW Name') ($($disk.Type)) (Size: $($disk.Size)) (SSD: $($disk.SSD)) ($($disk.'Health Status'))"
            $mdContent += "| Марка/модель диска [$i] | $diskInfo |"
            
            # Логические диски
            if ($disk.Volumes) {
                $volumesInfo = @()
                foreach ($volume in $disk.Volumes) {
                    $driveLetter = $volume.'Volume Path Names'.'Drive Letter'
					    if ([string]::IsNullOrWhiteSpace($driveLetter)) {
							$driveLetter = "no label"
						}
                    $volumesInfo += "$driveLetter ($($volume.'Total Space')) ($($volume.Usage))"
                }
                $mdContent += "| Логический диски | $($volumesInfo -join '<br>') |"
            }
        }
    }
    $mdContent += "| | |"
    
    # Сетевые устройства (только активные)
    if ($parsedJson.Network) {
        foreach ($net in $parsedJson.Network) {
			if ($net.Status -eq "Active") {
                $mdContent += "| Сетевое устройство | $($net.Description) |"
                $mdContent += "| DHCP | $($net.'DHCP Enabled') |"
                
                # Поиск IPv4 адреса
                $ipv4 = ""
                if ($net.Unicasts) {
                    foreach ($unicast in $net.Unicasts) {
                        if ($unicast.IPv4) {
                            $ipv4 = $unicast.IPv4
                            break
                        }
                    }
                }
                $mdContent += "| ip Адресс | $ipv4 |"
                $mdContent += "| mac адрес | $($net.'MAC Address') |"
                $mdContent += "| Тип | $($net.Type) |"
                $mdContent += "| | |"
            }
        }
    }
    Write-Host "Generation list installed programms"
	# Установленные программы
	$mdContent += "`n## Установленные программы`n"
	$mdContent += "| Наименование | Дата установки |"
	$mdContent += "| ------------ | -------------- |"

	# Получаем список установленных программ из реестра (64-bit и 32-bit)
	$programs = @()
	$registryPaths = @(
		"HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
		"HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
	)

	foreach ($path in $registryPaths) {
		$programs += Get-ItemProperty $path -ErrorAction SilentlyContinue | 
			Where-Object { $_.DisplayName -and $_.DisplayName -notlike "Update for*" } |
			Select-Object DisplayName, InstallDate
	}

	# Удаляем дубликаты и сортируем по дате установки (от поздней к ранней)
	$programs = $programs | Sort-Object DisplayName -Unique | Sort-Object {
		if ($_.InstallDate) {
			try {
				[DateTime]::ParseExact($_.InstallDate, "yyyyMMdd", $null)
			} catch {
				[DateTime]::MinValue  # Программы без даты в конец списка
			}
		} else {
			[DateTime]::MinValue
		}
	} -Descending

	foreach ($program in $programs) {
		$installDate = if ($program.InstallDate) {
			try {
				[DateTime]::ParseExact($program.InstallDate, "yyyyMMdd", $null).ToString("dd.MM.yyyy")
			} catch {
				$program.InstallDate
			}
		} else {
			"Не указана"
		}
		$mdContent += "| $($program.DisplayName) | $installDate |"
	}

	$mdContent += "`n"
	Write-Host "Generation list service"
	# Службы и их статус
	$mdContent += "`n## Службы и их статус`n"
	$mdContent += "| Наименование службы | Статус | Способ запуска |"
	$mdContent += "| ------------------- | ------ | -------------- |"

	# Получаем все службы одним запросом WMI
	$wmiServices = Get-WmiObject -Class Win32_Service | Select-Object Name, StartMode

	# Создаем хэш-таблицу для быстрого поиска
	$startModeHash = @{}
	foreach ($wmiService in $wmiServices) {
		$startModeHash[$wmiService.Name] = $wmiService.StartMode
	}

	# Сортируем: сначала запущенные (Running), потом остановленные
	$services = Get-Service | Sort-Object @{Expression = {$_.Status -ne 'Running'}}, DisplayName

	foreach ($service in $services) {
		$startType = $startModeHash[$service.Name]
		$status = $service.Status
		$mdContent += "| $($service.DisplayName) | $status | $startType |"
	}

	$mdContent += "`n"
	Write-Host "Generation list rinters"
	# Принтеры
	$mdContent += "`n## Принтеры`n"
	$mdContent += "| Имя принтера | Порт принтера | IP адрес (или имя) порта принтера |"
	$mdContent += "| ------------ | ------------- | --------------------------------- |"

	# Получаем список принтеров
	$printers = Get-Printer | Sort-Object Name

	# Получаем все порты принтеров одним запросом для ускорения
	$allPorts = Get-PrinterPort

	foreach ($printer in $printers) {
		$portName = $printer.PortName
		
		# Находим соответствующий порт
		$port = $allPorts | Where-Object { $_.Name -eq $portName } | Select-Object -First 1
		
		# Определяем IP адрес или имя порта
		$portAddress = if ($port) {
			if ($port.PrinterHostAddress) {
				$port.PrinterHostAddress
			} elseif ($port.HostName) {
				$port.HostName
			} elseif ($port.Description) {
				$port.Description
			} else {
				"Локальный порт"
			}
		} else {
			"Не определен"
		}
		
		$mdContent += "| $($printer.Name) | $portName | $portAddress |"
	}

	$mdContent += "`n"
	
	
    # Сохранение в файл
    $outputFile = Join-Path $nwinfoDir "$computerName`_$(Get-Date -Format 'dd-MM-yyyy').md"
    $mdContent | Out-File -FilePath $outputFile -Encoding UTF8
    
    Write-Host "Report saved to: $outputFile"
    
	
	if ($smtpServer -and $smtpPort -and $from -and $to) {
    try {
        $subject = "Отчет: $computerName - $(Get-Date -Format 'dd-MM-yyyy')"
        $body = "Автоматический отчет о конфигурации компьютера $computerName"
        
        # Если указаны учетные данные в параметрах
        if ($smtpUser -and $smtpPassword) {
            $securePassword = ConvertTo-SecureString $smtpPassword -AsPlainText -Force
            $credential = New-Object System.Management.Automation.PSCredential ($smtpUser, $securePassword)
        } else {
            # Иначе запрашиваем интерактивно
            $credential = Get-Credential -Message "Введите учетные данные для SMTP"
        }
        Write-Host "Send repport to email"
        Send-MailMessage -From $from `
                         -To $to `
                         -Subject $subject `
                         -Body $body `
                         -SmtpServer $smtpServer `
                         -Port $smtpPort `
                         -UseSsl `
                         -Credential $credential `
                         -Attachments $outputFile `
                         -Encoding UTF8
        
        Write-Host "Отчет успешно отправлен на $to" -ForegroundColor Green
    } catch {
        Write-Host "Ошибка отправки email: $_" -ForegroundColor Red
    }
}
	
	
    # Шаг 6: Очистка временных файлов
    Write-Host "`nШаг 6: Очистка временных файлов..." -ForegroundColor Yellow
    
    if (Test-Path $nwinfoZip) {
        Remove-Item $nwinfoZip -Force
        Write-Host "? Удален архив NWinfo.zip" -ForegroundColor Gray
    }
    
    if (Test-Path $nwinfoExtractPath) {
        Remove-Item $nwinfoExtractPath -Recurse -Force
        Write-Host "? Удалена папка NWinfo" -ForegroundColor Gray
    }
    
    #if (Test-Path $reportPath) {
    #    Remove-Item $reportPath -Force
    #    Write-Host "? Удален временный отчет" -ForegroundColor Gray
    #}
    
    Write-Host "`n=== Скрипт успешно завершен ===" -ForegroundColor Green
    
}
catch {
    Write-Host "`n!!! ОШИБКА !!!" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor DarkRed
    
    # Очистка при ошибке
    if (Test-Path $nwinfoZip) { Remove-Item $nwinfoZip -Force -ErrorAction SilentlyContinue }
    if (Test-Path $nwinfoExtractPath) { Remove-Item $nwinfoExtractPath -Recurse -Force -ErrorAction SilentlyContinue }
    
    exit 1
}

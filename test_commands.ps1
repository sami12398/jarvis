# JARVIS Test Suite
$baseUrl = "http://localhost:5000"

function Test-Command($name, $command) {
    Write-Host "`n[Test] $name" -ForegroundColor Cyan
    Write-Host "Command: $command" -ForegroundColor DarkGray
    
    try {
        $response = Invoke-RestMethod -Uri "$baseUrl/api/command" -Method POST -ContentType "application/json" -Body "{`"command`": `"$command`"}"
        
        if ($response.success) {
            Write-Host "Result: $($response.message)" -ForegroundColor Green
        } else {
            Write-Host "Result: $($response.message)" -ForegroundColor Red
        }
        return $response.success
    } catch {
        Write-Host "Error: $_" -ForegroundColor Red
        return $false
    }
}

Write-Host "JARVIS Command Test Suite" -ForegroundColor Magenta
Write-Host "=========================" -ForegroundColor Magenta

# Check server first
try {
    $status = Invoke-RestMethod -Uri "$baseUrl/api/status" -Method GET
    Write-Host "Server Status: $($status.status)`n" -ForegroundColor Green
} catch {
    Write-Host "ERROR: Cannot connect to server at $baseUrl" -ForegroundColor Red
    Write-Host "Make sure jarvis_server.py is running"
    exit
}

$passed = 0
$total = 0

# Basic Tests
$total++; if (Test-Command "Greeting" "hello") { $passed++ }
$total++; if (Test-Command "Time" "what time is it") { $passed++ }
$total++; if (Test-Command "IP Address" "ip address") { $passed++ }
$total++; if (Test-Command "System Info" "system info") { $passed++ }

# Application Tests
Write-Host "`n[Application Tests]" -ForegroundColor Yellow
$total++; if (Test-Command "Open Notepad" "open notepad") { $passed++ }
Start-Sleep 2
$total++; if (Test-Command "Close Notepad" "close notepad") { $passed++ }

# Window Management Tests
Write-Host "`n[Window Tests]" -ForegroundColor Yellow
$total++; if (Test-Command "Minimize All" "minimize all") { $passed++ }
Start-Sleep 1

# Clipboard Tests
Write-Host "`n[Clipboard Tests]" -ForegroundColor Yellow
$total++; if (Test-Command "Copy Text" "copy Hello JARVIS") { $passed++ }

# Summary
Write-Host "`n=========================" -ForegroundColor Magenta
Write-Host "Tests Passed: $passed / $total" -ForegroundColor $(if($passed -eq $total){"Green"}else{"Yellow"})
Write-Host "=========================" -ForegroundColor Magenta

if ($passed -eq $total) {
    Write-Host "All systems operational!" -ForegroundColor Green
} else {
    Write-Host "Some tests failed. Check the server logs." -ForegroundColor Yellow
}
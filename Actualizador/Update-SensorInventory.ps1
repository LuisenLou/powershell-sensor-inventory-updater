# --- REFERENCIAS PARA USAR MODALES ---
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

# --- IMPORTAR CONFIGURACIÓN ---
# La configuración de $carpetaAbbott, $maestroCSV, etc., se cargará desde un archivo externo.
. .\Config-Paths.ps1


# --- MAPEO DE MESES A COLUMNAS ---
$mesColumnas = @{
    "enero" = 4; "febrero" = 5; "marzo" = 6; "abril" = 7; "mayo" = 8;
    "junio" = 9; "julio" = 10; "agosto" = 11; "septiembre" = 12;
    "octubre" = 13; "noviembre" = 14; "diciembre" = 15
}

# --- FUNCIONES AUXILIARES ---
function ConfirmProcess($mensaje, $titulo) {
    $r = [System.Windows.Forms.MessageBox]::Show($mensaje, $titulo,
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question)
    return ($r -eq [System.Windows.Forms.DialogResult]::Yes)
}

function OpenSecureExcel {
    param($ruta)
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $workbook = $excel.Workbooks.Open($ruta, $null, $false)
        return @{Excel=$excel; Workbook=$workbook}
    } catch {
        if ($excel) { $excel.Quit() | Out-Null }
        throw $_
    }
}

# --- FUNCIÓN REUTILIZABLE PARA ACTUALIZAR ---
function ProcessManufacturer {
    param (
        [string]$nombre,
        [string]$carpeta,
        [string]$maestroCSV,
        [string]$excelPath
    )

    [System.Windows.Forms.MessageBox]::Show(
        "Iniciando actualización de $nombre...",
        "Inicio $nombre",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null

    if (-not (Test-Path $maestroCSV)) {
        [System.Windows.Forms.MessageBox]::Show("No se encontró el CSV maestro para $nombre.`nRuta: $maestroCSV",
            "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        return
    }
    $maestro = Import-Csv $maestroCSV

    if (-not (Test-Path $excelPath)) {
        [System.Windows.Forms.MessageBox]::Show("No se encontró el Excel para $nombre.`nRuta: $excelPath",
            "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        return
    }

    try {
        $exObj = OpenSecureExcel -ruta $excelPath
        $excel = $exObj.Excel
        $workbook = $exObj.Workbook
        $sheet = $workbook.Sheets.Item(1)
    } catch {
        [System.Windows.Forms.MessageBox]::Show("No se puede abrir el Excel `$excelPath`. Puede estar en uso. Ciérralo e inténtalo de nuevo.`nError: $($_.Exception.Message)",
            "Error al abrir Excel", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        return
    }

    $totalFabricante = 0
    $cipaGlobal = @{}

    $txtFiles = Get-ChildItem $carpeta -Filter *.txt -File -ErrorAction SilentlyContinue
    if (-not $txtFiles) {
        Write-Host "No hay archivos .txt en $carpeta"
    }

    foreach ($txt in $txtFiles) {
        $filePath = $txt.FullName
        $contenido = Get-Content $filePath -ErrorAction Stop

        $mesActual = $null
        foreach ($linea in $contenido) {
            if ($linea -match "Fecha de envío.*\b(Enero|Febrero|Marzo|Abril|Mayo|Junio|Julio|Agosto|Septiembre|Octubre|Noviembre|Diciembre)\b") {
                $mesActual = $matches[1].ToLower()
                break
            }
        }
        if (-not $mesActual) {
            [System.Windows.Forms.MessageBox]::Show(
                "No se pudo detectar el mes en el archivo $filePath.`nSe omitirá este archivo.",
                "Error de detección", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning
            ) | Out-Null
            continue
        }

        $seguir = ConfirmProcess "Se ha detectado el mes '$mesActual' en `n$filePath`n¿Deseas continuar con la actualización de este archivo?" "Confirmación de mes"
        if (-not $seguir) { Write-Host "Archivo omitido por el usuario: $filePath"; continue }

        $colMes = $mesColumnas[$mesActual]
        if (-not $colMes) {
            [System.Windows.Forms.MessageBox]::Show("Mes '$mesActual' no mapeado a columna. Omitiendo archivo $filePath",
                "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
            continue
        }

        $items = @()
        foreach ($linea in $contenido) {
            $ln = $linea -replace '➔','>'
            if ($ln -match '.*?([0-9\s''\.\-A-Za-z]+)>(\d+)\b') {
                $rawText = $matches[1].Trim()
                $digits = ($rawText -replace '[^\d]', '')
                $units = [int]$matches[2]
                $items += [pscustomobject]@{Raw=$rawText;Digits=$digits;Units=$units}
            }
        }

        if ($items.Count -eq 0) {
            Write-Host "No se encontraron líneas válidas en $filePath"
            continue
        }

        $cipaDict = @{}
        foreach ($it in $items) {
            if ($it.Digits.Length -eq 10) {
                $key = $it.Digits
            } else {
                $key = $it.Raw
            }
            if ($cipaDict.ContainsKey($key)) { $cipaDict[$key] += $it.Units } else { $cipaDict[$key] = $it.Units }
        }

        $descartados = $cipaDict.Keys | Where-Object { ($_ -as [string]).Length -ne 10 -and -not ($_ -match '^\d{10}$') }

        foreach ($rawKey in $descartados) {
            $units = $cipaDict[$rawKey]
            $promptMsg = "Se detectó entrada no estándar: `n'$rawKey' (unidades: $units)`nIntroduce el CIPA correcto (10 dígitos) o deja vacío para dejarlo como NO ASIGNADO (se insertará la cadena leída en columna CIPA)."
            $respuesta = [Microsoft.VisualBasic.Interaction]::InputBox($promptMsg, "Corrección CIPA ($nombre)", "")
            if ($respuesta -match '^\d{10}$') {
                if ($cipaDict.ContainsKey($respuesta)) { $cipaDict[$respuesta] += $units } else { $cipaDict[$respuesta] = $units }
                $cipaDict.Remove($rawKey) | Out-Null
            } else {
                Write-Host "Dejando sin asignar: '$rawKey' con $units unidades"
            }
        }

        Write-Host "Archivo: $filePath"
        foreach ($k in $cipaDict.Keys) {
            $paciente = $maestro | Where-Object { $_.CIPA -and $_.CIPA.Trim() -eq $k }
            $nombrePaciente = if ($paciente) { $paciente.'Nombre paciente' } else { 'No asignado' }
            $nombreEnfermera = if ($paciente) { $paciente.'Nombre enfermera' } else { 'No asignado' }
            Write-Host "CIPA: $k , Paciente: $nombrePaciente , Enfermera: $nombreEnfermera , Unidades: $($cipaDict[$k])"
        }

        $confirmar = ConfirmProcess "¿Deseas actualizar el Excel con los datos anteriores?`nUnidades totales (archivo): $((($cipaDict.Values)|Measure-Object -Sum).Sum)" "Confirmación final ($nombre)"
        if (-not $confirmar) { Write-Host "Usuario canceló volcado a Excel para $filePath"; continue }

        # --- ACTUALIZAR EXCEL ---
        $xlUp = -4162
        try {
            $lastUsedRow = $sheet.Cells($sheet.Rows.Count,1).End($xlUp).Row
        } catch {
            $lastUsedRow = $startRow - 1
        }
        if ($lastUsedRow -lt $startRow) { $lastUsedRow = $startRow - 1 }

        foreach ($k in $cipaDict.Keys) {
            $unidades = $cipaDict[$k]

            $cipaStr = if ($k -match '^\d{10}$') { $k } else { $k }
            $found = $false
            for ($i = $startRow; $i -le $lastUsedRow; $i++) {
                $cellValue = $sheet.Cells.Item($i,1).Value2
                if ($cellValue) { $cellValue = $cellValue.ToString().Trim() }
                if ($cellValue -eq $cipaStr) {
                    $current = $sheet.Cells.Item($i,$colMes).Value2
                    if (-not [int]::TryParse($current,[ref]$current)) { $current = 0 }
                    $sheet.Cells.Item($i,$colMes).Value2 = $current + $unidades
                    $found = $true
                    break
                }
            }

            if (-not $found) {

                $newRow = $lastUsedRow + 6
                $sheet.Cells.Item($newRow,1).NumberFormat = "@"
                try {
                    $sheet.Cells.Item($newRow,1).Formula = "'" + [string]$cipaStr
                } catch {
                    try { $sheet.Cells.Item($newRow,1).Value2 = [string]$cipaStr } catch { $sheet.Cells.Item($newRow,1).Formula = "'" + ([string]::Empty) }
                }

                $paciente = $null
           
                if ($cipaStr -match '^\d{10}$') {
                    $paciente = $maestro | Where-Object { $_.CIPA -and ($_.CIPA.ToString().Trim() -eq $cipaStr) } | Select-Object -First 1
                } else {
                    $paciente = $null
                }

                if ($paciente) {
                    $sheet.Cells.Item($newRow,2).Value2 = $paciente.'Nombre paciente'
                    $sheet.Cells.Item($newRow,3).Value2 = $paciente.'Nombre enfermera'
                } else {
                    $sheet.Cells.Item($newRow,2).Value2 = 'No asignado'
                    $sheet.Cells.Item($newRow,3).Value2 = 'No asignado'
                }
                $sheet.Cells.Item($newRow,$colMes).Value2 = [int]$unidades
                $lastUsedRow = $newRow
            }

            $totalFabricante += $unidades
            if ($cipaGlobal.ContainsKey($cipaStr)) { $cipaGlobal[$cipaStr] += $unidades } else { $cipaGlobal[$cipaStr] = $unidades }
        }
        $historicoFolder = Join-Path $carpeta ("{0}Historico" -f (Get-Culture).TextInfo.ToTitleCase($mesActual))
        if (-not (Test-Path $historicoFolder)) { New-Item -Path $historicoFolder -ItemType Directory | Out-Null }

        try {
            Move-Item -Path $filePath -Destination (Join-Path $historicoFolder $txt.Name) -Force -ErrorAction Stop
        } catch {
            Write-Host "No se pudo mover $filePath a histórico: $($_.Exception.Message)"
        }

        Get-ChildItem $carpeta -Filter *.pdf -File -ErrorAction SilentlyContinue | ForEach-Object {
            try { Move-Item -Path $_.FullName -Destination (Join-Path $historicoFolder $_.Name) -Force -ErrorAction Stop } catch {}
        }

    } 

    try {
        $workbook.Save()
    } catch {
        Write-Host "Advertencia: no se pudo guardar workbook: $($_.Exception.Message)"
    }
    try { $excel.Quit() } catch {}

    # --- Resumen final ---
    Write-Host "----------------------------------------"
    Write-Host "Resumen de $nombre"
    foreach ($cipa in $cipaGlobal.Keys) {
        $pac = $maestro | Where-Object { $_.CIPA -and $_.CIPA.Trim() -eq $cipa } | Select-Object -First 1
        $nombrePaciente = if ($pac) { $pac.'Nombre paciente' } else { 'No asignado' }
        $nombreEnfermera = if ($pac) { $pac.'Nombre enfermera' } else { 'No asignado' }
        Write-Host "CIPA: $cipa , Paciente: $nombrePaciente , Enfermera: $nombreEnfermera , Total Unidades: $($cipaGlobal[$cipa])"
    }
    Write-Host "Unidades totales de sensores: $totalFabricante"
    Write-Host "----------------------------------------"

    [System.Windows.Forms.MessageBox]::Show(
        "$nombre finalizado correctamente.`nUnidades totales de sensores: $totalFabricante",
        "Fin $nombre",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null
}

# --- EJECUCIÓN ---
ProcessManufacturer -nombre "Abbott" -carpeta $carpetaAbbott -maestroCSV $maestroCSV -excelPath $excelPath
ProcessManufacturer -nombre "Dexcom" -carpeta $carpetaDexcom -maestroCSV $maestroCSVDexcom -excelPath $excelPathDexcom

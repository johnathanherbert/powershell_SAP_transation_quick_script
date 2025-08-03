# Define o caminho completo do Script1.vbs
$scriptPath = 'C:\Users\user\AppData\Roaming\SAP\SAP GUI\Scripts\Script1.vbs'

# 1) Entrada e valicacao do codigo de 6 digitos
do {
    $code = Read-Host 'Informe o codigo do material: '
} while ($code -notmatch '^\d{6}$')

# 2) Gera automaticamente a data de hoje no formato DD.MM.AAAA
$date = Get-Date -Format 'dd.MM.yyyy'
Write-Host "Data atual utilizada: $date"

# 4) Processa o arquivo linha a linha
$novoConteudo = Get-Content -Path $scriptPath -Encoding Default | ForEach-Object {

    # Se for a linha do material number, substitui o valor entre aspas
    if ($_ -match 'ctxtT2_MATNR-LOW') {
        $_ -replace '(?<=ctxtT2_MATNR-LOW"\)\.text = ")(\d{6})(?=")', $code
    }
    # Se for a linha da data, substitui o valor entre aspas
    elseif ($_ -match 'ctxt%%DYN003-LOW') {
        $_ -replace '(?<=ctxt%%DYN003-LOW"\)\.text = ")(\d{2}\.\d{2}\.\d{4})(?=")', $date
    }
    else {
        $_
    }
}

# 6) Grava o novo conteudo em ANSI (ASCII)
$novoConteudo | Out-File -FilePath $scriptPath -Encoding ASCII -Force
Write-Host "Arquivo salvo em ANSI:`n$scriptPath`n"

# 7) Executa o script no SAP GUI via cscript.exe
Write-Host 'Executando cscript.exe...'
Start-Process -FilePath 'cscript.exe' `
              -ArgumentList "//nologo `"$scriptPath`"" `
              -NoNewWindow -Wait

Write-Host "`nExecucao concluida com sucesso."

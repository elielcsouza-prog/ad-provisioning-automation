<#
.SYNOPSIS
    Automação de Provisionamento em Lote (Active Directory + Outlook Notification).
    
.DESCRIPTION
    Script robusto para criação de usuários no AD a partir de dados tabulados.
    Inclui:
    - Tratamento de caracteres especiais (Normalização Unicode).
    - Lógica de atribuição de OUs e Proxy Addresses baseada em categorias.
    - Geração automatizada de e-mail de boas-vindas via Outlook.
#>

Import-Module ActiveDirectory

# --- FUNÇÃO PARA NORMALIZAÇÃO DE TEXTO ---
function Set-NormalizedText ($texto) {
    if ($null -eq $texto) { return "" }
    $normalizado = $texto.Normalize([System.Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($c in $normalizado.ToCharArray()) {
        $unicodeCategory = [System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($c)
        if ($unicodeCategory -ne [System.Globalization.UnicodeCategory]::NonSpacingMark) {
            $null = $sb.Append($c)
        }
    }
    return $sb.ToString().Replace('ç', 'c').Replace('Ç', 'C')
}

# --- BUSCA DINÂMICA DE ASSINATURA LOCAL ---
$appData = [System.Environment]::GetFolderPath('ApplicationData')
$sigPath = "$appData\Microsoft\Signatures"
$assinaturaHTML = ""

if (Test-Path $sigPath) {
    $lastSig = Get-ChildItem -Path $sigPath -Filter "*.htm" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($lastSig) {
        $assinaturaHTML = Get-Content -Path $lastSig.FullName -Raw -Encoding UTF8
    }
}

# --- CONFIGURAÇÕES DE LOG ---
$logPath = Join-Path $env:USERPROFILE "Documents\Logs_Provisionamento"
if (!(Test-Path $logPath)) { New-Item -ItemType Directory -Path $logPath }
$logFile = Join-Path $logPath "Log_Criacao_Lote.txt"

$sucessos = 0
$erros = 0

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "   AUTOMAÇÃO DE PROVISIONAMENTO: AD + NOTIFICAÇÃO         " -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "`nInstruções: Cole as linhas do Excel e pressione ENTER duas vezes." -ForegroundColor Yellow

$inputLines = @()
while ($line = Read-Host "Dados") {
    if ($line -ne "") { $inputLines += $line }
}

foreach ($inputRaw in $inputLines) {
    if ($inputRaw -like "*`t*") {
        $d = $inputRaw -split "`t"
        $primeiroNomeOrig = $d[0]; $sobreNomeOrig = $d[1]; $nomeCompletoOrig = $d[2]; $samAccount = $d[3]
        $cargo = $d[4]; $depto = $d[5]; $empresa = $d[6]; $matricula = $d[7]
        $rg = $d[8]; $cpf = $d[9]; $enderecoCom = $d[10]; $gerente = $d[11]
        $opcao = $d[13]

        $primeiroNomeLimpo = Set-NormalizedText $primeiroNomeOrig
        $sobreNomeLimpo    = Set-NormalizedText $sobreNomeOrig
    } else { continue }

    # --- LÓGICA DE ORGANIZAÇÃO (OU) ---
    $domainSuffix = "empresa.com.br"
    $targetOU = ""
    $mailPrimario = "$samAccount@$domainSuffix"

    # Switch simplificado sem referências específicas
    switch ($opcao) {
        "1" { $targetOU = "OU=Terceiros,DC=dominio,DC=local" }
        "2" { $targetOU = "OU=Funcionarios,DC=dominio,DC=local" }
        default { $targetOU = "OU=Usuarios_Geral,DC=dominio,DC=local" }
    }

    $cpfLimpo = $cpf -replace "[^0-9]", ""
    $senhaTxt = "SenhaProvisoria-" + $cpfLimpo.Substring(0,4)
    $senhaSec = ConvertTo-SecureString $senhaTxt -AsPlainText -Force

    $parametros = @{
        Name = $nomeCompletoOrig
        DisplayName = $nomeCompletoOrig
        GivenName = $primeiroNomeLimpo
        Surname = $sobreNomeLimpo
        SamAccountName = $samAccount
        UserPrincipalName = $mailPrimario
        EmailAddress = $mailPrimario
        Title = $cargo
        Department = $depto
        Company = $empresa
        Path = $targetOU
        AccountPassword = $senhaSec
        Enabled = $true
        ChangePasswordAtLogon = $true
        Description = "Ativo - Provisionamento Automatizado"
    }

    try {
        New-ADUser @parametros -ErrorAction Stop
        $sucessos++
        Write-Host "[OK] -> Identidade $($samAccount) criada." -ForegroundColor Green
        "$(Get-Date): SUCESSO - $($samAccount)" | Out-File $logFile -Append

        # --- GERAÇÃO DE E-MAIL (RASCUNHO) ---
        try {
            $outlook = New-Object -ComObject Outlook.Application
            $mail = $outlook.CreateItem(0)
            $mail.Subject = "Acessos de Rede: $nomeCompletoOrig"
            $corpoHTML = "<html><body style='font-family: Calibri;'>
                          <p>Olá,</p>
                          <p>A conta para <b>$nomeCompletoOrig</b> foi criada.</p>
                          <p>Senha Provisória: SenhaProvisoria-XXXX (Primeiros 4 dígitos do CPF)</p>
                          </body></html>"
            $mail.HTMLBody = $corpoHTML + $assinaturaHTML
            $mail.Save()
        } catch { Write-Host "  [!] Erro na notificação Outlook." -ForegroundColor Yellow }

    } catch {
        $erros++
        Write-Host "[ERRO] -> $($samAccount): $($_.Exception.Message)" -ForegroundColor Red
        "$(Get-Date): ERRO - $($samAccount)" | Out-File $logFile -Append
    }
}

Write-Host "`nConcluído! Sucessos: $sucessos | Erros: $erros" -ForegroundColor Cyan

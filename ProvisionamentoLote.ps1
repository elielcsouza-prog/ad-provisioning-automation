<#
.SYNOPSIS
    Provisionamento de Identidades em Lote e Integração Híbrida (AD + Graph API + Outlook).
    
.DESCRIPTION
    Script avançado em PowerShell para processamento de Joiners (novas contas) e migrações.
    Realiza saneamento de texto (Unicode/FormD), verificação condicional de limpeza de
    OnPremisesImmutableId via Microsoft Graph API e criação de rascunhos de onboarding no Outlook.
#>

Import-Module ActiveDirectory

# --- FUNÇÃO PARA REMOVER ACENTOS E Ç (UNICODE / FORMD) ---
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

# --- BUSCA DINÂMICA DE ASSINATURA (PORTABILIDADE) ---
$appData = [System.Environment]::GetFolderPath('ApplicationData')
$sigPath = "$appData\Microsoft\Signatures"
$assinaturaHTML = ""

if (Test-Path $sigPath) {
    $lastSig = Get-ChildItem -Path $sigPath -Filter "*.htm" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($lastSig) {
        $assinaturaHTML = Get-Content -Path $lastSig.FullName -Raw -Encoding UTF8
        $sigFolder = $lastSig.Name.Replace(".htm", "_files")
        $assinaturaHTML = $assinaturaHTML -replace $sigFolder, "$sigPath\$sigFolder"
    }
}

# --- CONFIGURAÇÕES DE LOGS ---
$logPath = Join-Path $env:USERPROFILE "Documents\Logs_Automacao"
if (!(Test-Path $logPath)) { New-Item -ItemType Directory -Path $logPath }
$logFile = Join-Path $logPath "Log_Criacao_Lote.txt"

$sucessos = 0
$erros = 0

Write-Host "--- CADASTRO DE IDENTIDADES EM LOTE (IAM / IGA) ---" -ForegroundColor Cyan
Write-Host "`nInstruções: Copie os dados tabulados do Excel e cole abaixo." -ForegroundColor Yellow
Write-Host "Colunas esperadas (17 colunas separadas por TAB):" -ForegroundColor Gray
Write-Host "PrimeiroNome | Sobrenome | NomeCompleto | SamAccountName | Cargo | Depto | Empresa | Matricula | RG | CPF | EnderecoCom | Gerente | TemEmail | Opcao | MalhaOp | CloudTimestamp | TipoProcesso (nova/migracao)" -ForegroundColor Gray
Write-Host "`nPressione ENTER em branco para iniciar o processamento." -ForegroundColor Yellow

$inputLines = @()
while ($line = Read-Host "Dados") {
    if ($line -ne "") { $inputLines += $line }
}

$registrosProcessados = @()
$precisaDeGraph = $false

foreach ($inputRaw in $inputLines) {
    if ($inputRaw -like "*`t*") {
        $d = $inputRaw -split "`t"
        
        $tipo = if ($d.Count -gt 16) { $d[16].Trim().ToLower() } else { "nova" }

        $registro = [PSCustomObject]@{
            PrimeiroNomeOrig = $d[0]; SobreNomeOrig = $d[1]; NomeCompletoOrig = $d[2]; SamAccount = $d[3]
            Cargo = $d[4]; Depto = $d[5]; Empresa = $d[6]; Matricula = $d[7]
            Rg = $d[8]; Cpf = $d[9]; EnderecoCom = $d[10]; Gerente = $d[11]
            TemEmail = $d[12]; Opcao = $d[13]; MalhaOp = $d[14]; CloudTimestamp = $d[15]
            TipoProcesso = $tipo
        }
        $registrosProcessados += $registro
        
        if ($tipo -eq "migracao") {
            $precisaDeGraph = $true
        }
    }
}

# --- CONEXÃO COM MICROSOFT GRAPH (CONDICIONAL) ---
if ($precisaDeGraph) {
    try {
        Write-Host "`n[GRAPH] Detectadas contas de migração. Solicitando autenticação..." -ForegroundColor Yellow
        Connect-MgGraph -Scopes "User.ReadWrite.All" -ErrorAction Stop
        Write-Host "[GRAPH] Conectado com sucesso!" -ForegroundColor Green
    }
    catch {
        Write-Host "[ERRO CRÍTICO] Falha ao conectar ao Microsoft Graph: $($_.Exception.Message)" -ForegroundColor Red
        Read-Host "Pressione ENTER para sair"
        exit
    }
} else {
    Write-Host "`nApenas contas novas detectadas. Conexão Graph ignorada (Processamento Local)." -ForegroundColor Green
}

Write-Host "`nIniciando processamento de $($registrosProcessados.Count) registros..." -ForegroundColor White

foreach ($reg in $registrosProcessados) {
    $primeiroNomeLimpo = Set-NormalizedText $reg.PrimeiroNomeOrig
    $sobreNomeLimpo    = Set-NormalizedText $reg.SobreNomeOrig
    $baseEmail         = $reg.SamAccount.ToLower()
    $matricula         = $reg.Matricula

    Write-Host "`n--------------------------------------------------" -ForegroundColor DarkGray
    Write-Host "Processando: $($reg.NomeCompletoOrig) ($matricula) | Tipo: $($reg.TipoProcesso.ToUpper())" -ForegroundColor White

    # ==========================================================================
    # PASSO 1: Mapeamento de OUs, Proxies e Domínios Genéricos
    # ==========================================================================
    $targetOU = ""; $listaProxys = @(); $mailPrimario = ""; $ext3 = ""; $descricao = ""; $upnFinal = ""

    # Domínios genéricos para portabilidade do portfólio
    $domainTerceiros = "terceiros.empresa.com.br"
    $domainInterno   = "empresa.com.br"
    $domainCloudFix  = "cloud.empresa.com.br"

    switch ($reg.Opcao) {
        "1" {
            $targetOU = "OU=Terceiros,OU=Usuarios,DC=empresa,DC=local"
            $mailPrimario = "$baseEmail@$domainTerceiros"
            $listaProxys = @("sip:$mailPrimario", "SMTP:$mailPrimario", "smtp:$baseEmail@ext.empresa.com.br")
            $descricao = "Ativo - Terceiro"
            $upnFinal = if ($reg.TipoProcesso -eq "migracao") { "$baseEmail@$domainCloudFix" } else { $mailPrimario }
        }
        "2" {
            $targetOU = "OU=Internos,OU=Usuarios,DC=empresa,DC=local"
            $mailPrimario = "$baseEmail@$domainInterno"
            $ext3 = if ($reg.MalhaOp -eq "1") { "Regiao Norte" } else { "Regiao Sul" }
            $descricao = "Ativo - $($ext3)"
            $listaProxys = @("sip:$mailPrimario", "SMTP:$mailPrimario", "smtp:$baseEmail@alias.empresa.com.br")
            $upnFinal = if ($reg.TipoProcesso -eq "migracao") { "$baseEmail@$domainCloudFix" } else { $mailPrimario }
        }
        default {
            $targetOU = "OU=Geral,OU=Usuarios,DC=empresa,DC=local"
            $mailPrimario = "$baseEmail@$domainInterno"
            $listaProxys = @("sip:$mailPrimario", "SMTP:$mailPrimario")
            $descricao = "Ativo - Colaborador"
            $upnFinal = if ($reg.TipoProcesso -eq "migracao") { "$baseEmail@$domainCloudFix" } else { $mailPrimario }
        }
    }

    # ==========================================================================
    # PASSO 2: Limpeza do ImmutableID no Microsoft Graph (Migrações)
    # ==========================================================================
    if ($reg.TipoProcesso -eq "migracao") {
        Write-Host "[GRAPH] [$upnFinal] Resetando OnPremisesImmutableId..." -ForegroundColor Yellow
        try {
            $bodyPatch = @{ OnPremisesImmutableId = $null }
            $uriGraph  = "https://graph.microsoft.com/v1.0/users/$upnFinal"

            Invoke-MgGraphRequest -Method PATCH -Uri $uriGraph -Body $bodyPatch -ErrorAction Stop
            Start-Sleep -Seconds 1
            Invoke-MgGraphRequest -Method PATCH -Uri $uriGraph -Body $bodyPatch -ErrorAction Stop
            
            Write-Host "   [OK] Limpeza concluída no Microsoft Entra ID." -ForegroundColor Green
        }
        catch {
            Write-Host "   [ERRO/GRAPH] Falha ao resetar ImmutableId: $($_.Exception.Message)" -ForegroundColor Red
            $(Get-Date).ToString() + " : ERRO GRAPH - $upnFinal - $($_.Exception.Message)" | Out-File $logFile -Append
        }
    }

    # ==========================================================================
    # PASSO 3: Regras de Senha Provisória e Atributos AD
    # ==========================================================================
    $cpfLimpo = $reg.Cpf -replace "[^0-9]", ""
    $senhaTxt = "Bemvindo-" + if ($cpfLimpo.Length -ge 4) { $cpfLimpo.Substring(0,4) } else { "1234" }
    $senhaSec = ConvertTo-SecureString $senhaTxt -AsPlainText -Force

    $outrosAtributos = @{ "employeeID" = $matricula; "msDS-cloudExtensionAttribute1" = $reg.CloudTimestamp; "RG" = $reg.Rg; "CPF" = $reg.Cpf }
    if ($ext3) { $null = $outrosAtributos.Add("extensionAttribute3", $ext3) }
    if ($reg.TemEmail -eq "s") { $null = $outrosAtributos.Add("proxyAddresses", $listaProxys) }

    $parametros = @{
        Name = $reg.NomeCompletoOrig; DisplayName = $reg.NomeCompletoOrig; GivenName = $primeiroNomeLimpo; Surname = $sobreNomeLimpo
        SamAccountName = $matricula 
        UserPrincipalName = $upnFinal; EmailAddress = $mailPrimario; Title = $reg.Cargo; Department = $reg.Depto; Company = $reg.Empresa
        Office = $reg.EnderecoCom 
        Manager = $reg.Gerente; Path = $targetOU; AccountPassword = $senhaSec; Enabled = $true
        ChangePasswordAtLogon = $true; Description = $descricao; OtherAttributes = $outrosAtributos
    }

    # ==========================================================================
    # PASSO 4: Criação Local & Geração de Rascunho no Outlook
    # ==========================================================================
    try {
        New-ADUser @parametros
        $sucessos++
        Write-Host "[OK] -> Account $($matricula) criada no AD. UPN: $upnFinal" -ForegroundColor Green
        $(Get-Date).ToString() + " : SUCESSO - $($matricula) ($($reg.TipoProcesso))" | Out-File $logFile -Append

        Start-Sleep -Seconds 1
        try {
            $userData = Get-ADUser -Identity $matricula -Properties DisplayName, EmailAddress
            $nomeParaEmail = $userData.DisplayName
            $emailParaEmail = if ($userData.EmailAddress) { $userData.EmailAddress } else { $mailPrimario }
            $destinatarioFinal = if ($reg.Gerente) { (Get-ADUser -Identity $reg.Gerente -Properties EmailAddress).EmailAddress } else { "" }
            
            $assunto = "Solicitação de Acessos - Onboarding de Colaborador"
            $licenciado = if ($reg.TemEmail -eq "s") { "Sim" } else { "Não Solicitado" }

            $corpoHTML = @"
<html>
<body style="font-family: 'Calibri', sans-serif;">
    <p>Prezado(a),</p>
    <p>O usuário de rede corporativo foi provisionado com sucesso conforme solicitado.</p>
    
    <table style="border-collapse: collapse; width: 100%; border: 1px solid #002060;">
        <thead>
            <tr style="background-color: #002060; color: white; text-align: center; font-weight: bold;">
                <td style="padding: 5px; border: 1px solid #002060;">Matrícula</td>
                <td style="padding: 5px; border: 1px solid #002060;">Nome</td>
                <td style="padding: 5px; border: 1px solid #002060;">Licença E-mail</td>
                <td style="padding: 5px; border: 1px solid #002060;">Endereço de E-mail</td>
                <td style="padding: 5px; border: 1px solid #002060;">Senha Provisória</td>
            </tr>
        </thead>
        <tbody>
            <tr style="text-align: center; font-weight: bold; color: #0070C0;">
                <td style="padding: 5px; border: 1px solid #002060;">$matricula</td>
                <td style="padding: 5px; border: 1px solid #002060;">$nomeParaEmail</td>
                <td style="padding: 5px; border: 1px solid #002060;">$licenciado</td>
                <td style="padding: 5px; border: 1px solid #002060; font-weight: normal; text-decoration: underline;">$emailParaEmail</td>
                <td style="padding: 5px; border: 1px solid #002060;">Bemvindo-XXXX</td>
            </tr>
        </tbody>
    </table>

    <p style="color: red; font-weight: bold; margin-top: 20px;">Atenção:</p>
    <ul style="color: red; font-weight: bold;">
        <li>Os dígitos 'XXXX' na senha provisória correspondem aos 4 primeiros dígitos do CPF do colaborador.</li>
    </ul>

    <ul>
        <li>Primeiro acesso para redefinição de senha: <a href="https://portal.empresa.com.br">https://portal.empresa.com.br</a></li>
        <li>O registro de MFA é obrigatório no primeiro logon.</li>
    </ul>

    <p>Central de Atendimento / Gestão de Acessos IAM</p>
</body>
</html>
"@
            $outlook = New-Object -ComObject Outlook.Application
            $mail = $outlook.CreateItem(0)
            $mail.To = $destinatarioFinal
            $mail.CC = "suporte.iam@empresa.com.br"
            $mail.Subject = $assunto
            $mail.HTMLBody = $corpoHTML + $assinaturaHTML
            $mail.Save()
        } 
        catch { 
            Write-Host "    [!] Erro ao gerar rascunho no Outlook." -ForegroundColor Yellow 
        }

    } 
    catch {
        $erros++
        Write-Host "[ERRO] -> $($matricula): $($_.Exception.Message)" -ForegroundColor Red
        $(Get-Date).ToString() + " : ERRO - $($matricula) - $($_.Exception.Message)" | Out-File $logFile -Append
    }
}

Write-Host "`n==================================================" -ForegroundColor DarkGray
Write-Host "Processamento Concluído! Sucessos: $sucessos | Erros: $erros" -ForegroundColor Cyan

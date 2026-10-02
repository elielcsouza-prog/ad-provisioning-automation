<#
.SYNOPSIS
    Central de Provisionamento & Auditoria de Acessos - IAM (WPF nativo em PowerShell)

.DESCRIPTION
    Interface corporativa em WPF para automação do ciclo de vida de identidades (Joiners e Migrações).
    - Execução desacoplada via Runspace STA (sem travamento de UI ou conflito de COM).
    - Validação prévia de homônimos e conformidade cadastral no Active Directory.
    - Notificação corporativa padronizada via Microsoft Outlook.
    - Compatível com compilação direta via PS2EXE.

    Compilar com PS2EXE (NÃO utilizar -requireAdmin):
        Invoke-ps2exe .\IAM_Provisionamento_WPF.ps1 .\IAM_Provisionamento.exe -STA -noConsole -title "Central IAM - Provisionamento"

    Requisitos: Windows PowerShell 5.1, RSAT ActiveDirectory, Microsoft Outlook desktop.
#>

# --- Garantia de Thread STA (WPF + COM Outlook) ---
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne [System.Threading.ApartmentState]::STA) {
    $self = $MyInvocation.MyCommand.Path
    if ($self -and $self.ToLower().EndsWith('.ps1')) {
        Start-Process -FilePath 'powershell.exe' -ArgumentList @('-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass', '-WindowStyle', 'Hidden', '-File', ('"{0}"' -f $self))
        exit
    }
}

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem

# =====================================================================================
# 1. MODELO DE DADOS (C# 5 compatível)
# =====================================================================================
if (-not ('Registro' -as [type])) {
Add-Type -TypeDefinition @'
using System;
using System.ComponentModel;
using System.Text.RegularExpressions;

public class Registro : INotifyPropertyChanged
{
    public event PropertyChangedEventHandler PropertyChanged;
    public static volatile bool Sujo;

    private string _id = Guid.NewGuid().ToString("N");
    public string Id { get { return _id; } set { _id = value; } }

    private void N(string p) { Sujo = true; PropertyChangedEventHandler h = PropertyChanged; if (h != null) h(this, new PropertyChangedEventArgs(p)); }
    private bool S(ref string f, string v, string n, bool upn) { if (v == null) v = ""; if (f == v) return false; f = v; N(n); if (upn) N("Upn"); return true; }
    private bool B(ref bool f, bool v, string n) { if (f == v) return false; f = v; N(n); return true; }

    private bool _sel, _ok, _mfa, _lic;
    private string _status = "nova", _prim = "", _sob = "", _nome = "", _logon = "", _cargo = "", _depto = "", _empresa = "";
    private string _mat = "", _rg = "", _cpf = "", _end = "", _ger = "", _email = "s", _perfil = "2", _malha = "1";
    private string _cloud = "", _cham = "", _cc = "", _notas = "", _res = "", _det = "", _rasc = "";

    public bool Sel { get { return _sel; } set { B(ref _sel, value, "Sel"); } }
    public bool Ok { get { return _ok; } set { B(ref _ok, value, "Ok"); } }
    public bool Mfa { get { return _mfa; } set { B(ref _mfa, value, "Mfa"); } }
    public bool Licenca { get { return _lic; } set { B(ref _lic, value, "Licenca"); } }

    public string Status { get { return _status; } set { S(ref _status, value == null ? "" : value.Trim().ToLower(), "Status", true); } }
    public string PrimeiroNome { get { return _prim; } set { S(ref _prim, value, "PrimeiroNome", true); } }
    public string Sobrenome { get { return _sob; } set { S(ref _sob, value, "Sobrenome", true); } }
    public string NomeCompleto { get { return _nome; } set { S(ref _nome, value, "NomeCompleto", false); } }
    public string Logon { get { return _logon; } set { S(ref _logon, NormLogon(value), "Logon", true); } }
    public string Cargo { get { return _cargo; } set { S(ref _cargo, value, "Cargo", false); } }
    public string Depto { get { return _depto; } set { S(ref _depto, value, "Depto", false); } }
    public string Empresa { get { return _empresa; } set { S(ref _empresa, value, "Empresa", true); } }
    public string Matricula { get { return _mat; } set { S(ref _mat, value == null ? "" : value.Trim(), "Matricula", false); } }
    public string Rg { get { return _rg; } set { S(ref _rg, value, "Rg", false); } }
    public string Cpf { get { return _cpf; } set { S(ref _cpf, value, "Cpf", false); } }
    public string Endereco { get { return _end; } set { S(ref _end, value, "Endereco", false); } }
    public string Gerente { get { return _ger; } set { S(ref _ger, value, "Gerente", false); } }
    public string EmailLic { get { return _email; } set { S(ref _email, value, "EmailLic", false); } }
    public string Perfil { get { return _perfil; } set { S(ref _perfil, value, "Perfil", true); } }
    public string Malha { get { return _malha; } set { S(ref _malha, value, "Malha", false); } }
    public string CloudTimestamp { get { return _cloud; } set { S(ref _cloud, value, "CloudTimestamp", false); } }
    public string Chamado { get { return _cham; } set { S(ref _cham, value, "Chamado", false); } }
    public string CentroCusto { get { return _cc; } set { S(ref _cc, value, "CentroCusto", false); } }
    public string Notas { get { return _notas; } set { S(ref _notas, value, "Notas", false); } }
    public string Resultado { get { return _res; } set { S(ref _res, value, "Resultado", false); } }
    public string Detalhe { get { return _det; } set { S(ref _det, value, "Detalhe", false); } }
    public string Rascunho { get { return _rasc; } set { S(ref _rasc, value, "Rascunho", false); } }

    public string Upn
    {
        get
        {
            string baseEmail = !string.IsNullOrEmpty(_logon) ? _logon.ToLower() : (Limpar(_prim) + "." + Limpar(_sob)).ToLower();
            if (baseEmail.Trim('.').Length == 0) return "";
            string dom;
            if (_perfil == "1") dom = "ext.empresa.com.br";
            else if (_perfil == "3") dom = "subsidiaria.com.br";
            else if (_perfil == "4") dom = "parceiro.empresa.com.br";
            else dom = "empresa.com.br";
            
            if (_status == "migracao") dom = "cloud.empresa.com.br";
            return baseEmail + "@" + dom;
        }
    }

    public static string Limpar(string t)
    {
        if (string.IsNullOrWhiteSpace(t)) return "";
        string s = Regex.Replace(t.Normalize(System.Text.NormalizationForm.FormD), "\\p{M}", "");
        return s.Replace("ç", "c").Replace("Ç", "C");
    }

    public static string NormLogon(string s)
    {
        if (string.IsNullOrWhiteSpace(s)) return "";
        return Regex.Replace(Limpar(s).ToLower().Trim(), "\\s+", ".");
    }
}

public class FiltroRegistro
{
    public string Texto = "";
    public string Visao = "todos";

    public Predicate<object> GetPredicate() { return new Predicate<object>(Aceita); }

    public bool Aceita(object o)
    {
        Registro r = o as Registro;
        if (r == null) return false;
        switch (Visao)
        {
            case "pendentes": if (r.Ok || r.Resultado == "Criado") return false; break;
            case "migracao": if (r.Status != "migracao") return false; break;
            case "divergencia": if (r.Status != "pendencia") return false; break;
            case "falhas": if (r.Resultado != "Falha") return false; break;
        }
        if (string.IsNullOrEmpty(Texto)) return true;
        string t = Texto.ToLower();
        string dig = Regex.Replace(t, "\\D", "");
        if (Has(r.Matricula, t) || Has(r.NomeCompleto, t) || Has(r.Cpf, t) || Has(r.Logon, t) ||
            Has(r.Chamado, t) || Has(r.CentroCusto, t) || Has(r.Notas, t)) return true;
        if (dig.Length > 0 && Regex.Replace(r.Cpf ?? "", "\\D", "").Contains(dig)) return true;
        return false;
    }

    private static bool Has(string s, string t) { return s != null && s.ToLower().Contains(t); }
}
'@
}

# =====================================================================================
# 2. MOTOR DE EXECUÇÃO (Background STA Runspace)
# =====================================================================================
function Limpar-Texto {
    param([string]$texto)
    if ([string]::IsNullOrWhiteSpace($texto)) { return "" }
    $semAcento = [System.Text.RegularExpressions.Regex]::Replace($texto.Normalize([System.Text.NormalizationForm]::FormD), '\p{M}', '')
    return $semAcento.Replace('ç', 'c').Replace('Ç', 'C')
}

function Test-Elevado {
    try {
        $p = New-Object System.Security.Principal.WindowsPrincipal([System.Security.Principal.WindowsIdentity]::GetCurrent())
        return $p.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch { return $false }
}

function Send-Fila { param([hashtable]$Msg) $null = $Q.Enqueue([pscustomobject]$Msg) }
function Send-Log { param([string]$Nivel, [string]$Texto) Send-Fila @{ T = 'log'; Nivel = $Nivel; Texto = $Texto } }
function Write-LogArquivo {
    param([string]$Linha)
    try { "$(Get-Date) : $Linha" | Out-File $global:LogFile -Append } catch { }
}

function Get-AssinaturaHtml {
    $appData = [System.Environment]::GetFolderPath('ApplicationData')
    $sigPath = [System.IO.Path]::Combine($appData, "Microsoft\Signatures")
    $assinaturaHTML = ""
    if ([System.IO.Directory]::Exists($sigPath)) {
        $arquivosHtm = [System.IO.Directory]::GetFiles($sigPath, "*.htm")
        if ($arquivosHtm -and $arquivosHtm.Length -gt 0) {
            $arquivoMaisRecente = $null
            $ultimaData = [datetime]::MinValue
            for ($idx = 0; $idx -lt $arquivosHtm.Length; $idx++) {
                $dataCriacao = [System.IO.File]::GetLastWriteTime($arquivosHtm[$idx])
                if ($dataCriacao -gt $ultimaData) {
                    $ultimaData = $dataCriacao
                    $arquivoMaisRecente = $arquivosHtm[$idx]
                }
            }
            if ($null -ne $arquivoMaisRecente) {
                $assinaturaHTML = [System.IO.File]::ReadAllText($arquivoMaisRecente, [System.Text.Encoding]::UTF8)
                $nomeBase = [System.IO.Path]::GetFileNameWithoutExtension($arquivoMaisRecente)
                $sigFolder = $nomeBase + "_files"
                $folderCompleto = [System.IO.Path]::Combine($sigPath, $sigFolder)
                $assinaturaHTML = [regex]::Replace($assinaturaHTML, [regex]::Escape($sigFolder), $folderCompleto.Replace('$', '$$'))
            }
        }
    }
    return $assinaturaHTML
}

function Resolve-PerfilAD {
    param($reg)

    $primeiroNomeLimpo = (Limpar-Texto -texto ($reg.PrimeiroNomeOrig))
    $sobreNomeLimpo    = (Limpar-Texto -texto ($reg.SobreNomeOrig))
    $baseEmail = if ($reg.SamAccount) { $reg.SamAccount.ToLower() } else { ($primeiroNomeLimpo + "." + $sobreNomeLimpo).ToLower() }

    $dominioBase    = "DC=empresa,DC=corp"
    $ouTercNorte    = "OU=Regiao Norte,OU=Terceiros,OU=Usuarios,$dominioBase"
    $ouTercSul      = "OU=Regiao Sul,OU=Terceiros,OU=Usuarios,$dominioBase"
    $ouTercEspecial = "OU=Especial,OU=Terceiros,OU=Usuarios,$dominioBase"
    $ouTercPadrao   = "OU=Geral,OU=Terceiros,OU=Usuarios,$dominioBase"
    $ouInternos     = "OU=Internos,OU=Usuarios,$dominioBase"
    $ouEspecial     = "OU=UnidadeEspecial,OU=Usuarios,$dominioBase"

    $ext3 = if ($reg.MalhaOp -eq "1" -or $reg.MalhaOp -match "norte") {
        "Regiao Norte"
    } elseif ($reg.MalhaOp -eq "2" -or $reg.MalhaOp -match "sul") {
        "Regiao Sul"
    } elseif ($reg.MalhaOp -eq "3" -or $reg.MalhaOp -match "especial") {
        "Especial"
    } else {
        "Regiao Norte"
    }

    if ($reg.Opcao -eq "1") {
        if ($ext3 -eq "Especial") { $targetOU = $ouTercEspecial }
        elseif ($ext3 -eq "Regiao Sul") { $targetOU = $ouTercSul }
        else { $targetOU = $ouTercNorte }
        
        $fallbackOU   = $ouTercPadrao
        $mailPrimario = "$baseEmail@ext.empresa.com.br"
        $listaProxys  = @("sip:$mailPrimario", "SMTP:$mailPrimario", "smtp:$baseEmail@ext.subsidiaria.com.br")
        $descricao    = "Ativo - Prestador de Servicos"
        $upnFinal     = if ($reg.TipoProcesso -eq "migracao") { "$baseEmail@cloud.empresa.com.br" } else { $mailPrimario }
    }
    elseif ($reg.Opcao -eq "3") {
        $targetOU     = $ouEspecial
        $fallbackOU   = $ouInternos
        $mailPrimario = "$baseEmail@subsidiaria.com.br"
        $listaProxys  = @("sip:$mailPrimario", "SMTP:$mailPrimario")
        $descricao    = "Ativo - Unidade Especial"
        $upnFinal     = if ($reg.TipoProcesso -eq "migracao") { "$baseEmail@cloud.empresa.com.br" } else { $mailPrimario }
    }
    elseif ($reg.Opcao -eq "4") {
        $targetOU     = "OU=Parceiros,OU=Usuarios,$dominioBase"
        $fallbackOU   = $ouInternos
        $mailPrimario = "$baseEmail@parceiro.empresa.com.br"
        $descricao    = "Ativo - Parceiro"
        $listaProxys  = @("sip:$mailPrimario", "SMTP:$mailPrimario")
        $upnFinal     = if ($reg.TipoProcesso -eq "migracao") { "$baseEmail@cloud.empresa.com.br" } else { $mailPrimario }
    }
    else {
        $targetOU     = $ouInternos
        $fallbackOU   = $ouTercPadrao
        $mailPrimario = "$baseEmail@empresa.com.br"
        $descricao    = "Ativo - $($ext3)"
        $listaProxys  = @("sip:$mailPrimario", "SMTP:$mailPrimario", "smtp:$baseEmail@alias.empresa.com.br")
        $upnFinal     = if ($reg.TipoProcesso -eq "migracao") { "$baseEmail@cloud.empresa.com.br" } else { $mailPrimario }
    }

    return [pscustomobject]@{
        Ext3 = $ext3; TargetOU = $targetOU; FallbackOU = $fallbackOU; MailPrimario = $mailPrimario
        ListaProxys = $listaProxys; Descricao = $descricao; UpnFinal = $upnFinal; BaseEmail = $baseEmail
    }
}

function Resolve-Gestor {
    param($reg)
    $res = @{ ManagerDN = $null; Destinatario = ""; Aviso = "" }
    if ($reg.Gerente) {
        try {
            $mgrObj = Get-ADUser -Identity $reg.Gerente -Properties DistinguishedName, EmailAddress -ErrorAction SilentlyContinue
            if ($mgrObj) {
                $res.ManagerDN = $mgrObj.DistinguishedName
                if ($mgrObj.EmailAddress) { $res.Destinatario = $mgrObj.EmailAddress }
            } else {
                $res.Aviso = "[!] Gestor '$($reg.Gerente)' não localizado no AD. Manager ignorado."
            }
        } catch {
            $res.Aviso = "[!] Erro ao buscar gestor '$($reg.Gerente)'."
        }
    }
    return $res
}

function New-UsuarioAD {
    param($reg, $perfil, $managerDN)

    $primeiroNomeLimpo  = (Limpar-Texto -texto ($reg.PrimeiroNomeOrig))
    $sobreNomeLimpo     = (Limpar-Texto -texto ($reg.SobreNomeOrig))
    $nomeCompletoLimpo  = (Limpar-Texto -texto ($reg.NomeCompletoOrig))
    $nomeExibicaoAcento = $reg.NomeCompletoOrig
    $matricula          = $reg.Matricula
    $ext3               = $perfil.Ext3
    $targetOU           = $perfil.TargetOU
    $fallbackOU         = $perfil.FallbackOU

    $cpfLimpo = $reg.Cpf -replace "[^0-9]", ""
    $senhaTxt = if ($cpfLimpo.Length -ge 4) { "Init-" + $cpfLimpo.Substring(0, 4) } else { "Init-1234" }
    $senhaSec = ConvertTo-SecureString $senhaTxt -AsPlainText -Force

    $outrosAtributos = @{ "employeeID" = $matricula; "RG" = $reg.Rg; "CPF" = $reg.Cpf }
    if ($reg.CloudTimestamp) { $null = $outrosAtributos.Add("msDS-cloudExtensionAttribute1", $reg.CloudTimestamp) }
    if ($ext3) { $null = $outrosAtributos.Add("extensionAttribute3", $ext3) }
    if ($reg.TemEmail -eq "s") { $null = $outrosAtributos.Add("proxyAddresses", $perfil.ListaProxys) }

    $parametros = @{
        Name                  = $nomeCompletoLimpo
        DisplayName           = $nomeExibicaoAcento
        GivenName             = $primeiroNomeLimpo
        Surname               = $sobreNomeLimpo
        SamAccountName        = $matricula
        UserPrincipalName     = $perfil.UpnFinal
        EmailAddress          = $perfil.MailPrimario
        Title                 = $reg.Cargo
        Department            = $reg.Depto
        Company               = $reg.Empresa
        Office                = $reg.EnderecoCom
        Path                  = $targetOU
        AccountPassword       = $senhaSec
        Enabled               = $true
        ChangePasswordAtLogon = $true
        Description           = $perfil.Descricao
        OtherAttributes       = $outrosAtributos
    }
    if ($managerDN) { $parametros.Add("Manager", $managerDN) }

    $usuarioCriado = $false
    $msgErro = ""
    $cnUsado = $nomeCompletoLimpo
    try {
        New-ADUser @parametros -ErrorAction Stop
        $usuarioCriado = $true
        Write-LogArquivo "SUCESSO - $($matricula) em ($targetOU)"
    }
    catch {
        $msgErro = $_.Exception.Message
        Send-Log 'info' "$matricula - Tentativa primária falhou ($targetOU): $msgErro"

        if ($targetOU -ne $fallbackOU) {
            Send-Log 'info' "$matricula - Redirecionando para OU alternativa: $fallbackOU"
            $parametros["Path"] = $fallbackOU
            try {
                New-ADUser @parametros -ErrorAction Stop
                $usuarioCriado = $true
                $targetOU = $fallbackOU
                Write-LogArquivo "SUCESSO FALLBACK - $($matricula) em ($targetOU)"
            }
            catch {
                $msgErro = $_.Exception.Message
            }
        }

        if (-not $usuarioCriado -and ($msgErro -match "já está em uso|already exists|em uso")) {
            $nomeComMatricula = "$nomeCompletoLimpo - $matricula"
            Send-Log 'info' "$matricula - CN em conflito. Tentando CN composto: '$nomeComMatricula'"
            $parametros["Name"] = $nomeComMatricula
            $parametros["Path"] = $fallbackOU
            try {
                New-ADUser @parametros -ErrorAction Stop
                $usuarioCriado = $true
                $targetOU = $fallbackOU
                $cnUsado = $nomeComMatricula
                Write-LogArquivo "SUCESSO CN COMPOSTO - $($matricula)"
            }
            catch {
                $msgErro = $_.Exception.Message
                Write-LogArquivo "ERRO - $($matricula) - $msgErro"
            }
        }
        elseif (-not $usuarioCriado) {
            Write-LogArquivo "ERRO - $($matricula) - $msgErro"
        }
    }

    return [pscustomobject]@{ Criado = $usuarioCriado; Erro = $msgErro; OU = $targetOU; CN = $cnUsado }
}

function Get-CorpoEmail {
    param($matricula, $nomeExibicaoAcento, $licenciado, $mailPrimario)
    $corpoHTML = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
</head>
<body style="font-family: 'Segoe UI', Tahoma, Arial, sans-serif; background-color: #f8fafc; margin: 0; padding: 20px; color: #1e293b;">
    <table align="center" border="0" cellpadding="0" cellspacing="0" width="650" style="background-color: #ffffff; border: 1px solid #e2e8f0; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);">
        <tr>
            <td style="background-color: #002060; padding: 24px 30px;">
                <table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                        <td>
                            <h2 style="color: #ffffff; margin: 0; font-size: 19px; font-weight: 600;">Central de Governança de Identidades &amp; Acessos</h2>
                            <p style="color: #93c5fd; margin: 4px 0 0 0; font-size: 13px;">Onboarding de Colaborador | Processo de Joiner Concluído</p>
                        </td>
                        <td align="right">
                            <span style="background-color: #0284c7; color: #ffffff; padding: 4px 10px; border-radius: 12px; font-size: 11px; font-weight: bold; text-transform: uppercase;">Acesso Ativo</span>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td style="padding: 28px 30px;">
                <p style="font-size: 14px; margin: 0 0 16px 0; line-height: 1.5;">Prezado(a) Gestor(a),</p>
                <p style="font-size: 14px; margin: 0 0 20px 0; line-height: 1.5;">Informamos que a conta de rede e os acessos corporativos básicos para o(a) colaborador(a) abaixo foram configurados com sucesso no domínio institucional.</p>

                <table width="100%" border="0" cellpadding="10" cellspacing="0" style="border-collapse: collapse; margin-bottom: 24px; border: 1px solid #cbd5e1; font-size: 13px;">
                    <thead>
                        <tr style="background-color: #f1f5f9; text-align: left; color: #475569; font-size: 11px; text-transform: uppercase;">
                            <th style="border: 1px solid #cbd5e1; padding: 10px;">Matrícula</th>
                            <th style="border: 1px solid #cbd5e1; padding: 10px;">Nome Completo</th>
                            <th style="border: 1px solid #cbd5e1; padding: 10px;">Licença E-mail</th>
                            <th style="border: 1px solid #cbd5e1; padding: 10px;">Logon Principal (UPN)</th>
                            <th style="border: 1px solid #cbd5e1; padding: 10px;">Senha Provisória</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr style="background-color: #ffffff; font-weight: 500;">
                            <td style="border: 1px solid #cbd5e1; color: #002060; font-weight: bold;">$matricula</td>
                            <td style="border: 1px solid #cbd5e1;">$nomeExibicaoAcento</td>
                            <td style="border: 1px solid #cbd5e1; text-align: center;">$licenciado</td>
                            <td style="border: 1px solid #cbd5e1; color: #0284c7; font-weight: 600;">$mailPrimario</td>
                            <td style="border: 1px solid #cbd5e1; background-color: #fef2f2; color: #991b1b; font-weight: bold; text-align: center;">Init-XXXX</td>
                        </tr>
                    </tbody>
                </table>

                <div style="background-color: #fef2f2; border-left: 4px solid #dc2626; padding: 12px 16px; margin-bottom: 24px; border-radius: 0 4px 4px 0;">
                    <p style="margin: 0; font-size: 13px; color: #991b1b; line-height: 1.4;">
                        <strong>Atenção à Senha Provisória:</strong> Os dígitos <strong>XXXX</strong> correspondem aos <strong>4 primeiros dígitos do CPF</strong> cadastrado. Exemplo: CPF 123.456.789-00 utilizará a credencial <code>Init-1234</code>.
                    </p>
                </div>

                <h3 style="font-size: 14px; color: #002060; margin: 0 0 12px 0; text-transform: uppercase;">Diretrizes de Ativação e Segurança</h3>
                <ol style="font-size: 13px; margin: 0 0 20px 20px; padding: 0; line-height: 1.6; color: #334155;">
                    <li><strong>Troca de Senha Obrigatória:</strong> Redefinição mandatória no primeiro acesso via portal: <a href="https://autoatendimento.empresa.com.br" style="color: #0284c7; text-decoration: none; font-weight: 600;">portal.empresa.com.br</a>.</li>
                    <li><strong>Registro de MFA (Multifator):</strong> Configuração obrigatória no primeiro logon via Microsoft Authenticator.</li>
                    <li><strong>Complexidade:</strong> Mínimo de 10 caracteres contendo maiúsculas, minúsculas, números e caracteres especiais.</li>
                </ol>

                <div style="background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 6px; padding: 14px; font-size: 12px; color: #64748b;">
                    <strong style="color: #1e293b;">Suporte Técnico:</strong><br>
                    Em caso de dúvidas, consulte o catálogo de identidades ou contate a Central de Atendimento no portal de serviços.
                </div>

                <p style="font-size: 13px; margin: 24px 0 0 0; color: #475569;">
                    Atenciosamente,<br>
                    <strong>Equipe de Governança de Identidades &amp; Acessos (IAM/IGA)</strong>
                </p>
            </td>
        </tr>
        <tr>
            <td style="background-color: #0f172a; padding: 12px 30px; text-align: center;">
                <p style="color: #94a3b8; font-size: 11px; margin: 0;">Mensagem automática gerada pela Central de Provisionamento Corporativo.</p>
            </td>
        </tr>
    </table>
</body>
</html>
"@
    return $corpoHTML
}

function New-RascunhoOutlook {
    param($outlook, $reg, $mailPrimario, $destinatarioFinal, $assinaturaHTML)
    $matricula = $reg.Matricula
    $nomeExibicaoAcento = $reg.NomeCompletoOrig
    $numChamado = $reg.Chamado

    $assunto = if (-not [string]::IsNullOrWhiteSpace($numChamado)) {
        "Solicitação $numChamado - Provisionamento de Conta Corporativa"
    } else {
        "Provisionamento de Conta Corporativa - Onboarding"
    }

    $licenciado = if ($reg.TemEmail -eq "s") { "Sim" } else { "Não Solicitado" }
    $corpoHTML = Get-CorpoEmail -matricula $matricula -nomeExibicaoAcento $nomeExibicaoAcento -licenciado $licenciado -mailPrimario $mailPrimario

    $mail = $outlook.CreateItem(0)
    try {
        if ($destinatarioFinal -ne "") { $mail.To = $destinatarioFinal }
        $mail.CC = "gestaodeacessos@empresa.com.br"
        $mail.Subject = $assunto
        $mail.HTMLBody = $corpoHTML + $assinaturaHTML
        $mail.Save()
        $mail.Close(0)
    }
    finally {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($mail) | Out-Null
    }
}

function Open-Outlook {
    try {
        return (New-Object -ComObject Outlook.Application)
    } catch {
        $m = $_.Exception.Message
        $dica = ""
        if (Test-Elevado) {
            $dica = " O aplicativo está ELEVADO (Administrador): o COM do Outlook exige o mesmo nível de integridade. Reabra sem 'Executar como administrador'."
        } elseif ($m -match '80080005|Server execution failed|Falha na execução do servidor') {
            $dica = " Verifique se o Outlook não está em nível de privilégio divergente."
        }
        Send-Log 'aviso' ("Outlook COM indisponível. E-mails não serão gerados. " + $m + $dica)
        return $null
    }
}

function ConvertTo-FiltroAD { param([string]$s) return $s.Replace("'", "''") }

function Get-CpfDigitos {
    param([object[]]$textos)
    foreach ($t in $textos) {
        if ($t) {
            $m = [regex]::Match([string]$t, '(?<!\d)(\d{3}\.?\d{3}\.?\d{3}-?\d{2})(?!\d)')
            if ($m.Success) { return ($m.Value -replace '\D', '') }
        }
    }
    return ""
}

function Find-ContasAD {
    param([string]$matricula, [string]$logon)
    $props = @('Description', 'extensionAttribute3', 'employeeID', 'Enabled', 'UserPrincipalName', 'DistinguishedName', 'SamAccountName')
    $filtros = @()
    if ($matricula) {
        $m = ConvertTo-FiltroAD $matricula
        $filtros += "SamAccountName -eq '$m'"
        $filtros += "employeeID -eq '$m'"
    }
    if ($logon) {
        $l = ConvertTo-FiltroAD $logon
        $filtros += "UserPrincipalName -like '$l@*'"
        $filtros += "SamAccountName -eq '$l'"
    }
    if ($filtros.Count -eq 0) { return @() }
    $filtro = $filtros -join ' -or '
    try { $r = Get-ADUser -Filter $filtro -Properties ($props + 'CPF') -ErrorAction Stop }
    catch { $r = Get-ADUser -Filter $filtro -Properties $props -ErrorAction Stop }
    if ($r) { return @($r) } else { return @() }
}

# =============== ROTINAS DO WORKER ===============
function Invoke-LoteCriacao {
    $sucessos = 0
    $erros = 0
    $pastaDocs = [System.Environment]::GetFolderPath('MyDocuments')
    $logPath = [System.IO.Path]::Combine($pastaDocs, "Logs_Automacao")
    if (-not ([System.IO.Directory]::Exists($logPath))) { $null = [System.IO.Directory]::CreateDirectory($logPath) }
    $global:LogFile = [System.IO.Path]::Combine($logPath, "Log_Criacao_Lote.txt")

    try { Import-Module ActiveDirectory -ErrorAction Stop }
    catch {
        Send-Log 'erro' ("Módulo ActiveDirectory (RSAT) indisponível: " + $_.Exception.Message)
        Send-Fila @{ T = 'fim'; Texto = 'Criação abortada: módulo ActiveDirectory indisponível.' }
        return
    }

    $assinaturaHTML = Get-AssinaturaHtml
    $outlook = Open-Outlook
    $total = @($Snaps).Count
    $k = 0

    foreach ($reg in $Snaps) {
        if ($Ctl.Cancel) { Send-Log 'aviso' 'Processamento cancelado pelo operador.'; break }
        $k++
        $matricula = $reg.Matricula
        $nomeExibicaoAcento = $reg.NomeCompletoOrig
        Send-Fila @{ T = 'prog'; Atual = $k; Total = $total; Texto = ("Criando no AD ({0}/{1}): {2}" -f $k, $total, $matricula) }

        try {
            if ($reg.TipoProcesso -eq 'pendencia') {
                Send-Fila @{ T = 'res'; Id = $reg.Id; Resultado = 'Bloqueado'; Detalhe = 'Status PENDÊNCIA: criação bloqueada por segurança.' }
                Send-Log 'erro' "$matricula - $nomeExibicaoAcento : BLOQUEADO (status PENDÊNCIA). Resolva a inconsistência cadastral antes de processar."
                continue
            }
            if ([string]::IsNullOrWhiteSpace($matricula)) {
                $erros++
                Send-Fila @{ T = 'res'; Id = $reg.Id; Resultado = 'Falha'; Detalhe = 'Matrícula ausente (SamAccountName mandatório).' }
                Send-Log 'erro' "[SEM MATRÍCULA] $nomeExibicaoAcento - Matrícula obrigatória."
                continue
            }

            $perfil = Resolve-PerfilAD -reg $reg
            $g = Resolve-Gestor -reg $reg
            if ($g.Aviso) { Send-Log 'aviso' ("$matricula - " + $g.Aviso) }

            $r = New-UsuarioAD -reg $reg -perfil $perfil -managerDN $g.ManagerDN
            if ($r.Criado) {
                $sucessos++
                $rasc = ''
                if ($null -ne $outlook) {
                    try {
                        New-RascunhoOutlook -outlook $outlook -reg $reg -mailPrimario $perfil.MailPrimario -destinatarioFinal $g.Destinatario -assinaturaHTML $assinaturaHTML
                        $rasc = 'Salvo'
                    } catch {
                        $rasc = 'Falha'
                        Send-Log 'aviso' "$matricula - Usuário criado no AD, mas falha ao salvar rascunho no Outlook: $($_.Exception.Message)"
                    }
                } else { $rasc = 'Sem Outlook' }
                Send-Fila @{ T = 'res'; Id = $reg.Id; Resultado = 'Criado'; Detalhe = ("OU: {0} | CN: {1} | UPN: {2}" -f $r.OU, $r.CN, $perfil.UpnFinal); Rascunho = $rasc }
            } else {
                $erros++
                Send-Fila @{ T = 'res'; Id = $reg.Id; Resultado = 'Falha'; Detalhe = $r.Erro }
                Send-Log 'erro' "$matricula - $($r.Erro)"
            }
        } catch {
            $erros++
            Send-Fila @{ T = 'res'; Id = $reg.Id; Resultado = 'Falha'; Detalhe = $_.Exception.Message }
            Send-Log 'erro' "$matricula - Exceção inesperada: $($_.Exception.Message)"
        }
    }

    if ($null -ne $outlook) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null }
    Send-Fila @{ T = 'fim'; Texto = ("Processamento concluído! Sucessos: {0} | Erros: {1} | Log: {2}" -f $sucessos, $erros, $global:LogFile) }
}

function Invoke-LoteRascunho {
    $adOk = $true
    try { Import-Module ActiveDirectory -ErrorAction Stop } catch { $adOk = $false; Send-Log 'aviso' 'AD indisponível: rascunhos serão gerados sem destinatário gestor.' }
    $assinaturaHTML = Get-AssinaturaHtml
    $outlook = Open-Outlook
    if ($null -eq $outlook) { Send-Fila @{ T = 'fim'; Texto = 'Rascunhos não gerados: Outlook COM indisponível.' }; return }

    $total = @($Snaps).Count
    $k = 0; $ok = 0
    foreach ($reg in $Snaps) {
        if ($Ctl.Cancel) { Send-Log 'aviso' 'Geração de rascunhos cancelada.'; break }
        $k++
        Send-Fila @{ T = 'prog'; Atual = $k; Total = $total; Texto = ("Gerando rascunho ({0}/{1}): {2}" -f $k, $total, $reg.Matricula) }
        try {
            $perfil = Resolve-PerfilAD -reg $reg
            $dest = ""
            if ($adOk) {
                $g = Resolve-Gestor -reg $reg
                if ($g.Aviso) { Send-Log 'aviso' ("$($reg.Matricula) - " + $g.Aviso) }
                $dest = $g.Destinatario
            }
            New-RascunhoOutlook -outlook $outlook -reg $reg -mailPrimario $perfil.MailPrimario -destinatarioFinal $dest -assinaturaHTML $assinaturaHTML
            $ok++
            Send-Fila @{ T = 'rasc'; Id = $reg.Id; Rascunho = 'Salvo' }
        } catch {
            Send-Fila @{ T = 'rasc'; Id = $reg.Id; Rascunho = 'Falha' }
            Send-Log 'erro' "$($reg.Matricula) - Falha ao salvar rascunho: $($_.Exception.Message)"
        }
    }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null
    Send-Fila @{ T = 'fim'; Texto = ("Rascunhos processados: {0} de {1}." -f $ok, $total) }
}

function Invoke-LoteAuditoria {
    try { Import-Module ActiveDirectory -ErrorAction Stop }
    catch {
        Send-Log 'erro' ("Módulo ActiveDirectory (RSAT) indisponível: " + $_.Exception.Message)
        Send-Fila @{ T = 'fim'; Texto = 'Auditoria abortada: módulo ActiveDirectory indisponível.' }
        return
    }
    $total = @($Snaps).Count
    $i = 0; $mig = 0; $pend = 0; $livres = 0; $falhas = 0

    foreach ($reg in $Snaps) {
        if ($Ctl.Cancel) { Send-Log 'aviso' 'Auditoria cancelada pelo operador.'; break }
        $i++
        $mat = $reg.Matricula
        Send-Fila @{ T = 'prog'; Atual = $i; Total = $total; Texto = ("Validando no AD ({0}/{1}): {2}" -f $i, $total, $mat) }
        try {
            $baseLogon = if ($reg.SamAccount) { $reg.SamAccount.ToLower() } else { ((Limpar-Texto -texto $reg.PrimeiroNomeOrig) + "." + (Limpar-Texto -texto $reg.SobreNomeOrig)).ToLower() }
            $hits = @(Find-ContasAD -matricula $mat -logon $baseLogon | Where-Object { $_ })
            $cpfBase = ($reg.Cpf -replace '\D', '')

            if ($hits.Count -eq 0) {
                $livres++
                Send-Fila @{ T = 'aud'; Id = $reg.Id; Achou = $false; Status = ''; Nota = 'Identidade e logon livres no AD: apto à criação.' }
                continue
            }

            $idem = $null; $cpfADIdem = ""
            $primeiro = $hits[0]; $cpfADPrimeiro = Get-CpfDigitos -textos @($primeiro.Description, $primeiro.extensionAttribute3, $primeiro.CPF)
            foreach ($u in $hits) {
                $cpfAD = Get-CpfDigitos -textos @($u.Description, $u.extensionAttribute3, $u.CPF)
                if ($cpfBase -and $cpfAD -and ($cpfAD -eq $cpfBase)) { $idem = $u; break }
            }

            if ($idem) {
                $mig++
                $estado = if ($idem.Enabled) { 'CONTA ATIVA' } else { 'CONTA DESATIVADA' }
                $nota = "CPF idêntico (mesmo colaborador): migração/recontratação. Conta {0} ({1}) - {2}." -f $idem.SamAccountName, $idem.UserPrincipalName, $estado
                Send-Fila @{ T = 'aud'; Id = $reg.Id; Achou = $true; Status = 'migracao'; Nota = $nota }
                Send-Log 'info' "$mat - Migração confirmada por CPF idêntico ($estado)."
            } else {
                $pend++
                $u = $primeiro
                $estado = if ($u.Enabled) { 'CONTA ATIVA' } else { 'CONTA DESATIVADA' }
                $porMat = ($u.SamAccountName -ieq $mat) -or ($u.employeeID -ieq $mat)
                $conflito = if ($porMat) { "matrícula $mat" } else { "logon '$baseLogon'" }
                $motivoCpf = if (-not $cpfBase) { 'CPF da base está vazio' }
                             elseif (-not $cpfADPrimeiro) { 'CPF não localizado no AD para cruzamento' }
                             else { 'CPF diverge do registrado no AD' }
                $nota = "Homônimo/Conflito: colisão por {0} com a conta {1} (UPN {2}) - {3}. {4}. Criação bloqueada: ajuste o Logon." -f $conflito, $u.SamAccountName, $u.UserPrincipalName, $estado, $motivoCpf
                Send-Fila @{ T = 'aud'; Id = $reg.Id; Achou = $true; Status = 'pendencia'; Nota = $nota }
                Send-Log 'aviso' "$mat - DIVERGÊNCIA: $nota"
            }
        } catch {
            $falhas++
            Send-Log 'erro' "$mat - Falha na auditoria: $($_.Exception.Message)"
        }
    }
    Send-Fila @{ T = 'fim'; Texto = ("Auditoria concluída. Migrações: {0} | Divergências: {1} | Livres: {2} | Falhas: {3}" -f $mig, $pend, $livres, $falhas) }
}

$script:Lib = (@(
    'Limpar-Texto', 'Test-Elevado', 'Send-Fila', 'Send-Log', 'Write-LogArquivo', 'Get-AssinaturaHtml', 'Resolve-PerfilAD',
    'Resolve-Gestor', 'New-UsuarioAD', 'Get-CorpoEmail', 'New-RascunhoOutlook', 'Open-Outlook', 'ConvertTo-FiltroAD',
    'Get-CpfDigitos', 'Find-ContasAD', 'Invoke-LoteCriacao', 'Invoke-LoteRascunho', 'Invoke-LoteAuditoria'
) | ForEach-Object { "function $_ {`n" + (Get-Item "function:$_").ScriptBlock.ToString() + "`n}" }) -join "`n"

# =====================================================================================
# 3. ESTADO E APOIO DA INTERFACE
# =====================================================================================
$script:AppDir = Join-Path ([System.Environment]::GetFolderPath('ApplicationData')) 'IAM_Provisionamento'
$script:BaseFile = Join-Path $script:AppDir 'base_usuarios.json'
$script:Dados = New-Object 'System.Collections.ObjectModel.ObservableCollection[Registro]'
$script:Fila = New-Object 'System.Collections.Concurrent.ConcurrentQueue[object]'
$script:Ctl = [hashtable]::Synchronized(@{ Cancel = $false })
$script:Worker = $null
$script:EditId = $null
$script:LogonManual = $false
$script:Atualizando = $false
$script:RefreshPendente = $false
$script:RefreshForcar = $false
$script:SalvarEm = $null
$script:Filtro = New-Object FiltroRegistro
$script:CamposTexto = @('Status', 'PrimeiroNome', 'Sobrenome', 'NomeCompleto', 'Logon', 'Cargo', 'Depto', 'Empresa', 'Matricula', 'Rg', 'Cpf', 'Endereco', 'Gerente', 'EmailLic', 'Perfil', 'Malha', 'CloudTimestamp', 'Chamado', 'CentroCusto', 'Notas', 'Resultado', 'Detalhe', 'Rascunho')
$script:CamposBool = @('Ok', 'Mfa', 'Licenca')

function Format-NomeProprio {
    param([string]$s)
    if ([string]::IsNullOrWhiteSpace($s)) { return "" }
    $prep = @('de', 'da', 'do', 'dos', 'das', 'e', 'del', 'di', 'van', 'von')
    $partes = $s.Trim().ToLower() -split '\s+'
    $out = for ($i = 0; $i -lt $partes.Count; $i++) {
        $p = $partes[$i]
        if ($p.Length -eq 0) { continue }
        if ($i -gt 0 -and ($prep -contains $p)) { $p } else { $p.Substring(0, 1).ToUpper() + $p.Substring(1) }
    }
    return ($out -join ' ')
}

function Format-Cpf {
    param([string]$s)
    $v = ($s -replace '\D', '')
    if ($v.Length -gt 11) { $v = $v.Substring(0, 11) }
    if ($v.Length -gt 9) { return ('{0}.{1}.{2}-{3}' -f $v.Substring(0, 3), $v.Substring(3, 3), $v.Substring(6, 3), $v.Substring(9)) }
    if ($v.Length -gt 6) { return ('{0}.{1}.{2}' -f $v.Substring(0, 3), $v.Substring(3, 3), $v.Substring(6)) }
    if ($v.Length -gt 3) { return ('{0}.{1}' -f $v.Substring(0, 3), $v.Substring(3)) }
    return $v
}

function Get-TimestampUtc { return ([DateTime]::UtcNow.ToString('yyyyMMdd') + '080000.0Z') }

function Get-Faltantes {
    param($r)
    $f = @()
    if (-not $r.NomeCompleto.Trim()) { $f += 'Nome Completo' }
    if (-not $r.Matricula.Trim()) { $f += 'Matrícula' }
    if (-not $r.Cargo.Trim()) { $f += 'Cargo' }
    if (-not $r.Depto.Trim()) { $f += 'Departamento' }
    if (-not $r.Empresa.Trim()) { $f += 'Empresa' }
    if (-not $r.Rg.Trim()) { $f += 'RG' }
    if (-not $r.Cpf.Trim()) { $f += 'CPF' }
    if (-not $r.Endereco.Trim()) { $f += 'Endereço' }
    if (-not $r.Gerente.Trim()) { $f += 'Gerente' }
    return $f
}

function Set-PendenciaSeFaltar {
    param($r)
    $f = @(Get-Faltantes $r)
    if ($f.Count -gt 0) {
        $r.Status = 'pendencia'
        $aviso = "[Falta: " + ($f -join ', ') + "]"
        if ($r.Notas -notlike "*$aviso*") { $r.Notas = if ($r.Notas) { "$aviso $($r.Notas)" } else { $aviso } }
    }
    return $f.Count
}

function ConvertTo-StatusValido {
    param([string]$s)
    $v = if ($s) { $s.Trim().ToLower() } else { '' }
    if ($v -match 'pend|diverg') { return 'pendencia' }
    if ($v -match 'migra') { return 'migracao' }
    return 'nova'
}

function Set-NotaAD {
    param([string]$notas, [string]$msg)
    $limpa = [regex]::Replace([string]$notas, '\[AD[^\]]*\][^¦]*¦\s*', '').Trim()
    $nova = "[AD {0}] {1} ¦" -f (Get-Date -Format 'dd/MM HH:mm'), $msg
    if ($limpa) { return "$nova $limpa" } else { return $nova }
}

function ConvertTo-Snap {
    param($r)
    $t = { param($x) if ($null -ne $x) { ([string]$x).Trim() } else { '' } }
    return [pscustomobject]@{
        Id = $r.Id
        PrimeiroNomeOrig = (& $t $r.PrimeiroNome); SobreNomeOrig = (& $t $r.Sobrenome); NomeCompletoOrig = (& $t $r.NomeCompleto)
        SamAccount = (& $t $r.Logon); Cargo = (& $t $r.Cargo); Depto = (& $t $r.Depto); Empresa = (& $t $r.Empresa)
        Matricula = (& $t $r.Matricula); Rg = (& $t $r.Rg); Cpf = (& $t $r.Cpf); EnderecoCom = (& $t $r.Endereco)
        Gerente = (& $t $r.Gerente); TemEmail = ((& $t $r.EmailLic).ToLower())
        Opcao = $(if ($r.Perfil) { (& $t $r.Perfil) } else { '2' })
        MalhaOp = $(if ($r.Malha) { ((& $t $r.Malha).ToLower()) } else { '1' })
        CloudTimestamp = (& $t $r.CloudTimestamp); TipoProcesso = $r.Status; Chamado = (& $t $r.Chamado)
    }
}

function Save-Base {
    try {
        if (-not (Test-Path $script:AppDir)) { $null = New-Item -ItemType Directory -Path $script:AppDir -Force }
        $arr = @(foreach ($r in $script:Dados) {
            $o = [ordered]@{}
            foreach ($c in $script:CamposTexto) { $o[$c] = [string]$r.$c }
            foreach ($c in $script:CamposBool) { $o[$c] = [bool]$r.$c }
            [pscustomobject]$o
        })
        $tmp = $script:BaseFile + '.tmp'
        ConvertTo-Json -InputObject $arr -Depth 3 | Set-Content -Path $tmp -Encoding UTF8
        Move-Item -Path $tmp -Destination $script:BaseFile -Force
    } catch { }
}

function Load-Base {
    try {
        if (Test-Path $script:BaseFile) {
            $itens = @(Get-Content -Path $script:BaseFile -Raw -Encoding UTF8 | ConvertFrom-Json)
            foreach ($o in $itens) {
                if (-not $o) { continue }
                $r = New-Object Registro
                foreach ($c in $script:CamposTexto) { if ($null -ne $o.$c) { $r.$c = [string]$o.$c } }
                foreach ($c in $script:CamposBool) { if ($null -ne $o.$c) { $r.$c = [bool]$o.$c } }
                # FILTRO RIGOROSO: descarta linhas sem matrícula ou nome válido
                if (-not [string]::IsNullOrWhiteSpace($r.Matricula) -and -not [string]::IsNullOrWhiteSpace($r.NomeCompleto)) {
                    $script:Dados.Add($r)
                }
            }
        }
    } catch { }
}

function Get-ColLetra {
    param([int]$n)
    $s = ''
    while ($n -gt 0) {
        $m = ($n - 1) % 26
        $s = ([string][char](65 + $m)) + $s
        $n = [int][math]::Floor(($n - 1) / 26)
    }
    return $s
}

function ConvertFrom-ColLetra {
    param([string]$ref)
    $l = ($ref -replace '\d', '')
    $n = 0
    foreach ($ch in $l.ToCharArray()) { $n = $n * 26 + ([int][char]::ToUpper($ch) - 64) }
    return $n
}

function Export-XlsxSimples {
    param([string]$Caminho, [string[]]$Cabecalho, [object[]]$Linhas)
    if (Test-Path $Caminho) { Remove-Item $Caminho -Force }
    $fs = [System.IO.File]::Open($Caminho, [System.IO.FileMode]::Create)
    $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        $addEntry = {
            param($nome, $conteudo)
            $e = $zip.CreateEntry($nome)
            $sw = New-Object System.IO.StreamWriter($e.Open(), (New-Object System.Text.UTF8Encoding($false)))
            $sw.Write($conteudo)
            $sw.Dispose()
        }
        $esc = {
            param($t)
            $s = if ($null -eq $t) { '' } else { [string]$t }
            $s = [regex]::Replace($s, '[\x00-\x08\x0B\x0C\x0E-\x1F]', '')
            return [System.Security.SecurityElement]::Escape($s)
        }
        $hdr = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        & $addEntry '[Content_Types].xml' ($hdr + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')
        & $addEntry '_rels/.rels' ($hdr + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        & $addEntry 'xl/workbook.xml' ($hdr + '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Base_Auditoria_IAM" sheetId="1" r:id="rId1"/></sheets></workbook>')
        & $addEntry 'xl/_rels/workbook.xml.rels' ($hdr + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>')

        $sb = New-Object System.Text.StringBuilder
        [void]$sb.Append($hdr + '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>')
        $todas = New-Object System.Collections.ArrayList
        [void]$todas.Add($Cabecalho)
        foreach ($l in $Linhas) { [void]$todas.Add($l) }
        $rn = 0
        foreach ($linha in $todas) {
            $rn++
            [void]$sb.Append('<row r="' + $rn + '">')
            for ($c = 0; $c -lt $linha.Count; $c++) {
                $ref = (Get-ColLetra ($c + 1)) + $rn
                [void]$sb.Append('<c r="' + $ref + '" t="inlineStr"><is><t xml:space="preserve">' + (& $esc $linha[$c]) + '</t></is></c>')
            }
            [void]$sb.Append('</row>')
        }
        [void]$sb.Append('</sheetData></worksheet>')
        & $addEntry 'xl/worksheets/sheet1.xml' $sb.ToString()
    }
    finally {
        $zip.Dispose()
        $fs.Dispose()
    }
}

function Import-XlsxSimples {
    param([string]$Caminho)
    $zip = [System.IO.Compression.ZipFile]::OpenRead($Caminho)
    try {
        $lerEntry = {
            param($entry)
            $sr = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
            $t = $sr.ReadToEnd()
            $sr.Dispose()
            return $t
        }
        $ss = New-Object System.Collections.ArrayList
        $eSS = $zip.GetEntry('xl/sharedStrings.xml')
        if ($eSS) {
            $xs = New-Object System.Xml.XmlDocument
            $xs.LoadXml((& $lerEntry $eSS))
            foreach ($si in $xs.SelectNodes('//*[local-name()="si"]')) { [void]$ss.Add($si.InnerText) }
        }
        $planilha = $zip.Entries | Where-Object { $_.FullName -match '^xl/worksheets/sheet\d+\.xml$' } | Sort-Object FullName | Select-Object -First 1
        if (-not $planilha) { throw 'Nenhuma planilha encontrada no arquivo .xlsx.' }
        $xml = New-Object System.Xml.XmlDocument
        $xml.LoadXml((& $lerEntry $planilha))

        $matriz = New-Object System.Collections.ArrayList
        foreach ($row in $xml.SelectNodes('//*[local-name()="sheetData"]/*[local-name()="row"]')) {
            $cells = @{}
            foreach ($c in $row.SelectNodes('*[local-name()="c"]')) {
                $idx = ConvertFrom-ColLetra $c.GetAttribute('r')
                $t = $c.GetAttribute('t')
                $txt = ''
                if ($t -eq 'inlineStr') { $txt = $c.InnerText }
                else {
                    $v = $c.SelectSingleNode('*[local-name()="v"]')
                    if ($v) { if ($t -eq 's') { $txt = [string]$ss[[int]$v.InnerText] } else { $txt = $v.InnerText } }
                }
                $cells[$idx] = $txt
            }
            [void]$matriz.Add($cells)
        }
        $res = New-Object System.Collections.ArrayList
        if ($matriz.Count -lt 2) { return @() }
        $cab = $matriz[0]
        for ($i = 1; $i -lt $matriz.Count; $i++) {
            $h = @{}
            foreach ($k in $cab.Keys) {
                $nome = ([string]$cab[$k]).Trim()
                if ($nome) { $h[$nome] = $matriz[$i][$k] }
            }
            [void]$res.Add($h)
        }
        return $res.ToArray()
    }
    finally { $zip.Dispose() }
}

function Import-CsvFlex {
    param([string]$Caminho)
    $bytes = [System.IO.File]::ReadAllBytes($Caminho)
    try { $txt = (New-Object System.Text.UTF8Encoding($false, $true)).GetString($bytes) }
    catch { $txt = [System.Text.Encoding]::GetEncoding(1252).GetString($bytes) }
    $txt = $txt.TrimStart([char]0xFEFF)
    $linhas = $txt -split "\r?\n"
    $primeira = $linhas[0]
    $d = ';'
    if (($primeira -split ',').Count -gt ($primeira -split ';').Count) { $d = ',' }
    if ($primeira -match "`t") { $d = "`t" }
    $out = New-Object System.Collections.ArrayList
    foreach ($o in ($linhas | Where-Object { $_.Trim() -ne '' } | ConvertFrom-Csv -Delimiter $d)) {
        $h = @{}
        foreach ($p in $o.PSObject.Properties) { $h[$p.Name.Trim()] = $p.Value }
        [void]$out.Add($h)
    }
    return $out.ToArray()
}

function Get-Val {
    param($h, [string[]]$nomes)
    foreach ($n in $nomes) {
        if ($h.ContainsKey($n) -and $null -ne $h[$n] -and ([string]$h[$n]).Trim() -ne '') { return ([string]$h[$n]).Trim() }
    }
    return ''
}

function Test-Sim { param([string]$v) return (@('sim', 'true', '1', 's', 'x', 'yes') -contains $v.Trim().ToLower()) }

# Converte registro validando se não é linha vazia/incompleta (Requisito 3)
function ConvertFrom-Hash {
    param($h)
    $nome = Format-NomeProprio (Get-Val $h @('C (Nome Completo)', 'NomeCompleto', 'Nome Completo', 'Nome', 'C'))
    $mat = Get-Val $h @('H (Matrícula)', 'H (Matricula)', 'Matricula', 'Matrícula', 'H')

    # DESCARTA SILENCIOSAMENTE LINHAS SEM NOME OU MATRÍCULA
    if ([string]::IsNullOrWhiteSpace($nome) -or [string]::IsNullOrWhiteSpace($mat)) {
        return $null
    }

    $r = New-Object Registro
    $partes = $nome -split '\s+'
    $prim = Get-Val $h @('A (Nome)', 'PrimeiroNome', 'A')
    $sob = Get-Val $h @('B (Sobrenome)', 'Sobrenome', 'B')
    $prim = if ($prim) { Format-NomeProprio $prim } elseif ($partes.Count -gt 0) { $partes[0] } else { '' }
    $sob = if ($sob) { Format-NomeProprio $sob } elseif ($partes.Count -gt 1) { $partes[$partes.Count - 1] } else { $prim }
    $r.NomeCompleto = $nome; $r.PrimeiroNome = $prim; $r.Sobrenome = $sob; $r.Matricula = $mat

    $logon = (Get-Val $h @('D (Logon)', 'Logon', 'Logon (UPN)', 'SamAccount', 'D')).ToLower()
    if (-not $logon -and $prim -and $sob) { $logon = ([Registro]::Limpar($prim) + '.' + [Registro]::Limpar($sob)).ToLower() }
    $r.Logon = $logon
    $r.Cargo = Get-Val $h @('E (Cargo)', 'Cargo', 'E')
    $r.Depto = Get-Val $h @('F (Depto)', 'Depto', 'Departamento', 'F')
    $r.Empresa = Get-Val $h @('G (Empresa)', 'Empresa', 'G')
    $r.Rg = Get-Val $h @('I (RG)', 'RG', 'I')
    $cpf = Get-Val $h @('J (CPF)', 'CPF', 'J')
    $dig = $cpf -replace '\D', ''
    if ($dig -and $dig.Length -ge 9 -and $dig.Length -lt 11) { $dig = $dig.PadLeft(11, '0') }
    $r.Cpf = if ($dig.Length -eq 11) { Format-Cpf $dig } else { $cpf }
    $r.Endereco = Get-Val $h @('K (Endereço)', 'K (Endereco)', 'Endereco', 'Endereço', 'K')
    $r.Gerente = Get-Val $h @('L (Gerente)', 'Gerente', 'L')
    $em = (Get-Val $h @('M (Email)', 'Email(s/n)', 'Email', 'M')).ToLower()
    $r.EmailLic = if ($em) { $em } else { 's' }
    $pf = Get-Val $h @('N (Perfil)', 'Perfil(1,2,3,4)', 'Perfil', 'N')
    $r.Perfil = if ($pf) { $pf.Substring(0, 1) } else { '2' }
    $mr = (Get-Val $h @('O (Malha)', 'Malha', 'O')).ToLower()
    $r.Malha = if ($mr -match 'especial' -or $mr -eq '3') { '3' } elseif ($mr -match 'sul' -or $mr -eq '2') { '2' } elseif ($mr -eq '0' -or $mr -match 'n[aã]o') { '0' } else { '1' }
    $ts = Get-Val $h @('P (Timestamp)', 'CloudTimestamp', 'P')
    $r.CloudTimestamp = if ($ts) { $ts } else { Get-TimestampUtc }
    $r.Status = ConvertTo-StatusValido (Get-Val $h @('tipoProcesso', 'Tipo de Processo', 'TipoProcesso', 'Processo', 'Status / Processo', 'Status'))
    $r.Chamado = Get-Val $h @('R (Chamado)', 'Chamado', 'RITM', 'R')
    $r.CentroCusto = Get-Val $h @('Centro de Custo', 'CentroDeCusto', 'CC')
    $r.Notas = Get-Val $h @('Notas_Lembretes', 'Notas', 'Lembrete / Notas')
    $r.Ok = Test-Sim (Get-Val $h @('Concluido', 'Concluído', 'OK?', 'OK'))
    $r.Mfa = Test-Sim (Get-Val $h @('MFA', 'MFA?'))
    $r.Licenca = Test-Sim (Get-Val $h @('Licenca', 'Licença', 'Licença?'))
    [void](Set-PendenciaSeFaltar $r)
    return $r
}

# Converte linha colada via TAB validando se não é linha vazia/incompleta (Requisito 3)
function ConvertFrom-LinhaTab {
    param([string]$linha)
    $d = $linha -split "`t"
    $g = { param($i) if ($d.Count -gt $i -and $d[$i]) { $d[$i].Trim() } else { '' } }
    
    $nome = Format-NomeProprio (& $g 2)
    $mat  = & $g 7

    # DESCARTA SILENCIOSAMENTE LINHAS SEM DADOS ESSENCIAIS
    if ([string]::IsNullOrWhiteSpace($nome) -or [string]::IsNullOrWhiteSpace($mat)) {
        return $null
    }

    $r = New-Object Registro
    $r.PrimeiroNome = Format-NomeProprio (& $g 0)
    $r.Sobrenome    = Format-NomeProprio (& $g 1)
    $r.NomeCompleto = $nome
    $r.Logon        = (& $g 3)
    $r.Cargo        = (& $g 4)
    $r.Depto        = (& $g 5)
    $r.Empresa      = (& $g 6)
    $r.Matricula    = $mat
    $r.Rg           = (& $g 8)
    
    $cpf = (& $g 9)
    $dig = $cpf -replace '\D', ''
    $r.Cpf = if ($dig.Length -eq 11) { Format-Cpf $dig } else { $cpf }
    
    $r.Endereco = (& $g 10)
    $r.Gerente  = (& $g 11)
    $em = (& $g 12).ToLower()
    $r.EmailLic = if ($em) { $em } else { 'n' }
    $op = (& $g 13)
    $r.Perfil = if ($op) { $op } else { '2' }
    $mr = (& $g 14).ToLower()
    $r.Malha = if ($mr -match 'especial' -or $mr -eq '3') { '3' } elseif ($mr -match 'sul' -or $mr -eq '2') { '2' } elseif ($mr -eq '0') { '0' } else { '1' }
    $ts = (& $g 15)
    $r.CloudTimestamp = if ($ts) { $ts } else { Get-TimestampUtc }
    $r.Status = ConvertTo-StatusValido (& $g 16)
    $r.Chamado = (& $g 17)
    
    if (-not $r.Logon -and $r.PrimeiroNome -and $r.Sobrenome) { 
        $r.Logon = ([Registro]::Limpar($r.PrimeiroNome) + '.' + [Registro]::Limpar($r.Sobrenome)).ToLower() 
    }
    [void](Set-PendenciaSeFaltar $r)
    return $r
}

function Copy-RegistroDados {
    param($dest, $src)
    foreach ($c in @('Status', 'PrimeiroNome', 'Sobrenome', 'NomeCompleto', 'Logon', 'Cargo', 'Depto', 'Empresa', 'Matricula', 'Rg', 'Cpf', 'Endereco', 'Gerente', 'EmailLic', 'Perfil', 'Malha', 'CloudTimestamp', 'Chamado', 'CentroCusto', 'Notas')) { $dest.$c = $src.$c }
    $dest.Ok = ($dest.Ok -or $src.Ok); $dest.Mfa = ($dest.Mfa -or $src.Mfa); $dest.Licenca = ($dest.Licenca -or $src.Licenca)
}

# =====================================================================================
# 4. INTERFACE GRÁFICA (XAML COM DATAGRID ROWSTYLE DE ALTO CONTRASTE)
# =====================================================================================
$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Central de Provisionamento &amp; Auditoria de Acessos - IAM"
        Width="1440" Height="900" MinWidth="1200" MinHeight="700"
        WindowStartupLocation="CenterScreen" Background="#F8FAFC"
        FontFamily="Segoe UI" FontSize="12" UseLayoutRounding="True">
  <Window.Resources>
    <Style x:Key="Btn" TargetType="Button">
      <Setter Property="Foreground" Value="White"/><Setter Property="Background" Value="#002060"/>
      <Setter Property="FontWeight" Value="SemiBold"/><Setter Property="FontSize" Value="11.5"/>
      <Setter Property="Padding" Value="10,5"/><Setter Property="Margin" Value="0,0,6,4"/>
      <Setter Property="Cursor" Value="Hand"/><Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Template"><Setter.Value>
        <ControlTemplate TargetType="Button">
          <Border x:Name="bd" Background="{TemplateBinding Background}" CornerRadius="4" Padding="{TemplateBinding Padding}">
            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
          </Border>
          <ControlTemplate.Triggers>
            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="Opacity" Value="0.88"/></Trigger>
            <Trigger Property="IsPressed" Value="True"><Setter TargetName="bd" Property="Opacity" Value="0.72"/></Trigger>
            <Trigger Property="IsEnabled" Value="False"><Setter TargetName="bd" Property="Opacity" Value="0.4"/></Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value></Setter>
    </Style>
    <Style x:Key="Lbl" TargetType="TextBlock">
      <Setter Property="FontSize" Value="10"/><Setter Property="FontWeight" Value="Bold"/>
      <Setter Property="Foreground" Value="#64748B"/><Setter Property="Margin" Value="0,3,0,1"/>
    </Style>
    <Style TargetType="DataGridColumnHeader">
      <Setter Property="Background" Value="#002060"/><Setter Property="Foreground" Value="White"/>
      <Setter Property="FontWeight" Value="SemiBold"/><Setter Property="Padding" Value="6,5"/>
      <Setter Property="BorderBrush" Value="#1E3A8A"/><Setter Property="BorderThickness" Value="0,0,1,0"/>
    </Style>
    <Style x:Key="Cartao" TargetType="Border">
      <Setter Property="Background" Value="White"/><Setter Property="BorderBrush" Value="#CBD5E1"/>
      <Setter Property="BorderThickness" Value="1"/><Setter Property="CornerRadius" Value="5"/>
      <Setter Property="Margin" Value="2"/><Setter Property="Padding" Value="8,5"/>
    </Style>
    <Style x:Key="Painel" TargetType="Border">
      <Setter Property="Background" Value="White"/><Setter Property="BorderBrush" Value="#CBD5E1"/>
      <Setter Property="BorderThickness" Value="1"/><Setter Property="CornerRadius" Value="6"/>
      <Setter Property="Margin" Value="0,0,0,8"/><Setter Property="Padding" Value="10,8"/>
    </Style>
  </Window.Resources>

  <Grid>
    <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>

    <!-- CABEÇALHO -->
    <Border Grid.Row="0" Padding="14,7">
      <Border.Background><LinearGradientBrush StartPoint="0,0" EndPoint="1,1"><GradientStop Color="#002060" Offset="0"/><GradientStop Color="#001030" Offset="1"/></LinearGradientBrush></Border.Background>
      <Grid>
        <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
          <TextBlock Text="🛡️ Central de Provisionamento &amp; Auditoria de Acessos" Foreground="White" FontSize="15" FontWeight="SemiBold"/>
          <Border Background="#0070C0" CornerRadius="10" Padding="8,1" Margin="10,0,0,0" VerticalAlignment="Center">
            <TextBlock x:Name="txtContador" Text="0 Registros" Foreground="White" FontSize="11" FontWeight="SemiBold"/>
          </Border>
        </StackPanel>
        <StackPanel Grid.Column="1" Orientation="Horizontal">
          <Button x:Name="btnImportar" Style="{StaticResource Btn}" Background="#475569" Content="📤 Importar (.xlsx/.csv)"/>
          <Button x:Name="btnExportar" Style="{StaticResource Btn}" Background="#16A34A" Content="📥 Exportar (.xlsx/.csv)"/>
          <Button x:Name="btnMfa" Style="{StaticResource Btn}" Background="#0284C7" Content="🔐 Copiar E-mails p/ MFA"/>
          <Button x:Name="btnLimparOk" Style="{StaticResource Btn}" Background="#DC2626" Content="🗑️ Limpar Concluídos" Margin="0,0,0,4"/>
        </StackPanel>
      </Grid>
    </Border>

    <!-- CORPO -->
    <Grid Grid.Row="1" Margin="10">
      <Grid.ColumnDefinitions><ColumnDefinition Width="360"/><ColumnDefinition Width="10"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>

      <!-- PAINEL ESQUERDO -->
      <ScrollViewer Grid.Column="0" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
        <StackPanel Margin="0,0,6,0">
          <StackPanel.Resources>
            <Style TargetType="TextBox"><Setter Property="Height" Value="26"/><Setter Property="Padding" Value="4,3"/><Setter Property="BorderBrush" Value="#CBD5E1"/><Setter Property="VerticalContentAlignment" Value="Center"/></Style>
            <Style TargetType="ComboBox"><Setter Property="Height" Value="26"/><Setter Property="VerticalContentAlignment" Value="Center"/></Style>
          </StackPanel.Resources>

          <Border Style="{StaticResource Painel}">
            <StackPanel>
              <TextBlock Text="📝 FORMULÁRIO DE PROVISIONAMENTO" FontWeight="Bold" Foreground="#002060" FontSize="12" Margin="0,0,0,4"/>
              <TextBlock Text="C - NOME COMPLETO" Style="{StaticResource Lbl}" Foreground="#0070C0"/>
              <TextBox x:Name="txtNome"/>
              <UniformGrid Columns="2" Margin="0,2,0,0">
                <StackPanel Margin="0,0,3,0"><TextBlock Text="A - PRIMEIRO NOME" Style="{StaticResource Lbl}"/><TextBox x:Name="txtPrimeiro"/></StackPanel>
                <StackPanel Margin="3,0,0,0"><TextBlock Text="B - SOBRENOME" Style="{StaticResource Lbl}"/><TextBox x:Name="txtSobrenome"/></StackPanel>
                <StackPanel Margin="0,0,3,0"><TextBlock Text="D - LOGON (BASE SAMACCOUNT/UPN)" Style="{StaticResource Lbl}" Foreground="#B45309"/><TextBox x:Name="txtLogon"/></StackPanel>
                <StackPanel Margin="3,0,0,0"><TextBlock Text="H - MATRÍCULA" Style="{StaticResource Lbl}"/><TextBox x:Name="txtMatricula" ToolTip="Identificador do colaborador"/></StackPanel>
                <StackPanel Margin="0,0,3,0"><TextBlock Text="E - CARGO" Style="{StaticResource Lbl}"/><TextBox x:Name="txtCargo"/></StackPanel>
                <StackPanel Margin="3,0,0,0"><TextBlock Text="F - DEPTO" Style="{StaticResource Lbl}"/><TextBox x:Name="txtDepto"/></StackPanel>
                <StackPanel Margin="0,0,3,0"><TextBlock Text="G - EMPRESA" Style="{StaticResource Lbl}"/><TextBox x:Name="txtEmpresa"/></StackPanel>
                <StackPanel Margin="3,0,0,0"><TextBlock Text="I - RG" Style="{StaticResource Lbl}"/><TextBox x:Name="txtRg"/></StackPanel>
                <StackPanel Margin="0,0,3,0"><TextBlock Text="J - CPF (BASE DA SENHA)" Style="{StaticResource Lbl}"/><TextBox x:Name="txtCpf"/></StackPanel>
                <StackPanel Margin="3,0,0,0"><TextBlock Text="K - ENDEREÇO (OFFICE)" Style="{StaticResource Lbl}"/><TextBox x:Name="txtEndereco"/></StackPanel>
                <StackPanel Margin="0,0,3,0"><TextBlock Text="L - GERENTE (MATRÍCULA)" Style="{StaticResource Lbl}"/><TextBox x:Name="txtGerente"/></StackPanel>
                <StackPanel Margin="3,0,0,0"><TextBlock Text="M - E-MAIL LICENCIADO?" Style="{StaticResource Lbl}"/>
                  <ComboBox x:Name="cmbEmailLic" SelectedValuePath="Tag" SelectedIndex="0">
                    <ComboBoxItem Tag="s" Content="s (Sim - Licenciar)"/><ComboBoxItem Tag="n" Content="n (Não)"/></ComboBox></StackPanel>
                <StackPanel Margin="0,0,3,0"><TextBlock Text="N - PERFIL" Style="{StaticResource Lbl}"/>
                  <ComboBox x:Name="cmbPerfil" SelectedValuePath="Tag" SelectedIndex="0">
                    <ComboBoxItem Tag="2" Content="2 (Interno / Padrão)"/><ComboBoxItem Tag="1" Content="1 (Terceiro)"/><ComboBoxItem Tag="3" Content="3 (Filial Especial)"/><ComboBoxItem Tag="4" Content="4 (Parceiro)"/></ComboBox></StackPanel>
                <StackPanel Margin="3,0,0,0"><TextBlock Text="O - REGIÃO / SEGMENTO" Style="{StaticResource Lbl}"/>
                  <ComboBox x:Name="cmbMalha" SelectedValuePath="Tag" SelectedIndex="0">
                    <ComboBoxItem Tag="1" Content="1 (Região Norte)"/><ComboBoxItem Tag="2" Content="2 (Região Sul)"/><ComboBoxItem Tag="3" Content="3 (Especial)"/><ComboBoxItem Tag="0" Content="0 (Não Aplicável)"/></ComboBox></StackPanel>
                <StackPanel Margin="0,0,3,0"><TextBlock Text="P - CLOUD TIMESTAMP (UTC)" Style="{StaticResource Lbl}"/>
                  <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                    <TextBox x:Name="txtCloud"/>
                    <Button x:Name="btnCloudUtc" Grid.Column="1" Style="{StaticResource Btn}" Padding="6,3" Margin="3,0,0,0" Height="26" Background="#0070C0" Content="UTC" ToolTip="Gerar timestamp UTC (AAAAMMDD080000.0Z)"/></Grid></StackPanel>
                <StackPanel Margin="3,0,0,0"><TextBlock Text="Q - STATUS / PROCESSO" Style="{StaticResource Lbl}"/>
                  <ComboBox x:Name="cmbStatus" SelectedValuePath="Tag" SelectedIndex="0">
                    <ComboBoxItem Tag="nova" Content="nova"/><ComboBoxItem Tag="migracao" Content="migracao"/><ComboBoxItem Tag="pendencia" Content="🛑 pendencia"/></ComboBox></StackPanel>
                <StackPanel Margin="0,0,3,0"><TextBlock Text="R - CHAMADO (ID)" Style="{StaticResource Lbl}"/><TextBox x:Name="txtChamado"/></StackPanel>
                <StackPanel Margin="3,0,0,0"><TextBlock Text="CENTRO DE CUSTO" Style="{StaticResource Lbl}" Foreground="#0369A1"/><TextBox x:Name="txtCC"/></StackPanel>
              </UniformGrid>
              <TextBlock Text="📌 LEMBRETES / NOTAS / PENDÊNCIAS" Style="{StaticResource Lbl}"/>
              <TextBox x:Name="txtNotas" Height="54" TextWrapping="Wrap" AcceptsReturn="True" VerticalScrollBarVisibility="Auto" VerticalContentAlignment="Top"/>
              <TextBlock x:Name="txtDup" Visibility="Collapsed" Foreground="#DC2626" FontWeight="SemiBold" FontSize="11" TextWrapping="Wrap" Margin="0,4,0,0"/>
              <WrapPanel Margin="0,8,0,0">
                <Button x:Name="btnAdd" Style="{StaticResource Btn}" Content="➕ Adicionar à Fila"/>
                <Button x:Name="btnLimparForm" Style="{StaticResource Btn}" Background="#E2E8F0" Foreground="#1E293B" Content="🔄 Limpar Formulário"/>
              </WrapPanel>
              <WrapPanel x:Name="pnlEdicao" Visibility="Collapsed">
                <Button x:Name="btnSalvarEd" Style="{StaticResource Btn}" Background="#16A34A" Content="💾 Salvar Alterações"/>
                <Button x:Name="btnCancelarEd" Style="{StaticResource Btn}" Background="#D97706" Content="✖ Cancelar Edição"/>
              </WrapPanel>
              <Button x:Name="btnColar" Style="{StaticResource Btn}" Background="#0284C7" HorizontalAlignment="Stretch" Content="📋 Colar Lote do Excel (18 colunas via TAB)"/>
            </StackPanel>
          </Border>

          <Border Style="{StaticResource Painel}">
            <StackPanel>
              <TextBlock Text="📊 DASHBOARD OPERACIONAL" FontWeight="Bold" Foreground="#002060" FontSize="12" Margin="0,0,0,4"/>
              <UniformGrid Columns="3">
                <Border Style="{StaticResource Cartao}"><StackPanel><TextBlock x:Name="numTotal" Text="0" FontSize="20" FontWeight="Bold" Foreground="#002060"/><TextBlock Text="Total inserido" FontSize="10" Foreground="#64748B"/></StackPanel></Border>
                <Border Style="{StaticResource Cartao}"><StackPanel><TextBlock x:Name="numSel" Text="0" FontSize="20" FontWeight="Bold" Foreground="#0070C0"/><TextBlock Text="Selecionados [✔]" FontSize="10" Foreground="#64748B"/></StackPanel></Border>
                <Border Style="{StaticResource Cartao}"><StackPanel><TextBlock x:Name="numCriados" Text="0" FontSize="20" FontWeight="Bold" Foreground="#16A34A"/><TextBlock Text="Criados c/ sucesso" FontSize="10" Foreground="#64748B"/></StackPanel></Border>
                <Border Style="{StaticResource Cartao}"><StackPanel><TextBlock x:Name="numPend" Text="0" FontSize="20" FontWeight="Bold" Foreground="#D97706"/><TextBlock Text="Pend./Divergências" FontSize="10" Foreground="#64748B"/></StackPanel></Border>
                <Border Style="{StaticResource Cartao}"><StackPanel><TextBlock x:Name="numFalhas" Text="0" FontSize="20" FontWeight="Bold" Foreground="#DC2626"/><TextBlock Text="Falhas" FontSize="10" Foreground="#64748B"/></StackPanel></Border>
                <Border Style="{StaticResource Cartao}"><StackPanel><TextBlock x:Name="numOk" Text="0" FontSize="20" FontWeight="Bold" Foreground="#475569"/><TextBlock Text="Concluídos (OK?)" FontSize="10" Foreground="#64748B"/></StackPanel></Border>
              </UniformGrid>
            </StackPanel>
          </Border>

          <Border Style="{StaticResource Painel}">
            <StackPanel>
              <Grid>
                <TextBlock Text="🛑 LOG DE FALHAS E AUDITORIA" FontWeight="Bold" Foreground="#991B1B" FontSize="12" VerticalAlignment="Center"/>
                <Button x:Name="btnLimparLog" Style="{StaticResource Btn}" Background="#E2E8F0" Foreground="#1E293B" Padding="6,2" Margin="0" HorizontalAlignment="Right" Content="Limpar log"/>
              </Grid>
              <TextBox x:Name="txtLog" Height="210" Margin="0,4,0,0" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" FontFamily="Consolas" FontSize="11" Background="#FEF2F2" VerticalContentAlignment="Top" Padding="4"/>
            </StackPanel>
          </Border>
        </StackPanel>
      </ScrollViewer>

      <!-- ÁREA PRINCIPAL DA GRADE -->
      <Border Grid.Column="2" Background="White" BorderBrush="#CBD5E1" BorderThickness="1" CornerRadius="6">
        <Grid>
          <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
          <WrapPanel Grid.Row="0" Margin="8,8,8,2">
            <Button x:Name="btnMarcar" Style="{StaticResource Btn}" Background="#475569" Content="☑️ Marcar/Desmarcar Todos"/>
            <TextBox x:Name="txtFiltro" Width="230" Height="28" Margin="0,0,6,4" Padding="4,4" VerticalContentAlignment="Center" BorderBrush="#CBD5E1" ToolTip="Filtra por nome, matrícula, chamado, CPF, logon, centro de custo ou notas"/>
            <ComboBox x:Name="cmbVisao" Width="150" Height="28" Margin="0,0,10,4" VerticalContentAlignment="Center" SelectedValuePath="Tag" SelectedIndex="0">
              <ComboBoxItem Tag="todos" Content="Exibir: Todos"/><ComboBoxItem Tag="pendentes" Content="⏳ Pendentes de criação"/><ComboBoxItem Tag="migracao" Content="⚡ Migrações"/><ComboBoxItem Tag="divergencia" Content="🛑 Divergências"/><ComboBoxItem Tag="falhas" Content="❌ Falhas"/></ComboBox>
            <Button x:Name="btnCarregar" Style="{StaticResource Btn}" Background="#0070C0" Content="✏️ Carregar no Form"/>
            <Button x:Name="btnValidar" Style="{StaticResource Btn}" Background="#D97706" Content="🔍 Validar no AD"/>
            <Button x:Name="btnProcessar" Style="{StaticResource Btn}" Background="#16A34A" Content="⚡ Processar Criação no AD"/>
            <Button x:Name="btnRascunho" Style="{StaticResource Btn}" Background="#0284C7" Content="✉️ Gerar Rascunho"/>
            <Button x:Name="btnCancelar" Style="{StaticResource Btn}" Background="#DC2626" Content="⏹ Cancelar" Visibility="Collapsed"/>
          </WrapPanel>

          <DataGrid x:Name="grid" Grid.Row="1" Margin="6,2,6,6" AutoGenerateColumns="False" CanUserAddRows="False" CanUserDeleteRows="False"
                    SelectionUnit="Cell" SelectionMode="Extended" ClipboardCopyMode="ExcludeHeader" HeadersVisibility="Column"
                    GridLinesVisibility="All" HorizontalGridLinesBrush="#E2E8F0" VerticalGridLinesBrush="#E2E8F0"
                    RowHeight="25" ColumnHeaderHeight="28" FontSize="11" FrozenColumnCount="5"
                    ScrollViewer.HorizontalScrollBarVisibility="Auto" ScrollViewer.VerticalScrollBarVisibility="Auto"
                    EnableRowVirtualization="True" Background="White" BorderThickness="0">
            <DataGrid.CellStyle>
              <Style TargetType="DataGridCell">
                <Setter Property="Padding" Value="4,2"/><Setter Property="BorderThickness" Value="0"/><Setter Property="VerticalContentAlignment" Value="Center"/>
                <Style.Triggers>
                  <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#BFDBFE"/><Setter Property="Foreground" Value="#0F172A"/>
                    <Setter Property="BorderBrush" Value="#0070C0"/><Setter Property="BorderThickness" Value="1"/>
                  </Trigger>
                </Style.Triggers>
              </Style>
            </DataGrid.CellStyle>
            
            <!-- REQUISITO 2: CONTRASTE VISUAL DAS LINHAS (ROWSTYLE) -->
            <DataGrid.RowStyle>
              <Style TargetType="DataGridRow">
                <Setter Property="Background" Value="#FFFFFF"/>
                <Setter Property="Foreground" Value="#1E293B"/>
                <Style.Triggers>
                  <!-- Status pendencia: Fundo vermelho suave e texto vermelho escuro -->
                  <DataTrigger Binding="{Binding Status}" Value="pendencia">
                    <Setter Property="Background" Value="#FEE2E2"/>
                    <Setter Property="Foreground" Value="#991B1B"/>
                    <Setter Property="FontWeight" Value="SemiBold"/>
                  </DataTrigger>
                  
                  <!-- Status migracao: Fundo âmbar suave e texto marrom/laranja escuro -->
                  <DataTrigger Binding="{Binding Status}" Value="migracao">
                    <Setter Property="Background" Value="#FEF3C7"/>
                    <Setter Property="Foreground" Value="#92400E"/>
                    <Setter Property="FontWeight" Value="SemiBold"/>
                  </DataTrigger>
                  
                  <!-- Criados com sucesso: Fundo verde suave e texto verde escuro -->
                  <DataTrigger Binding="{Binding Resultado}" Value="Criado">
                    <Setter Property="Background" Value="#DCFCE7"/>
                    <Setter Property="Foreground" Value="#166534"/>
                  </DataTrigger>
                  
                  <!-- Falha operacional: Fundo vermelho com texto escuro -->
                  <DataTrigger Binding="{Binding Resultado}" Value="Falha">
                    <Setter Property="Background" Value="#FECACA"/>
                    <Setter Property="Foreground" Value="#7F1D1D"/>
                  </DataTrigger>
                  
                  <!-- Concluído OK: Fundo verde suave sem opacidade excessiva para preservar nitidez -->
                  <DataTrigger Binding="{Binding Ok}" Value="True">
                    <Setter Property="Background" Value="#DCFCE7"/>
                    <Setter Property="Foreground" Value="#166534"/>
                  </DataTrigger>
                </Style.Triggers>
              </Style>
            </DataGrid.RowStyle>

            <DataGrid.Columns>
              <DataGridTemplateColumn Header="Sel." Width="42" SortMemberPath="Sel" ClipboardContentBinding="{Binding Sel}">
                <DataGridTemplateColumn.CellTemplate><DataTemplate><CheckBox IsChecked="{Binding Sel, UpdateSourceTrigger=PropertyChanged}" HorizontalAlignment="Center" VerticalAlignment="Center" ToolTip="Marcar para processamento"/></DataTemplate></DataGridTemplateColumn.CellTemplate>
              </DataGridTemplateColumn>
              <DataGridTemplateColumn Header="OK?" Width="42" SortMemberPath="Ok" ClipboardContentBinding="{Binding Ok}">
                <DataGridTemplateColumn.CellTemplate><DataTemplate><CheckBox IsChecked="{Binding Ok, UpdateSourceTrigger=PropertyChanged}" HorizontalAlignment="Center" VerticalAlignment="Center" ToolTip="Concluído"/></DataTemplate></DataGridTemplateColumn.CellTemplate>
              </DataGridTemplateColumn>
              <DataGridTemplateColumn Header="MFA?" Width="46" SortMemberPath="Mfa" ClipboardContentBinding="{Binding Mfa}">
                <DataGridTemplateColumn.CellTemplate><DataTemplate><CheckBox IsChecked="{Binding Mfa, UpdateSourceTrigger=PropertyChanged}" HorizontalAlignment="Center" VerticalAlignment="Center" ToolTip="MFA configurado"/></DataTemplate></DataGridTemplateColumn.CellTemplate>
              </DataGridTemplateColumn>
              <DataGridTemplateColumn Header="Licença?" Width="58" SortMemberPath="Licenca" ClipboardContentBinding="{Binding Licenca}">
                <DataGridTemplateColumn.CellTemplate><DataTemplate><CheckBox IsChecked="{Binding Licenca, UpdateSourceTrigger=PropertyChanged}" HorizontalAlignment="Center" VerticalAlignment="Center" ToolTip="Licença atribuída"/></DataTemplate></DataGridTemplateColumn.CellTemplate>
              </DataGridTemplateColumn>
              <DataGridTemplateColumn Header="Status 🔄" Width="104" SortMemberPath="Status" ClipboardContentBinding="{Binding Status}">
                <DataGridTemplateColumn.CellTemplate><DataTemplate>
                  <ComboBox SelectedValuePath="Content" SelectedValue="{Binding Status, UpdateSourceTrigger=PropertyChanged}" BorderThickness="0" FontSize="11" Padding="3,1">
                    <ComboBoxItem Content="nova"/><ComboBoxItem Content="migracao"/><ComboBoxItem Content="pendencia"/>
                  </ComboBox>
                </DataTemplate></DataGridTemplateColumn.CellTemplate>
              </DataGridTemplateColumn>
              <DataGridTextColumn Header="A (Nome)" Binding="{Binding PrimeiroNome, Mode=OneWay}" IsReadOnly="True" Width="90"/>
              <DataGridTextColumn Header="B (Sobrenome)" Binding="{Binding Sobrenome, Mode=OneWay}" IsReadOnly="True" Width="100"/>
              <DataGridTextColumn Header="C (Nome Completo)" Binding="{Binding NomeCompleto, Mode=OneWay}" IsReadOnly="True" Width="200"/>
              <DataGridTextColumn Header="D (Logon) ✏️" Binding="{Binding Logon, Mode=TwoWay, UpdateSourceTrigger=LostFocus}" Width="150"/>
              <DataGridTextColumn Header="SamAccount (Matrícula H) ✏️" Binding="{Binding Matricula, Mode=TwoWay, UpdateSourceTrigger=LostFocus}" Width="150"/>
              <DataGridTextColumn Header="E (Cargo)" Binding="{Binding Cargo, Mode=OneWay}" IsReadOnly="True" Width="160"/>
              <DataGridTextColumn Header="F (Depto)" Binding="{Binding Depto, Mode=OneWay}" IsReadOnly="True" Width="140"/>
              <DataGridTextColumn Header="G (Empresa)" Binding="{Binding Empresa, Mode=OneWay}" IsReadOnly="True" Width="110"/>
              <DataGridTextColumn Header="I (RG)" Binding="{Binding Rg, Mode=OneWay}" IsReadOnly="True" Width="100"/>
              <DataGridTextColumn Header="J (CPF)" Binding="{Binding Cpf, Mode=OneWay}" IsReadOnly="True" Width="110"/>
              <DataGridTextColumn Header="K (Endereço)" Binding="{Binding Endereco, Mode=OneWay}" IsReadOnly="True" Width="90"/>
              <DataGridTextColumn Header="L (Gerente)" Binding="{Binding Gerente, Mode=OneWay}" IsReadOnly="True" Width="100"/>
              <DataGridTextColumn Header="M (E-mail)" Binding="{Binding EmailLic, Mode=OneWay}" IsReadOnly="True" Width="64"/>
              <DataGridTextColumn Header="N (Perfil)" Binding="{Binding Perfil, Mode=OneWay}" IsReadOnly="True" Width="64"/>
              <DataGridTextColumn Header="O (Região)" Binding="{Binding Malha, Mode=OneWay}" IsReadOnly="True" Width="64"/>
              <DataGridTextColumn Header="P (Cloud Timestamp)" Binding="{Binding CloudTimestamp, Mode=OneWay}" IsReadOnly="True" Width="140"/>
              <DataGridTextColumn Header="R (Chamado)" Binding="{Binding Chamado, Mode=OneWay}" IsReadOnly="True" Width="100"/>
              <DataGridTextColumn Header="Centro Custo" Binding="{Binding CentroCusto, Mode=OneWay}" IsReadOnly="True" Width="95"/>
              <DataGridTextColumn Header="E-mail (UPN)" Binding="{Binding Upn, Mode=OneWay}" IsReadOnly="True" Width="250" Foreground="#0070C0"/>
              <DataGridTextColumn Header="Resultado AD" Binding="{Binding Resultado, Mode=OneWay}" IsReadOnly="True" Width="95">
                <DataGridTextColumn.ElementStyle><Style TargetType="TextBlock"><Setter Property="ToolTip" Value="{Binding Detalhe}"/><Setter Property="FontWeight" Value="SemiBold"/></Style></DataGridTextColumn.ElementStyle>
              </DataGridTextColumn>
              <DataGridTextColumn Header="Rascunho" Binding="{Binding Rascunho, Mode=OneWay}" IsReadOnly="True" Width="90"/>
              <DataGridTextColumn Header="Lembrete / Pendência" Binding="{Binding Notas, Mode=OneWay}" IsReadOnly="True" Width="420">
                <DataGridTextColumn.ElementStyle><Style TargetType="TextBlock"><Setter Property="TextTrimming" Value="CharacterEllipsis"/><Setter Property="ToolTip" Value="{Binding Notas}"/></Style></DataGridTextColumn.ElementStyle>
              </DataGridTextColumn>
            </DataGrid.Columns>
          </DataGrid>
        </Grid>
      </Border>
    </Grid>

    <!-- STATUS BAR -->
    <Border Grid.Row="2" Background="#E2E8F0" Padding="12,4">
      <Grid>
        <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="260"/></Grid.ColumnDefinitions>
        <TextBlock x:Name="txtStatus" Text="Pronto." Foreground="#1E293B" VerticalAlignment="Center" TextTrimming="CharacterEllipsis"/>
        <ProgressBar x:Name="pb" Grid.Column="1" Height="12" Visibility="Collapsed" Minimum="0" Maximum="100"/>
      </Grid>
    </Border>
  </Grid>
</Window>
'@

$win = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader ([xml]$xaml)))
$ui = @{}
foreach ($m in [regex]::Matches($xaml, 'x:Name="([^"]+)"')) {
    $n = $m.Groups[1].Value
    $c = $win.FindName($n)
    if ($null -ne $c) { $ui[$n] = $c }
}
$script:BrushNormal = $ui.txtCpf.BorderBrush
$script:BrushErro = [System.Windows.Media.Brushes]::Red

$script:View = [System.Windows.Data.CollectionViewSource]::GetDefaultView($script:Dados)
$script:View.Filter = $script:Filtro.GetPredicate()
$ui.grid.ItemsSource = $script:Dados

# =====================================================================================
# 5. CONTROLE DA UI E EVENTOS
# =====================================================================================
function Show-Msg {
    param([string]$Texto, [string]$Titulo = 'IAM - Central de Provisionamento', [string]$Botoes = 'OK', [string]$Icone = 'Information')
    return [System.Windows.MessageBox]::Show($win, $Texto, $Titulo, $Botoes, $Icone)
}

function Add-LogUI {
    param([string]$Nivel, [string]$Texto)
    if ($Nivel -eq 'info') { $ui.txtStatus.Text = $Texto; return }
    $pref = if ($Nivel -eq 'erro') { '[ERRO] ' } else { '[AVISO] ' }
    $ui.txtLog.AppendText(("{0} {1}{2}`r`n" -f (Get-Date -Format 'HH:mm:ss'), $pref, $Texto))
    $ui.txtLog.ScrollToEnd()
}

function Commit-Grid {
    try { $null = $ui.grid.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Row, $true) } catch { }
}

function Request-Refresh {
    param([switch]$Forcar)
    $script:RefreshPendente = $true
    if ($Forcar) { $script:RefreshForcar = $true }
}

function Invoke-RefreshSeguro {
    if (-not $script:RefreshPendente) { return }
    $v = $script:View
    if ($v.IsEditingItem -or $v.IsAddingNew) {
        if ($script:RefreshForcar) {
            Commit-Grid
            try { if ($v.IsEditingItem) { $v.CommitEdit() } } catch { }
            try { if ($v.IsAddingNew) { $v.CommitNew() } } catch { }
        } else { return }
    }
    try {
        $v.Refresh()
        $script:RefreshPendente = $false
        $script:RefreshForcar = $false
    } catch { }
}

function Update-Dashboard {
    $t = 0; $s = 0; $c = 0; $p = 0; $f = 0; $o = 0
    foreach ($r in $script:Dados) {
        $t++
        if ($r.Sel) { $s++ }
        if ($r.Resultado -eq 'Criado') { $c++ }
        if ($r.Status -eq 'pendencia') { $p++ }
        if ($r.Resultado -eq 'Falha') { $f++ }
        if ($r.Ok) { $o++ }
    }
    $ui.numTotal.Text = "$t"; $ui.numSel.Text = "$s"; $ui.numCriados.Text = "$c"
    $ui.numPend.Text = "$p"; $ui.numFalhas.Text = "$f"; $ui.numOk.Text = "$o"
    $ui.txtContador.Text = "$t Registros"
}

function Get-RegPorId {
    param([string]$Id)
    foreach ($r in $script:Dados) { if ($r.Id -eq $Id) { return $r } }
    return $null
}

function Get-LinhaAtual {
    $c = $ui.grid.CurrentCell
    if ($c.Item -is [Registro]) { return $c.Item }
    $s = $ui.grid.SelectedCells
    if ($s.Count -gt 0 -and $s[0].Item -is [Registro]) { return $s[0].Item }
    return $null
}

function Get-LinhasSelecionadasGrid {
    $lista = New-Object System.Collections.ArrayList
    foreach ($cell in $ui.grid.SelectedCells) {
        if ($cell.Item -is [Registro] -and -not $lista.Contains($cell.Item)) { [void]$lista.Add($cell.Item) }
    }
    return $lista.ToArray()
}

function Test-Dups {
    $cpf = ($ui.txtCpf.Text -replace '\D', '')
    $cham = $ui.txtChamado.Text.Trim().ToLower()
    $nome = [Registro]::Limpar($ui.txtNome.Text).ToLower()
    $mat = $ui.txtMatricula.Text.Trim().ToLower()
    $logon = $ui.txtLogon.Text.Trim().ToLower()
    $dCpf = $false; $dCham = $false; $dNome = $false; $dMat = $false; $dLogon = $false
    foreach ($r in $script:Dados) {
        if ($script:EditId -and $r.Id -eq $script:EditId) { continue }
        if ($cpf -and (($r.Cpf -replace '\D', '') -eq $cpf)) { $dCpf = $true }
        if ($cham -and $r.Chamado -and $r.Chamado.Trim().ToLower() -eq $cham) { $dCham = $true }
        if ($nome -and ([Registro]::Limpar($r.NomeCompleto).ToLower() -eq $nome)) { $dNome = $true }
        if ($mat -and $r.Matricula -and $r.Matricula.Trim().ToLower() -eq $mat) { $dMat = $true }
        if ($logon -and $r.Logon -and $r.Logon.Trim().ToLower() -eq $logon) { $dLogon = $true }
    }
    $set = { param($ctl, $flag) $ctl.BorderBrush = if ($flag) { $script:BrushErro } else { $script:BrushNormal } }
    & $set $ui.txtCpf $dCpf; & $set $ui.txtChamado $dCham; & $set $ui.txtNome $dNome; & $set $ui.txtMatricula $dMat; & $set $ui.txtLogon $dLogon
    $msgs = @()
    if ($dCham) { $msgs += '🛑 Chamado duplicado!' }
    if ($dCpf) { $msgs += '⚠️ CPF já existente!' }
    if ($dNome) { $msgs += '⚠️ Nome já existente!' }
    if ($dMat) { $msgs += '⚠️ Matrícula já existente!' }
    if ($dLogon) { $msgs += '⚠️ Logon já existente!' }
    $ui.txtDup.Text = ($msgs -join '    ')
    $ui.txtDup.Visibility = if ($msgs.Count -gt 0) { 'Visible' } else { 'Collapsed' }
    return @{ Cham = $dCham; Cpf = $dCpf; Nome = $dNome; Mat = $dMat; Logon = $dLogon }
}

function Clear-Form {
    $script:Atualizando = $true
    foreach ($n in 'txtNome', 'txtPrimeiro', 'txtSobrenome', 'txtLogon', 'txtMatricula', 'txtCargo', 'txtDepto', 'txtEmpresa', 'txtRg', 'txtCpf', 'txtEndereco', 'txtGerente', 'txtChamado', 'txtCC', 'txtNotas') { $ui[$n].Text = '' }
    $ui.cmbEmailLic.SelectedValue = 's'; $ui.cmbPerfil.SelectedValue = '2'; $ui.cmbMalha.SelectedValue = '1'; $ui.cmbStatus.SelectedValue = 'nova'
    $ui.txtCloud.Text = Get-TimestampUtc
    $script:Atualizando = $false
    $script:EditId = $null
    $script:LogonManual = $false
    $ui.btnAdd.Visibility = 'Visible'; $ui.pnlEdicao.Visibility = 'Collapsed'
    [void](Test-Dups)
}

function Read-FormInto {
    param($r)
    $nome = Format-NomeProprio $ui.txtNome.Text
    $partes = if ($nome) { $nome -split '\s+' } else { @() }
    $prim = Format-NomeProprio $(if ($ui.txtPrimeiro.Text.Trim()) { $ui.txtPrimeiro.Text } elseif ($partes.Count -gt 0) { $partes[0] } else { '' })
    $sob = Format-NomeProprio $(if ($ui.txtSobrenome.Text.Trim()) { $ui.txtSobrenome.Text } elseif ($partes.Count -gt 1) { $partes[$partes.Count - 1] } else { $prim })
    $logon = $ui.txtLogon.Text.Trim().ToLower()
    if (-not $logon -and $prim -and $sob) { $logon = ([Registro]::Limpar($prim) + '.' + [Registro]::Limpar($sob)).ToLower() }
    $r.NomeCompleto = $nome; $r.PrimeiroNome = $prim; $r.Sobrenome = $sob; $r.Logon = $logon
    $r.Cargo = $ui.txtCargo.Text.Trim(); $r.Depto = $ui.txtDepto.Text.Trim(); $r.Empresa = $ui.txtEmpresa.Text.Trim()
    $r.Matricula = $ui.txtMatricula.Text.Trim(); $r.Rg = $ui.txtRg.Text.Trim(); $r.Cpf = $ui.txtCpf.Text.Trim()
    $r.Endereco = $ui.txtEndereco.Text.Trim(); $r.Gerente = $ui.txtGerente.Text.Trim()
    $r.EmailLic = [string]$ui.cmbEmailLic.SelectedValue; $r.Perfil = [string]$ui.cmbPerfil.SelectedValue; $r.Malha = [string]$ui.cmbMalha.SelectedValue
    $ts = $ui.txtCloud.Text.Trim(); $r.CloudTimestamp = if ($ts) { $ts } else { Get-TimestampUtc }
    $r.Status = [string]$ui.cmbStatus.SelectedValue
    $r.Chamado = $ui.txtChamado.Text.Trim(); $r.CentroCusto = $ui.txtCC.Text.Trim(); $r.Notas = $ui.txtNotas.Text.Trim()
    $falt = Set-PendenciaSeFaltar $r
    if ($falt -gt 0) { $ui.txtStatus.Text = '⚠️ Registro marcado como PENDÊNCIA: pendências cadastrais identificadas.' }
}

function Load-FormFrom {
    param($r)
    $script:Atualizando = $true
    $ui.txtNome.Text = $r.NomeCompleto; $ui.txtPrimeiro.Text = $r.PrimeiroNome; $ui.txtSobrenome.Text = $r.Sobrenome; $ui.txtLogon.Text = $r.Logon
    $ui.txtMatricula.Text = $r.Matricula; $ui.txtCargo.Text = $r.Cargo; $ui.txtDepto.Text = $r.Depto; $ui.txtEmpresa.Text = $r.Empresa
    $ui.txtRg.Text = $r.Rg; $ui.txtCpf.Text = $r.Cpf; $ui.txtEndereco.Text = $r.Endereco; $ui.txtGerente.Text = $r.Gerente
    $ui.cmbEmailLic.SelectedValue = $r.EmailLic; $ui.cmbPerfil.SelectedValue = $r.Perfil; $ui.cmbMalha.SelectedValue = $r.Malha
    $ui.txtCloud.Text = $r.CloudTimestamp; $ui.cmbStatus.SelectedValue = $r.Status
    $ui.txtChamado.Text = $r.Chamado; $ui.txtCC.Text = $r.CentroCusto; $ui.txtNotas.Text = $r.Notas
    $script:Atualizando = $false
    $script:EditId = $r.Id
    $script:LogonManual = $true
    $ui.btnAdd.Visibility = 'Collapsed'; $ui.pnlEdicao.Visibility = 'Visible'
    [void](Test-Dups)
}

function Set-Ocupado {
    param([bool]$on)
    foreach ($n in 'btnValidar', 'btnProcessar', 'btnRascunho', 'btnImportar', 'btnLimparOk') { $ui[$n].IsEnabled = -not $on }
    $ui.btnCancelar.Visibility = if ($on) { 'Visible' } else { 'Collapsed' }
    $ui.pb.Visibility = if ($on) { 'Visible' } else { 'Collapsed' }
    $ui.pb.IsIndeterminate = $on
}

function Start-Worker {
    param([string]$Funcao, [object[]]$Snaps)
    if ($script:Worker) { return }
    $script:Ctl.Cancel = $false
    $rs = [runspacefactory]::CreateRunspace()
    $rs.ApartmentState = [System.Threading.ApartmentState]::STA
    $rs.ThreadOptions = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
    $rs.Open()
    $rs.SessionStateProxy.SetVariable('Q', $script:Fila)
    $rs.SessionStateProxy.SetVariable('Ctl', $script:Ctl)
    $rs.SessionStateProxy.SetVariable('Snaps', $Snaps)
    $ps = [powershell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript($script:Lib + "`n" + $Funcao)
    Set-Ocupado $true
    $script:Worker = @{ PS = $ps; RS = $rs; Handle = $ps.BeginInvoke() }
}

function Complete-Worker {
    if (-not $script:Worker) { return }
    if (-not $script:Worker.Handle.IsCompleted) { return }
    if (-not $script:Fila.IsEmpty) { return }
    $w = $script:Worker
    $script:Worker = $null
    try { [void]$w.PS.EndInvoke($w.Handle) } catch { Add-LogUI 'erro' ('Falha no processamento: ' + $_.Exception.Message) }
    foreach ($e in $w.PS.Streams.Error) { Add-LogUI 'erro' ($e.ToString()) }
    try { $w.PS.Dispose(); $w.RS.Close(); $w.RS.Dispose() } catch { }
    Set-Ocupado $false
    Request-Refresh
    [Registro]::Sujo = $true
}

function Process-Fila {
    $msg = $null
    $n = 0
    while ($n -lt 300 -and $script:Fila.TryDequeue([ref]$msg)) {
        $n++
        switch ($msg.T) {
            'log' { Add-LogUI $msg.Nivel $msg.Texto }
            'prog' {
                $ui.pb.IsIndeterminate = $false
                $ui.pb.Maximum = [double]$msg.Total
                $ui.pb.Value = [double]$msg.Atual
                $ui.txtStatus.Text = $msg.Texto
            }
            'res' {
                $r = Get-RegPorId $msg.Id
                if ($r) {
                    $r.Resultado = [string]$msg.Resultado
                    $r.Detalhe = [string]$msg.Detalhe
                    if ($msg.Rascunho) { $r.Rascunho = [string]$msg.Rascunho }
                    if ($msg.Resultado -eq 'Criado') { $r.Sel = $false; $r.Ok = $true }
                }
            }
            'rasc' {
                $r = Get-RegPorId $msg.Id
                if ($r) { $r.Rascunho = [string]$msg.Rascunho }
            }
            'aud' {
                $r = Get-RegPorId $msg.Id
                if ($r) {
                    $teveHomonimo = ($r.Notas -match 'Homônimo|Conflito')
                    $r.Notas = Set-NotaAD $r.Notas ([string]$msg.Nota)
                    if ($msg.Status) { $r.Status = [string]$msg.Status }
                    elseif (-not $msg.Achou -and $r.Status -eq 'pendencia' -and $teveHomonimo -and (@(Get-Faltantes $r).Count -eq 0)) { $r.Status = 'nova' }
                }
            }
            'fim' {
                Add-LogUI 'aviso' ([string]$msg.Texto)
                $ui.txtStatus.Text = [string]$msg.Texto
            }
        }
    }
}

# =====================================================================================
# REQUISITO 1: TRAVA ESTRITA DE SELEÇÃO NO PROCESSAMENTO
# =====================================================================================
function Get-Alvos {
    param([ValidateSet('criacao', 'auditoria')][string]$Modo)
    Commit-Grid
    
    $marcados = @($script:Dados | Where-Object { $_.Sel })
    
    if ($Modo -eq 'criacao') {
        # TRAVA ESTRITA: Exige [✔] Sel marcado obrigatoriamente. Sem fallback automático!
        return $marcados
    }
    
    # Para auditoria, se nada estiver marcado com Sel, permite validar a fila não concluída
    if ($marcados.Count -gt 0) { return $marcados }
    return @($script:Dados | Where-Object { -not $_.Ok -and $_.Resultado -ne 'Criado' })
}

# =============== EVENTOS DOS CONTROLES ===============
$ui.txtNome.Add_TextChanged({
    if ($script:Atualizando) { return }
    $nome =$ui.txtNome.Text.Trim()
    if (-not $nome) { $ui.txtPrimeiro.Text = '';$ui.txtSobrenome.Text = ''; if (-not $script:LogonManual) {$ui.txtLogon.Text = '' } }
    else {
        $partes = $nome -split '\s+'$prim = Format-NomeProprio $partes[0]$sob = if ($partes.Count -gt 1) { Format-NomeProprio$partes[$partes.Count - 1] } else {$prim }
        $ui.txtPrimeiro.Text =$prim; $ui.txtSobrenome.Text =$sob
        if (-not $script:LogonManual -or -not $ui.txtLogon.Text) {$script:Atualizando = $true$ui.txtLogon.Text = ([Registro]::Limpar($prim) + '.' + [Registro]::Limpar($sob)).ToLower()
            $script:Atualizando =$false
        }
    }
    [void](Test-Dups)
})
$ui.txtNome.Add_LostFocus({
    if ($ui.txtNome.Text) {
        $script:Atualizando =$true
        $ui.txtNome.Text = Format-NomeProprio$ui.txtNome.Text
        $script:Atualizando =$false
    }
})
$ui.txtLogon.Add_TextChanged({
    if ($script:Atualizando) { return }
    if ($ui.txtLogon.IsKeyboardFocused) { $script:LogonManual =$true }
    [void](Test-Dups)
})
$ui.txtCpf.Add_TextChanged({
    if ($script:Atualizando) { return }
    $script:Atualizando =$true
    $f = Format-Cpf$ui.txtCpf.Text
    if ($f -ne$ui.txtCpf.Text) { $ui.txtCpf.Text =$f; $ui.txtCpf.CaretIndex =$f.Length }
    $script:Atualizando =$false
    [void](Test-Dups)
})
foreach ($n in 'txtMatricula', 'txtChamado') { $ui[$n].Add_TextChanged({ if (-not $script:Atualizando) { [void](Test-Dups) } }) }$ui.txtCloud.Add_LostFocus({
    $d = ($ui.txtCloud.Text -replace '\D', '')
    if ($d.Length -ge 8) { $ui.txtCloud.Text =$d.Substring(0, 8) + '080000.0Z' }
})
$ui.btnCloudUtc.Add_Click({$ui.txtCloud.Text = Get-TimestampUtc })
$ui.btnLimparForm.Add_Click({ Clear-Form })$ui.btnCancelarEd.Add_Click({ Clear-Form })

$ui.btnAdd.Add_Click({$d = Test-Dups
    if ($d.Cham) { [void](Show-Msg '🛑 ERRO DE AUDITORIA: este chamado já foi cadastrado na base.' 'Chamado duplicado' 'OK' 'Error'); return }
    $r = New-Object Registro
    Read-FormInto $r
    if ([string]::IsNullOrWhiteSpace($r.Matricula) -or [string]::IsNullOrWhiteSpace($r.NomeCompleto)) {
        [void](Show-Msg 'Preencha ao menos Nome Completo e Matrícula.' 'Validação' 'OK' 'Warning'); return
    }
    $script:Dados.Add($r)
    [Registro]::Sujo = $true
    Clear-Form
    $ui.txtStatus.Text = 'Registro adicionado à fila.'
})

$ui.btnSalvarEd.Add_Click({
    $r = Get-RegPorId$script:EditId
    if (-not $r) { Clear-Form; return }$d = Test-Dups
    if ($d.Cham) { [void](Show-Msg '🛑 ERRO DE AUDITORIA: este chamado já foi cadastrado em outro registro.' 'Chamado duplicado' 'OK' 'Error'); return }
    Read-FormInto $r
    Clear-Form
    Request-Refresh -Forcar
    $ui.txtStatus.Text = 'Registro atualizado.'
})

# REQUISITO 3: COLAGEM COM FILTRO RIGOROSO DE LINHAS VAZIAS
$ui.btnColar.Add_Click({$txt = ''
    try { $txt = [System.Windows.Clipboard]::GetText() } catch { }
    if ([string]::IsNullOrWhiteSpace($txt)) { [void](Show-Msg 'A área de transferência está vazia. Copie as 18 colunas do Excel e tente novamente.' 'Colar Lote' 'OK' 'Warning'); return }
    $linhas = @($txt -split "\r?\n" | Where-Object { $_.Trim() -ne '' -and$_ -like "*`t*" })
    $add = 0; $dup = 0
    foreach ($l in $linhas) {
        $primeiraCol = ($l -split "`t")[0].Trim()
        if ($primeiraCol -match '^(PrimeiroNome|A \(Nome\))$') { continue }
        
        $r = ConvertFrom-LinhaTab$l
        if ($null -eq$r) { continue } # Ignora silenciosamente linhas vazias/incompletas

        $existe =$false
        if ($r.Matricula) { 
            foreach ($x in$script:Dados) { 
                if ($x.Matricula -ieq$r.Matricula) { $existe =$true; break } 
            } 
        }
        if ($existe) {$dup++; Add-LogUI 'aviso' "$($r.Matricula) - já existe na fila; registro ignorado." }
        else { $script:Dados.Add($r);$add++ }
    }
    [Registro]::Sujo = $true
    Request-Refresh -Forcar
    $ui.txtStatus.Text = "Lote colado: $add adicionado(s),$dup ignorado(s) por duplicidade."
})

$ui.txtFiltro.Add_TextChanged({$script:Filtro.Texto = $ui.txtFiltro.Text.Trim(); Request-Refresh -Forcar })$ui.cmbVisao.Add_SelectionChanged({ $script:Filtro.Visao = [string]$ui.cmbVisao.SelectedValue; Request-Refresh -Forcar })

$ui.btnMarcar.Add_Click({
    Commit-Grid
    $visiveis = @($script:View \vert{} ForEach-Object {$_ })
    if ($visiveis.Count -eq 0) { return }$todos = (@($visiveis \vert{} Where-Object { -not$_.Sel }).Count -eq 0)
    foreach ($r in$visiveis) { $r.Sel = -not$todos }
})

$ui.btnCarregar.Add_Click({
    Commit-Grid
    $r = Get-LinhaAtual
    if (-not $r) { [void](Show-Msg 'Selecione uma linha na tabela para carregar no formulário.' 'Carregar no Form' 'OK' 'Warning'); return }
    Load-FormFrom $r$ui.txtStatus.Text = "A editar: $($r.Matricula) - $($r.NomeCompleto)"
})

$ui.btnValidar.Add_Click({$alvos = @(Get-Alvos 'auditoria')
    if ($alvos.Count -eq 0) { [void](Show-Msg 'Nenhum registro elegível para validar no AD.' 'Validar no AD' 'OK' 'Information'); return }
    $snaps = @($alvos | ForEach-Object { ConvertTo-Snap $_ })$ui.txtStatus.Text = "Validando $($snaps.Count) registro(s) no AD..."
    Start-Worker 'Invoke-LoteAuditoria' $snaps
})

# EXECUÇÃO DO BOTÃO PROCESSAR COM A TRAVA ESTRITA (Requisito 1)
$ui.btnProcessar.Add_Click({$alvos = @(Get-Alvos 'criacao')
    
    # SE NENHUM REGISTRO ESTIVER COM [✔] Sel MARCADO, ABORTA IMEDIATAMENTE
    if ($alvos.Count -eq 0) {
        [void](Show-Msg "Nenhum registro selecionado! Marque a caixa [✔] Sel dos colaboradores que deseja processar." "Seleção Obrigatória" "OK" "Warning")
        return
    }

    $alvosNaoCriados = @($alvos \vert{} Where-Object {$_.Resultado -ne 'Criado' })
    if ($alvosNaoCriados.Count -eq 0) {
        [void](Show-Msg "Todos os registros marcados com [✔] já foram criados anteriormente." "Processar Criação" "OK" "Information")
        return
    }

    $bloq = @($alvosNaoCriados \vert{} Where-Object {$_.Status -eq 'pendencia' }).Count
    $txt = "Criar $($alvosNaoCriados.Count) usuário(s) selecionado(s) no Active Directory e gerar e-mails no Outlook?"
    if ($bloq -gt 0) {$txt += "`n`nATENÇÃO: $bloq registro(s) marcado(s) com status PENDÊNCIA serão BLOQUEADOS por governança." 
    }
    
    if ((Show-Msg $txt 'Confirmação de Provisionamento' 'YesNo' 'Question') -ne 'Yes') { return }
    
    $snaps = @($alvosNaoCriados \vert{} ForEach-Object { ConvertTo-Snap$_ })
    foreach ($a in$alvosNaoCriados) { $a.Resultado = '';$a.Detalhe = '' }
    Start-Worker 'Invoke-LoteCriacao' $snaps
})

$ui.btnRascunho.Add_Click({
    Commit-Grid
    $alvos = @($script:Dados \vert{} Where-Object {$_.Sel })
    if ($alvos.Count -eq 0) {$alvos = @(Get-LinhasSelecionadasGrid) }
    if ($alvos.Count -eq 0) { [void](Show-Msg 'Marque [✔] ou selecione as linhas para gerar os rascunhos.' 'Gerar Rascunho' 'OK' 'Information'); return }$snaps = @($alvos \vert{} ForEach-Object { ConvertTo-Snap$_ })
    Start-Worker 'Invoke-LoteRascunho' $snaps
})

$ui.btnCancelar.Add_Click({$script:Ctl.Cancel = $true; $ui.txtStatus.Text = 'Cancelando após o registro atual...' })
$ui.btnLimparLog.Add_Click({$ui.txtLog.Clear() })

$ui.btnMfa.Add_Click({
    Commit-Grid
    $alvos = @($script:Dados \vert{} Where-Object {$_.Sel })
    if ($alvos.Count -eq 0) { $alvos = @($script:Dados | Where-Object { $_.Resultado -eq 'Criado' -and -not$_.Mfa }) }
    $emails = @($alvos | ForEach-Object { $_.Upn } \vert{} Where-Object {$_ })
    if ($emails.Count -eq 0) { [void](Show-Msg 'Nenhum e-mail disponível para cópia (marque [✔] ou processe as criações primeiro).' 'MFA' 'OK' 'Information'); return }
    [System.Windows.Clipboard]::SetText(($emails -join "`r`n"))
    $ui.txtStatus.Text = "$($emails.Count) e-mail(s) copiado(s) para a área de transferência."
})

$ui.btnLimparOk.Add_Click({
    Commit-Grid
    $ok = @($script:Dados \vert{} Where-Object {$_.Ok })
    if ($ok.Count -eq 0) {$ui.txtStatus.Text = 'Nenhum registro marcado como OK?.'; return }
    if ((Show-Msg "Remover apenas os $($ok.Count) registro(s) marcados como OK?`n`nRegistros pendentes continuarão na fila." 'Limpar Concluídos' 'YesNo' 'Question') -ne 'Yes') { return }
    foreach ($r in$ok) { [void]$script:Dados.Remove($r) }
    [Registro]::Sujo = $true$ui.txtStatus.Text = "$($ok.Count) concluído(s) removido(s)."
})

$ui.btnExportar.Add_Click({
    Commit-Grid
    if ($script:Dados.Count -eq 0) { [void](Show-Msg 'Nenhum dado para exportar.' 'Exportar' 'OK' 'Information'); return }$dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Filter = 'Excel (*.xlsx)\vert{}*.xlsx\vert{}CSV (*.csv)\vert{}*.csv'$dlg.FileName = 'Base_Auditoria_IAM_' + (Get-Date -Format 'yyyy-MM-dd')
    if ($dlg.ShowDialog($win) -ne$true) { return }
    $cab = @('Concluido', 'MFA', 'Licenca', 'PrimeiroNome', 'Sobrenome', 'NomeCompleto', 'Logon', 'Cargo', 'Depto', 'Empresa', 'Matricula', 'RG', 'CPF', 'Endereco', 'Gerente', 'Email(s/n)', 'Perfil(1,2,3,4)', 'Malha', 'CloudTimestamp', 'tipoProcesso', 'Chamado', 'CentroDeCusto', 'Notas_Lembretes', 'UPN')$sim = { param($b) if ($b) { 'Sim' } else { 'Não' } }
    $linhas = @(foreach ($r in$script:Dados) {
        , @((& $sim $r.Ok), (&$sim $r.Mfa), (&$sim $r.Licenca),$r.PrimeiroNome, $r.Sobrenome, $r.NomeCompleto, $r.Logon, $r.Cargo, $r.Depto, $r.Empresa, $r.Matricula, $r.Rg, $r.Cpf, $r.Endereco, $r.Gerente, $r.EmailLic, $r.Perfil, $r.Malha, $r.CloudTimestamp, $r.Status, $r.Chamado, $r.CentroCusto, $r.Notas, $r.Upn)
    })
    try {
        if ($dlg.FileName.ToLower().EndsWith('.csv')) {
            $objs = foreach ($l in $linhas) {$o = [ordered]@{}
                for ($i = 0; $i -lt$cab.Count; $i++) {$o[$cab[$i]] = [string]$l[$i] }
                [pscustomobject]$o
            }
            $objs \vert{} Export-Csv -Path$dlg.FileName -Delimiter ';' -NoTypeInformation -Encoding UTF8
        } else {
            Export-XlsxSimples -Caminho $dlg.FileName -Cabecalho $cab -Linhas$linhas
        }
        $ui.txtStatus.Text = "Exportado com sucesso: $($dlg.FileName)"
    } catch { [void](Show-Msg ("Falha ao exportar: " + $_.Exception.Message) 'Exportar' 'OK' 'Error') }
})

# REQUISITO 3: IMPORTAÇÃO VIA ARQUIVO COM FILTRO RIGOROSO DE LINHAS VAZIAS
$ui.btnImportar.Add_Click({
    Commit-Grid
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Filter = 'Planilhas (*.xlsx;*.csv)|*.xlsx;*.csv'
    if ($dlg.ShowDialog($win) -ne$true) { return }
    try {
        $linhas = if ($dlg.FileName.ToLower().EndsWith('.csv')) { @(Import-CsvFlex $dlg.FileName) } else { @(Import-XlsxSimples$dlg.FileName) }
    } catch { [void](Show-Msg ("Falha ao ler o arquivo: " + $_.Exception.Message) 'Importar' 'OK' 'Error'); return }
    if ($linhas.Count -eq 0) { [void](Show-Msg 'Nenhuma linha válida encontrada no arquivo.' 'Importar' 'OK' 'Warning'); return }

    $substituir =$false
    if ($script:Dados.Count -gt 0) {$resp = Show-Msg ("Você já possui $($script:Dados.Count) registro(s) na base.`n`n[Sim] = ACRESCENTAR/MESCLAR`n[Não] = SUBSTITUIR a fila atual`n[Cancelar] = Abortar") 'Importar' 'YesNoCancel' 'Question'
        if ($resp -eq 'Cancel') { return }
        if ($resp -eq 'No') {
            if ((Show-Msg 'Deseja realmente apagar os registros atuais da tela?' 'Confirmar substituição' 'YesNo' 'Warning') -ne 'Yes') { return }
            $substituir =$true
        }
    }

    # Processa filtrando linhas nulas/vazias
    $novos = @($linhas \vert{} ForEach-Object { ConvertFrom-Hash$_ } | Where-Object { $null -ne$_ })
    if ($substituir) {$script:Dados.Clear() }
    $add = 0; $upd = 0
    foreach ($n in$novos) {
        $alvo =$null
        $cpfN = ($n.Cpf -replace '\D', '')
        foreach ($x in$script:Dados) {
            if (($n.Matricula -and $x.Matricula -ieq$n.Matricula) -or (-not $n.Matricula -and$cpfN -and (($x.Cpf -replace '\D', '') -eq$cpfN))) { $alvo =$x; break }
        }
        if ($alvo) { Copy-RegistroDados $alvo$n; $upd++ } else {$script:Dados.Add($n);$add++ }
    }
    [Registro]::Sujo = $true
    Request-Refresh -Forcar
    $ui.txtStatus.Text = "Importação concluída: $add adicionado(s),$upd atualizado(s)."
})

# =============== TIMER DO DISPATCHER ===============
$script:Timer = New-Object System.Windows.Threading.DispatcherTimer
$script:Timer.Interval = [TimeSpan]::FromMilliseconds(150)$script:Timer.Add_Tick({
    Process-Fila
    Complete-Worker
    Invoke-RefreshSeguro
    Update-Dashboard
    if ([Registro]::Sujo) {
        [Registro]::Sujo = $false$script:SalvarEm = (Get-Date).AddSeconds(2)
    }
    if ($script:SalvarEm -and (Get-Date) -ge$script:SalvarEm) {
        $script:SalvarEm =$null
        Save-Base
    }
})

[System.Windows.Threading.Dispatcher]::CurrentDispatcher.add_UnhandledException({
    param($sender,$e)
    try { Add-LogUI 'erro' ('Exceção de interface: ' + $e.Exception.Message) } catch { }
    $e.Handled =$true
})

$win.Add_Closing({
    param($sender,$e)
    if ($script:Worker) {
        if ((Show-Msg 'Há um processamento em andamento. Deseja realmente sair?' 'Sair' 'YesNo' 'Warning') -ne 'Yes') { $e.Cancel =$true; return }
        $script:Ctl.Cancel =$true
    }
    Commit-Grid
    $script:Timer.Stop()
    Save-Base
})

# =============== INICIALIZAÇÃO ===============
Load-Base
Clear-Form
Update-Dashboard
if (Test-Elevado) {
    Add-LogUI 'aviso' "Aplicativo em modo ADMINISTRADOR: o COM do Outlook pode falhar se o Outlook estiver em modo normal. Execute sem elevação se necessário."
}
$ui.txtStatus.Text = "Central pronta. $($script:Dados.Count) registro(s) carregado(s)."
$script:Timer.Start()
[void]$win.ShowDialog()

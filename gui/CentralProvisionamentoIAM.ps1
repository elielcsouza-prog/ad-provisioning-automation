<#
.SYNOPSIS
    Central Gráfica de Provisionamento, Auditoria e Governança de Identidades (IAM/IGA).

.DESCRIPTION
    Interface WPF (XAML) integrada com Active Directory e Microsoft Outlook para
    orquestração do ciclo de vida de identidades (Joiners, Migrações e Auditoria).
    Implementa validação contra homónimos, processamento em lote, verificação de duplicidade
    e geração de notificações corporativas em HTML.
#>

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms, System.Drawing;
Import-Module ActiveDirectory -ErrorAction SilentlyContinue;

# ------------------------------------------------------------------------------
# 0. CONFIGURAÇÃO DE AMBIENTE E ÍCONE DINÂMICO
# ------------------------------------------------------------------------------
$pastaTrabalho = Join-Path $env:TEMP "IAM_Automacao"
if (-not (Test-Path $pastaTrabalho)) {
    New-Item -ItemType Directory -Path $pastaTrabalho -Force -ErrorAction SilentlyContinue | Out-Null
}
$caminhoIco = Join-Path $pastaTrabalho "app.ico"

if (-not (Test-Path $caminhoIco)) {
    try {
        $bmp = New-Object System.Drawing.Bitmap 256, 256
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias

        $pincelBorda = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255, 0, 180, 240))
        $pontosEscudoBorda = @(
            (New-Object System.Drawing.Point 128, 30),
            (New-Object System.Drawing.Point 218, 55),
            (New-Object System.Drawing.Point 218, 140),
            (New-Object System.Drawing.Point 128, 226),
            (New-Object System.Drawing.Point 38, 140),
            (New-Object System.Drawing.Point 38, 55)
        )
        $g.FillPolygon($pincelBorda, $pontosEscudoBorda)

        $pincelFundo = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255, 0, 32, 96))
        $pontosEscudoInterno = @(
            (New-Object System.Drawing.Point 128, 38),
            (New-Object System.Drawing.Point 210, 60),
            (New-Object System.Drawing.Point 210, 136),
            (New-Object System.Drawing.Point 128, 216),
            (New-Object System.Drawing.Point 46, 136),
            (New-Object System.Drawing.Point 46, 60)
        )
        $g.FillPolygon($pincelFundo, $pontosEscudoInterno)

        $pincelId = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::White)
        $g.FillEllipse($pincelId, 106, 75, 44, 44)

        $pontosCorpo = @(
            (New-Object System.Drawing.Point 128, 126),
            (New-Object System.Drawing.Point 170, 168),
            (New-Object System.Drawing.Point 86, 168)
        )
        $g.FillPolygon($pincelId, $pontosCorpo)
        $g.FillRectangle($pincelId, 120, 165, 16, 22)

        $hIcon = $bmp.GetHicon()
        $icon = [System.Drawing.Icon]::FromHandle($hIcon)
        $stream = [System.IO.File]::OpenWrite($caminhoIco)
        $icon.Save($stream)
        $stream.Close()

        $g.Dispose()
        $bmp.Dispose()
    } catch {}
}

# ------------------------------------------------------------------------------
# 1. MAPEAMENTO DE OUs E ASSINATURA (.HTM)
# ------------------------------------------------------------------------------
$dominioBase    = "DC=empresa,DC=corp"
$ouTercNorte    = "OU=Regiao Norte,OU=Terceiros,OU=Usuarios,$dominioBase"
$ouTercSul      = "OU=Regiao Sul,OU=Terceiros,OU=Usuarios,$dominioBase"
$ouTercEspecial = "OU=Especial,OU=Terceiros,OU=Usuarios,$dominioBase"
$ouTercPadrao   = "OU=Geral,OU=Terceiros,OU=Usuarios,$dominioBase"

$pastaDocs = [System.Environment]::GetFolderPath('MyDocuments')
$logPath   = [System.IO.Path]::Combine($pastaDocs, "Logs_Automacao")
if (-not ([System.IO.Directory]::Exists($logPath))) {
    $null = [System.IO.Directory]::CreateDirectory($logPath)
}
$logFile = [System.IO.Path]::Combine($logPath, "Log_Criacao_Lote.txt")

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
            $assinaturaHTML = $assinaturaHTML -replace $sigFolder, $folderCompleto
        }
    }
}

# ------------------------------------------------------------------------------
# 2. INTERFACE VISUAL EM XAML
# ------------------------------------------------------------------------------
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Central de Cadastro &amp; Auditoria de Acessos - IAM" 
        Height="850" Width="1440" 
        WindowStartupLocation="CenterScreen" 
        Background="#F8FAFC" FontFamily="Segoe UI">
    
    <Window.Resources>
        <Style TargetType="DataGridRow">
            <Style.Triggers>
                <DataTrigger Binding="{Binding Concluido}" Value="True">
                    <Setter Property="Background" Value="#F0FDF4"/>
                    <Setter Property="Foreground" Value="#166534"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding TipoProcesso}" Value="migracao">
                    <Setter Property="Background" Value="#FFFBEB"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding TipoProcesso}" Value="pendencia">
                    <Setter Property="Background" Value="#FEF2F2"/>
                </DataTrigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>

    <Grid Margin="8">
        <Grid.RowDefinitions>
            <RowDefinition Height="46"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="32"/>
        </Grid.RowDefinitions>

        <!-- CABEÇALHO -->
        <Border Grid.Row="0" Background="#002060" CornerRadius="4" Margin="0,0,0,6" Padding="12,0">
            <Grid>
                <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Text="🛡️ Central de Provisionamento &amp; Auditoria de Acessos" Foreground="White" FontSize="15" FontWeight="SemiBold" VerticalAlignment="Center"/>
                    <Border Background="#0070C0" CornerRadius="12" Margin="14,0,0,0" Padding="8,2">
                        <TextBlock x:Name="lblTotal" Text="0 Registros" Foreground="White" FontSize="11" FontWeight="Bold"/>
                    </Border>
                </StackPanel>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Center">
                    <Button x:Name="btnCopiarEmailsMfa" Content="📋 Copiar E-mails p/ MFA" Background="#0284C7" Foreground="White" FontWeight="Bold" Padding="8,4" Margin="0,0,6,0" BorderThickness="0" Cursor="Hand"/>
                    <Button x:Name="btnImportar" Content="📤 Importar (.csv)" Background="#E2E8F0" Foreground="#1E293B" FontWeight="Bold" Padding="8,4" Margin="0,0,6,0" BorderThickness="0" Cursor="Hand"/>
                    <Button x:Name="btnExportar" Content="📥 Exportar (.csv)" Background="#16A34A" Foreground="White" FontWeight="Bold" Padding="8,4" Margin="0,0,6,0" BorderThickness="0" Cursor="Hand"/>
                    <Button x:Name="btnLimparOk" Content="🗑️ Limpar Concluídos" Background="#DC2626" Foreground="White" FontWeight="Bold" Padding="8,4" BorderThickness="0" Cursor="Hand"/>
                </StackPanel>
            </Grid>
        </Border>

        <!-- CORPO PRINCIPAL -->
        <Grid Grid.Row="1">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="360"/>
                <ColumnDefinition Width="8"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <!-- FORMULÁRIO + PAINEL DE RESULTADO -->
            <Border Grid.Column="0" Background="White" BorderBrush="#CBD5E1" BorderThickness="1" CornerRadius="4" Padding="10">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <StackPanel>
                        <Grid Margin="0,0,0,8">
                            <TextBlock x:Name="lblTituloForm" Text="📝 FORMULÁRIO DE PROVISIONAMENTO" Foreground="#002060" FontWeight="Bold" FontSize="11" VerticalAlignment="Center"/>
                            <Button x:Name="btnNovoCadastro" Content="🔄 Limpar" Background="#E2E8F0" Foreground="#1E293B" FontWeight="Bold" FontSize="10" Padding="6,2" HorizontalAlignment="Right" BorderThickness="0" Cursor="Hand"/>
                        </Grid>

                        <TextBlock Text="C - Nome Completo" FontSize="10" FontWeight="Bold" Foreground="#0070C0"/>
                        <TextBox x:Name="txtNome" Height="24" Margin="0,2,0,5"/>

                        <Grid Margin="0,0,0,5">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="6"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <StackPanel Grid.Column="0">
                                <TextBlock Text="H - Matrícula" FontSize="10" FontWeight="Bold" Foreground="#64748B"/>
                                <TextBox x:Name="txtMatricula" Height="24" Margin="0,2,0,0"/>
                            </StackPanel>
                            <StackPanel Grid.Column="2">
                                <TextBlock Text="D - Logon UPN" FontSize="10" FontWeight="Bold" Foreground="#B45309"/>
                                <TextBox x:Name="txtLogon" Height="24" Margin="0,2,0,0"/>
                            </StackPanel>
                        </Grid>

                        <Grid Margin="0,0,0,5">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="6"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <StackPanel Grid.Column="0">
                                <TextBlock Text="E - Cargo" FontSize="10" FontWeight="Bold" Foreground="#64748B"/>
                                <TextBox x:Name="txtCargo" Height="24" Margin="0,2,0,0"/>
                            </StackPanel>
                            <StackPanel Grid.Column="2">
                                <TextBlock Text="F - Departamento" FontSize="10" FontWeight="Bold" Foreground="#64748B"/>
                                <TextBox x:Name="txtDepto" Height="24" Margin="0,2,0,0"/>
                            </StackPanel>
                        </Grid>

                        <Grid Margin="0,0,0,5">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="6"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <StackPanel Grid.Column="0">
                                <TextBlock Text="G - Empresa" FontSize="10" FontWeight="Bold" Foreground="#64748B"/>
                                <TextBox x:Name="txtEmpresa" Height="24" Margin="0,2,0,0"/>
                            </StackPanel>
                            <StackPanel Grid.Column="2">
                                <TextBlock Text="I - Documento RG" FontSize="10" FontWeight="Bold" Foreground="#64748B"/>
                                <TextBox x:Name="txtRg" Height="24" Margin="0,2,0,0"/>
                            </StackPanel>
                        </Grid>

                        <Grid Margin="0,0,0,5">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="6"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <StackPanel Grid.Column="0">
                                <TextBlock Text="J - CPF (Base Senha)" FontSize="10" FontWeight="Bold" Foreground="#64748B"/>
                                <TextBox x:Name="txtCpf" Height="24" Margin="0,2,0,0"/>
                            </StackPanel>
                            <StackPanel Grid.Column="2">
                                <TextBlock Text="K - Endereço (Office)" FontSize="10" FontWeight="Bold" Foreground="#64748B"/>
                                <TextBox x:Name="txtEndereco" Height="24" Margin="0,2,0,0"/>
                            </StackPanel>
                        </Grid>

                        <Grid Margin="0,0,0,5">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="6"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <StackPanel Grid.Column="0">
                                <TextBlock Text="L - Gerente (Matrícula)" FontSize="10" FontWeight="Bold" Foreground="#64748B"/>
                                <TextBox x:Name="txtGerente" Height="24" Margin="0,2,0,0"/>
                            </StackPanel>
                            <StackPanel Grid.Column="2">
                                <TextBlock Text="M - E-mail Licenciado?" FontSize="10" FontWeight="Bold" Foreground="#64748B"/>
                                <ComboBox x:Name="cbEmailLic" Height="24" Margin="0,2,0,0">
                                    <ComboBoxItem Content="s (Sim - Licenciar)" IsSelected="True"/>
                                    <ComboBoxItem Content="n (Não)"/>
                                </ComboBox>
                            </StackPanel>
                        </Grid>

                        <Grid Margin="0,0,0,5">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="6"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <StackPanel Grid.Column="0">
                                <TextBlock Text="N - Opção Perfil" FontSize="10" FontWeight="Bold" Foreground="#64748B"/>
                                <ComboBox x:Name="cbPerfil" Height="24" Margin="0,2,0,0">
                                    <ComboBoxItem Content="1 (Terceiro)" IsSelected="True"/>
                                    <ComboBoxItem Content="2 (Interno / Padrão)"/>
                                    <ComboBoxItem Content="3 (Filial Especial)"/>
                                    <ComboBoxItem Content="4 (Parceiro)"/>
                                </ComboBox>
                            </StackPanel>
                            <StackPanel Grid.Column="2">
                                <TextBlock Text="O - Região / Segmento" FontSize="10" FontWeight="Bold" Foreground="#64748B"/>
                                <ComboBox x:Name="cbMalha" Height="24" Margin="0,2,0,0">
                                    <ComboBoxItem Content="1 (Região Norte)" IsSelected="True"/>
                                    <ComboBoxItem Content="2 (Região Sul)"/>
                                    <ComboBoxItem Content="3 (Especial)"/>
                                </ComboBox>
                            </StackPanel>
                        </Grid>

                        <!-- P - CLOUD TIMESTAMP -->
                        <Grid Margin="0,0,0,5">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="6"/>
                                <ColumnDefinition Width="80"/>
                            </Grid.ColumnDefinitions>
                            <StackPanel Grid.Column="0">
                                <TextBlock Text="P - CLOUD TIMESTAMP" FontSize="10" FontWeight="Bold" Foreground="#0070C0"/>
                                <TextBox x:Name="txtCloudTimestamp" Height="24" Margin="0,2,0,0"/>
                            </StackPanel>
                            <StackPanel Grid.Column="2" VerticalAlignment="Bottom">
                                <Button x:Name="btnGerarTimestamp" Content="🕒 Gerar UTC" Background="#0284C7" Foreground="White" FontWeight="Bold" FontSize="10" Height="24" BorderThickness="0" Cursor="Hand"/>
                            </StackPanel>
                        </Grid>

                        <Grid Margin="0,0,0,5">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="6"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <StackPanel Grid.Column="0">
                                <TextBlock Text="Q - Processo" FontSize="10" FontWeight="Bold" Foreground="#64748B"/>
                                <ComboBox x:Name="cbTipoProcesso" Height="24" Margin="0,2,0,0">
                                    <ComboBoxItem Content="nova" IsSelected="True"/>
                                    <ComboBoxItem Content="migracao"/>
                                    <ComboBoxItem Content="pendencia"/>
                                </ComboBox>
                            </StackPanel>
                            <StackPanel Grid.Column="2">
                                <TextBlock Text="R - Chamado (ID)" FontSize="10" FontWeight="Bold" Foreground="#64748B"/>
                                <TextBox x:Name="txtChamado" Height="24" Margin="0,2,0,0"/>
                            </StackPanel>
                        </Grid>

                        <TextBlock Text="📌 Observações / Motivo da Pendência" FontSize="10" FontWeight="Bold" Foreground="#64748B" Margin="0,2,0,2"/>
                        <TextBox x:Name="txtNotas" Height="36" TextWrapping="Wrap" Margin="0,0,0,8"/>

                        <Button x:Name="btnAdicionar" Content="➕ Adicionar à Fila" Background="#002060" Foreground="White" FontWeight="Bold" Height="28" BorderThickness="0" Cursor="Hand" Margin="0,0,0,6"/>
                        <Button x:Name="btnColarLote" Content="📋 Colar Lote do Excel (18 Cols - TAB)" Background="#475569" Foreground="White" FontWeight="Bold" Height="24" BorderThickness="0" Cursor="Hand" Margin="0,0,0,10"/>

                        <!-- PAINEL DE MONITORAMENTO -->
                        <Border Background="#F1F5F9" BorderBrush="#CBD5E1" BorderThickness="1" CornerRadius="4" Padding="8" Margin="0,2,0,4">
                            <StackPanel>
                                <TextBlock Text="📊 MONITORAMENTO OPERACIONAL" Foreground="#002060" FontSize="10" FontWeight="Bold" Margin="0,0,0,6"/>
                                
                                <Grid Margin="0,0,0,6">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="4"/>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="4"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>

                                    <Border Grid.Column="0" Background="#DCFCE7" BorderBrush="#86EFAC" BorderThickness="1" CornerRadius="3" Padding="4,3">
                                        <StackPanel HorizontalAlignment="Center">
                                            <TextBlock Text="CRIADOS" FontSize="9" FontWeight="Bold" Foreground="#166534" HorizontalAlignment="Center"/>
                                            <TextBlock x:Name="lblDashSucesso" Text="0" FontSize="14" FontWeight="Bold" Foreground="#15803D" HorizontalAlignment="Center"/>
                                        </StackPanel>
                                    </Border>

                                    <Border Grid.Column="2" Background="#FEF9C3" BorderBrush="#FDE047" BorderThickness="1" CornerRadius="3" Padding="4,3">
                                        <StackPanel HorizontalAlignment="Center">
                                            <TextBlock Text="PENDENTES" FontSize="9" FontWeight="Bold" Foreground="#854D0E" HorizontalAlignment="Center"/>
                                            <TextBlock x:Name="lblDashPendentes" Text="0" FontSize="14" FontWeight="Bold" Foreground="#A16207" HorizontalAlignment="Center"/>
                                        </StackPanel>
                                    </Border>

                                    <Border Grid.Column="4" Background="#FEE2E2" BorderBrush="#FCA5A5" BorderThickness="1" CornerRadius="3" Padding="4,3">
                                        <StackPanel HorizontalAlignment="Center">
                                            <TextBlock Text="FALHAS" FontSize="9" FontWeight="Bold" Foreground="#991B1B" HorizontalAlignment="Center"/>
                                            <TextBlock x:Name="lblDashErros" Text="0" FontSize="14" FontWeight="Bold" Foreground="#DC2626" HorizontalAlignment="Center"/>
                                        </StackPanel>
                                    </Border>
                                </Grid>

                                <TextBlock Text="Log de Ocorrências / Auditoria:" FontSize="9" FontWeight="Bold" Foreground="#64748B" Margin="0,2,0,2"/>
                                <TextBox x:Name="txtLogErros" Height="68" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" IsReadOnly="True" FontSize="9" Background="White" Foreground="#475569" Text="Nenhum erro registrado até o momento."/>
                            </StackPanel>
                        </Border>

                    </StackPanel>
                </ScrollViewer>
            </Border>

            <!-- TABELA COM SCROLLBAR HORIZONTAL -->
            <Border Grid.Column="2" Background="White" BorderBrush="#CBD5E1" BorderThickness="1" CornerRadius="4" Padding="6">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="32"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>

                    <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,4">
                        <TextBlock Text="🔍 Filtro:" VerticalAlignment="Center" FontSize="11" Margin="0,0,6,0" Foreground="#64748B"/>
                        <TextBox x:Name="txtFiltro" Width="180" Height="22" VerticalAlignment="Center"/>
                        
                        <Button x:Name="btnValidarAD" Content="🔍 Validar no AD" Background="#0F766E" Foreground="White" FontWeight="Bold" FontSize="11" Padding="10,0" Margin="8,0,0,0" BorderThickness="0" Height="24" Cursor="Hand"/>
                        <Button x:Name="btnCriarLote" Content="⚡ Processar Criação no AD" Background="#16A34A" Foreground="White" FontWeight="Bold" FontSize="11" Padding="10,0" Margin="8,0,0,0" BorderThickness="0" Height="24" Cursor="Hand"/>
                        
                        <TextBlock Text="(➡️ Role para a direita para ver todas as colunas | Duplo clique para editar)" Foreground="#64748B" FontSize="10" FontStyle="Italic" VerticalAlignment="Center" Margin="12,0,0,0"/>
                    </StackPanel>

                    <DataGrid x:Name="gridUsuarios" Grid.Row="1" AutoGenerateColumns="False" CanUserAddRows="False" 
                              IsReadOnly="False" SelectionMode="Extended" SelectionUnit="CellOrRowHeader"
                              HeadersVisibility="Column" GridLinesVisibility="Horizontal" Background="White" FontSize="11"
                              HorizontalGridLinesBrush="#E2E8F0" VerticalGridLinesBrush="#E2E8F0"
                              ScrollViewer.HorizontalScrollBarVisibility="Auto"
                              ScrollViewer.VerticalScrollBarVisibility="Auto">
                        <DataGrid.ColumnHeaderStyle>
                            <Style TargetType="DataGridColumnHeader">
                                <Setter Property="Background" Value="#002060"/>
                                <Setter Property="Foreground" Value="White"/>
                                <Setter Property="FontWeight" Value="SemiBold"/>
                                <Setter Property="Padding" Value="8,4"/>
                                <Setter Property="BorderBrush" Value="#003380"/>
                                <Setter Property="BorderThickness" Value="0,0,1,0"/>
                            </Style>
                        </DataGrid.ColumnHeaderStyle>

                        <DataGrid.Columns>
                            <DataGridCheckBoxColumn Header="OK?" Binding="{Binding Concluido}" Width="38"/>
                            <DataGridCheckBoxColumn Header="MFA?" Binding="{Binding MfaAplicado}" Width="45"/>
                            <DataGridCheckBoxColumn Header="Lic.?" Binding="{Binding LicencaAplicada}" Width="45"/>
                            
                            <DataGridTextColumn Header="Status" Binding="{Binding TipoProcesso}" Width="90" FontWeight="Bold" IsReadOnly="True"/>
                            <DataGridTextColumn Header="H (Matrícula)" Binding="{Binding Matricula}" FontWeight="Bold" Width="95" IsReadOnly="True"/>
                            <DataGridTextColumn Header="C (Nome Completo)" Binding="{Binding NomeCompletoOrig}" Width="210" FontWeight="SemiBold" IsReadOnly="True"/>
                            <DataGridTextColumn Header="D (E-mail UPN)" Binding="{Binding UpnFinal}" Width="210" FontWeight="Bold" Foreground="#0070C0" IsReadOnly="True"/>
                            <DataGridTextColumn Header="SamAccount" Binding="{Binding SamAccount}" Width="100" IsReadOnly="True"/>

                            <DataGridTextColumn Header="E (Cargo)" Binding="{Binding Cargo}" Width="140" IsReadOnly="True"/>
                            <DataGridTextColumn Header="F (Departamento)" Binding="{Binding Depto}" Width="130" IsReadOnly="True"/>
                            <DataGridTextColumn Header="G (Empresa)" Binding="{Binding Empresa}" Width="110" IsReadOnly="True"/>
                            <DataGridTextColumn Header="J (CPF)" Binding="{Binding Cpf}" Width="110" IsReadOnly="True"/>
                            <DataGridTextColumn Header="I (RG)" Binding="{Binding Rg}" Width="100" IsReadOnly="True"/>
                            <DataGridTextColumn Header="K (Endereço/Office)" Binding="{Binding EnderecoCom}" Width="140" IsReadOnly="True"/>
                            <DataGridTextColumn Header="L (Gerente)" Binding="{Binding Gerente}" Width="100" IsReadOnly="True"/>
                            <DataGridTextColumn Header="M (Email Lic)" Binding="{Binding TemEmail}" Width="85" IsReadOnly="True"/>
                            <DataGridTextColumn Header="N (Opção)" Binding="{Binding Opcao}" Width="75" IsReadOnly="True"/>
                            <DataGridTextColumn Header="O (Região)" Binding="{Binding MalhaOp}" Width="75" IsReadOnly="True"/>
                            <DataGridTextColumn Header="P (Cloud Timestamp)" Binding="{Binding CloudTimestamp}" Width="160" IsReadOnly="True"/>
                            <DataGridTextColumn Header="R (Chamado)" Binding="{Binding Chamado}" FontWeight="Bold" Foreground="#002060" Width="120" IsReadOnly="True"/>
                            <DataGridTextColumn Header="Notas / Auditoria AD" Binding="{Binding Notas}" Width="320" IsReadOnly="True"/>
                        </DataGrid.Columns>
                    </DataGrid>
                </Grid>
            </Border>
        </Grid>

        <!-- STATUS BAR -->
        <Border Grid.Row="2" Background="#0F172A" CornerRadius="3" Margin="0,4,0,0" Padding="10,0">
            <TextBlock x:Name="lblStatus" Text="Central IAM pronta. Active Directory e Outlook conectados." Foreground="#CBD5E1" FontSize="11" VerticalAlignment="Center"/>
        </Border>
    </Grid>
</Window>
"@

# ------------------------------------------------------------------------------
# 3. CARREGAMENTO DOS CONTROLOS
# ------------------------------------------------------------------------------
$reader = New-Object System.Xml.XmlNodeReader $xaml
$janela = [System.Windows.Markup.XamlReader]::Load($reader)

if (Test-Path $caminhoIco) {
    try {
        $janela.Icon = [System.Windows.Media.Imaging.BitmapFrame]::Create([System.Uri]::new($caminhoIco))
    } catch {}
}

$lblTituloForm       = $janela.FindName("lblTituloForm")
$btnNovoCadastro     = $janela.FindName("btnNovoCadastro")
$txtNome             = $janela.FindName("txtNome")
$txtMatricula        = $janela.FindName("txtMatricula")
$txtLogon            = $janela.FindName("txtLogon")
$txtCargo            = $janela.FindName("txtCargo")
$txtDepto            = $janela.FindName("txtDepto")
$txtEmpresa          = $janela.FindName("txtEmpresa")
$txtRg               = $janela.FindName("txtRg")
$txtCpf              = $janela.FindName("txtCpf")
$txtEndereco         = $janela.FindName("txtEndereco")
$txtGerente          = $janela.FindName("txtGerente")
$cbEmailLic          = $janela.FindName("cbEmailLic")
$cbPerfil            = $janela.FindName("cbPerfil")
$cbMalha             = $janela.FindName("cbMalha")
$txtCloudTimestamp   = $janela.FindName("txtCloudTimestamp")
$btnGerarTimestamp   = $janela.FindName("btnGerarTimestamp")
$cbTipoProcesso      = $janela.FindName("cbTipoProcesso")
$txtChamado          = $janela.FindName("txtChamado")
$txtNotas            = $janela.FindName("txtNotas")
$btnAdicionar        = $janela.FindName("btnAdicionar")
$btnColarLote        = $janela.FindName("btnColarLote")
$btnValidarAD        = $janela.FindName("btnValidarAD")
$btnCriarLote        = $janela.FindName("btnCriarLote")
$btnCopiarEmailsMfa  = $janela.FindName("btnCopiarEmailsMfa")
$btnImportar         = $janela.FindName("btnImportar")
$btnExportar         = $janela.FindName("btnExportar")
$btnLimparOk         = $janela.FindName("btnLimparOk")
$txtFiltro           = $janela.FindName("txtFiltro")
$gridUsuarios        = $janela.FindName("gridUsuarios")
$lblTotal            = $janela.FindName("lblTotal")
$lblStatus           = $janela.FindName("lblStatus")

$lblDashSucesso      = $janela.FindName("lblDashSucesso")
$lblDashPendentes    = $janela.FindName("lblDashPendentes")
$lblDashErros        = $janela.FindName("lblDashErros")
$txtLogErros         = $janela.FindName("txtLogErros")

$script:indiceEdicao = -1
$listaUsuarios = New-Object System.Collections.ObjectModel.ObservableCollection[PSCustomObject]
$gridUsuarios.ItemsSource = $listaUsuarios

function AtualizarDashboard {
    $total = $listaUsuarios.Count
    $lblTotal.Text = "$total Registros"

    $sucessos = 0
    $pendentes = 0
    $erros = 0
    $mensagensErro = New-Object System.Collections.Generic.List[string]

    for ($i = 0; $i -lt $total; $i++) {
        $item = $listaUsuarios[$i]
        if ($item.Concluido) {
            $sucessos++
        } elseif ($item.Notas -like "*ERRO*") {
            $erros++
            $mensagensErro.Add("• $($item.Matricula) ($($item.NomeCompletoOrig)): $($item.Notas)")
        } elseif ($item.TipoProcesso -eq "pendencia" -or $item.Notas -like "*CONFLITO*") {
            $pendentes++
            if (-not [string]::IsNullOrWhiteSpace($item.Notas)) {
                $mensagensErro.Add("• $($item.Matricula) [Pendência]: $($item.Notas)")
            }
        }
    }

    $lblDashSucesso.Text = "$sucessos"
    $lblDashPendentes.Text = "$pendentes"
    $lblDashErros.Text = "$erros"

    if ($mensagensErro.Count -gt 0) {
        $txtLogErros.Text = $mensagensErro -join "`r`n"
    } else {
        $txtLogErros.Text = "Nenhuma falha cadastral ou de processamento registrada."
    }
}

# ------------------------------------------------------------------------------
# 4. FUNÇÕES DE SANEAMENTO E REGRAS DE NEGÓCIO
# ------------------------------------------------------------------------------
function Set-NormalizedText {
    param([string]$texto)
    if ([string]::IsNullOrWhiteSpace($texto)) { return "" }
    $semAcento = [System.Text.RegularExpressions.Regex]::Replace($texto.Normalize([System.Text.NormalizationForm]::FormD), '\p{M}', '')
    return $semAcento.Replace('ç', 'c').Replace('Ç', 'C')
}

function Limpar-Numeros ($str) {
    if ([string]::IsNullOrWhiteSpace($str)) { return "" }
    return [System.Text.RegularExpressions.Regex]::Replace($str, '\D', '')
}

function Gerar-TimestampUTC {
    return (Get-Date).ToUniversalTime().ToString("yyyyMMdd080000.0Z")
}

$btnGerarTimestamp.Add_Click({
    $txtCloudTimestamp.Text = Gerar-TimestampUTC
})

function Limpar-Formulario {
    $txtNome.Text = ""
    $txtMatricula.Text = ""
    $txtLogon.Text = ""
    $txtCargo.Text = ""
    $txtDepto.Text = ""
    $txtEmpresa.Text = ""
    $txtRg.Text = ""
    $txtCpf.Text = ""
    $txtEndereco.Text = ""
    $txtGerente.Text = ""
    $txtCloudTimestamp.Text = ""
    $txtChamado.Text = ""
    $txtNotas.Text = ""
    
    $cbEmailLic.SelectedIndex = 0
    $cbPerfil.SelectedIndex = 0
    $cbMalha.SelectedIndex = 0
    $cbTipoProcesso.SelectedIndex = 0

    $script:indiceEdicao = -1
    $btnAdicionar.Content = "➕ Adicionar à Fila"
    $btnAdicionar.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0, 32, 96))
    $lblTituloForm.Text = "📝 FORMULÁRIO DE PROVISIONAMENTO"
    $lblTituloForm.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0, 32, 96))
}

$btnNovoCadastro.Add_Click({
    Limpar-Formulario
    $lblStatus.Text = "Formulário pronto para inserção."
})

$gridUsuarios.Add_MouseDoubleClick({
    $selecionado = $gridUsuarios.SelectedItem
    if ($null -ne $selecionado) {
        $idx = $listaUsuarios.IndexOf($selecionado)
        if ($idx -ge 0) {
            $script:indiceEdicao = $idx
            
            $txtNome.Text           = $selecionado.NomeCompletoOrig
            $txtMatricula.Text      = $selecionado.Matricula
            $txtLogon.Text          = $selecionado.SamAccount
            $txtCargo.Text          = $selecionado.Cargo
            $txtDepto.Text          = $selecionado.Depto
            $txtEmpresa.Text        = $selecionado.Empresa
            $txtRg.Text             = $selecionado.Rg
            $txtCpf.Text            = $selecionado.Cpf
            $txtEndereco.Text       = $selecionado.EnderecoCom
            $txtGerente.Text        = $selecionado.Gerente
            $txtCloudTimestamp.Text = $selecionado.CloudTimestamp
            $txtChamado.Text        = $selecionado.Chamado
            $txtNotas.Text          = $selecionado.Notas

            if ($selecionado.TipoProcesso -eq "migracao") { $cbTipoProcesso.SelectedIndex = 1 }
            elseif ($selecionado.TipoProcesso -eq "pendencia") { $cbTipoProcesso.SelectedIndex = 2 }
            else { $cbTipoProcesso.SelectedIndex = 0 }

            $btnAdicionar.Content = "💾 Salvar Alterações"
            $btnAdicionar.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(217, 119, 6))
            $lblTituloForm.Text = "A EDITAR REGISTO: " + $selecionado.Matricula
            $lblTituloForm.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(217, 119, 6))
            $lblStatus.Text = "A editar: " + $selecionado.Matricula + " (" + $selecionado.UpnFinal + ")."
        }
    }
})

$txtNome.Add_LostFocus({
    if (-not [string]::IsNullOrWhiteSpace($txtNome.Text) -and [string]::IsNullOrWhiteSpace($txtLogon.Text)) {
        $partes = $txtNome.Text.Trim().Split(" ")
        $pNome = (Set-NormalizedText -texto ($partes[0])).ToLower()
        $sNome = if ($partes.Length -gt 1) { (Set-NormalizedText -texto ($partes[$partes.Length - 1])).ToLower() } else { $pNome }
        $txtLogon.Text = "$pNome.$sNome"
    }
})

$btnAdicionar.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtNome.Text) -and [string]::IsNullOrWhiteSpace($txtMatricula.Text)) {
        [System.Windows.MessageBox]::Show("Preencha pelo menos o Nome ou a Matrícula.", "Aviso", "OK", "Warning")
        return
    }

    $statusEscolhido = $cbTipoProcesso.Text
    $notas = $txtNotas.Text

    if ($statusEscolhido -ne "pendencia") {
        $notas = $notas.Replace("[🛑 CONFLITO UPN: HOMONIMO (CONTA ATIVA) - AJUSTE MANUAL DO LOGON]", "").Replace("[🛑 CONFLITO UPN: HOMONIMO (CONTA DESATIVADA) - AJUSTE MANUAL DO LOGON]", "").Replace("[Falta dados]", "").Trim()
    } elseif ([string]::IsNullOrWhiteSpace($txtCpf.Text) -or [string]::IsNullOrWhiteSpace($txtCargo.Text) -or [string]::IsNullOrWhiteSpace($txtGerente.Text)) {
        if (-not $notas.Contains("[Falta dados]")) { 
            $notas = "[Falta dados] " + $notas 
        }
    }

    $partesN = $txtNome.Text.Trim().Split(" ")
    $pNome = $partesN[0]
    $sNome = if ($partesN.Length -gt 1) { $partesN[$partesN.Length - 1] } else { $pNome }

    $op = $cbPerfil.Text.Substring(0,1)
    $malha = $cbMalha.Text.Substring(0,1)
    $sam = $txtLogon.Text.Trim()

    # Sufixos de e-mail parametrizáveis
    $domExt    = "ext.empresa.com.br"
    $domCorp   = "empresa.com.br"
    $domAlt    = "subsidiaria.com.br"
    $domCloud  = "cloud.empresa.com.br"

    $upnCalc = switch ($op) {
        "1" { "$sam@$domExt" }
        "3" { "$sam@$domAlt" }
        default { "$sam@$domCorp" }
    }

    if ($statusEscolhido -eq "migracao") {
        $upnCalc = "$sam@$domCloud"
    }

    $item = [PSCustomObject]@{
        Concluido        = $false
        MfaAplicado      = $false
        LicencaAplicada  = $false
        PrimeiroNomeOrig = $pNome
        SobreNomeOrig    = $sNome
        NomeCompletoOrig = $txtNome.Text.Trim()
        SamAccount       = $sam
        UpnFinal         = $upnCalc
        Cargo            = $txtCargo.Text.Trim()
        Depto            = $txtDepto.Text.Trim()
        Empresa          = $txtEmpresa.Text.Trim()
        Matricula        = $txtMatricula.Text.Trim()
        Rg               = $txtRg.Text.Trim()
        Cpf              = $txtCpf.Text.Trim()
        EnderecoCom      = $txtEndereco.Text.Trim()
        Gerente          = $txtGerente.Text.Trim()
        TemEmail         = if ($cbEmailLic.Text -like "*Sim*") { "s" } else { "n" }
        Opcao            = $op
        MalhaOp          = $malha
        CloudTimestamp   = $txtCloudTimestamp.Text.Trim()
        TipoProcesso     = $statusEscolhido
        Chamado          = $txtChamado.Text.Trim()
        Notas            = $notas
    }

    if ($script:indiceEdicao -ge 0 -and $script:indiceEdicao -lt $listaUsuarios.Count) {
        $item.Concluido = $listaUsuarios[$script:indiceEdicao].Concluido
        $item.MfaAplicado = $listaUsuarios[$script:indiceEdicao].MfaAplicado
        $item.LicencaAplicada = $listaUsuarios[$script:indiceEdicao].LicencaAplicada
        $listaUsuarios[$script:indiceEdicao] = $item
        $lblStatus.Text = "Registo " + $item.Matricula + " atualizado para: " + $statusEscolhido.ToUpper()
    } else {
        $listaUsuarios.Add($item)
        $lblStatus.Text = "Colaborador " + $item.Matricula + " adicionado à fila."
    }

    AtualizarDashboard
    Limpar-Formulario
})

# ------------------------------------------------------------------------------
# 5. AUDITORIA NO AD: CRUZA UPN, MATRÍCULA E CPF
# ------------------------------------------------------------------------------
$btnValidarAD.Add_Click({
    if ($listaUsuarios.Count -eq 0) {
        [System.Windows.MessageBox]::Show("A fila está vazia para validação.", "Aviso", "OK", "Information")
        return
    }

    $lblStatus.Text = "A auditar contas e a cruzar dados no Active Directory..."
    $qtdMigracao = 0
    $qtdHomonimos = 0

    for ($i = 0; $i -lt $listaUsuarios.Count; $i++) {
        $u = $listaUsuarios[$i]
        if ($u.Concluido) { continue }

        $cpfPlanilhaLimpo = Limpar-Numeros $u.Cpf
        $adUserEncontrado = $null

        if (-not [string]::IsNullOrWhiteSpace($u.Matricula)) {
            $adUserEncontrado = Get-ADUser -Filter "SamAccountName -eq '$($u.Matricula)' -or EmployeeID -eq '$($u.Matricula)'" `
                                           -Properties Enabled, Description, employeeID, extensionAttribute3 -ErrorAction SilentlyContinue
        }

        if ($null -eq $adUserEncontrado -and -not [string]::IsNullOrWhiteSpace($u.SamAccount)) {
            $adUserEncontrado = Get-ADUser -Filter "UserPrincipalName -like '$($u.SamAccount)*' -or SamAccountName -eq '$($u.SamAccount)'" `
                                           -Properties Enabled, Description, employeeID, extensionAttribute3 -ErrorAction SilentlyContinue
        }

        if ($null -ne $adUserEncontrado) {
            $statusConta = if ($adUserEncontrado.Enabled) { "CONTA ATIVA" } else { "CONTA DESATIVADA" }
            $cpfAdLimpo = Limpar-Numeros ($adUserEncontrado.Description + " " + $adUserEncontrado.extensionAttribute3)

            if (-not [string]::IsNullOrWhiteSpace($cpfPlanilhaLimpo) -and -not [string]::IsNullOrWhiteSpace($cpfAdLimpo) -and $cpfAdLimpo.Contains($cpfPlanilhaLimpo)) {
                $u.TipoProcesso = "migracao"
                $u.UpnFinal = "$($u.SamAccount)@cloud.empresa.com.br"
                $u.Notas = "[MIGRAÇÃO CONFIRMADA POR CPF | AD: $statusConta] " + $u.Notas
                $qtdMigracao++
            }
            elseif ([string]::IsNullOrWhiteSpace($cpfPlanilhaLimpo)) {
                $u.TipoProcesso = "pendencia"
                $u.Notas = "[CONTA NO AD ($statusConta) - NECESSÁRIO VALIDAR DOCUMENTO] " + $u.Notas
                $qtdMigracao++
            }
            else {
                $u.TipoProcesso = "pendencia"
                $u.Notas = "[🛑 CONFLITO UPN: HOMÓNIMO ($statusConta) - AJUSTE MANUAL DO LOGON] " + $u.Notas
                $qtdHomonimos++
            }
        }
    }

    $gridUsuarios.Items.Refresh()
    AtualizarDashboard
    
    $resumo = "Auditoria de AD Concluída:`n`n" +
              "• Migrações Confirmadas: $qtdMigracao`n" +
              "• Conflitos de Homónimos: $qtdHomonimos`n`n" +
              "Os valores foram atualizados no Painel de Monitorização."

    [System.Windows.MessageBox]::Show($resumo, "Auditoria IAM", "OK", "Information")
    $lblStatus.Text = "Auditoria finalizada: $qtdMigracao migração(ões) e $qtdHomonimos homónimo(s)."
})

function Validar-Migracao {
    param(
        [string]$processo,
        [string]$matricula,
        [string]$nome
    )

    if ($processo -eq "migracao") {
        $msg = "ATENÇÃO: PROCESSO DE MIGRAÇÃO DETETADO!`n`n" +
               "Colaborador: " + $matricula + " - " + $nome + "`n`n" +
               "CHECKLIST DE GOVERNANÇA:`n" +
               "1. Ações no Entra ID / Cloud concluídas?`n" +
               "2. Limpeza de ImmutableId executada?`n" +
               "3. Licenciamento e UPN validados?`n`n" +
               "Deseja prosseguir com a criação no AD?"

        $resultado = [System.Windows.MessageBox]::Show($msg, "Confirmação de Migração", "YesNo", "Warning")
        return ($resultado -eq "Yes")
    }
    return $true
}

# ------------------------------------------------------------------------------
# 6. CRIAÇÃO NO AD E INTEGRAÇÃO OUTLOOK
# ------------------------------------------------------------------------------
$btnCriarLote.Add_Click({
    $pendentes = New-Object System.Collections.Generic.List[PSCustomObject]
    for ($i = 0; $i -lt $listaUsuarios.Count; $i++) {
        if (-not $listaUsuarios[$i].Concluido) {
            $pendentes.Add($listaUsuarios[$i])
        }
    }

    if ($pendentes.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Nenhum utilizador pendente na fila.", "Aviso", "OK", "Information")
        return
    }

    $confirmarLote = [System.Windows.MessageBox]::Show("Deseja iniciar o provisionamento de $($pendentes.Count) conta(s) no Active Directory?", "Confirmação", "YesNo", "Question")
    if ($confirmarLote -ne "Yes") { return }

    $outlook = $null
    try {
        $outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
    } catch {
        try {
            $outlook = New-Object -ComObject Outlook.Application
        } catch {
            $outlook = $null
        }
    }

    $sucessos = 0
    $erros = 0
    $emailsGerados = 0

    for ($idxP = 0; $idxP -lt $pendentes.Count; $idxP++) {
        $reg = $pendentes[$idxP]

        if ($reg.TipoProcesso -eq "pendencia" -and $reg.Notas -like "*CONFLITO UPN: HOMONIMO*") {
            [System.Windows.MessageBox]::Show("O colaborador $($reg.Matricula) ($($reg.NomeCompletoOrig)) possui conflito de Homónimo!`n`nAjuste o logon antes de criar.", "Bloqueio de Homónimo", "OK", "Error")
            continue
        }

        if ($reg.TipoProcesso -eq "migracao") {
            if (-not (Validar-Migracao -processo "migracao" -matricula $reg.Matricula -nome $reg.NomeCompletoOrig)) {
                $lblStatus.Text = "Criação de $($reg.Matricula) suspensa."
                continue
            }
        }

        if ($reg.TipoProcesso -eq "pendencia") {
            $aviso = "O colaborador $($reg.Matricula) está marcado como PENDÊNCIA.`n`nDeseja criar mesmo assim?"
            $resp = [System.Windows.MessageBox]::Show($aviso, "Alerta", "YesNo", "Warning")
            if ($resp -ne "Yes") { continue }
        }

        $primeiroNomeLimpo  = (Set-NormalizedText -texto ($reg.PrimeiroNomeOrig))
        $sobreNomeLimpo     = (Set-NormalizedText -texto ($reg.SobreNomeOrig))
        $nomeCompletoLimpo  = (Set-NormalizedText -texto ($reg.NomeCompletoOrig))
        $nomeExibicaoAcento = $reg.NomeCompletoOrig

        $baseEmail = if ($reg.SamAccount) { $reg.SamAccount.ToLower() } else { ($primeiroNomeLimpo + "." + $sobreNomeLimpo).ToLower() }
        $matricula = $reg.Matricula

        $ext3 = switch ($reg.MalhaOp) {
            "1" { "Regiao Norte" }
            "2" { "Regiao Sul" }
            "3" { "Especial" }
            default { "Geral" }
        }

        $ouPadraoCorp = "OU=Usuarios,OU=Corporativo,$dominioBase"

        switch ($reg.Opcao) {
            "1" {
                $targetOU = switch ($ext3) {
                    "Especial"     { $ouTercEspecial }
                    "Regiao Sul"   { $ouTercSul }
                    default        { $ouTercNorte }
                }
                $fallbackOU   = $ouTercPadrao
                $mailPrimario = "$baseEmail@ext.empresa.com.br"
                $listaProxys  = @("sip:$mailPrimario", "SMTP:$mailPrimario", "smtp:$baseEmail@ext.subsidiaria.com.br")
                $descricao    = "Ativo - Prestador de Serviços"
                $upnFinal     = if ($reg.TipoProcesso -eq "migracao") { "$baseEmail@cloud.empresa.com.br" } else { $mailPrimario }
            }
            "3" {
                $targetOU     = "OU=UnidadeEspecial,OU=Usuarios,$dominioBase"
                $fallbackOU   = $ouPadraoCorp
                $mailPrimario = "$baseEmail@subsidiaria.com.br"
                $listaProxys  = @("sip:$mailPrimario", "SMTP:$mailPrimario")
                $descricao    = "Ativo - Unidade Especial"
                $upnFinal     = if ($reg.TipoProcesso -eq "migracao") { "$baseEmail@cloud.empresa.com.br" } else { $mailPrimario }
            }
            default {
                $targetOU     = "OU=Internos,OU=Usuarios,$dominioBase"
                $fallbackOU   = $ouPadraoCorp
                $mailPrimario = "$baseEmail@empresa.com.br"
                $descricao    = "Ativo - $ext3"
                $listaProxys  = @("sip:$mailPrimario", "SMTP:$mailPrimario", "smtp:$baseEmail@alias.empresa.com.br")
                $upnFinal     = if ($reg.TipoProcesso -eq "migracao") { "$baseEmail@cloud.empresa.com.br" } else { $mailPrimario }
            }
        }

        $reg.UpnFinal = $upnFinal

        $managerDN = $null
        $destinatarioFinal = ""
        if ($reg.Gerente) {
            try {
                $mgrObj = Get-ADUser -Identity $reg.Gerente -Properties DistinguishedName, EmailAddress -ErrorAction SilentlyContinue
                if ($mgrObj) {
                    $managerDN = $mgrObj.DistinguishedName
                    if ($mgrObj.EmailAddress) { $destinatarioFinal = $mgrObj.EmailAddress }
                }
            } catch {}
        }

        $cpfLimpo = Limpar-Numeros $reg.Cpf
        $senhaTxt = if ($cpfLimpo.Length -ge 4) { "Init-" + $cpfLimpo.Substring(0,4) } else { "Init-1234" }
        $senhaSec = ConvertTo-SecureString $senhaTxt -AsPlainText -Force

        $outrosAtributos = @{ "employeeID" = $matricula; "RG" = $reg.Rg; "CPF" = $reg.Cpf }
        if (-not [string]::IsNullOrWhiteSpace($reg.CloudTimestamp)) {
            $null = $outrosAtributos.Add("msDS-cloudExtensionAttribute1", $reg.CloudTimestamp)
        }
        if ($ext3) { $null = $outrosAtributos.Add("extensionAttribute3", $ext3) }
        if ($reg.TemEmail -eq "s") { $null = $outrosAtributos.Add("proxyAddresses", $listaProxys) }

        $parametros = @{
            Name                  = $nomeCompletoLimpo
            DisplayName           = $nomeExibicaoAcento
            GivenName             = $primeiroNomeLimpo
            Surname               = $sobreNomeLimpo
            SamAccountName        = $matricula 
            UserPrincipalName     = $upnFinal
            EmailAddress          = $mailPrimario
            Title                 = $reg.Cargo
            Department            = $reg.Depto
            Company               = $reg.Empresa
            Office                = $reg.EnderecoCom 
            Path                  = $targetOU
            AccountPassword       = $senhaSec
            Enabled               = $true
            ChangePasswordAtLogon = $true 
            Description           = $descricao
            OtherAttributes       = $outrosAtributos
        }

        if ($managerDN) {
            $parametros.Add("Manager", $managerDN)
        }

        $usuarioCriado = $false
        try {
            New-ADUser @parametros -ErrorAction Stop
            $usuarioCriado = $true
            $sucessos++
            [System.IO.File]::AppendAllText($logFile, "$(Get-Date) : SUCESSO - $($matricula) em ($targetOU)`r`n", [System.Text.Encoding]::UTF8)
        }
        catch {
            $msgErro = $_.Exception.Message

            if ($targetOU -ne $fallbackOU) {
                $parametros["Path"] = $fallbackOU
                try {
                    New-ADUser @parametros -ErrorAction Stop
                    $usuarioCriado = $true
                    $sucessos++
                    $targetOU = $fallbackOU
                    [System.IO.File]::AppendAllText($logFile, "$(Get-Date) : SUCESSO OU ALTERNATIVA - $($matricula)`r`n", [System.Text.Encoding]::UTF8)
                }
                catch {
                    $msgErro =$_.Exception.Message
                }
            }

            if (-not $usuarioCriado -and ($msgErro -match "já está em uso|already exists|em uso")) {
                $parametros["Name"] = "$nomeCompletoLimpo -$matricula"
                $parametros["Path"] = $fallbackOU

                try {
                    New-ADUser @parametros -ErrorAction Stop
                    $usuarioCriado = $true$sucessos++
                    [System.IO.File]::AppendAllText($logFile, "$(Get-Date) : SUCESSO CN COMPOSTO - $($matricula)`r`n", [System.Text.Encoding]::UTF8)
                }
                catch {
                    $erros++$reg.TipoProcesso = "pendencia"
                    $reg.Notas = "[ERRO CRÍTICO AD: " + $_.Exception.Message + "] " + $reg.Notas
                    [System.IO.File]::AppendAllText($logFile, "$(Get-Date) : ERRO - $($matricula) - $($_.Exception.Message)`r`n", [System.Text.Encoding]::UTF8)
                }
            }
            elseif (-not $usuarioCriado) {
                $erros++$reg.TipoProcesso = "pendencia"
                $reg.Notas = "[ERRO AD: $msgErro] " + $reg.Notas
                [System.IO.File]::AppendAllText($logFile, "$(Get-Date) : ERRO -$($matricula) -$msgErro`r`n", [System.Text.Encoding]::UTF8)
            }
        }

        if ($usuarioCriado) {
            try {
                if ($null -eq $outlook) {$outlook = New-Object -ComObject Outlook.Application -ErrorAction SilentlyContinue
                }

                if ($null -ne$outlook) {
                    $numChamado =$reg.Chamado
                    $assunto = if (-not [string]::IsNullOrWhiteSpace($numChamado)) {
                        "Solicitação $numChamado - Provisionamento de Conta Corporativa"
                    } else {
                        "Provisionamento de Conta Corporativa - Onboarding"
                    }

                    $licenciado = if ($reg.TemEmail -eq "s") { "Sim" } else { "Não Solicitado" }

                    $corpoHTML = @"
<html>
<body style="font-family: 'Segoe UI', Calibri, sans-serif;">
    <p>Prezado(a) Gestor(a),</p>
    <p>O utilizador corporativo foi criado com sucesso conforme os requisitos de conformidade e acessos.</p>
    
    <table style="border-collapse: collapse; width: 100%; border: 1px solid #002060;">
        <thead>
            <tr style="background-color: #002060; color: white; text-align: center; font-weight: bold;">
                <td style="padding: 5px; border: 1px solid #002060;">Matrícula</td>
                <td style="padding: 5px; border: 1px solid #002060;">Nome</td>
                <td style="padding: 5px; border: 1px solid #002060;">E-mail Ativo?</td>
                <td style="padding: 5px; border: 1px solid #002060;">Endereço de E-mail</td>
                <td style="padding: 5px; border: 1px solid #002060;">Senha Provisória</td>
            </tr>
        </thead>
        <tbody>
            <tr style="text-align: center; font-weight: bold; color: #0070C0;">
                <td style="padding: 5px; border: 1px solid #002060;">$matricula</td>
                <td style="padding: 5px; border: 1px solid #002060;">$nomeExibicaoAcento</td>
                <td style="padding: 5px; border: 1px solid #002060;">$licenciado</td>
                <td style="padding: 5px; border: 1px solid #002060; font-weight: normal; text-decoration: underline;">$mailPrimario</td>
                <td style="padding: 5px; border: 1px solid #002060;">Init-XXXX</td>
            </tr>
        </tbody>
    </table>

    <p style="color: #B91C1C; font-weight: bold; margin-top: 18px;">Regras de Ativação:</p>
    <ul>
        <li>A senha provisória utiliza o padrão <b>Init-XXXX</b>, onde XXXX são os primeiros 4 dígitos do documento registado.</li>
        <li>O primeiro acesso exige alteração imediata de credenciais via portal institucional: <a href="https://autoatendimento.empresa.com.br">https://autoatendimento.empresa.com.br</a></li>
        <li>A autenticação multifator (MFA) é obrigatória desde o primeiro logon.</li>
    </ul>

    <p>Atenciosamente,<br><b>Equipa de Gestão de Identidades e Acessos (IAM)</b></p>
</body>
</html>
"@
                    $mail =$outlook.CreateItem(0)
                    if ($destinatarioFinal -ne "") { $mail.To = $destinatarioFinal }$mail.CC = "gestaodeacessos@empresa.com.br"
                    $mail.Subject =$assunto
                    $mail.HTMLBody =$corpoHTML + $assinaturaHTML$mail.Save()
                    $mail.Close(0)$null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($mail)$emailsGerados++
                }
            } 
            catch {
                [System.IO.File]::AppendAllText($logFile, "$(Get-Date) : AVISO OUTLOOK - $($matricula) - $($_.Exception.Message)`r`n", [System.Text.Encoding]::UTF8)
            }

            $reg.Concluido =$true
            $reg.Notas = "[CRIADO NO AD COM SUCESSO] " + $reg.Notas
        }
    }

    $gridUsuarios.Items.Refresh()
    AtualizarDashboard

    $msgFinal = "Processamento Concluído!`n`n• Criados no AD: $sucessos`n• Rascunhos no Outlook: $emailsGerados`n• Ocorrências: $erros`n`nLog arquivado em:`n$logFile"
    [System.Windows.MessageBox]::Show($msgFinal, "Resultado do Processamento", "OK", "Information")
    $lblStatus.Text = "$sucessos conta(s) criadas. $emailsGerados notificação(ões) geradas."
})

# ------------------------------------------------------------------------------
# 7. IMPORTAÇÃO, EXPORTAÇÃO E TRATAMENTO DE CLIPBOARD
# ------------------------------------------------------------------------------
$btnColarLote.Add_Click({
    $textoClip = [System.Windows.Clipboard]::GetText()
    if ([string]::IsNullOrWhiteSpace($textoClip) -or (-not $textoClip.Contains("`t"))) {
        [System.Windows.MessageBox]::Show("Copie as colunas da folha de cálculo separadas por TAB.", "Aviso", "OK", "Information")
        return
    }

    $linhas =$textoClip -split "`r?\n"
    $adicionados = 0

    for ($idxL = 0; $idxL -lt $linhas.Length; $idxL++) {
        $l = $linhas[$idxL]
        if ([string]::IsNullOrWhiteSpace($l)) { continue }
        $d = $l -split "`t"
        if ($d.Length -ge 2) {$pNome = if ($d.Length -gt 0) {$d[0].Trim() } else { "" }
            $sNome = if ($d.Length -gt 1) {$d[1].Trim() } else { "" }
            $nComp = if ($d.Length -gt 2) {$d[2].Trim() } else { "$pNome$sNome" }
            $sam   = if ($d.Length -gt 3) {$d[3].Trim() } else { "" }
            $mat   = if ($d.Length -gt 7) {$d[7].Trim() } else { "" }
            $cTime = if ($d.Length -gt 15) {$d[15].Trim() } else { "" }
            $tipo  = if ($d.Length -gt 16 -and $d[16]) {$d[16].Trim().ToLower() } else { "nova" }
            $cham  = if ($d.Length -gt 17 -and $d[17]) {$d[17].Trim() } else { "" }

            $status =$tipo
            $notas = ""
            if ([string]::IsNullOrWhiteSpace($mat) -or $d.Length -lt 10) {$status = "pendencia"
                $notas = "[Dados cadastrais incompletos]"
            }

            $op    = if ($d.Length -gt 13 -and $d[13]) {$d[13].Trim() } else { "2" }
            $baseE = if ($sam) {$sam.ToLower() } else { ((Set-NormalizedText -texto $pNome) + "." + (Set-NormalizedText -texto $sNome)).ToLower() }

            $upnC = switch ($op) {
                "1" { "$baseE@ext.empresa.com.br" }
                "3" { "$baseE@subsidiaria.com.br" }
                default { "$baseE@empresa.com.br" }
            }

            if ($status -eq "migracao") {
                $upnC = "$baseE@cloud.empresa.com.br"
            }

            $obj = [PSCustomObject]@{
                Concluido        = $false
                MfaAplicado      = $false
                LicencaAplicada  = $false
                PrimeiroNomeOrig = $pNome
                SobreNomeOrig    = $sNome
                NomeCompletoOrig = $nComp
                SamAccount       = $sam
                UpnFinal         = $upnC
                Cargo            = if ($d.Length -gt 4) {$d[4].Trim() } else { "" }
                Depto            = if ($d.Length -gt 5) {$d[5].Trim() } else { "" }
                Empresa          = if ($d.Length -gt 6) {$d[6].Trim() } else { "" }
                Matricula        = $mat
                Rg               = if ($d.Length -gt 8) {$d[8].Trim() } else { "" }
                Cpf              = if ($d.Length -gt 9) {$d[9].Trim() } else { "" }
                EnderecoCom      = if ($d.Length -gt 10) {$d[10].Trim() } else { "" }
                Gerente          = if ($d.Length -gt 11) {$d[11].Trim() } else { "" }
                TemEmail         = if ($d.Length -gt 12) {$d[12].Trim().ToLower() } else { "" }
                Opcao            = $op
                MalhaOp          = if ($d.Length -gt 14) {$d[14].Trim().ToLower() } else { "1" }
                CloudTimestamp   = $cTime
                TipoProcesso     = $status
                Chamado          = $cham
                Notas            = $notas
            }

            $encontrado =$false
            for ($k = 0; $k -lt $listaUsuarios.Count; $k++) {
                if ($listaUsuarios[$k].Matricula -eq $mat -and -not [string]::IsNullOrWhiteSpace($mat)) {
                    $encontrado =$true
                    break
                }
            }

            if (-not $encontrado) {$listaUsuarios.Add($obj)$adicionados++
            }
        }
    }

    AtualizarDashboard
    $lblStatus.Text = "$adicionados registo(s) importados via colagem."
})

$btnCopiarEmailsMfa.Add_Click({$emailsCopiar = New-Object System.Collections.Generic.List[string]

    for ($i = 0; $i -lt$listaUsuarios.Count; $i++) {$u = $listaUsuarios[$i]
        if ($u.Concluido -and (-not$u.MfaAplicado -or ($u.TemEmail -eq "s" -and -not $u.LicencaAplicada))) {
            if (-not [string]::IsNullOrWhiteSpace($u.UpnFinal)) {
                $emailsCopiar.Add($u.UpnFinal)
            }
        }
    }

    if ($emailsCopiar.Count -eq 0) {
        for ($i = 0; $i -lt$listaUsuarios.Count; $i++) {$u = $listaUsuarios[$i]
            if ($u.Concluido -and -not [string]::IsNullOrWhiteSpace($u.UpnFinal)) {
                $emailsCopiar.Add($u.UpnFinal)
            }
        }
    }

    if ($emailsCopiar.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Nenhum e-mail de conta concluída disponível.", "Aviso", "OK", "Information")
        return
    }

    $textoClipboard =$emailsCopiar -join "`r`n"
    [System.Windows.Clipboard]::SetText($textoClipboard)
    [System.Windows.MessageBox]::Show("$($emailsCopiar.Count) e-mail(s) copiados para a Área de Transferência.", "Sucesso", "OK", "Information")
    $lblStatus.Text = "$($emailsCopiar.Count) e-mail(s) preparados para MFA/Licenciamento."
})

$btnExportar.Add_Click({
    if ($listaUsuarios.Count -eq 0) {
        [System.Windows.MessageBox]::Show("A fila está vazia para exportação.", "Aviso", "OK", "Information")
        return
    }

    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "Ficheiro CSV (*.csv)|*.csv"
    $saveDialog.FileName = "Base_Fila_IAM_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"

    if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $linhasCsv = New-Object System.Collections.Generic.List[string]$linhasCsv.Add("Concluido;MfaAplicado;LicencaAplicada;PrimeiroNome;Sobrenome;NomeCompleto;SamAccount;UpnFinal;Cargo;Depto;Empresa;Matricula;Rg;Cpf;EnderecoCom;Gerente;TemEmail;Opcao;MalhaOp;CloudTimestamp;TipoProcesso;Chamado;Notas")
        for ($i = 0; $i -lt $listaUsuarios.Count; $i++) {
            $u =$listaUsuarios[$i]$linha = "$($u.Concluido);$($u.MfaAplicado);$($u.LicencaAplicada);$($u.PrimeiroNomeOrig);$($u.SobreNomeOrig);$($u.NomeCompletoOrig);$($u.SamAccount);$($u.UpnFinal);$($u.Cargo);$($u.Depto);$($u.Empresa);$($u.Matricula);$($u.Rg);$($u.Cpf);$($u.EnderecoCom);$($u.Gerente);$($u.TemEmail);$($u.Opcao);$($u.MalhaOp);$($u.CloudTimestamp);$($u.TipoProcesso);$($u.Chamado);$($u.Notas)"
            $linhasCsv.Add($linha)
        }
        [System.IO.File]::WriteAllLines($saveDialog.FileName, $linhasCsv, [System.Text.Encoding]::UTF8)
        [System.Windows.MessageBox]::Show("Base exportada com sucesso!", "Sucesso", "OK", "Information")
        $lblStatus.Text = "Base arquivada em: $($saveDialog.FileName)"
    }
})

$btnImportar.Add_Click({$openDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openDialog.Filter = "Ficheiros CSV (*.csv)|*.csv"

    if ($openDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {$linhasArquivo = [System.IO.File]::ReadAllLines($openDialog.FileName, [System.Text.Encoding]::UTF8)$novos = 0

        for ($idx = 1; $idx -lt$linhasArquivo.Length; $idx++) {$linhaRaw = $linhasArquivo[$idx]
            if ([string]::IsNullOrWhiteSpace($linhaRaw)) { continue }
            $c =$linhaRaw -split ";"
            if ($c.Length -ge 12) {
                $mat =$c[11].Trim()
                
                $jaExiste =$false
                for ($m = 0; $m -lt $listaUsuarios.Count; $m++) {
                    if ($listaUsuarios[$m].Matricula -eq $mat -and -not [string]::IsNullOrWhiteSpace($mat)) {
                        $jaExiste =$true
                        break
                    }
                }

                if (-not $jaExiste) {$obj = [PSCustomObject]@{
                        Concluido        = [bool]($c[0] -like "*True*")
                        MfaAplicado      = [bool]($c[1] -like "*True*")
                        LicencaAplicada  = [bool]($c[2] -like "*True*")
                        PrimeiroNomeOrig = $c[3]
                        SobreNomeOrig    = $c[4]
                        NomeCompletoOrig = $c[5]
                        SamAccount       = $c[6]
                        UpnFinal         = $c[7]
                        Cargo            = $c[8]
                        Depto            = $c[9]
                        Empresa          = $c[10]
                        Matricula        = $mat
                        Rg               = $c[12]
                        Cpf              = $c[13]
                        EnderecoCom      = if ($c.Length -gt 14) {$c[14] } else { "" }
                        Gerente          = if ($c.Length -gt 15) {$c[15] } else { "" }
                        TemEmail         = if ($c.Length -gt 16) {$c[16] } else { "" }
                        Opcao            = if ($c.Length -gt 17) {$c[17] } else { "2" }
                        MalhaOp          = if ($c.Length -gt 18) {$c[18] } else { "1" }
                        CloudTimestamp   = if ($c.Length -gt 19) {$c[19] } else { "" }
                        TipoProcesso     = if ($c.Length -gt 20) {$c[20] } else { "nova" }
                        Chamado          = if ($c.Length -gt 21) {$c[21] } else { "" }
                        Notas            = if ($c.Length -gt 22) {$c[22] } else { "" }
                    }
                    $listaUsuarios.Add($obj)$novos++
                }
            }
        }

        AtualizarDashboard
        [System.Windows.MessageBox]::Show("$novos registos importados!", "Importação Concluída", "OK", "Information")
        $lblStatus.Text = "$novos registos carregados com sucesso."
    }
})

$btnLimparOk.Add_Click({$paraRemover = New-Object System.Collections.Generic.List[PSCustomObject]
    for ($j = 0; $j -lt $listaUsuarios.Count; $j++) {
        if ($listaUsuarios[$j].Concluido -eq $true) {$paraRemover.Add($listaUsuarios[$j])
        }
    }

    if ($paraRemover.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Nenhum registo com status de concluído.", "Informação", "OK", "Information")
        return
    }

    for ($n = 0; $n -lt $paraRemover.Count; $n++) {
        $null =$listaUsuarios.Remove($paraRemover[$n])
    }
    AtualizarDashboard
    $lblStatus.Text = "$($paraRemover.Count) registos concluídos removidos da vista."
})

# ------------------------------------------------------------------------------
# 8. EXIBIÇÃO DA JANELA PRINCIPAL
# ------------------------------------------------------------------------------
$janela.ShowDialog() | Out-Null

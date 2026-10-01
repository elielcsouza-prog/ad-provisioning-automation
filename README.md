# 🛡️ Central de Provisionamento e Auditoria de Identidades (IAM / IGA)

Solução modular em PowerShell para governança e automação do ciclo de vida de identidades (**Joiners**, **Migrações Híbridas** e **Auditoria Cadastral**), integrando Active Directory local, Microsoft Entra ID e notificações corporativas no Microsoft Outlook.

---

## 🏛️ Arquitetura do Projeto

O repositório está dividido em duas abordagens operacionais:

* 📁 **/cli (Versão 1.0 - Linha de Comando):** Script ágil em terminal para processamento em lote via colagem direta de dados tabulados, focado em operações rotineiras de Service Desk / N2.
* 📁 **/gui (Versão 2.0 - Central Gráfica Enterprise):** Aplicação visual completa construída em WPF (XAML), com dashboard analítico em tempo real, prevenção contra homônimos, validação cruzada no AD antes da criação e controles de importação/exportação.

---

## 🚀 Diferenciais Técnicos e Governança

* **Interface Visual com Dashboard em Tempo Real (WPF/XAML):** Monitoramento visual de contas criadas, falhas operacionais e pendências cadastrais com log de erros integrado.
* **Prevenção Ativa Contra Homônimos:** Validação em tempo de execução via cruzamento de UPN, Matrícula e CPF diretamente no Active Directory, bloqueando a criação duplicada de credenciais ativas ou inativas.
* **Normalização Unicode:** Higienização de acentuação e caracteres especiais (`FormD`) garantindo compatibilidade entre diretórios locais e nuvem.
* **Comunicação Automatizada (Outlook COM):** Geração de rascunhos de e-mail em formato HTML corporativo com orientações de primeiro acesso, políticas de MFA e diretrizes de senhas seguras.
* **Exportação e Portabilidade:** Suporte a importação e exportação de filas em formato `.csv` e cópia automatizada de e-mails para ativação de MFA no portal do Entra ID.

---

## 🛠️ Tecnologias Utilizadas

* **PowerShell 5.1 / 7+**
* **ActiveDirectory Module**
* **Windows Presentation Foundation (WPF / XAML)**
* **Microsoft Outlook COM Object**
* **System.Drawing & GDI+** (Geração dinâmica de identidade visual)

---

## 📋 Como Executar

### Versão Gráfica (Recomendada):
1. Navegue até a pasta `gui/`.
2. Execute o arquivo `CentralProvisionamentoIAM.ps1` no PowerShell como Administrador:
   ```powershell
   powershell.exe -ExecutionPolicy Bypass -File .\gui\CentralProvisionamentoIAM.ps1

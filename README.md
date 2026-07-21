# Automação de Provisionamento Híbrido: AD, Microsoft Graph & Notificação Corporativa

## 📌 Visão Geral (Contexto IAM/IGA)
Este projeto automatiza o fluxo de **Provisionamento de Identidades (Joiner Process)** e **Migração Híbrida**, integrando a criação de contas no **Active Directory local**, a gestão de atributos no **Microsoft Entra ID via Graph API** e a comunicação operacional via **Outlook**. Desenvolvido para padronizar a entrada e transição de colaboradores, garantindo governança, conformidade técnica e integridade de dados desde o primeiro dia.

## 🚀 Diferenciais Técnicos
* **Sincronização Híbrida Condicional (Entra ID / Graph API):** Identifica automaticamente processos de migração e realiza chamadas à API Microsoft Graph para limpeza do atributo `OnPremisesImmutableId`, evitando conflitos de sincronização no Azure AD Connect.
* **Normalização e Saneamento de Dados:** Função avançada (`Set-NormalizedText`) baseada em decomposição Unicode (FormD) para remoção de acentuação e caracteres especiais, garantindo compatibilidade com sistemas legados.
* **Segurança de Senha Provisória:** Lógica dinâmica para geração de credenciais provisórias baseadas em atributos seguros do usuário, com alteração obrigatória no primeiro logon (`ChangePasswordAtLogon = $true`).
* **Lógica Dinâmica de Unidades Organizacionais (OUs):** Alocação automática baseada na categoria do perfil (Interno vs. Terceiro/Prestador) e definição personalizada de proxies SMTP/SIP.
* **Comunicação Automatizada (UX):** Geração automática de e-mails em HTML formatados com tabela de dados e instruções de segurança (MFA), salvando diretamente nos **Rascunhos (Drafts)** do Outlook para validação prévia do analista.

## 🛠️ Tecnologias Utilizadas
* **PowerShell (Core):** Engine principal da automação e manipulação de objetos.
* **Módulo Active Directory:** Gestão de objetos de usuário e atributos estendidos (`extensionAttributes`, `employeeID`, etc.).
* **Microsoft Graph SDK (`Connect-MgGraph` / `Invoke-MgGraphRequest`):** Comunicação via REST API (método PATCH) com o Microsoft Entra ID.
* **Microsoft Outlook COM Object:** Automação e integração com o cliente de e-mail local.

## 📋 Como Utilizar
1. Copie a massa de dados do Excel (17 colunas separadas por Tab, incluindo a sinalização do tipo de processo: `nova` ou `migracao`).
2. Execute o script `ProvisionamentoLote.ps1`.
3. Cole os dados diretamente no console quando solicitado e pressione `Enter`.
4. O script identificará os cenários:
   * Se houver contas de **migração**, solicitará a conexão com o **Microsoft Graph** para reset do `ImmutableID`.
   * Se forem apenas **contas novas**, executará o fluxo 100% local.
5. As contas serão criadas no AD e os e-mails de onboarding serão salvos na pasta de **Rascunhos** do Outlook.

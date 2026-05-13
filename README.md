Automação de Provisionamento: AD & Notificação Corporativa
📌 Visão Geral (Contexto IAM/IGA)
Este projeto automatiza o fluxo de Provisionamento de Identidades (Joiner Process), integrando a criação técnica de contas no Active Directory com a comunicação operacional via Outlook. Como Analista de Gestão de Acessos II, desenvolvi esta solução para padronizar a entrada de novos colaboradores, garantindo que os atributos de governança sejam populados corretamente desde o primeiro dia.

🚀 Diferenciais Técnicos
Normalização de Dados: Função integrada para limpeza de caracteres especiais e acentuação (Unicode), evitando erros de sincronização em sistemas legados.

Segurança de Senha Provisória: Lógica para geração de senhas baseadas em atributos do usuário (ex: parte do CPF), forçando a alteração obrigatória no primeiro logon.

Lógica de Unidades Organizacionais (OU): Atribuição dinâmica baseada na categoria do colaborador (Interno ou Externo/Terceiro).

Experiência do Usuário (UX): Geração automática de e-mail em HTML para o gestor, contendo todas as instruções de acesso e suporte, salvando-o nos rascunhos para revisão.

🛠️ Tecnologias Utilizadas
PowerShell: Core da automação.

Módulo Active Directory: Gestão de objetos de usuário e atributos estendidos.

Microsoft Outlook COM Object: Interface para automação de notificações corporativas.

📋 Como Utilizar
Prepare os dados dos novos usuários (formato tabulado/Excel).

Execute o script ProvisionamentoLote.ps1.

Cole os dados diretamente no console quando solicitado.

O script criará as contas no AD e gerará os e-mails de boas-vindas na sua pasta de Rascunhos (Drafts).

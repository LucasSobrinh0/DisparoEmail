# Disparador de E mails do UOL

## Funcionalidade
Dispara e mails em lote utilizando contas do domínio UOL, a partir de um arquivo Excel.

## Como usar
Antes de tudo, converse com o criador do disparador.  
É necessário liberar o SMTP do e mail do UOL.  
O criador irá orientar e auxiliar nesse processo.

Depois disso:

• Informe o e mail do UOL  
• Informe a senha do e mail  
• Preencha o título do e mail  
• Selecione o arquivo Excel no formato XLSX  
• Escreva a mensagem que será enviada  

## Formato obrigatório do arquivo Excel
O arquivo deve estar no formato XLSX e seguir exatamente este modelo.

Linha 1 é o cabeçalho  
Os dados começam a partir da linha 2  

Exemplo de preenchimento:

| E MAILS           | ID ou MOV        | N° PROCESSO     |
|-------------------|------------------|-----------------|
| email1@uol.com.br | 66               | 0001234562024   |
| email2@uol.com.br | 105              | 0009876542024   |

O template oficial será disponibilizado para download.

## Variáveis disponíveis na mensagem
Para inserir dados do Excel no corpo do e mail, utilize:

Id ou movimentação do processo  
{{ movimento }}

Número do processo  
{{ numero_processo }}

## Exemplo de mensagem para despacho jurídico
Sugestão de texto objetivo e profissional para uso jurídico:

Prezados,

Informamos que houve a seguinte movimentação processual: {{ movimento }}.  
Processo de número {{ numero_processo }}.

Permanecemos à disposição para eventuais esclarecimentos.

Atenciosamente.

## Envio
Após preencher todos os campos e selecionar o arquivo correto, clique em Enviar.  
O sistema realizará o disparo automático dos e mails.

## Downloads
Baixar o programa: [DisparadorDeEmail.exe](https://raw.githubusercontent.com/LucasSobrinh0/DisparoEmail/main/files/DisparadorDeEmail.exe)  
Baixar o template Excel: [TEMPLATE.xlsx](https://raw.githubusercontent.com/LucasSobrinh0/DisparoEmail/main/files/TEMPLATE.xlsx)








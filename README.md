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

| E MAILS               | ID ou MOV        | N° PROCESSO        |
|-----------------------|------------------|--------------------|
| email1@uol.com.br     | Movimentação 01  | 0001234562024      |
| email2@uol.com.br     | Movimentação 02  | 0009876542024      |

O template oficial será disponibilizado para download.

## Variáveis disponíveis na mensagem
Para inserir dados do Excel no corpo do e mail, utilize:

Para id ou movimentação  
{{ movimento }}

Para número do processo  
{{ numero_processo }}

## Envio
Após preencher todos os campos e selecionar o arquivo correto, clique em Enviar.  
O sistema fará o disparo automático dos e mails.

## Downloads
Baixar o template Excel  
Baixar o programa


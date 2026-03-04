# Mapa indicador de Suprimentos 2026

Este projeto reune a interface web e a API Node responsavel por ler os dados dos fornecedores, gerar os feedbacks com IA (usando Google Gemini) e enviar automaticamente os e-mails.

## Como executar localmente
1. Instale as dependencias:
   ```bash
   npm install
   ```
2. Copie o arquivo `.env.example` para `.env` e ajuste as credenciais SMTP.
3. Inicie o servidor (ele tambem publica os arquivos estaticos):
   ```bash
   npm start
   ```
4. Acesse `http://localhost:4173/analise.html`, selecione um fornecedor e utilize o cartao "Envio automatico do feedback" para disparar o e-mail.

## Variaveis de ambiente
- `PORT`: porta HTTP usada pelo servidor (padrao 4173).
- `HOST`: interface que o Express vai escutar (`0.0.0.0` em hospedagens, `127.0.0.1` para uso local).
- `SMTP_HOST`/`SMTP_PORT`/`SMTP_SECURE`: configuracoes do servidor de e-mail.
- `SMTP_USER` e `SMTP_PASS` (ou `SMTP_PASSWORD`): credenciais de autenticacao.
- `MAIL_FROM`: remetente exibido no e-mail; se vazio, usa `SMTP_USER`.
- `EMAIL_API_TOKEN`: opcional. Quando definido, o endpoint `/api/send-email` exige `Authorization: Bearer <token>`.

## Hospedagem
- Defina `PORT` e `HOST` conforme o provedor (a maioria injeta `PORT` automaticamente e espera `HOST=0.0.0.0`).
- O servidor exposto por `npm start` ja serve todo o conteudo estatico e a API; basta apontar a raiz do repositorio como pasta de trabalho.
- Para ambientes que importam o aplicativo (ex.: tests), utilize `require('./server').startServer()`.

## Layout do e-mail
Os modelos enviados são objetivos e independem dos cards da interface. Cada um contém:
- Nome do fornecedor.
- Status consolidado.
- Data e valor da ultima avaliacao registrada.
- Feedback produzido pela IA.
- Lista de ocorrencias mais recentes (ou indicacao de ausencia).

Para os indicadores mensais, o relatório reúne a média global, distribuição por status, destaques (reprovados/atenção/excelência) e o resumo estratégico gerado pela IA. Ambos os fluxos podem ser disparados automaticamente pela própria tela de análise.

## Configuração da API de IA

O sistema utiliza a API do Google Gemini (gratuita) para gerar análises de fornecedores. Para habilitar as funcionalidades de IA:

1. Obtenha uma chave de API gratuita em: https://makersuite.google.com/app/apikey
2. Acesse a tela de análise (`analise.html`)
3. Clique no ícone de engrenagem (⚙️) para abrir as configurações
4. Cole sua chave do Google Gemini no campo "Chave Google Gemini"
5. Clique em "Aplicar chave"

A chave é armazenada apenas no navegador (localStorage) e não é enviada para nenhum servidor externo além da API do Google Gemini.

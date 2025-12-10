'use strict';

const path = require('path');
const express = require('express');
const nodemailer = require('nodemailer');
const dotenv = require('dotenv');

dotenv.config();

const app = express();
const PORT = Number(process.env.PORT || 4173);
const HOST = process.env.HOST || '0.0.0.0';
const EMAIL_TOKEN = (process.env.EMAIL_API_TOKEN || '').trim();
const EMAIL_REGEX = /^[^\s@]+@[^\s@]+\.[^\s@]+$/i;
const STATIC_DIR = __dirname;

app.use(express.json({ limit: '5mb' }));
app.use(express.static(STATIC_DIR));

let transporter;

function createTransporter() {
  if (process.env.SMTP_URL) {
    return nodemailer.createTransport(process.env.SMTP_URL);
  }

  const host = process.env.SMTP_HOST || process.env.SMTP_SERVER || process.env.MTP_SERVER;
  if (!host) {
    throw new Error('Configure SMTP_HOST (ou SMTP_SERVER/MTP_SERVER) ou SMTP_URL para habilitar o envio automatico.');
  }

  const port = Number(process.env.SMTP_PORT || 587);
  const secure = process.env.SMTP_SECURE ? process.env.SMTP_SECURE === 'true' : port === 465;
  const user = process.env.SMTP_USER;
  const pass = process.env.SMTP_PASS || process.env.SMTP_PASSWORD;
  const rejectUnauthorized = process.env.SMTP_TLS_REJECT_UNAUTHORIZED !== 'false';

  return nodemailer.createTransport({
    host,
    port,
    secure,
    auth: user && pass ? { user, pass } : undefined,
    tls: { rejectUnauthorized }
  });
}

function getTransporter() {
  if (!transporter) {
    transporter = createTransporter();
  }
  return transporter;
}

function getMailFrom() {
  return (process.env.MAIL_FROM || process.env.SMTP_USER || '').trim();
}

function buildPlainText(html) {
  if (!html) {
    return '';
  }
  return String(html)
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n\n')
    .replace(/<[^>]+>/g, '')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function validateToken(req) {
  if (!EMAIL_TOKEN) {
    return true;
  }
  const provided = req.get('authorization');
  return provided === 'Bearer ' + EMAIL_TOKEN;
}

app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

app.post('/api/send-email', async (req, res) => {
  try {
    if (!validateToken(req)) {
      return res.status(401).json({ error: 'Nao autorizado' });
    }

    const { to, subject, html, text, supplier, generatedAt } = req.body || {};

    if (!to || !EMAIL_REGEX.test(to)) {
      return res.status(400).json({ error: 'Destinatario invalido.' });
    }
    if (!subject || !html) {
      return res.status(400).json({ error: 'Assunto e conteudo sao obrigatorios.' });
    }

    const from = getMailFrom();
    if (!from) {
      return res
        .status(500)
        .json({ error: 'Configure MAIL_FROM ou SMTP_USER para definir o remetente dos e-mails.' });
    }

    const transporterInstance = getTransporter();
    const info = await transporterInstance.sendMail({
      from,
      to,
      subject,
      html,
      text: text && text.trim() ? text : buildPlainText(html),
      headers: {
        'X-Engeman-Supplier-Id': supplier?.id ? String(supplier.id) : undefined,
        'X-Engeman-Supplier-Code': supplier?.code ? String(supplier.code) : undefined,
        'X-Engeman-Generated-At': generatedAt || new Date().toISOString()
      }
    });

    res.json({ message: 'Email enviado com sucesso.', id: info.messageId });
  } catch (error) {
    console.error('[server:send-email]', error);
    res.status(500).json({
      error: 'Falha ao enviar e-mail automaticamente.',
      details: process.env.NODE_ENV === 'development' ? String(error.message || error) : undefined
    });
  }
});

app.get(/^(?!\/api\/).*/, (req, res) => {
  res.sendFile(path.join(STATIC_DIR, 'analise.html'));
});

function warmupTransporter() {
  try {
    const instance = getTransporter();
    instance
      .verify()
      .then(() => {
        console.log('[server] Servidor SMTP pronto para uso.');
      })
      .catch((error) => {
        console.warn('[server] Nao foi possivel verificar o SMTP:', error.message);
      });
  } catch (error) {
    console.warn('[server] SMTP ainda nao configurado:', error.message);
  }
}

function startServer(customPort) {
  const portToUse = Number(customPort || PORT);
  const host = HOST;
  const server = app.listen(portToUse, host, () => {
    const displayHost = host === '0.0.0.0' ? 'localhost' : host;
    console.log(`[server] Servidor iniciado em http://${displayHost}:${portToUse}`);
    warmupTransporter();
  });
  return server;
}

if (require.main === module) {
  startServer();
}

module.exports = app;
module.exports.startServer = startServer;

const https = require('https');

function buildSummaryMessage({ title, summary, runName }) {
  const lines = [];
  lines.push(`**${title}**`);
  lines.push(`- Run: ${runName}`);
  lines.push(`- Total checks: ${summary.total}`);
  lines.push(`- Passed: ${summary.passed}`);
  lines.push(`- Failed: ${summary.failed}`);

  if (summary.failed > 0) {
    lines.push('');
    lines.push('**Failure categories**');
    Object.entries(summary.categories || {}).forEach(([category, count]) => {
      lines.push(`- ${category}: ${count}`);
    });

    lines.push('');
    lines.push('**Top reasons**');
    Object.entries(summary.reasons || {})
      .sort((a, b) => b[1] - a[1])
      .slice(0, 8)
      .forEach(([reason, count]) => {
        lines.push(`- ${reason}: ${count}`);
      });
  }

  return lines.join('\n');
}

function postJson(urlString, payload) {
  return new Promise((resolve, reject) => {
    const url = new URL(urlString);
    const req = https.request(
      {
        protocol: url.protocol,
        hostname: url.hostname,
        port: url.port || (url.protocol === 'https:' ? 443 : 80),
        path: `${url.pathname}${url.search}`,
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Content-Length': Buffer.byteLength(payload),
        },
      },
      (res) => {
        let body = '';
        res.on('data', (chunk) => {
          body += chunk;
        });
        res.on('end', () => {
          if (res.statusCode && res.statusCode >= 200 && res.statusCode < 300) {
            resolve(body);
          } else {
            reject(new Error(`Request failed with status ${res.statusCode}: ${body}`));
          }
        });
      }
    );

    req.on('error', reject);
    req.write(payload);
    req.end();
  });
}

async function sendWebhookMessage({ webhookUrl, title, message }) {
  if (!webhookUrl) {
    throw new Error('Webhook URL is not configured.');
  }

  const payload = JSON.stringify({
    title,
    message,
    summary: title,
    timestamp: new Date().toISOString(),
  });

  console.log(`[NOTIFICATION] Sending webhook notification to ${webhookUrl}`);
  return postJson(webhookUrl, payload);
}

async function sendEmailMessage({ smtpConfig, to, subject, message }) {
  try {
    const nodemailer = require('nodemailer');
    const transporter = nodemailer.createTransport(smtpConfig);
    await transporter.sendMail({
      from: smtpConfig.from || smtpConfig.auth?.user,
      to,
      subject,
      text: message,
    });
  } catch (error) {
    throw new Error(`Email sending failed: ${error.message}`);
  }
}

async function sendMessage({ provider = 'webhook', webhookUrl, title, message, smtpConfig, to, subject }) {
  if (provider === 'webhook') {
    return sendWebhookMessage({ webhookUrl, title, message });
  }

  if (provider === 'email') {
    return sendEmailMessage({ smtpConfig, to, subject, message });
  }

  throw new Error(`Unsupported notification provider: ${provider}`);
}

module.exports = {
  buildSummaryMessage,
  sendMessage,
  sendWebhookMessage,
  sendEmailMessage,
};

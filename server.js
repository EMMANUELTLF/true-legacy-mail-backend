/**
 * TRUE LEGACY FINANCE LLC
 * Backend seguro para Microsoft Graph
 *
 * Requisitos:
 * - Node.js 18+
 * - Aplicación en Microsoft Entra ID
 * - Application permissions:
 *     Mail.Send
 *     Mail.ReadWrite
 * - Admin consent
 *
 * IMPORTANTE:
 * CLIENT_SECRET vive SOLO en .env / variables secretas del servidor.
 */

"use strict";

require("dotenv").config();

const express = require("express");
const cors = require("cors");

const app = express();

const PORT = Number(process.env.PORT || 3000);
const TENANT_ID = process.env.MS_TENANT_ID;
const CLIENT_ID = process.env.MS_CLIENT_ID;
const CLIENT_SECRET = process.env.MS_CLIENT_SECRET;
const SENDER_EMAIL =
  process.env.MS_SENDER_EMAIL || "servicios@truelegacyfinancepr.com";

const ALLOWED_ORIGIN = process.env.ALLOWED_ORIGIN || "*";

// Puede necesitar aumentarse si un agente lleva varios PDF grandes.
// En producción conviene limitarlo según su uso real.
app.use(express.json({ limit: "220mb" }));

app.use(
  cors({
    origin(origin, callback) {
      if (ALLOWED_ORIGIN === "*") return callback(null, true);
      if (!origin || origin === ALLOWED_ORIGIN) return callback(null, true);
      return callback(new Error("Origen no autorizado por CORS."));
    }
  })
);

app.get("/health", (req, res) => {
  res.json({
    ok: true,
    service: "true-legacy-mailer",
    sender: SENDER_EMAIL
  });
});

app.post("/api/send-agent-email", async (req, res) => {
  try {
    validateConfiguration();

    const payload = validatePayload(req.body);
    const token = await getGraphToken();

    // Se crea un draft para poder manejar correctamente tanto archivos
    // pequeños como archivos grandes mediante upload session.
    const draft = await createDraft(token, payload);

    try {
      for (const attachment of payload.attachments) {
        const bytes = Buffer.from(attachment.contentBase64, "base64");

        if (bytes.length !== attachment.size) {
          throw new Error(
            `El tamaño recibido de "${attachment.name}" no coincide con el declarado.`
          );
        }

        if (bytes.length >= 150 * 1024 * 1024) {
          throw new Error(
            `"${attachment.name}" alcanza o excede 150 MB y no puede procesarse con este módulo.`
          );
        }

        if (bytes.length < 3 * 1024 * 1024) {
          await addSmallAttachment(token, draft.id, attachment, bytes);
        } else {
          await addLargeAttachment(token, draft.id, attachment, bytes);
        }
      }

      await sendDraft(token, draft.id);

      res.json({
        ok: true,
        messageId: draft.id,
        accepted: true,
        to: payload.to,
        attachmentCount: payload.attachments.length
      });
    } catch (error) {
      // Si falla antes de enviar, intentamos borrar el draft para no dejar
      // borradores incompletos acumulándose.
      await deleteDraftQuietly(token, draft.id);
      throw error;
    }
  } catch (error) {
    console.error("SEND ERROR:", error);

    res.status(error.statusCode || 500).json({
      ok: false,
      error: error.message || "Error interno al enviar el correo."
    });
  }
});

function validateConfiguration() {
  const missing = [];

  if (!TENANT_ID) missing.push("MS_TENANT_ID");
  if (!CLIENT_ID) missing.push("MS_CLIENT_ID");
  if (!CLIENT_SECRET) missing.push("MS_CLIENT_SECRET");
  if (!SENDER_EMAIL) missing.push("MS_SENDER_EMAIL");

  if (missing.length) {
    throw new Error(
      `Faltan variables de entorno del servidor: ${missing.join(", ")}`
    );
  }
}

function validatePayload(body) {
  if (!body || typeof body !== "object") {
    throw badRequest("Payload inválido.");
  }

  const to = String(body.to || "").trim();
  const cc = String(body.cc || "").trim();
  const subject = String(body.subject || "").trim();
  const bodyText = String(body.bodyText || "").trim();
  const attachments = Array.isArray(body.attachments) ? body.attachments : [];

  if (!isValidEmail(to)) throw badRequest("Destinatario inválido.");
  if (cc && !isValidEmail(cc)) throw badRequest("CC inválido.");
  if (!subject) throw badRequest("El asunto está vacío.");
  if (!bodyText) throw badRequest("El cuerpo del mensaje está vacío.");
  if (!attachments.length) throw badRequest("No se recibieron PDF para adjuntar.");

  const cleanAttachments = attachments.map((item, index) => {
    const name = sanitizeAttachmentName(item.name || `documento-${index + 1}.pdf`);
    const contentBase64 = String(item.contentBase64 || "");
    const size = Number(item.size || 0);

    if (!name.toLowerCase().endsWith(".pdf")) {
      throw badRequest(`El archivo "${name}" no es PDF.`);
    }
    if (!contentBase64) {
      throw badRequest(`El archivo "${name}" llegó sin contenido.`);
    }
    if (!Number.isFinite(size) || size <= 0) {
      throw badRequest(`El archivo "${name}" tiene un tamaño inválido.`);
    }

    return {
      name,
      contentType: "application/pdf",
      contentBase64,
      size
    };
  });

  return {
    to,
    cc,
    subject: subject.slice(0, 255),
    bodyText,
    attachments: cleanAttachments
  };
}

function badRequest(message) {
  const error = new Error(message);
  error.statusCode = 400;
  return error;
}

async function getGraphToken() {
  const url =
    `https://login.microsoftonline.com/${encodeURIComponent(TENANT_ID)}` +
    "/oauth2/v2.0/token";

  const form = new URLSearchParams({
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials"
  });

  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: form
  });

  const data = await response.json();

  if (!response.ok || !data.access_token) {
    throw new Error(
      `No se pudo obtener token de Microsoft Graph: ${
        data.error_description || data.error || response.status
      }`
    );
  }

  return data.access_token;
}

async function createDraft(token, payload) {
  const message = {
    subject: payload.subject,
    body: {
      contentType: "Text",
      content: payload.bodyText
    },
    toRecipients: [
      {
        emailAddress: {
          address: payload.to
        }
      }
    ]
  };

  if (payload.cc) {
    message.ccRecipients = [
      {
        emailAddress: {
          address: payload.cc
        }
      }
    ];
  }

  return graphJson(
    token,
    `/users/${encodeURIComponent(SENDER_EMAIL)}/messages`,
    {
      method: "POST",
      body: message
    }
  );
}

async function addSmallAttachment(token, messageId, attachment, bytes) {
  await graphJson(
    token,
    `/users/${encodeURIComponent(SENDER_EMAIL)}/messages/` +
      `${encodeURIComponent(messageId)}/attachments`,
    {
      method: "POST",
      body: {
        "@odata.type": "#microsoft.graph.fileAttachment",
        name: attachment.name,
        contentType: attachment.contentType,
        contentBytes: bytes.toString("base64")
      }
    }
  );
}

async function addLargeAttachment(token, messageId, attachment, bytes) {
  const session = await graphJson(
    token,
    `/users/${encodeURIComponent(SENDER_EMAIL)}/messages/` +
      `${encodeURIComponent(messageId)}/attachments/createUploadSession`,
    {
      method: "POST",
      body: {
        AttachmentItem: {
          attachmentType: "file",
          name: attachment.name,
          size: bytes.length,
          contentType: attachment.contentType
        }
      }
    }
  );

  if (!session?.uploadUrl) {
    throw new Error(
      `Microsoft Graph no devolvió uploadUrl para "${attachment.name}".`
    );
  }

  // 10 x 320 KiB. Los fragmentos intermedios mantienen una
  // granularidad adecuada para cargas reanudables.
  const chunkSize = 320 * 1024 * 10;

  for (let start = 0; start < bytes.length; start += chunkSize) {
    const endExclusive = Math.min(start + chunkSize, bytes.length);
    const chunk = bytes.subarray(start, endExclusive);
    const endInclusive = endExclusive - 1;

    const response = await fetch(session.uploadUrl, {
      method: "PUT",
      headers: {
        "Content-Length": String(chunk.length),
        "Content-Range": `bytes ${start}-${endInclusive}/${bytes.length}`,
        "Content-Type": "application/octet-stream"
      },
      body: chunk
    });

    if (![200, 201, 202].includes(response.status)) {
      const text = await response.text();
      throw new Error(
        `Falló la carga de "${attachment.name}" ` +
        `(HTTP ${response.status}): ${text.slice(0, 500)}`
      );
    }
  }
}

async function sendDraft(token, messageId) {
  const url =
    "https://graph.microsoft.com/v1.0" +
    `/users/${encodeURIComponent(SENDER_EMAIL)}/messages/` +
    `${encodeURIComponent(messageId)}/send`;

  const response = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`
    }
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(
      `Microsoft Graph no aceptó el envío (HTTP ${response.status}): ${text.slice(0, 600)}`
    );
  }
}

async function deleteDraftQuietly(token, messageId) {
  try {
    const url =
      "https://graph.microsoft.com/v1.0" +
      `/users/${encodeURIComponent(SENDER_EMAIL)}/messages/` +
      `${encodeURIComponent(messageId)}`;

    await fetch(url, {
      method: "DELETE",
      headers: {
        Authorization: `Bearer ${token}`
      }
    });
  } catch (_) {
    // No interrumpe el error principal.
  }
}

async function graphJson(token, path, options = {}) {
  const response = await fetch(
    `https://graph.microsoft.com/v1.0${path}`,
    {
      method: options.method || "GET",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json"
      },
      body: options.body ? JSON.stringify(options.body) : undefined
    }
  );

  let data = null;
  const text = await response.text();

  if (text) {
    try {
      data = JSON.parse(text);
    } catch (_) {
      data = { raw: text };
    }
  }

  if (!response.ok) {
    const graphMessage =
      data?.error?.message ||
      data?.raw ||
      `HTTP ${response.status}`;

    const error = new Error(`Microsoft Graph: ${graphMessage}`);
    error.statusCode = response.status >= 400 && response.status < 500
      ? 400
      : 502;
    throw error;
  }

  return data;
}

function isValidEmail(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value).trim());
}

function sanitizeAttachmentName(value) {
  return String(value)
    .replace(/[\\/:*?"<>|]/g, "_")
    .trim()
    .slice(0, 180);
}

app.use((error, req, res, next) => {
  console.error("SERVER ERROR:", error);
  res.status(500).json({
    ok: false,
    error: error.message || "Error interno."
  });
});

app.listen(PORT, () => {
  console.log(`True Legacy mail backend activo en puerto ${PORT}`);
});

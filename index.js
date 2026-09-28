const { onRequest } = require("firebase-functions/v2/https");
const nodemailer = require("nodemailer");
const cors = require("cors")({ origin: true });

// Configuração do servidor SMTP do Outlook / Office 365
const transporter = nodemailer.createTransport({
  host: "smtp.office365.com",
  port: 587,
  secure: false, // TLS
  auth: {
    user: "henry.perissinotti@transjordano.com.br",
    pass: "SUA_SENHA_AQUI" // Substitua pela sua senha do Outlook
  },
  tls: {
    ciphers: "SSLv3",
    rejectUnauthorized: false
  }
});

exports.sendEmail = onRequest((req, res) => {
  cors(req, res, async () => {
    if (req.method !== "POST") {
      return res.status(405).json({ message: "Método não permitido" });
    }

    const { toEmail, subject, bodyContent } = req.body;

    if (!toEmail) {
      return res.status(400).json({ message: "O e-mail do destinatário é obrigatório." });
    }

    const mailOptions = {
      from: '"Henry Peressinotti" <henry.perissinotti@transjordano.com.br>',
      to: toEmail,
      subject: subject || "Notificação de Cadastro",
      html: bodyContent || "<p>Segue a documentação referente ao cadastro.</p>"
    };

    try {
      await transporter.sendMail(mailOptions);
      return res.status(200).json({ success: true, message: "E-mail enviado com sucesso!" });
    } catch (error) {
      console.error("Erro ao enviar e-mail:", error);
      return res.status(500).json({ success: false, error: error.toString() });
    }
  });
});
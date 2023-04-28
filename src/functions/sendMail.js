import nodemailer from "nodemailer";
import dotenv from "dotenv";
dotenv.config();
export async function sendEmail() {
  console.log("Enviando Email...");
  const transporter = nodemailer.createTransport({
    host: "smtp.gmail.com",
    port: 587,
    secure: false,
    auth: {
      user: process.env.EMAIL_SENDER,
      pass: process.env.EMAIL_SENDER_PASS,
    },
  });

  const today = new Date();
  const day = today.getDate().toString().padStart(2, "0");
  const month = (today.getMonth() + 1).toString().padStart(2, "0");
  const date = `${day}.${month}`;

  const opcoesEmail = {
    from: process.env.EMAIL_SENDER,
    to: process.env.EMAIL_ADDRESSEE,
    subject: "Arquivo Excel formatado",
    text: "Segue em anexo o TESTE do arquivo Excel formatado.",
    attachments: [
      {
        filename: `LUC - ${date}.xlsx`,
        path: "src/Excel/LUC.xlsx",
      },
    ],
  };

  transporter.sendMail(opcoesEmail, (error, info) => {
    if (error) {
      console.error("Erro ao enviar o e-mail:", error);
    } else {
      console.log("E-mail enviado com sucesso:", info.response);
    }
  });
}

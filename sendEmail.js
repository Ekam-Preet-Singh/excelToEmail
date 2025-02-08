require("dotenv").config();
const nodemailer = require("nodemailer");
const { google } = require("googleapis");
const fs = require("fs");
const path = require("path");

// Load environment variables
const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const REDIRECT_URI = process.env.REDIRECT_URI;
const REFRESH_TOKEN = process.env.REFRESH_TOKEN;
const SENDER_EMAIL = process.env.MY_EMAIL;
const SENDER_NAME = process.env.MY_NAME;

// Create an OAuth2 client
const oAuth2Client = new google.auth.OAuth2(
  CLIENT_ID,
  CLIENT_SECRET,
  REDIRECT_URI
);
oAuth2Client.setCredentials({ refresh_token: REFRESH_TOKEN });

const sendEmailsFromFile = async (filePath) => {
  console.log(`🚀 Starting email sending process from: ${filePath}`);

  try {
    if (!fs.existsSync(filePath)) {
      throw new Error(`File not found: ${filePath}`);
    }

    const accessToken = await oAuth2Client.getAccessToken();
    if (!accessToken.token) {
      throw new Error("Failed to retrieve access token");
    }

    const transporter = nodemailer.createTransport({
      service: "gmail",
      auth: {
        type: "OAuth2",
        user: SENDER_EMAIL,
        clientId: CLIENT_ID,
        clientSecret: CLIENT_SECRET,
        refreshToken: REFRESH_TOKEN,
        accessToken: accessToken.token,
      },
    });

    const jsonData = JSON.parse(fs.readFileSync(filePath, "utf-8"));
    if (!jsonData || typeof jsonData !== "object") {
      throw new Error("Invalid JSON file format");
    }

    let totalEmails = 0;

    for (const company in jsonData) {
      for (const candidate of jsonData[company]) {
        let { Name, Email, Subject, EmailBody } = candidate;
        if (!Name || !Email || !Subject || !EmailBody) {
          console.warn(
            `⚠️ Skipping invalid candidate data: ${JSON.stringify(candidate)}`
          );
          continue;
        }
        EmailBody = EmailBody.replace(/\n/g, "<br>");

        const mailOptions = {
          from: `${SENDER_NAME} <${SENDER_EMAIL}>`,
          to: Email,
          subject: Subject,
          html: EmailBody,
        };

        try {
          await transporter.sendMail(mailOptions);
          console.log(`✅ Email sent successfully to ${Name} at ${Email}`);
          totalEmails++;
        } catch (error) {
          console.error(
            `❌ Failed to send email to ${Name} (${Email}):`,
            error.message
          );
        }
      }
    }

    console.log(
      `🎉 Email sending process completed. Total Emails Sent: ${totalEmails}`
    );
  } catch (error) {
    console.error("❌ Error sending emails:", error.message || error);
  }
};

const filePath = path.join(__dirname, "output_files", "output_part_1.json");
sendEmailsFromFile(filePath);

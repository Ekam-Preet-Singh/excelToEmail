const nodemailer = require("nodemailer");
const { google } = require("googleapis");
const fs = require("fs");
const path = require("path");
require("dotenv").config();

// OAuth2 credentials from environment variables
const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const REDIRECT_URI = "https://developers.google.com/oauthplayground";
const REFRESH_TOKEN = process.env.REFRESH_TOKEN;
const USER_EMAIL = process.env.MY_EMAIL;
const USER_NAME = process.env.MY_NAME;

// Create an OAuth2 client
const oAuth2Client = new google.auth.OAuth2(
  CLIENT_ID,
  CLIENT_SECRET,
  REDIRECT_URI
);
oAuth2Client.setCredentials({ refresh_token: REFRESH_TOKEN });

/**
 * Sends emails from a given JSON file.
 * @param {string} filePath - The path to the JSON file containing email data.
 */
const sendEmailsFromFile = async (filePath) => {
  console.log(`🚀 Starting email sending process from: ${filePath}`);

  try {
    // Check if file exists
    if (!fs.existsSync(filePath)) {
      throw new Error(`File not found: ${filePath}`);
    }

    // Read and parse JSON file
    let jsonData;
    try {
      const fileContent = fs.readFileSync(filePath, "utf-8");
      jsonData = JSON.parse(fileContent);
    } catch (error) {
      throw new Error(`Failed to read or parse JSON file: ${error.message}`);
    }

    // Get a new access token
    let accessToken;
    try {
      accessToken = await oAuth2Client.getAccessToken();
      if (!accessToken.token) {
        throw new Error("Failed to retrieve access token");
      }
    } catch (error) {
      throw new Error(`OAuth2 authentication failed: ${error.message}`);
    }

    // Create the Nodemailer transporter with OAuth2
    const transporter = nodemailer.createTransport({
      service: "gmail",
      auth: {
        type: "OAuth2",
        user: USER_EMAIL,
        clientId: CLIENT_ID,
        clientSecret: CLIENT_SECRET,
        refreshToken: REFRESH_TOKEN,
        accessToken: accessToken.token,
      },
    });

    let totalEmails = 0;

    for (const company in jsonData) {
      for (const candidate of jsonData[company]) {
        try {
          let { Name, Email, Subject, EmailBody } = candidate;

          // Validate required fields
          if (!Email || !Subject || !EmailBody) {
            console.warn(
              `⚠️ Skipping email to ${
                Name || "Unknown"
              } due to missing required fields.`
            );
            continue;
          }

          EmailBody = EmailBody.replace(/\n/g, "<br>"); // Format email body

          const mailOptions = {
            from: `${USER_NAME} <${USER_EMAIL}>`,
            to: Email,
            subject: Subject,
            html: EmailBody,
          };

          await transporter.sendMail(mailOptions);
          console.log(`✅ Email sent successfully to ${Name} at ${Email}`);
          totalEmails++;
        } catch (error) {
          console.error(
            `❌ Failed to send email: ${error.message}`,
            error.stack
          );
        }
      }
    }

    console.log(
      `🎉 Email sending process completed. Total Emails Sent: ${totalEmails}`
    );
  } catch (error) {
    console.error("❌ Critical Error:", error.message, error.stack);
  }
};

// Run the function with the first output JSON file
const filePath = path.join(
  __dirname,
  process.env.OUTPUT_DIR,
  "output_part_1.json"
);
sendEmailsFromFile(filePath);

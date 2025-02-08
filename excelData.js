require("dotenv").config();
const xlsx = require("xlsx");
const fs = require("fs");
const path = require("path");

/**
 * Validates required environment variables
 */
const validateEnvVariables = () => {
  const requiredVars = [
    "MY_NAME",
    "MY_EMAIL",
    "MY_PHONE",
    "RESUME_LINK",
    "FILE_PATH",
    "OUTPUT_DIR",
  ];
  const missingVars = requiredVars.filter((key) => !process.env[key]);

  if (missingVars.length > 0) {
    throw new Error(
      `Missing required environment variables: ${missingVars.join(", ")}`
    );
  }
};

/**
 * Converts an Excel file to a structured JSON format for sending emails.
 * @param {string} filePath - Path to the Excel file.
 * @returns {object} - Grouped email data by company.
 */
const convertExcelToJson = (filePath) => {
  try {
    // Validate file existence
    if (!fs.existsSync(filePath)) {
      throw new Error(`Excel file not found: ${filePath}`);
    }

    // Read the Excel workbook
    const workbook = xlsx.readFile(filePath);

    // Ensure at least one sheet exists
    if (workbook.SheetNames.length < 2) {
      throw new Error("Excel file does not contain the expected sheets.");
    }

    const sheetName = workbook.SheetNames[1]; // Adjust this if necessary
    const sheet = workbook.Sheets[sheetName];

    // Convert sheet data to JSON
    const rawData = xlsx.utils.sheet_to_json(sheet);

    if (rawData.length === 0) {
      throw new Error("No data found in the selected Excel sheet.");
    }

    // Personal details from .env
    const { MY_NAME, MY_EMAIL, MY_PHONE, RESUME_LINK } = process.env;

    // Organize data by company
    const groupedData = rawData.reduce((accumulator, row) => {
      const { Company, Name, Email, Title } = row;

      if (!Company || !Name || !Email || !Title) {
        console.warn(
          `⚠️ Skipping entry due to missing fields: ${JSON.stringify(row)}`
        );
        return accumulator;
      }

      const subject = `Experienced Mobile Developer | React Native & Kotlin | Excited to Join ${Company}`;

      const emailBody = `Dear ${Name},

I hope this email finds you well.

My name is <b>${MY_NAME}</b>, a passionate Mobile Developer specializing in <b>React Native</b> and <b>Kotlin</b> with over a year of experience in building high-performance applications. I am eager to explore opportunities at <b>${Company}</b> where I can contribute my expertise to drive innovation in mobile technology.

<b>📌 Key Highlights:</b>
- <b>Experience:</b> 1+ years
- <b>Current CTC:</b> ₹1.3 LPA | <b>Expected CTC:</b> ₹4-6 LPA
- <b>Notice Period:</b> 30 days (Negotiable)
- <b>Location:</b> Mohali, Punjab

<b>💡 Technical Skills:</b>
✔ React Native | Kotlin | Firebase | SQLite | Supabase
✔ MVVM Architecture | API Integration | Performance Optimization
✔ Git | Android Studio | Play Store & App Store Deployments

<b>📂 Professional Experience:</b>
🚀 <b>Software Engineer | Macrew Technologies Pvt. Ltd. (Jan 2024 – Present)</b>
- Developed scalable cross-platform applications using React Native & Kotlin.
- Optimized API integrations for enhanced performance & security.
- Successfully deployed applications on Google Play & App Store.

<b>🌟 Notable Projects:</b>
- <b>Mining App</b> – Offline-first architecture with SQLite & Supabase.
- <b>Music Streaming App</b> – Built a native Android application with real-time streaming.
- <b>Estate Planning App</b> – Integrated in-app purchases in a React Native application.

I am enthusiastic about the opportunity to be a part of <b>${Company}</b> and would love to discuss how my skills align with your needs.

📄 <b>Resume:</b> <a href="${RESUME_LINK}">View Here</a>

Looking forward to your response.

Best Regards,  
<b>${MY_NAME}</b>  
📞 ${MY_PHONE}  
✉️ ${MY_EMAIL}`;

      if (!accumulator[Company]) {
        accumulator[Company] = [];
      }

      accumulator[Company].push({
        Name,
        Email,
        Title,
        ResumeLink: RESUME_LINK,
        Subject: subject,
        EmailBody: emailBody,
      });
      return accumulator;
    }, {});

    return groupedData;
  } catch (error) {
    console.error("❌ Error processing Excel file:", error.message);
    return {};
  }
};

/**
 * Splits JSON data into multiple files
 * @param {object} jsonData - The grouped data from Excel.
 * @param {string} outputDir - Directory to save JSON files.
 * @param {number} chunkSize - Number of companies per file.
 */
const splitAndSaveJsonFiles = (jsonData, outputDir, chunkSize = 10) => {
  try {
    const totalCompanies = Object.keys(jsonData).length;
    if (totalCompanies === 0) {
      throw new Error("No valid data found to process.");
    }

    console.log(`✅ Total Companies: ${totalCompanies}`);
    console.log(
      `✅ Total Candidates: ${Object.values(jsonData).flat().length}`
    );

    const companyNames = Object.keys(jsonData);
    const totalFiles = Math.ceil(totalCompanies / chunkSize);

    // Create output directory if it doesn't exist
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir);
    }

    for (let i = 0; i < totalFiles; i++) {
      const chunk = companyNames.slice(i * chunkSize, (i + 1) * chunkSize);
      const chunkData = chunk.reduce((acc, company) => {
        acc[company] = jsonData[company];
        return acc;
      }, {});

      const fileName = path.join(outputDir, `output_part_${i + 1}.json`);

      try {
        fs.writeFileSync(fileName, JSON.stringify(chunkData, null, 2));
        console.log(`✅ Created file: ${fileName} (${chunk.length} companies)`);
      } catch (fileError) {
        console.error(
          `❌ Failed to write JSON file: ${fileName}`,
          fileError.message
        );
      }
    }

    console.log("🎉 All JSON files have been successfully created!");
  } catch (error) {
    console.error("❌ Error splitting and saving JSON files:", error.message);
  }
};

/**
 * Main execution function
 */
(() => {
  try {
    validateEnvVariables(); // Validate environment variables

    const filePath = process.env.FILE_PATH;
    const outputDir = process.env.OUTPUT_DIR;

    const jsonData = convertExcelToJson(filePath);
    splitAndSaveJsonFiles(jsonData, outputDir);
  } catch (error) {
    console.error("❌ Critical Error:", error.message);
  }
})();

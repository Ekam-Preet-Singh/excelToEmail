require("dotenv").config();
const xlsx = require("xlsx");
const fs = require("fs");
const path = require("path");

const convertExcelToJson = (filePath) => {
  try {
    // Read the Excel workbook
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[1];
    const sheet = workbook.Sheets[sheetName];

    // Convert sheet data to JSON
    const rawData = xlsx.utils.sheet_to_json(sheet);

    // Personal details from .env
    const myName = process.env.MY_NAME;
    const myEmail = process.env.MY_EMAIL;
    const myPhone = process.env.MY_PHONE;
    const resumeLink = process.env.RESUME_LINK;

    // Organize data by company
    const groupedData = rawData.reduce((accumulator, row) => {
      const { Company, Name, Email, Title } = row;
      const subject = `Experienced Mobile Developer | React Native & Kotlin | Excited to Join ${Company}`;

      const emailBody = `Dear ${Name},

I hope this email finds you well.

My name is <b>${myName}</b>, a passionate Mobile Developer specializing in <b>React Native</b> and <b>Kotlin</b> with over a year of experience in building high-performance applications. I am eager to explore opportunities at <b>${Company}</b> where I can contribute my expertise to drive innovation in mobile technology.

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

📄 <b>Resume:</b> <a href="${resumeLink}">View Here</a>

Looking forward to your response.

Best Regards,
<b>${myName}</b>
📞 ${myPhone}
✉️ ${myEmail}`;

      if (!accumulator[Company]) {
        accumulator[Company] = [];
      }

      accumulator[Company].push({
        Name,
        Email,
        Title,
        ResumeLink: resumeLink,
        Subject: subject,
        EmailBody: emailBody,
      });
      return accumulator;
    }, {});

    return groupedData;
  } catch (error) {
    console.error("Error processing Excel file:", error);
    return {};
  }
};

// Main Execution
(() => {
  const filePath = process.env.FILE_PATH;
  const jsonData = convertExcelToJson(filePath);

  const totalCompanies = Object.keys(jsonData).length;
  console.log(`✅ Total Companies: ${totalCompanies}`);
  console.log(`✅ Total Candidates: ${Object.values(jsonData).flat().length}`);

  const chunkSize = 10; // Maximum companies per file
  const companyNames = Object.keys(jsonData);
  const totalFiles = Math.ceil(totalCompanies / chunkSize);

  // Create an output folder if it doesn't exist
  const outputDir = process.env.OUTPUT_DIR;
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir);
  }

  // Split JSON data into multiple files
  for (let i = 0; i < totalFiles; i++) {
    const chunk = companyNames.slice(i * chunkSize, (i + 1) * chunkSize);
    const chunkData = chunk.reduce((acc, company) => {
      acc[company] = jsonData[company];
      return acc;
    }, {});

    const fileName = path.join(outputDir, `output_part_${i + 1}.json`);
    fs.writeFileSync(fileName, JSON.stringify(chunkData, null, 2));

    console.log(`✅ Created file: ${fileName} (${chunk.length} companies)`);
  }

  console.log("🎉 All JSON files have been successfully created!");
})();

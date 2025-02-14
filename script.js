const fs = require("fs");
const path = require("path");
const pdf = require("pdf-parse");
const ExcelJS = require("exceljs");

const pdfFolderPath = "./pdf"; // Folder containing PDFs
const outputExcelPath = "./output.xlsx"; // Output Excel file

if (!fs.existsSync(pdfFolderPath)) {
    console.error(`Error: PDF folder does not exist at ${pdfFolderPath}`);
} else {
    console.log(`PDF folder found at: ${path.resolve(pdfFolderPath)}`);
}


// Function to extract specific values from text
const extractData = (text) => {
    const extracted = {};
    console.log("entered into the extractdata function")

    // Example regex patterns (modify as needed)
    const rollNumberMatch = text.match(/Roll\s*Number\s*(\d+)/);
    const nameMatch = text.match(/Name\s*([A-Z\s]+)/);
    const pwbdMatch = text.match(/Whether\s+PwBD\?\s*(Yes|No)/);
    const pwbdTypeMatch = text.match(/Type of PwBD\s*:\s*([^\n]+)/);
    const heightMatch = text.match(/Height\s*\(in\s*Cms\.\)\s*(\d+)/);
    const Chest_Not_Expanded_Match = text.match(/Chest\s+Measurement\s+in\s+cms\.\s+\(Not\s+Expanded\)\s*(\d+)?(?!\.\d)/);
    const Chest_Expanded_Match = text.match(/Chest\s+Measurement\s+in\s+cms\.\s+\(Expanded\)\s*(\d+)?(?!\.\d)/);
    const weightMatch = text.match(/Weight\s+\(in\s+Kgs\)\s*([\d.]+)/);
    const pstStatusMatch = text.match(/Status\s+of\s+Physical\s+Standard\/\s*Measurement\s+Test\s*(\w+)/);
    const petFinalStatusMatch = text.match(/PST\/\s*PET\s+Final\s+Status\s*(\w+)/);
    const finalRemarksMatch = text.match(/Final\s+Remarks\s*(.+)/);

    extracted.rollNumber = rollNumberMatch ? rollNumberMatch[1] : "Not Found rollnum";
    extracted.name = nameMatch ? nameMatch[1] : "Not Found";
    extracted.pwbd = pwbdMatch ? pwbdMatch[1] : "Not Found";
    extracted.pwbdType = pwbdTypeMatch ? pwbdTypeMatch[1].trim() : "Not Found";
    extracted.height = heightMatch ? heightMatch[1] : "Not Found";
    extracted.Chest_Not_Expanded = Chest_Not_Expanded_Match[1] ? Chest_Not_Expanded_Match[1] : "Not Found";
    extracted. Chest_Expanded = Chest_Expanded_Match[1] ? Chest_Expanded_Match[1] : "Not Found";
    extracted.weight = weightMatch ? weightMatch[1] : "Not Found";
    extracted.pstStatus = pstStatusMatch ? pstStatusMatch[1] : "Not Found";
    extracted.petFinalStatus = petFinalStatusMatch ? petFinalStatusMatch[1] : "Not Found";
    extracted.finalRemarks = finalRemarksMatch ? finalRemarksMatch[1].trim() : "Not Found"
    
    return extracted;
};

console.log("randow check")
const processPDFs = async () => {
    console.log("entered into the process PDFs function")
    const files = fs.readdirSync(pdfFolderPath).filter(file => file.endsWith(".pdf"));
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Extracted Data");

    // Add headers
    sheet.columns = [
        { header: "Roll Number", key: "RollNumber" },
        { header: "Name", key: "Name" },
        { header: "Whether PwBD", key: "WhetherPwBD" },
        { header: "Type of PwBD", key: "TypeofPwBD" },
        { header: "Height", key: "Height" },
        { header: "Chest-Not Expanded", key: "ChestNotExpanded" },
        { header: "Chest-Expanded", key: "ChestExpand" },
        { header: "Weight", key: "Weight" },
        { header: "Status of Physical Standard Test/Measurement Test", key: "StatusofPhysicalStandardTest_MeasurementTest" },
        { header: "PST/PET Final Status", key: "PST_PETFinalStatus" },
        { header: "Final Remarks", key: "FinalRemarks" },
        
    ];

    for (const file of files) {
        const filePath = path.join(pdfFolderPath, file);
        const dataBuffer = fs.readFileSync(filePath);
        const data = await pdf(dataBuffer);
        // console.log(data)
        const extracted = extractData(data.text);

        // Add row to Excel
        sheet.addRow({
            file,
            RollNumber: extracted.rollNumber,
            Name: extracted.name,
            WhetherPwBD: extracted.pwbd,
            TypeofPwBD: extracted.pwbdType,
            Height: extracted.height,
            ChestNotExpanded: extracted.Chest_Not_Expanded,
            ChestExpand:extracted.Chest_Expanded,
            Weight: extracted.weight,
            StatusofPhysicalStandardTest_MeasurementTest: extracted.pstStatus,
            PST_PETFinalStatus: extracted.petFinalStatus,
            FinalRemarks: extracted.finalRemarks,
            
        });

        console.log(`Processed: ${file}`);
    }
    // Save Excel file
    await workbook.xlsx.writeFile(outputExcelPath);
    console.log(`Data saved to ${outputExcelPath}`);
}
// Run the script
processPDFs().catch(console.error);
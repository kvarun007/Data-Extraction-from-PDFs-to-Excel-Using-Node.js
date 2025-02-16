📄 PDF Data Extraction to Excel using Node.js
This Node.js script extracts structured data from PDF files and saves it into an Excel spreadsheet. It automates the process of retrieving key details like Roll Number, Name, PwBD status, Height, Weight, and more.

🚀 Features
✅ Reads multiple PDFs from a folder
✅ Extracts relevant information using regex
✅ Saves extracted data into an Excel sheet
✅ Automates manual data entry, saving time and effort

📌 Prerequisites
Make sure you have Node.js installed on your system. You can download it from: Node.js Official Site

🛠 Installation
1️⃣ Clone this repository:

sh
Copy
Edit
git clone https://github.com/yourusername/pdf-to-excel-extractor.git
cd pdf-to-excel-extractor
2️⃣ Install required dependencies:

sh
Copy
Edit
npm install
📂 Folder Structure
graphql
Copy
Edit
pdf-to-excel-extractor/
│── pdf/ # Place all PDF files here  
│── output.xlsx # Output Excel file (generated after running the script)  
│── script.js # Main script  
│── package.json # Dependencies  
│── README.md # This file  
🚀 Usage
1️⃣ Place your PDFs in the pdf/ folder.
2️⃣ Run the script:

sh
Copy
Edit
node script.js
3️⃣ The extracted data will be saved in output.xlsx.

🛠 Dependencies
pdf-parse – Extracts text from PDFs
ExcelJS – Writes extracted data into Excel
fs & path – Handles file operations
🔍 How It Works
✔️ Reads all PDFs in the pdf/ folder
✔️ Extracts data using regex patterns
✔️ Stores extracted values in an Excel file

🏆 Future Enhancements
✅ Handle more complex PDF formats
✅ Integrate AI-based text extraction for better accuracy
✅ Add CLI options for better usability

🤝 Contributing
Want to improve this project? Feel free to fork and submit a PR! 🚀

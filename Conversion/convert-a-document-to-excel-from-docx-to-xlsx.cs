using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Input DOCX file path
        string inputPath = "Document.docx";

        // Output XLSX file path
        string outputPath = "Document.xlsx";

        // Load the DOCX document
        Document doc = new Document(inputPath);

        // Set up XLSX save options
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions();
        xlsxOptions.SaveFormat = SaveFormat.Xlsx; // Explicitly specify XLSX format
        // Optional: configure how sections are saved
        // xlsxOptions.SectionMode = XlsxSectionMode.SingleWorksheet;

        // Save the document as an XLSX spreadsheet
        doc.Save(outputPath, xlsxOptions);
    }
}

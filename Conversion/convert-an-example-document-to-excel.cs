using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertDocToExcel
{
    static void Main()
    {
        // Path to the source Word document.
        string inputPath = @"C:\Docs\Example.docx";

        // Path where the resulting Excel file will be saved.
        string outputPath = @"C:\Docs\Example.xlsx";

        // Load the Word document from the file system.
        Document doc = new Document(inputPath);

        // Create save options for XLSX format.
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions();

        // Optional: save each section of the Word document to a separate worksheet.
        // Uncomment the line below to change the behavior.
        // xlsxOptions.SectionMode = XlsxSectionMode.MultipleWorksheets;

        // Save the document as an Excel file using the specified options.
        doc.Save(outputPath, xlsxOptions);
    }
}

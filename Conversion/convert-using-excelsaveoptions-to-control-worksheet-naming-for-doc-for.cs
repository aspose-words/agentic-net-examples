using Aspose.Words;
using Aspose.Words.Saving;
using System;

class Program
{
    static void Main()
    {
        // Load the source Word document (DOC/DOCX).
        Document doc = new Document("Input.doc");

        // Create XlsxSaveOptions to influence how the document is saved to Excel.
        XlsxSaveOptions saveOptions = new XlsxSaveOptions();

        // Each section of the Word document will be exported to a separate worksheet.
        // This determines the worksheet creation; the default worksheet names will be used (Sheet1, Sheet2, ...).
        saveOptions.SectionMode = XlsxSectionMode.MultipleWorksheets;

        // Explicitly set the target format to XLSX.
        saveOptions.SaveFormat = SaveFormat.Xlsx;

        // Save the document as an Excel file using the configured options.
        doc.Save("Output.xlsx", saveOptions);
    }
}

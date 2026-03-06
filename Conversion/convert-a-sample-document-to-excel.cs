using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source Word document.
        Document doc = new Document("Sample.docx");

        // Configure XLSX save options.
        XlsxSaveOptions saveOptions = new XlsxSaveOptions();
        saveOptions.SectionMode = XlsxSectionMode.MultipleWorksheets;
        saveOptions.SaveFormat = SaveFormat.Xlsx; // Ensure the format is XLSX.

        // Save the document as an Excel file.
        doc.Save("Sample.xlsx", saveOptions);
    }
}

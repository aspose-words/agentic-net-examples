using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source Word document from a file.
        Document doc = new Document("input.docx");

        // Create XlsxSaveOptions to control how the document is saved as Excel.
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions();
        // Example: save each section of the Word document to a separate worksheet.
        xlsxOptions.SectionMode = XlsxSectionMode.MultipleWorksheets;

        // Save the document in Excel (XLSX) format using the specified options.
        doc.Save("output.xlsx", xlsxOptions);
    }
}

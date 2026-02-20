using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOC/DOCX document.
        Document doc = new Document("InputDocument.docx");

        // Create XlsxSaveOptions to control how sections are exported to worksheets.
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions();

        // Export each section of the Word document to a separate worksheet.
        // This is the only built‑in way to influence worksheet creation via the API.
        xlsxOptions.SectionMode = XlsxSectionMode.MultipleWorksheets;

        // NOTE:
        // Aspose.Words does not provide a direct property to set custom worksheet names
        // during the save operation. To assign specific names you would need to post‑process
        // the generated XLSX file (e.g., using Aspose.Cells) after this conversion.

        // Save the document as an XLSX file using the configured options.
        doc.Save("ConvertedDocument.xlsx", xlsxOptions);
    }
}

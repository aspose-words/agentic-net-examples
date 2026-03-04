using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOCX document.
        Document doc = new Document("input.docx");

        // Create XlsxSaveOptions to control the XLSX conversion.
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions();

        // Save the whole document to a single worksheet.
        xlsxOptions.SectionMode = XlsxSectionMode.SingleWorksheet;

        // Enable pretty formatting of the generated XLSX (makes the XML more readable).
        xlsxOptions.PrettyFormat = true;

        // Explicitly set the save format to Xlsx (required for XlsxSaveOptions).
        xlsxOptions.SaveFormat = SaveFormat.Xlsx;

        // Save the document as an XLSX file using the configured options.
        doc.Save("output.xlsx", xlsxOptions);
    }
}

using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertToExcel
{
    static void Main()
    {
        // Load an existing Word document from disk.
        // Replace "input.docx" with the path to your source document.
        Document doc = new Document("input.docx");

        // Create save options for the XLSX format.
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions();

        // Explicitly set the save format to XLSX (optional, as XlsxSaveOptions defaults to this format).
        xlsxOptions.SaveFormat = SaveFormat.Xlsx;

        // Optionally, choose how sections are handled.
        // xlsxOptions.SectionMode = XlsxSectionMode.SingleWorksheet; // or MultipleWorksheets

        // Save the document as an Excel workbook.
        // Replace "output.xlsx" with the desired output file path.
        doc.Save("output.xlsx", xlsxOptions);
    }
}

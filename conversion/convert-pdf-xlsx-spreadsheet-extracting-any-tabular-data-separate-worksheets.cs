using System;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfToXlsxConverter
{
    static void Main()
    {
        // Path to the source PDF file.
        string pdfPath = "input.pdf";

        // Path where the resulting XLSX file will be saved.
        string xlsxPath = "output.xlsx";

        // Load the PDF document. The Document constructor automatically detects the format.
        Document pdfDocument = new Document(pdfPath);

        // Configure XLSX save options.
        // SectionMode.MultipleWorksheets creates a separate worksheet for each document section.
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions();
        xlsxOptions.SectionMode = XlsxSectionMode.MultipleWorksheets;

        // Save the document as an XLSX spreadsheet.
        pdfDocument.Save(xlsxPath, xlsxOptions);
    }
}

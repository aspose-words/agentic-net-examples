using System;
using Aspose.Words;
using Aspose.Words.Saving;

class DocumentToExcelConverter
{
    static void Main()
    {
        // Path to the source document (any format supported by Aspose.Words)
        string inputFile = "input.docx";

        // Path where the Excel file will be saved
        string outputFile = "output.xlsx";

        // Load the source document
        Document doc = new Document(inputFile);

        // Configure Excel save options
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions
        {
            // Example: save all sections on a single worksheet.
            // Change to XlsxSectionMode.MultipleWorksheets to create a worksheet per section.
            SectionMode = XlsxSectionMode.SingleWorksheet
        };

        // Save the document as an Excel file
        doc.Save(outputFile, xlsxOptions);
    }
}

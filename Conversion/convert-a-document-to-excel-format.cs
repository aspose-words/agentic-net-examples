using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertToExcel
{
    static void Main()
    {
        // Input document (any format supported by Aspose.Words)
        string inputPath = "input.docx";

        // Desired output Excel file
        string outputPath = "output.xlsx";

        // Load the source document
        Document doc = new Document(inputPath);

        // Configure save options for XLSX format
        XlsxSaveOptions options = new XlsxSaveOptions();
        options.SectionMode = XlsxSectionMode.MultipleWorksheets; // each section -> separate worksheet
        options.SaveFormat = SaveFormat.Xlsx; // must be Xlsx when using XlsxSaveOptions

        // Save the document as an Excel workbook
        doc.Save(outputPath, options);
    }
}

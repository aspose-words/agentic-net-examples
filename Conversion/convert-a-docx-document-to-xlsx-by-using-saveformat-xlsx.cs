using System;
using Aspose.Words;
using Aspose.Words.Saving;

class DocxToXlsxConverter
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = @"C:\Docs\SourceDocument.docx";

        // Path where the resulting XLSX file will be saved.
        string outputPath = @"C:\Docs\ConvertedDocument.xlsx";

        // Load the DOCX document from the file system.
        Document doc = new Document(inputPath);

        // Save the document in XLSX format.
        // The Save method overload (string, SaveFormat) is used as defined in the Aspose.Words API.
        doc.Save(outputPath, SaveFormat.Xlsx);
    }
}

using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsToExcel
{
    class Program
    {
        static void Main()
        {
            // Path to the source Word document.
            string inputPath = "MyDir/ExampleDocument.docx";

            // Path where the resulting Excel file will be saved.
            string outputPath = "ArtifactsDir/ConvertedDocument.xlsx";

            // Load the Word document from the file system.
            Document doc = new Document(inputPath);

            // Create save options for XLSX format.
            XlsxSaveOptions xlsxOptions = new XlsxSaveOptions();

            // Optional: save each section of the Word document to a separate worksheet.
            xlsxOptions.SectionMode = XlsxSectionMode.MultipleWorksheets;

            // Save the document as an Excel file using the specified options.
            doc.Save(outputPath, xlsxOptions);
        }
    }
}

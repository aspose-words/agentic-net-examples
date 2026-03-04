using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace DocumentToExcelConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            // Example usage:
            // args[0] - path to the source document (any format supported by Aspose.Words)
            // args[1] - desired output Excel file path (should have .xlsx extension)
            if (args.Length != 2)
            {
                Console.WriteLine("Usage: DocumentToExcelConverter <inputFile> <outputXlsxFile>");
                return;
            }

            string inputPath = args[0];
            string outputPath = args[1];

            try
            {
                ConvertToExcel(inputPath, outputPath);
                Console.WriteLine($"Document converted successfully: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Conversion failed: {ex.Message}");
            }
        }

        /// <summary>
        /// Loads a document from any supported format and saves it as an Excel workbook (XLSX).
        /// </summary>
        /// <param name="inputFile">Path to the source document.</param>
        /// <param name="outputFile">Path where the XLSX file will be saved.</param>
        static void ConvertToExcel(string inputFile, string outputFile)
        {
            // Load the source document. Aspose.Words automatically detects the format.
            Document doc = new Document(inputFile);

            // Optionally, configure XlsxSaveOptions (e.g., single worksheet mode).
            XlsxSaveOptions saveOptions = new XlsxSaveOptions
            {
                // SaveFormat is implicitly Xlsx, but set explicitly for clarity.
                SaveFormat = SaveFormat.Xlsx,
                // Save all sections on a single worksheet (change to MultipleWorksheets if needed).
                SectionMode = XlsxSectionMode.SingleWorksheet
            };

            // Save the document as an Excel file using the specified options.
            doc.Save(outputFile, saveOptions);
        }
    }
}

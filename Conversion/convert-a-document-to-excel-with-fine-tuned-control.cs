using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsToExcelDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source Word document.
            string sourcePath = Path.Combine("Input", "SampleDocument.docx");

            // Load the Word document using the provided constructor.
            Document doc = new Document(sourcePath);

            // Create XlsxSaveOptions to control the Excel conversion.
            XlsxSaveOptions xlsxOptions = new XlsxSaveOptions();

            // Fine‑tuned control: each Word section becomes a separate worksheet.
            xlsxOptions.SectionMode = XlsxSectionMode.MultipleWorksheets;

            // Optional: maximize compression for the resulting XLSX file.
            xlsxOptions.CompressionLevel = CompressionLevel.Maximum;

            // Optional: produce a human‑readable XML structure.
            xlsxOptions.PrettyFormat = true;

            // Path for the output Excel file.
            string outputPath = Path.Combine("Output", "ConvertedDocument.xlsx");

            // Ensure the output directory exists.
            Directory.CreateDirectory(Path.GetDirectoryName(outputPath));

            // Save the document as XLSX using the configured options.
            doc.Save(outputPath, xlsxOptions);

            Console.WriteLine($"Document successfully converted to Excel: {outputPath}");
        }
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

namespace PdfToDocxBatch
{
    public class Converter
    {
        /// <summary>
        /// Converts each PDF file in <paramref name="pdfFiles"/> to DOCX and writes the result to <paramref name="outputFolder"/>.
        /// </summary>
        /// <param name="pdfFiles">Full paths of the PDF files to convert.</param>
        /// <param name="outputFolder">Folder where the DOCX files will be saved.</param>
        public void ConvertAll(IEnumerable<string> pdfFiles, string outputFolder)
        {
            // Ensure the output directory exists.
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            foreach (string pdfPath in pdfFiles)
            {
                if (!File.Exists(pdfPath))
                {
                    Console.WriteLine($"Skipping missing file: {pdfPath}");
                    continue;
                }

                // Load the PDF document.
                Document pdfDoc = new Document(pdfPath);

                // Build the output DOCX file name (same base name, .docx extension).
                string docxFileName = Path.GetFileNameWithoutExtension(pdfPath) + ".docx";
                string docxPath = Path.Combine(outputFolder, docxFileName);

                // Save the document as DOCX.
                pdfDoc.Save(docxPath);
                Console.WriteLine($"Converted '{pdfPath}' to '{docxPath}'.");
            }
        }

        // Example usage.
        public static void Main()
        {
            // Create temporary input and output folders.
            string tempRoot = Path.Combine(Path.GetTempPath(), "PdfToDocxDemo");
            string inputFolder = Path.Combine(tempRoot, "Input");
            string outputFolder = Path.Combine(tempRoot, "Output");

            Directory.CreateDirectory(inputFolder);
            Directory.CreateDirectory(outputFolder);

            // Generate sample PDF files.
            var pdfFiles = new List<string>();
            for (int i = 1; i <= 3; i++)
            {
                string pdfPath = Path.Combine(inputFolder, $"file{i}.pdf");
                // Create a simple Word document and save it as PDF.
                Document doc = new Document();
                DocumentBuilder builder = new DocumentBuilder(doc);
                builder.Writeln($"This is sample PDF number {i}.");
                doc.Save(pdfPath, SaveFormat.Pdf);
                pdfFiles.Add(pdfPath);
            }

            // Perform conversion.
            Converter converter = new Converter();
            converter.ConvertAll(pdfFiles, outputFolder);

            Console.WriteLine("Conversion completed.");
        }
    }
}

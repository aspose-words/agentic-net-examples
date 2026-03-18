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
                // Load the PDF document using the Document(string) constructor.
                Document pdfDoc = new Document(pdfPath);

                // Build the output DOCX file name (same base name, .docx extension).
                string docxFileName = Path.GetFileNameWithoutExtension(pdfPath) + ".docx";
                string docxPath = Path.Combine(outputFolder, docxFileName);

                // Save the document as DOCX using the Save(string) overload.
                pdfDoc.Save(docxPath);
            }
        }

        // Example usage.
        public static void Main()
        {
            // Example list of PDF files – replace with your actual file paths.
            List<string> pdfFiles = new List<string>
            {
                @"C:\Input\file1.pdf",
                @"C:\Input\file2.pdf",
                @"C:\Input\file3.pdf"
            };

            string outputFolder = @"C:\Output";

            Converter converter = new Converter();
            converter.ConvertAll(pdfFiles, outputFolder);

            Console.WriteLine("Conversion completed.");
        }
    }
}

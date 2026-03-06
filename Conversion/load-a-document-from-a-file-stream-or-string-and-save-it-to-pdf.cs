using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving; // Added for PdfSaveOptions

namespace AsposeWordsPdfConversion
{
    class Program
    {
        static void Main(string[] args)
        {
            // -----------------------------------------------------------------
            // 1. Load a document from a file path and save it as PDF.
            // -----------------------------------------------------------------
            // The Document constructor that takes a file name automatically detects the format.
            Document docFromFile = new Document("input.docx");
            // Saving to a file with a .pdf extension automatically uses the PDF format.
            docFromFile.Save("output_from_file.pdf");

            // -----------------------------------------------------------------
            // 2. Load a document from a FileStream and save it as PDF.
            // -----------------------------------------------------------------
            using (FileStream fileStream = File.OpenRead("input.docx"))
            {
                // The Document constructor that takes a Stream loads the document.
                Document docFromStream = new Document(fileStream);
                docFromStream.Save("output_from_stream.pdf");
            }

            // -----------------------------------------------------------------
            // 3. Load a document from a string (in‑memory text) and save it as PDF.
            // -----------------------------------------------------------------
            // For demonstration, read the source file into a string.
            string documentText = File.ReadAllText("input.txt", Encoding.UTF8);
            // Convert the string to a byte array and wrap it in a MemoryStream.
            byte[] textBytes = Encoding.UTF8.GetBytes(documentText);
            using (MemoryStream memoryStream = new MemoryStream(textBytes))
            {
                // Load the document from the memory stream.
                Document docFromString = new Document(memoryStream);
                docFromString.Save("output_from_string.pdf");
            }

            // -----------------------------------------------------------------
            // 4. Load a document from a stream and explicitly specify PDF save options.
            // -----------------------------------------------------------------
            using (FileStream fileStream = File.OpenRead("input.docx"))
            {
                Document doc = new Document(fileStream);
                // Create a PdfSaveOptions instance.
                PdfSaveOptions pdfOptions = new PdfSaveOptions();
                // Example: enable high‑quality rendering.
                pdfOptions.UseHighQualityRendering = true;
                // Save using the options object.
                doc.Save("output_with_options.pdf", pdfOptions);
            }

            Console.WriteLine("Documents have been converted to PDF successfully.");
        }
    }
}

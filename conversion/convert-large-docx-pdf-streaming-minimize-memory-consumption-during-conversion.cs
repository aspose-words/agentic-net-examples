using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Paths to the source DOCX and the destination PDF.
        const string inputFile = @"C:\Docs\LargeDocument.docx";
        const string outputFile = @"C:\Docs\LargeDocument.pdf";

        // Open the source document as a read‑only stream to avoid loading the whole file into memory.
        using (FileStream inputStream = File.OpenRead(inputFile))
        {
            // Load the document from the stream.
            Document doc = new Document(inputStream);

            // Create PDF save options and enable memory optimization to reduce RAM usage.
            SaveOptions pdfOptions = SaveOptions.CreateSaveOptions(SaveFormat.Pdf);
            pdfOptions.MemoryOptimization = true;

            // Save the document to a write‑only stream, which writes the PDF incrementally.
            using (FileStream outputStream = File.Create(outputFile))
            {
                doc.Save(outputStream, pdfOptions);
            }
        }
    }
}

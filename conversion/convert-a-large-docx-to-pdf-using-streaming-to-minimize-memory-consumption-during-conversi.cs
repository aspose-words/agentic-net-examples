using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a large DOCX document (simulating a large file).
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);

        // Add many pages to increase size.
        for (int i = 0; i < 1000; i++)
        {
            builder.Writeln($"This is page {i + 1} of a large document.");
            builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the document locally as DOCX (input file).
        const string inputPath = "large_input.docx";
        source.Save(inputPath, SaveFormat.Docx);

        // Load the DOCX document from the file system.
        Document doc = new Document(inputPath);

        // Create PDF save options with memory optimization enabled.
        SaveOptions pdfOptions = SaveOptions.CreateSaveOptions(SaveFormat.Pdf);
        pdfOptions.MemoryOptimization = true;

        // Convert the document to PDF using a memory stream to minimize memory usage.
        using (MemoryStream pdfStream = new MemoryStream())
        {
            doc.Save(pdfStream, pdfOptions);

            // Verify that data was written to the stream.
            if (pdfStream.Length == 0)
                throw new InvalidOperationException("PDF conversion produced an empty stream.");

            // Optionally write the PDF to a file for verification.
            const string outputPath = "output.pdf";
            pdfStream.Position = 0;
            using (FileStream file = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
            {
                pdfStream.CopyTo(file);
            }

            // Ensure the output file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("Expected output PDF file was not created.");
        }
    }
}

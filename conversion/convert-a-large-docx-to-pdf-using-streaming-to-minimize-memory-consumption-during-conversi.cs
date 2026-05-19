using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a large sample DOCX document.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        for (int i = 0; i < 1000; i++)
        {
            builder.Writeln($"Paragraph {i + 1}");
        }

        // Save the sample document to a local DOCX file (input bootstrap rule).
        const string inputPath = "input.docx";
        source.Save(inputPath, SaveFormat.Docx);

        // Step 2: Load the DOCX document from the file system.
        Document doc = new Document(inputPath);

        // Step 3: Prepare PDF save options with memory optimization enabled.
        SaveOptions pdfSaveOptions = SaveOptions.CreateSaveOptions(SaveFormat.Pdf);
        pdfSaveOptions.MemoryOptimization = true;

        // Step 4: Convert the document to PDF using a MemoryStream to minimize memory usage.
        using (MemoryStream pdfStream = new MemoryStream())
        {
            doc.Save(pdfStream, pdfSaveOptions);

            // Validate that data was written to the stream.
            if (pdfStream.Length == 0)
                throw new InvalidOperationException("No PDF data was written to the stream.");

            // Optional: write the stream to a physical PDF file for verification.
            const string outputPath = "output.pdf";
            pdfStream.Position = 0;
            using (FileStream file = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
            {
                pdfStream.CopyTo(file);
            }

            if (!File.Exists(outputPath))
                throw new InvalidOperationException("Expected output PDF file was not created.");
        }
    }
}

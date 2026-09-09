using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample document generated for streaming conversion.");
        // Add enough content to simulate a large document.
        for (int i = 0; i < 5000; i++)
        {
            builder.Writeln($"Paragraph {i + 1}: Lorem ipsum dolor sit amet, consectetur adipiscing elit.");
        }
        const string inputPath = "input.docx";
        sourceDoc.Save(inputPath, SaveFormat.Docx);

        // Load the DOCX document.
        Document doc = new Document(inputPath);

        // Prepare PDF save options with memory optimization enabled.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            MemoryOptimization = true
        };

        // Convert to PDF using a memory stream to keep memory usage low.
        using (MemoryStream pdfStream = new MemoryStream())
        {
            doc.Save(pdfStream, pdfOptions);

            // Verify that data was written to the stream.
            if (pdfStream.Length == 0)
                throw new InvalidOperationException("No PDF data was written to the stream.");

            // Write the stream to a file.
            const string outputPath = "output.pdf";
            pdfStream.Position = 0;
            using (FileStream fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
            {
                pdfStream.CopyTo(fileStream);
            }

            // Validate that the output file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The PDF file was not created.");
        }
    }
}

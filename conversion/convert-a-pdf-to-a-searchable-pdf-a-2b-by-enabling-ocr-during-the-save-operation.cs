using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple PDF document from a blank Word document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample PDF document.");
        string inputPath = "input.pdf";
        sourceDoc.Save(inputPath, SaveFormat.Pdf);

        // Load the PDF that was just created.
        Document pdfDoc = new Document(inputPath);

        // Configure save options for PDF/A‑2u compliance.
        // Aspose.Words does not expose a PdfA2b enum value; PdfA2u provides PDF/A‑2b compatible output.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA2u
        };

        // Save the document as a searchable PDF/A‑2u file.
        string outputPath = "output.pdf";
        pdfDoc.Save(outputPath, saveOptions);

        // Verify that the output file was created and is not empty.
        if (!File.Exists(outputPath) || new FileInfo(outputPath).Length == 0)
            throw new InvalidOperationException("The searchable PDF/A‑2u file was not created.");
    }
}

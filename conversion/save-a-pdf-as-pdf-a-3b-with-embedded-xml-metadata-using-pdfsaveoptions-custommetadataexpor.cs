using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample document for PDF/A‑3b conversion with embedded XML metadata.");

        // Add custom document properties – these will be exported as XMP metadata.
        doc.CustomDocumentProperties.Add("Company", "Acme Corp");
        doc.CustomDocumentProperties.Add("Project", "PDF/A‑3b Demo");

        // Configure PDF save options for PDF/A‑3b compliance and XMP metadata export.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            // PDF/A‑3b compliance – use the unrestricted variant.
            Compliance = PdfCompliance.PdfA3u,

            // Export custom properties as XMP metadata.
            CustomPropertiesExport = PdfCustomPropertiesExport.Metadata
        };

        // Define the output file path.
        string outputPath = "output_pdfa3b.pdf";

        // Save the document as PDF/A‑3b with embedded XML metadata.
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(outputPath) || new FileInfo(outputPath).Length == 0)
        {
            throw new InvalidOperationException("The PDF/A‑3b file was not created successfully.");
        }

        // Inform that the process completed.
        Console.WriteLine($"PDF/A‑3b file saved to: {Path.GetFullPath(outputPath)}");
    }
}

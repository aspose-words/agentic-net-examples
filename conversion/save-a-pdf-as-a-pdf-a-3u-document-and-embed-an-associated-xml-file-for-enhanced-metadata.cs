using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(
            "This is a sample document that will be saved as PDF/A‑3u with embedded XML metadata.");

        // Prepare XML metadata to embed.
        string xmlContent = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<metadata>
    <title>Sample Document</title>
    <author>Aspose Example</author>
    <created>" + DateTime.UtcNow.ToString("o") + @"</created>
</metadata>";

        // Store the XML as a custom document property (as a simple way to embed metadata).
        doc.CustomDocumentProperties.Add("MetadataXml", xmlContent);

        // Configure PDF/A‑3u save options.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA3u
        };

        // Save the document as PDF/A‑3u.
        string outputPath = "output.pdf";
        doc.Save(outputPath, saveOptions);

        // Validate that the PDF was created.
        if (!File.Exists(outputPath) || new FileInfo(outputPath).Length == 0)
        {
            throw new InvalidOperationException("The PDF/A‑3u file was not created successfully.");
        }

        // Optionally, clean up the generated file (comment out if you want to keep it).
        // File.Delete(outputPath);
    }
}

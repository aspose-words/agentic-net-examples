using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document and add some text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample content for PDF/A-3b conversion.");

        // Add a custom document property that contains XML metadata.
        string xmlMetadata = "<metadata><author>John Doe</author><description>Sample PDF/A-3b document</description></metadata>";
        doc.CustomDocumentProperties.Add("XmlMetadata", xmlMetadata);

        // Configure PDF save options:
        // - Set compliance to PDF/A-3b (represented by PdfA3u in Aspose.Words).
        // - Export custom properties as XMP metadata.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA3u,
            CustomPropertiesExport = PdfCustomPropertiesExport.Metadata
        };

        // Save the document as PDF/A-3b.
        string outputPath = "output.pdf";
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("PDF/A-3b file was not created.");
    }
}

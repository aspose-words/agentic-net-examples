using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple Word document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, PDF/A‑3u with attached XML metadata.");

        // Prepare XML metadata content (not embedded because the current Aspose.Words version
        // does not expose the FileSpecification/EmbeddedFiles API). The XML can still be saved
        // to a separate file for reference.
        string xmlContent = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><metadata><author>John Doe</author></metadata>";
        string xmlPath = "metadata.xml";
        File.WriteAllText(xmlPath, xmlContent);

        // Configure PDF/A‑3u save options.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA3u
        };

        // Save the document as PDF/A‑3u.
        string outputPath = "output_pdfa3u.pdf";
        doc.Save(outputPath, saveOptions);

        // Validate that the PDF was created.
        if (!File.Exists(outputPath) || new FileInfo(outputPath).Length == 0)
        {
            throw new InvalidOperationException("PDF/A‑3u file was not created correctly.");
        }

        Console.WriteLine($"PDF/A‑3u file created successfully at '{Path.GetFullPath(outputPath)}'.");
        Console.WriteLine($"XML metadata file saved at '{Path.GetFullPath(xmlPath)}'.");
    }
}

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX with placeholders.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("Report for {{Name}}");
        builder.Writeln("Date: {{Date}}");

        string inputPath = "input.docx";
        sampleDoc.Save(inputPath, SaveFormat.Docx);

        // Load the DOCX.
        Document doc = new Document(inputPath);

        // Replace placeholders with actual values.
        doc.Range.Replace("{{Name}}", "John Doe", new FindReplaceOptions());
        doc.Range.Replace("{{Date}}", DateTime.Today.ToString("yyyy-MM-dd"), new FindReplaceOptions());

        // Export to PDF.
        string outputPath = "output.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("Expected output PDF was not created.");
        }

        // Optional cleanup (commented out to keep files for inspection).
        // File.Delete(inputPath);
    }
}

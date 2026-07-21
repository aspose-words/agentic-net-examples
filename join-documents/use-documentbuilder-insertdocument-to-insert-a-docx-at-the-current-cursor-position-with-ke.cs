using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample source DOCX file.
        string sourcePath = Path.Combine(outputDir, "Source.docx");
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("This is the source document.");
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // Create the destination document.
        Document destDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(destDoc);
        builder.Writeln("This is the destination document before insertion.");
        builder.InsertBreak(BreakType.PageBreak);

        // Load the source document and insert it at the current cursor position,
        // preserving its original formatting.
        Document srcToInsert = new Document(sourcePath);
        builder.InsertDocument(srcToInsert, ImportFormatMode.KeepSourceFormatting);

        // Save the merged result as HTML.
        string htmlPath = Path.Combine(outputDir, "Merged.html");
        destDoc.Save(htmlPath, SaveFormat.Html);

        // Validate that the HTML file was created and contains content from both documents.
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("Merged HTML file was not created.");

        string htmlContent = File.ReadAllText(htmlPath);
        if (!htmlContent.Contains("destination document", StringComparison.OrdinalIgnoreCase) ||
            !htmlContent.Contains("source document", StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("Merged HTML does not contain expected content.");
    }
}

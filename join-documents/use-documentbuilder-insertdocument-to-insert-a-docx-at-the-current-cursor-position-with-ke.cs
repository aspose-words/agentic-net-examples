using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the source DOCX and the final HTML result.
        string sourceDocPath = Path.Combine(outputDir, "Source.docx");
        string resultHtmlPath = Path.Combine(outputDir, "Result.html");

        // ---------- Create a source DOCX document ----------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("This is the source document.");
        sourceDoc.Save(sourceDocPath, SaveFormat.Docx);

        // ---------- Create a destination document ----------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("This is the destination document before insertion.");

        // Load the source document to be inserted.
        Document docToInsert = new Document(sourceDocPath);

        // Insert the source document at the current cursor position,
        // preserving its original formatting.
        destBuilder.InsertDocument(docToInsert, ImportFormatMode.KeepSourceFormatting);

        // Add additional content after the insertion (optional).
        destBuilder.Writeln("This is the destination document after insertion.");

        // ---------- Save the merged document as HTML ----------
        destDoc.Save(resultHtmlPath, SaveFormat.Html);

        // ---------- Simple validation ----------
        if (!File.Exists(resultHtmlPath))
            throw new InvalidOperationException("The HTML output file was not created.");

        string html = File.ReadAllText(resultHtmlPath);
        if (!html.Contains("This is the source document.") ||
            !html.Contains("This is the destination document before insertion.") ||
            !html.Contains("This is the destination document after insertion."))
        {
            throw new InvalidOperationException("The merged HTML does not contain expected content.");
        }
    }
}

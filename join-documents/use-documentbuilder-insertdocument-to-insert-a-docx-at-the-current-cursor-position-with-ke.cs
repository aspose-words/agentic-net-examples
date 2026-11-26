using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for temporary source document and final HTML output
        string sourceDocPath = Path.Combine(Path.GetTempPath(), "Source.docx");
        string outputHtmlPath = Path.Combine(Path.GetTempPath(), "JoinedOutput.html");

        // Create source DOCX document
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);
        sourceBuilder.Writeln("This is the source document inserted with KeepSourceFormatting.");
        sourceDoc.Save(sourceDocPath, SaveFormat.Docx);

        // Create destination document
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("This is the destination document before insertion.");

        // Load the source document from file
        Document loadedSource = new Document(sourceDocPath);

        // Insert the source document at the current cursor position, preserving its formatting
        destBuilder.InsertDocument(loadedSource, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document as HTML
        destDoc.Save(outputHtmlPath, SaveFormat.Html);

        // Validate that the HTML file was created and contains content from both documents
        if (!File.Exists(outputHtmlPath))
        {
            throw new InvalidOperationException("The output HTML file was not created.");
        }

        string htmlContent = File.ReadAllText(outputHtmlPath);
        if (!htmlContent.Contains("destination document before insertion") ||
            !htmlContent.Contains("source document inserted with KeepSourceFormatting"))
        {
            throw new InvalidOperationException("The merged HTML does not contain expected content from both source documents.");
        }

        // Clean up temporary source document (optional)
        try { File.Delete(sourceDocPath); } catch { /* ignore cleanup errors */ }
    }
}

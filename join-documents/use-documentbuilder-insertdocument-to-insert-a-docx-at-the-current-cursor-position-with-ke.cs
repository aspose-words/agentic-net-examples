using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define paths for temporary files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        string sourceDocPath = Path.Combine(outputDir, "SourceDocument.docx");
        string resultHtmlPath = Path.Combine(outputDir, "MergedDocument.html");

        // -------------------------------------------------
        // Create the source DOCX document that will be inserted.
        // -------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);
        sourceBuilder.Writeln("This is the content of the source DOCX document.");
        sourceBuilder.Writeln("It will be inserted into another document.");
        // Save the source as DOCX as required by the rules.
        sourceDoc.Save(sourceDocPath, SaveFormat.Docx);

        // Load the source document from the file system to simulate an external file.
        Document loadedSourceDoc = new Document(sourceDocPath);

        // -------------------------------------------------
        // Create the destination document.
        // -------------------------------------------------
        Document destinationDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destinationDoc);
        destBuilder.Writeln("Destination document begins here.");
        destBuilder.Writeln("The source document will be inserted below.");

        // Move the cursor to the end of the destination document.
        destBuilder.MoveToDocumentEnd();

        // Insert the source document at the current cursor position,
        // preserving its original formatting.
        destBuilder.InsertDocument(loadedSourceDoc, ImportFormatMode.KeepSourceFormatting);

        // -------------------------------------------------
        // Save the merged document as HTML.
        // -------------------------------------------------
        destinationDoc.Save(resultHtmlPath, SaveFormat.Html);

        // Verify that the output file was created.
        if (!File.Exists(resultHtmlPath))
        {
            throw new InvalidOperationException("The merged HTML file was not created.");
        }

        // Optional: inform that the process completed successfully.
        Console.WriteLine("Document merged and saved as HTML at:");
        Console.WriteLine(resultHtmlPath);
    }
}

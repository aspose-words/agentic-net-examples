using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for the output files.
        const string outputHtmlPath = "MergedDocument.html";

        // ---------- Create the destination document ----------
        Document destination = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destination);
        destBuilder.Writeln("This is the destination document.");

        // Move the cursor to the end of the destination document.
        destBuilder.MoveToDocumentEnd();

        // Insert a page break to separate the documents (optional).
        destBuilder.InsertBreak(BreakType.PageBreak);

        // ---------- Create the source document ----------
        Document source = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(source);
        srcBuilder.Writeln("This is the source document that will be inserted.");

        // ---------- Insert the source document ----------
        // KeepSourceFormatting ensures the source formatting is preserved.
        destBuilder.InsertDocument(source, ImportFormatMode.KeepSourceFormatting);

        // ---------- Save the merged document as HTML ----------
        destination.Save(outputHtmlPath, SaveFormat.Html);

        // ---------- Validation ----------
        if (!File.Exists(outputHtmlPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputHtmlPath}");

        // Verify that both pieces of text are present in the merged document.
        string mergedText = destination.GetText();
        if (!mergedText.Contains("destination document") || !mergedText.Contains("source document"))
            throw new InvalidOperationException("The merged document does not contain expected content.");

        // Indicate successful completion.
        Console.WriteLine("Document merged and saved as HTML successfully.");
    }
}

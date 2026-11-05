using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Paths for the sample files.
        string destPath = "Destination.docx";
        string sourcePath = "Source.docx";
        string mergedPath = "Merged.docx";

        // -------------------------------------------------
        // Create the destination document with a bookmark named "Content".
        // -------------------------------------------------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("Start of Destination Document.");
        destBuilder.StartBookmark("Content");
        destBuilder.Writeln("Bookmark placeholder.");
        destBuilder.EndBookmark("Content");
        destBuilder.Writeln("End of Destination Document.");
        destDoc.Save(destPath, SaveFormat.Docx);

        // -------------------------------------------------
        // Create the source document that will be inserted.
        // -------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);
        sourceBuilder.Writeln("This is content from the source document.");
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // -------------------------------------------------
        // Load the destination document, move to the bookmark, and insert the source document
        // preserving its original formatting.
        // -------------------------------------------------
        Document loadedDest = new Document(destPath);
        DocumentBuilder insertBuilder = new DocumentBuilder(loadedDest);
        insertBuilder.MoveToBookmark("Content");
        insertBuilder.InsertDocument(sourceDoc, ImportFormatMode.KeepSourceFormatting);
        loadedDest.Save(mergedPath, SaveFormat.Docx);

        // -------------------------------------------------
        // Validation: ensure the merged file exists and contains source content.
        // -------------------------------------------------
        if (!File.Exists(mergedPath))
        {
            throw new Exception($"Merged document was not created at '{mergedPath}'.");
        }

        Document mergedDoc = new Document(mergedPath);
        string mergedText = mergedDoc.GetText();

        if (!mergedText.Contains("This is content from the source document."))
        {
            throw new Exception("Merged document does not contain the expected source content.");
        }

        // Optional: clean up sample files (comment out if you want to inspect them).
        // File.Delete(destPath);
        // File.Delete(sourcePath);
        // File.Delete(mergedPath);
    }
}

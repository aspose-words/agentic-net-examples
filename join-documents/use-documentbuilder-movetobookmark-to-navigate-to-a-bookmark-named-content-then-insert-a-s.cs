using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the documents.
        string destinationPath = Path.Combine(outputDir, "Destination.docx");
        string sourcePath = Path.Combine(outputDir, "Source.docx");
        string mergedPath = Path.Combine(outputDir, "Merged.docx");

        // ---------- Create destination document with a bookmark named "Content" ----------
        Document destinationDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destinationDoc);

        destBuilder.Writeln("Header before bookmark.");
        destBuilder.StartBookmark("Content");
        destBuilder.Writeln("Placeholder text that will be replaced.");
        destBuilder.EndBookmark("Content");
        destBuilder.Writeln("Footer after bookmark.");

        // Save the destination document (optional, just for reference).
        destinationDoc.Save(destinationPath);

        // ---------- Create source document (DOCX) that will be inserted ----------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);

        srcBuilder.Writeln("Inserted source document content line 1.");
        srcBuilder.Writeln("Inserted source document content line 2.");

        // Save the source document (optional, just for reference).
        sourceDoc.Save(sourcePath);

        // ---------- Insert the source document at the bookmark, preserving its formatting ----------
        // Move the builder cursor to the bookmark named "Content".
        bool bookmarkFound = destBuilder.MoveToBookmark("Content");
        if (!bookmarkFound)
        {
            throw new InvalidOperationException("Bookmark 'Content' was not found in the destination document.");
        }

        // Insert the source document inline, keeping its original formatting.
        destBuilder.InsertDocument(sourceDoc, ImportFormatMode.KeepSourceFormatting);

        // ---------- Save the merged result ----------
        destinationDoc.Save(mergedPath);

        // ---------- Validation ----------
        if (!File.Exists(mergedPath))
        {
            throw new FileNotFoundException("The merged document was not created.", mergedPath);
        }

        // Verify that the merged document contains content from the source document.
        string mergedText = destinationDoc.GetText();
        if (!mergedText.Contains("Inserted source document content line 1.") ||
            !mergedText.Contains("Inserted source document content line 2."))
        {
            throw new InvalidOperationException("The merged document does not contain the expected source content.");
        }

        // Indicate successful completion (no console output required by the task).
    }
}

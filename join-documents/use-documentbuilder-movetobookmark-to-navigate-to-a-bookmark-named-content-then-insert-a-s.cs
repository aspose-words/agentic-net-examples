using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define a folder for all temporary and output files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // Create a source DOCX document that will be inserted later.
        // -----------------------------------------------------------------
        string sourcePath = Path.Combine(outputDir, "Source.docx");
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);
        sourceBuilder.Writeln("This is the source document content.");
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Create the destination document with a bookmark named "Content".
        // -----------------------------------------------------------------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("Destination document start.");
        destBuilder.StartBookmark("Content");
        destBuilder.Writeln("Placeholder text inside bookmark.");
        destBuilder.EndBookmark("Content");
        destBuilder.Writeln("Destination document end.");

        // Load the source document from the file system.
        Document srcToInsert = new Document(sourcePath);

        // Move the builder cursor to the bookmark and insert the source document,
        // preserving its original formatting.
        bool moved = destBuilder.MoveToBookmark("Content");
        if (!moved)
        {
            throw new InvalidOperationException("Bookmark 'Content' was not found in the destination document.");
        }

        destBuilder.InsertDocument(srcToInsert, ImportFormatMode.KeepSourceFormatting);

        // -----------------------------------------------------------------
        // Save the merged document.
        // -----------------------------------------------------------------
        string mergedPath = Path.Combine(outputDir, "Merged.docx");
        destDoc.Save(mergedPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Simple validation to ensure the merge succeeded.
        // -----------------------------------------------------------------
        if (!File.Exists(mergedPath))
        {
            throw new FileNotFoundException("The merged document was not created.", mergedPath);
        }

        // Verify that the merged document contains text from both source and destination.
        string mergedText = destDoc.GetText();
        if (!mergedText.Contains("This is the source document content.") ||
            !mergedText.Contains("Destination document start.") ||
            !mergedText.Contains("Destination document end."))
        {
            throw new InvalidOperationException("Merged document does not contain expected content.");
        }

        // The program finishes without requiring any user interaction.
    }
}

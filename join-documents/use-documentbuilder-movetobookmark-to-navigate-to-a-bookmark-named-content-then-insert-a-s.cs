using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folder for generated files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // Create the source DOCX that will be inserted.
        // -----------------------------------------------------------------
        string sourcePath = Path.Combine(outputDir, "Source.docx");
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("This is the content of the source document.");
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Create the destination document with a bookmark named "Content".
        // -----------------------------------------------------------------
        string destPath = Path.Combine(outputDir, "Merged.docx");
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);

        destBuilder.Writeln("=== Destination Document Start ===");
        destBuilder.StartBookmark("Content");
        destBuilder.Writeln("[Placeholder for inserted content]");
        destBuilder.EndBookmark("Content");
        destBuilder.Writeln("=== Destination Document End ===");

        // -----------------------------------------------------------------
        // Move to the bookmark and insert the source document,
        // preserving its original formatting.
        // -----------------------------------------------------------------
        bool moved = destBuilder.MoveToBookmark("Content");
        if (!moved)
            throw new InvalidOperationException("Bookmark 'Content' was not found in the destination document.");

        // Insert the source document at the bookmark location.
        destBuilder.InsertDocument(sourceDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the merged result.
        destDoc.Save(destPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Validation: ensure the merged file exists and contains source text.
        // -----------------------------------------------------------------
        if (!File.Exists(destPath))
            throw new FileNotFoundException("The merged document was not created.", destPath);

        // Load the merged document to verify content.
        Document verificationDoc = new Document(destPath);
        string mergedText = verificationDoc.GetText();

        if (!mergedText.Contains("This is the content of the source document."))
            throw new Exception("The source content was not found in the merged document.");

        // Indicate successful completion (no console output required).
    }
}

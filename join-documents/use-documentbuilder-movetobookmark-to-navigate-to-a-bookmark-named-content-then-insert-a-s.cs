using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare file paths
        string outputFolder = Directory.GetCurrentDirectory();
        string destPath = Path.Combine(outputFolder, "Destination.docx");
        string sourcePath = Path.Combine(outputFolder, "Source.docx");
        string mergedPath = Path.Combine(outputFolder, "Merged.docx");

        // Create destination document with a bookmark named "Content"
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("This is the beginning of the destination document.");
        destBuilder.StartBookmark("Content");
        destBuilder.Writeln("[Content will be inserted here]");
        destBuilder.EndBookmark("Content");
        destBuilder.Writeln("This is the end of the destination document.");
        destDoc.Save(destPath, SaveFormat.Docx);

        // Create source document with some formatted text
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);
        sourceBuilder.Font.Bold = true;
        sourceBuilder.Writeln("Source Document Text - bold formatting.");
        sourceBuilder.Font.Bold = false;
        sourceBuilder.Writeln("Additional source content.");
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // Load the destination document for editing
        Document loadedDest = new Document(destPath);
        DocumentBuilder builder = new DocumentBuilder(loadedDest);

        // Move to the bookmark "Content" and insert the source document preserving its formatting
        builder.MoveToBookmark("Content");
        builder.InsertDocument(sourceDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the merged document
        loadedDest.Save(mergedPath, SaveFormat.Docx);

        // Validation: ensure merged file exists
        if (!File.Exists(mergedPath))
        {
            throw new InvalidOperationException("Merged document was not created.");
        }

        // Validation: ensure merged document contains text from the source document
        Document mergedDoc = new Document(mergedPath);
        string mergedText = mergedDoc.GetText();
        if (!mergedText.Contains("Source Document Text"))
        {
            throw new InvalidOperationException("Merged document does not contain expected source content.");
        }

        // Optional: indicate success (no interactive console required)
        Console.WriteLine("Document merging completed successfully.");
    }
}

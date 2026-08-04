using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare output folder
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // File paths for the documents
        string sourcePath = Path.Combine(outputDir, "Source.docx");
        string destinationPath = Path.Combine(outputDir, "Destination.docx");
        string mergedPath = Path.Combine(outputDir, "MergedDocument.docx");

        // ---------- Create source DOCX ----------
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);
        sourceBuilder.Writeln("This is the source document content.");
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // ---------- Create destination DOCX with a bookmark named "Content" ----------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("Header of destination document.");
        destBuilder.StartBookmark("Content");
        destBuilder.Writeln("Placeholder inside bookmark.");
        destBuilder.EndBookmark("Content");
        destBuilder.Writeln("Footer of destination document.");
        destDoc.Save(destinationPath, SaveFormat.Docx);

        // ---------- Load the source document ----------
        Document srcToInsert = new Document(sourcePath);

        // ---------- Move to the bookmark and insert the source document ----------
        bool moved = destBuilder.MoveToBookmark("Content");
        if (!moved)
            throw new InvalidOperationException("Bookmark 'Content' not found in the destination document.");

        destBuilder.InsertDocument(srcToInsert, ImportFormatMode.KeepSourceFormatting);

        // ---------- Save the merged document ----------
        destDoc.Save(mergedPath, SaveFormat.Docx);

        // ---------- Validate that the merged document exists and contains source content ----------
        if (!File.Exists(mergedPath))
            throw new FileNotFoundException("Merged document was not created.", mergedPath);

        Document mergedDoc = new Document(mergedPath);
        string mergedText = mergedDoc.GetText();

        if (!mergedText.Contains("This is the source document content."))
            throw new Exception("Merged document does not contain the source document content.");
    }
}

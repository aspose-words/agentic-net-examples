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

        // Create a source DOCX document.
        string sourcePath = Path.Combine(outputDir, "Source.docx");
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("This is the source document content.");
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // Create a destination document containing a bookmark named "Content".
        string destPath = Path.Combine(outputDir, "Destination.docx");
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("Header before bookmark.");
        destBuilder.StartBookmark("Content");
        destBuilder.Writeln("Placeholder inside bookmark.");
        destBuilder.EndBookmark("Content");
        destBuilder.Writeln("Footer after bookmark.");
        destDoc.Save(destPath, SaveFormat.Docx);

        // Load the documents (optional, they are already in memory).
        Document dest = new Document(destPath);
        Document src = new Document(sourcePath);

        // Move to the bookmark and insert the source document with KeepSourceFormatting.
        DocumentBuilder builder = new DocumentBuilder(dest);
        bool moved = builder.MoveToBookmark("Content");
        if (!moved)
        {
            throw new InvalidOperationException("Bookmark 'Content' not found.");
        }
        builder.InsertDocument(src, ImportFormatMode.KeepSourceFormatting);

        // Save the merged result.
        string mergedPath = Path.Combine(outputDir, "Merged.docx");
        dest.Save(mergedPath, SaveFormat.Docx);

        // Validate that the merged file exists and contains the source text.
        if (!File.Exists(mergedPath))
        {
            throw new FileNotFoundException("Merged document was not created.", mergedPath);
        }

        Document mergedDoc = new Document(mergedPath);
        string mergedText = mergedDoc.GetText();
        if (!mergedText.Contains("This is the source document content."))
        {
            throw new Exception("Merged document does not contain source content.");
        }
    }
}

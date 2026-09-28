using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        string destPath = "Destination.docx";
        string srcPath = "Source.docx";
        string mergedPath = "Merged.docx";

        // Create destination document with a bookmark named "Content"
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("Start of destination document.");
        destBuilder.StartBookmark("Content");
        destBuilder.Writeln("Placeholder text inside bookmark.");
        destBuilder.EndBookmark("Content");
        destBuilder.Writeln("End of destination document.");
        destDoc.Save(destPath, SaveFormat.Docx);

        // Create source document that will be inserted
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.Font.Bold = true;
        srcBuilder.Writeln("This is content from the source document.");
        srcBuilder.Font.Bold = false;
        srcBuilder.Writeln("Additional source paragraph.");
        srcDoc.Save(srcPath, SaveFormat.Docx);

        // Load the destination document for modification
        Document destination = new Document(destPath);
        Document source = new Document(srcPath);
        DocumentBuilder builder = new DocumentBuilder(destination);

        // Move to the bookmark named "Content"
        if (!builder.MoveToBookmark("Content"))
        {
            throw new InvalidOperationException("Bookmark 'Content' not found in the destination document.");
        }

        // Insert the source document at the bookmark, preserving its formatting
        builder.InsertDocument(source, ImportFormatMode.KeepSourceFormatting);

        // Save the merged document
        destination.Save(mergedPath, SaveFormat.Docx);

        // Validation: ensure the merged file exists
        if (!File.Exists(mergedPath))
        {
            throw new FileNotFoundException("Merged document was not created.", mergedPath);
        }

        // Validation: ensure content from source document is present
        Document result = new Document(mergedPath);
        string resultText = result.GetText();
        if (!resultText.Contains("This is content from the source document.") ||
            !resultText.Contains("Additional source paragraph."))
        {
            throw new InvalidOperationException("Source document content was not found in the merged result.");
        }

        // Clean up temporary files (optional)
        // File.Delete(destPath);
        // File.Delete(srcPath);
        // File.Delete(mergedPath);
    }
}

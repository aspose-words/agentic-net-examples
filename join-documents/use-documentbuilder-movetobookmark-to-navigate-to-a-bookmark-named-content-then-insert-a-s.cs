using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the destination document and add a bookmark named "Content".
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("Header before bookmark.");
        destBuilder.StartBookmark("Content");
        destBuilder.Writeln("Placeholder text.");
        destBuilder.EndBookmark("Content");
        destBuilder.Writeln("Footer after bookmark.");

        // Create the source document (DOCX) with sample content.
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.Writeln("This is the inserted source document content.");
        srcBuilder.Writeln("Second line of source.");

        // Move the builder cursor to the bookmark and insert the source document,
        // preserving its original formatting.
        if (!destBuilder.MoveToBookmark("Content"))
            throw new InvalidOperationException("Bookmark 'Content' not found in the destination document.");

        destBuilder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the merged document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedOutput.docx");
        destDoc.Save(outputPath, SaveFormat.Docx);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("Merged document was not saved.", outputPath);

        // Load the saved document and confirm that the source text is present.
        Document resultDoc = new Document(outputPath);
        string resultText = resultDoc.GetText();
        if (!resultText.Contains("This is the inserted source document content."))
            throw new Exception("Source content was not found in the merged document.");
    }
}

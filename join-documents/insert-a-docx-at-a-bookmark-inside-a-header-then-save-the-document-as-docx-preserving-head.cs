using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Folder for all generated files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a source DOCX that will be inserted into the header.
        // -----------------------------------------------------------------
        string sourcePath = Path.Combine(outputDir, "Source.docx");
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("This is the inserted document content.");
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Create the destination document with a header that contains a bookmark.
        // -----------------------------------------------------------------
        string resultPath = Path.Combine(outputDir, "Result.docx");
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);

        // Ensure the first section has a primary header.
        destBuilder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // Insert a bookmark named "HeaderBookmark" where the source will be placed.
        destBuilder.StartBookmark("HeaderBookmark");
        destBuilder.EndBookmark("HeaderBookmark");

        // Add some surrounding text to visualize the header.
        destBuilder.Write("Header before bookmark. ");
        destBuilder.MoveToBookmark("HeaderBookmark");
        destBuilder.Write(" [Inserted content will appear here] ");
        destBuilder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        destBuilder.Writeln(" Header after bookmark.");

        // -----------------------------------------------------------------
        // 3. Load the source document and insert it at the bookmark inside the header.
        // -----------------------------------------------------------------
        Document docToInsert = new Document(sourcePath);
        destBuilder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        destBuilder.MoveToBookmark("HeaderBookmark");
        destBuilder.InsertDocument(docToInsert, ImportFormatMode.KeepSourceFormatting);

        // -----------------------------------------------------------------
        // 4. Save the final document preserving header formatting.
        // -----------------------------------------------------------------
        destDoc.Save(resultPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 5. Simple validation that the output file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(resultPath))
        {
            throw new InvalidOperationException($"The result document was not saved to '{resultPath}'.");
        }
    }
}

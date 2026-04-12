using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a source document with a bookmark that encloses the content to extract.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        builder.Writeln("Sample Document");

        // Bookmark named "Extract".
        builder.StartBookmark("Extract");
        builder.Writeln("This paragraph will be extracted. Current date: ");
        // Insert a DATE field and update it immediately so the result text is stored.
        builder.InsertField(FieldType.FieldDate, true);
        builder.Writeln(" End of extracted content.");
        builder.EndBookmark("Extract");

        // Paragraph outside the bookmark.
        builder.Writeln("This paragraph remains in the original document.");

        // Ensure all fields have up‑to‑date results before extraction.
        sourceDoc.UpdateFields();

        // Locate the bookmark.
        Bookmark extractBookmark = sourceDoc.Range.Bookmarks["Extract"];
        if (extractBookmark == null)
            throw new InvalidOperationException("Bookmark 'Extract' was not found.");

        // The bookmark is inside a paragraph; get that paragraph.
        Paragraph sourceParagraph = extractBookmark.BookmarkStart.ParentNode as Paragraph;
        if (sourceParagraph == null)
            throw new InvalidOperationException("The bookmark does not reside within a paragraph.");

        // Create the destination document.
        Document extractedDoc = new Document();

        // Remove the default empty paragraph.
        extractedDoc.FirstSection.Body.RemoveAllChildren();

        // Import the paragraph from the source document into the destination document.
        NodeImporter importer = new NodeImporter(sourceDoc, extractedDoc, ImportFormatMode.KeepSourceFormatting);
        Node importedParagraph = importer.ImportNode(sourceParagraph, true);
        extractedDoc.FirstSection.Body.AppendChild(importedParagraph);

        // Save the extracted document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Extracted.docx");
        extractedDoc.Save(outputPath, SaveFormat.Docx);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The extracted document was not saved.", outputPath);

        Console.WriteLine("Extraction completed. File saved to: " + outputPath);
    }
}

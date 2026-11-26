using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a bookmark that will contain the table and its caption.
        const string bookmarkName = "TableBookmark";
        builder.StartBookmark(bookmarkName);

        // Insert a caption paragraph inside the bookmark and apply a heading style
        // so that it appears in the TOC.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Table 1: Sample Table");

        // Build a simple 2x2 table inside the bookmark.
        builder.StartTable();

        // First row - header cells.
        builder.InsertCell();
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.EndRow();

        // Second row - data cells.
        builder.InsertCell();
        builder.Writeln("Cell A1");
        builder.InsertCell();
        builder.Writeln("Cell A2");
        builder.EndRow();

        builder.EndTable();

        // End the bookmark.
        builder.EndBookmark(bookmarkName);

        // Move to the end of the document and insert a TOC field that references the bookmark.
        builder.MoveToDocumentEnd();
        builder.InsertParagraph();

        // Insert a TOC field with the \\b switch to include entries from the specified bookmark.
        // Use the string overload of InsertField.
        builder.InsertField($"TOC \\b \"{bookmarkName}\"");

        // Update fields so the TOC is generated.
        doc.UpdateFields();

        // Save the document.
        const string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception($"Failed to create the output file: {outputPath}");
    }
}

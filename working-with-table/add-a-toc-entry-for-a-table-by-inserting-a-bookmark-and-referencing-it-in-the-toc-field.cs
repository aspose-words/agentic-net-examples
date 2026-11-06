using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a Table of Contents field at the beginning of the document.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.Writeln(); // Add a blank line after TOC.

        // Insert a bookmark that will be referenced by the TOC entry.
        string bookmarkName = "Table1";
        builder.StartBookmark(bookmarkName);

        // Insert a TC (Table of Contents entry) field that points to the bookmark.
        // This field will be picked up by the TOC when fields are updated.
        builder.InsertField($"TC \"Table 1\" \\f {bookmarkName}");
        builder.Writeln(); // Ensure the TC field is on its own paragraph.

        // Build a simple table.
        builder.StartTable();

        // First row – header.
        builder.InsertCell();
        builder.Write("Header");
        builder.EndRow();

        // Second row – data.
        builder.InsertCell();
        builder.Write("Data");
        builder.EndRow();

        builder.EndTable();

        // End the bookmark after the table.
        builder.EndBookmark(bookmarkName);

        // Update all fields (including TOC) to reflect the new entry.
        doc.UpdateFields();

        // Save the document.
        string outputPath = "TableOfContents.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception($"Failed to create the output file: {outputPath}");
    }
}

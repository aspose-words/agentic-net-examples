using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a bookmark that will surround the table so we can reference it later.
        builder.StartBookmark("MyTable");

        // Build a simple 2x2 table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1,1");
        builder.InsertCell();
        builder.Write("Cell 1,2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Cell 2,1");
        builder.InsertCell();
        builder.Write("Cell 2,2");
        builder.EndRow();
        builder.EndTable();

        // End the bookmark after the table.
        builder.EndBookmark("MyTable");

        // Insert a caption paragraph directly below the table.
        // Use the built‑in "Caption" style.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Caption;
        builder.Writeln("Table 1: Sample table.");

        // Add some regular text and a cross‑reference to the table.
        builder.ParagraphFormat.ClearFormatting(); // reset to normal style
        builder.Writeln();
        builder.Writeln("Reference to the table above:");
        // Insert a REF field that points to the bookmark "MyTable" and makes it a hyperlink.
        builder.InsertField(@" REF MyTable \h ");

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableCaption.docx");
        doc.Save(outputPath);
    }
}

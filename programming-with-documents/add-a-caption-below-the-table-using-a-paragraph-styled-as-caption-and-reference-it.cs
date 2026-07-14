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

        // Insert a caption below the table.
        // The caption will be styled with the built‑in "Caption" style and wrapped in a bookmark.
        builder.StartBookmark("TableCaption");

        // Move to a new paragraph for the caption.
        builder.Writeln();

        // Apply the built‑in Caption style.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Caption;

        // Insert a SEQ field that generates the table number (e.g., "Table 1").
        builder.InsertField("SEQ Table \\* ARABIC");

        // Add the caption text after the number.
        builder.Write(" Sample Table");

        // End the bookmark that surrounds the caption.
        builder.EndBookmark("TableCaption");

        // Add a paragraph that references the caption using a REF field.
        builder.Writeln(); // Ensure we are on a new paragraph.
        builder.Write("See ");
        // The \\h switch makes the reference a hyperlink.
        builder.InsertField(" REF TableCaption \\h ");
        builder.Writeln(" for details.");

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "TableWithCaption.docx");
        doc.Save(outputPath);
    }
}

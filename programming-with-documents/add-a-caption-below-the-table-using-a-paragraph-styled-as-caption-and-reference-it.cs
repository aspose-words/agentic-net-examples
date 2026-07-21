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

        // Insert a simple 2x2 table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("R1C1");
        builder.InsertCell();
        builder.Write("R1C2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("R2C1");
        builder.InsertCell();
        builder.Write("R2C2");
        builder.EndRow();
        builder.EndTable();

        // Add a caption paragraph below the table.
        // The caption is placed inside a bookmark so it can be referenced later.
        builder.StartBookmark("TableCaption");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Caption;
        builder.Writeln("Table 1: Sample table.");
        builder.EndBookmark("TableCaption");

        // Insert some intervening text.
        builder.Writeln();
        builder.Writeln("The following reference points to the table above:");

        // Insert a cross‑reference (REF) field that points to the caption bookmark.
        // The \\h switch makes the reference a hyperlink.
        builder.InsertField("REF TableCaption \\h");

        // Update fields so that the reference shows the correct caption text.
        doc.UpdateFields();

        // Save the document to the local file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
        doc.Save(outputPath);
    }
}

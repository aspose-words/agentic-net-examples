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

        // Insert a TOC field that lists only entries with the label "Table".
        // Switches: \c "Table" – table of figures for label Table,
        // \h – hyperlink, \z – hide page numbers in web layout, \u – use outline levels.
        builder.InsertTableOfContents("\\c \"Table\" \\h \\z \\u");
        builder.InsertBreak(BreakType.PageBreak);

        // Create three sample tables with captions.
        for (int i = 1; i <= 3; i++)
        {
            // Insert a caption paragraph before the table.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Caption;

            // Insert a SEQ field that generates the table number.
            Field seqField = builder.InsertField(FieldType.FieldSequence, true);
            ((FieldSeq)seqField).SequenceIdentifier = "Table";

            // Write the rest of the caption text.
            builder.Write($" Table {i} Caption");
            builder.Writeln(); // End of caption paragraph.

            // Reset paragraph style for the table content.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;

            // Build a simple 2x2 table.
            Table table = builder.StartTable();

            // First row.
            builder.InsertCell();
            builder.Write($"R{i}C1");
            builder.InsertCell();
            builder.Write($"R{i}C2");
            builder.EndRow();

            // Second row.
            builder.InsertCell();
            builder.Write($"R{i}C3");
            builder.InsertCell();
            builder.Write($"R{i}C4");
            builder.EndRow();

            builder.EndTable();

            // Add a blank line after each table.
            builder.Writeln();
        }

        // Update all fields (captions and TOC).
        doc.UpdateFields();

        // Save the document.
        string fileName = "TableOfContentsForTables.docx";
        doc.Save(fileName);

        // Verify that the file was created.
        if (!File.Exists(fileName))
            throw new Exception("Failed to create the output document.");
    }
}

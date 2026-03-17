using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Saving;
using Aspose.Words.Tables; // Added for Table class

class TableOfContentsForTables
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert three tables, each preceded by a caption that uses a SEQ field with the identifier "Table".
        for (int i = 1; i <= 3; i++)
        {
            // Write the caption: "Table {SEQ Table \* ARABIC}: Description"
            builder.Writeln("Table ");
            FieldSeq seq = (FieldSeq)builder.InsertField(FieldType.FieldSequence, true);
            seq.SequenceIdentifier = "Table"; // This SEQ field will be picked up by the TOC.
            builder.Writeln($": Sample table {i}");
            builder.Writeln(); // Add a blank line before the table.

            // Build a simple 2x2 table.
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write($"R{i}C1");
            builder.InsertCell();
            builder.Write($"R{i}C2");
            builder.EndRow();

            builder.InsertCell();
            builder.Write($"R{i}C3");
            builder.InsertCell();
            builder.Write($"R{i}C4");
            builder.EndRow();

            builder.EndTable();
            builder.Writeln(); // Space after the table.
        }

        // Insert a page break before the Table of Figures (TOC for tables).
        builder.InsertBreak(BreakType.PageBreak);

        // Insert a TOC field that will list only the SEQ fields with identifier "Table".
        FieldToc toc = (FieldToc)builder.InsertField(FieldType.FieldTOC, true);
        toc.TableOfFiguresLabel = "Table"; // Restrict the TOC to entries from the "Table" sequence.
        toc.EntrySeparator = "\t"; // Use a tab between entry text and page number.
        toc.InsertHyperlinks = true; // Make entries clickable.

        // Update all fields so the TOC reflects the current document state.
        doc.UpdateFields();

        // Save the document.
        doc.Save("TableOfContentsForTables.docx", SaveFormat.Docx);
    }
}

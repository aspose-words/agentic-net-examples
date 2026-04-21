using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a TOC field that will list only entries with the "Table" label (table captions).
        // The \\c switch specifies the sequence identifier to use.
        FieldToc toc = (FieldToc)builder.InsertField(FieldType.FieldTOC, true);
        toc.TableOfFiguresLabel = "Table"; // Limit TOC to table captions.
        toc.InsertHyperlinks = true;       // Make entries clickable.
        toc.EntrySeparator = "\t";         // Use a tab between entry text and page number.

        // Add a page break after the TOC so the tables start on a new page.
        builder.InsertBreak(BreakType.PageBreak);

        // Insert three sample tables each with a caption that uses a SEQ field with the identifier "Table".
        for (int i = 1; i <= 3; i++)
        {
            // Write the caption: "Table X: Sample table Y"
            builder.Write("Table ");
            FieldSeq seq = (FieldSeq)builder.InsertField(FieldType.FieldSequence, true);
            seq.SequenceIdentifier = "Table";
            builder.Write($": Sample table {i}");
            builder.InsertBreak(BreakType.LineBreak);

            // Build a simple 2x2 table.
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write($"R{i}C1");
            builder.InsertCell();
            builder.Write($"R{i}C2");
            builder.EndRow();

            builder.InsertCell();
            builder.Write($"R{i+1}C1");
            builder.InsertCell();
            builder.Write($"R{i+1}C2");
            builder.EndRow();
            builder.EndTable();

            // Add a blank line after each table.
            builder.InsertBreak(BreakType.ParagraphBreak);
        }

        // Update all fields (TOC and SEQ numbers) so the document reflects the correct values.
        doc.UpdateFields();

        // Save the document to the local file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableOfContentsForTables.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        // The program ends automatically; no user interaction is required.
    }
}

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

        // Insert a TOC field that will list only entries from the "Table" sequence (table captions).
        FieldToc toc = (FieldToc)builder.InsertField(FieldType.FieldTOC, true);
        toc.TableOfFiguresLabel = "Table"; // Use the "Table" sequence identifier.
        toc.InsertHyperlinks = true;       // Make entries clickable.
        builder.Writeln(); // Move cursor after the TOC field.

        // Helper to insert a table with a caption.
        void InsertTableWithCaption(string captionText)
        {
            // Insert the caption using a SEQ field with identifier "Table".
            FieldSeq seq = (FieldSeq)builder.InsertField(FieldType.FieldSequence, true);
            seq.SequenceIdentifier = "Table";
            builder.Write(" "); // Space between the number and the caption text.
            builder.Writeln(captionText);
            builder.Writeln(); // Add a blank line before the table.

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
            builder.EndTable();

            builder.Writeln(); // Add spacing after the table.
        }

        // Insert several tables with captions.
        InsertTableWithCaption("First sample table.");
        InsertTableWithCaption("Second sample table.");
        InsertTableWithCaption("Third sample table.");

        // Update all fields so the TOC reflects the inserted captions.
        doc.UpdateFields();

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TablesTOC.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The document was not saved correctly.");
    }
}

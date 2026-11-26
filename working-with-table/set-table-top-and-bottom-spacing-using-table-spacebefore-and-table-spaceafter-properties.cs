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

        // Add a paragraph before the table.
        builder.Writeln("Paragraph before the table.");

        // Build a simple 2x2 table.
        builder.StartTable();

        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Writeln("Cell 3");
        builder.InsertCell();
        builder.Writeln("Cell 4");
        builder.EndRow();

        builder.EndTable();

        // Insert an empty paragraph after the table to host the bottom spacing.
        builder.Writeln();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // The spacing before/after a table is defined by the paragraphs that
        // precede and follow the table node.
        Paragraph paragraphBefore = table.PreviousSibling as Paragraph;
        Paragraph paragraphAfter = table.NextSibling as Paragraph;

        if (paragraphBefore != null)
        {
            // Space after the preceding paragraph = space below the table.
            paragraphBefore.ParagraphFormat.SpaceAfter = 24; // 24 points = 1/3 inch
        }

        if (paragraphAfter != null)
        {
            // Space before the following paragraph = space above the table.
            paragraphAfter.ParagraphFormat.SpaceBefore = 12; // 12 points = 1/6 inch
        }

        // Save the document.
        string outputPath = "TableSpacing.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");
    }
}

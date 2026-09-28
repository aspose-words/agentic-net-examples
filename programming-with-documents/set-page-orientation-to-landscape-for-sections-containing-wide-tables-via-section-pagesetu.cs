using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First section: will contain a wide table and be set to landscape orientation.
        builder.Writeln("Section with wide table (landscape):");

        // Create a table with several columns to make it wide.
        Table table = builder.StartTable();

        // Add header row.
        for (int col = 0; col < 5; col++)
        {
            builder.InsertCell();
            // Set a relatively wide cell width.
            builder.CellFormat.Width = 100; // points
            builder.Writeln($"Header {col + 1}");
        }
        builder.EndRow();

        // Add a few data rows.
        for (int row = 0; row < 3; row++)
        {
            for (int col = 0; col < 5; col++)
            {
                builder.InsertCell();
                builder.Writeln($"R{row + 1}C{col + 1}");
            }
            builder.EndRow();
        }

        builder.EndTable();

        // Set the orientation of the first section to landscape.
        doc.Sections[0].PageSetup.Orientation = Orientation.Landscape;

        // Insert a new section that will keep the default portrait orientation.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section with normal orientation (portrait).");

        // Save the document to disk.
        string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}

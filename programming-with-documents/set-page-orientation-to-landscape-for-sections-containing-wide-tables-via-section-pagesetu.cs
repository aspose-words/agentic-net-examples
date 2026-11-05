using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a wide table (10 columns, 100% width) into the first section.
        Table table = builder.StartTable();
        for (int col = 0; col < 10; col++)
        {
            builder.InsertCell();
            builder.Writeln($"Col {col + 1}");
        }
        builder.EndRow();
        builder.EndTable();

        // Ensure the table occupies the full page width.
        // Use the correct PreferredWidth factory method.
        table.PreferredWidth = PreferredWidth.FromPercent(100);

        // Add a second section with a normal table (3 columns) to demonstrate selective orientation.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        Table smallTable = builder.StartTable();
        for (int col = 0; col < 3; col++)
        {
            builder.InsertCell();
            builder.Writeln($"Small {col + 1}");
        }
        builder.EndRow();
        builder.EndTable();

        // Iterate through sections and set orientation to landscape for sections containing wide tables.
        foreach (Section section in doc.Sections)
        {
            bool containsWideTable = false;

            // Look for any table in the current section.
            NodeCollection tables = section.Body.GetChildNodes(NodeType.Table, true);
            foreach (Table tbl in tables)
            {
                // Consider a table wide if its preferred width is 100% of the page.
                if (tbl.PreferredWidth.Type == PreferredWidthType.Percent && tbl.PreferredWidth.Value >= 100)
                {
                    containsWideTable = true;
                    break;
                }
            }

            if (containsWideTable)
            {
                // Set page orientation to landscape.
                section.PageSetup.Orientation = Orientation.Landscape;
            }
        }

        // Save the document.
        string outputPath = "WideTableLandscape.docx";
        doc.Save(outputPath);

        // Optional verification that the file was created.
        if (File.Exists(outputPath))
        {
            // Document saved successfully.
        }
    }
}

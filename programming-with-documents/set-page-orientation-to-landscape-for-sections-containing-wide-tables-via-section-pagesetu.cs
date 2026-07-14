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

        // Build a wide table (10 columns, each 100 points wide).
        builder.StartTable();
        for (int col = 0; col < 10; col++)
        {
            builder.InsertCell();
            // Set a fixed width for each cell to make the table wide.
            builder.CellFormat.Width = 100;
            builder.Write($"Column {col + 1}");
        }
        builder.EndRow();
        builder.EndTable();

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "WideTableLandscape.docx");

        // Iterate through each section and check for wide tables.
        foreach (Section section in doc.Sections)
        {
            bool hasWideTable = false;

            // Get all tables in the current section.
            NodeCollection tables = section.GetChildNodes(NodeType.Table, true);
            foreach (Table table in tables)
            {
                // Calculate the total width of the first row (assumes uniform column widths).
                double totalWidth = 0;
                if (table.Rows.Count > 0)
                {
                    foreach (Cell cell in table.Rows[0].Cells)
                    {
                        totalWidth += cell.CellFormat.Width;
                    }
                }

                // If the table width exceeds the page width, mark the section.
                if (totalWidth > section.PageSetup.PageWidth)
                {
                    hasWideTable = true;
                    break;
                }
            }

            // Set orientation to landscape for sections that contain a wide table.
            if (hasWideTable)
            {
                section.PageSetup.Orientation = Orientation.Landscape;
            }
        }

        // Save the document.
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}

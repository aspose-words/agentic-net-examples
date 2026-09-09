using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class SetLandscapeForWideTables
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -------------------------------------------------
        // Section 1 – normal table (portrait orientation).
        // -------------------------------------------------
        builder.Writeln("Section 1 – normal table (portrait).");
        Table normalTable = builder.StartTable();
        // Two columns, default width.
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Start a new section for the wide table.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // -------------------------------------------------
        // Section 2 – wide table (will be set to landscape).
        // -------------------------------------------------
        builder.Writeln("Section 2 – wide table (should become landscape).");
        Table wideTable = builder.StartTable();

        // Create a table with many columns and explicit widths so that it exceeds the page width.
        const int columnCount = 10;
        const double cellWidth = 100; // points

        for (int i = 0; i < columnCount; i++)
        {
            builder.InsertCell();
            builder.CellFormat.Width = cellWidth;
            builder.Write($"Col {i + 1}");
        }
        builder.EndRow();
        builder.EndTable();

        // -------------------------------------------------
        // Detect sections that contain a table wider than the page
        // and set those sections to landscape orientation.
        // -------------------------------------------------
        foreach (Section section in doc.Sections)
        {
            // Retrieve all tables in the current section.
            NodeCollection tables = section.GetChildNodes(NodeType.Table, true);
            foreach (Table table in tables)
            {
                // Calculate the total width of the table by summing the widths of its first row's cells.
                double totalWidth = 0;
                if (table.Rows.Count > 0)
                {
                    foreach (Cell cell in table.Rows[0].Cells)
                    {
                        // If a cell has an explicit width, use it; otherwise use the default width.
                        totalWidth += cell.CellFormat.Width > 0 ? cell.CellFormat.Width : 0;
                    }
                }

                // If the table's width exceeds the page width, switch the section to landscape.
                if (totalWidth > section.PageSetup.PageWidth)
                {
                    section.PageSetup.Orientation = Orientation.Landscape;
                    // No need to check other tables in this section.
                    break;
                }
            }
        }

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "WideTableLandscape.docx");
        doc.Save(outputPath);
    }
}

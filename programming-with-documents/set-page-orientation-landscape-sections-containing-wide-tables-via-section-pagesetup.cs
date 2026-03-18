using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

class SetLandscapeForWideTables
{
    static void Main()
    {
        // Directory where the resulting document will be saved.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a wide table (each cell width = 500 points, which is wider than a default portrait page).
        Table wideTable = builder.StartTable();
        for (int col = 0; col < 3; col++)
        {
            builder.InsertCell();
            builder.CellFormat.Width = 500; // Wide cell.
            builder.Write($"Cell {col + 1}");
        }
        builder.EndRow();
        builder.EndTable();

        // After the table is inserted, examine each section.
        foreach (Section section in doc.Sections)
        {
            // Determine the maximum width of any table in this section.
            double maxTableWidth = 0;
            foreach (Node node in section.GetChildNodes(NodeType.Table, true))
            {
                Table table = (Table)node;
                // Table width is the sum of the widths of the cells in the first row.
                double tableWidth = 0;
                Row firstRow = table.FirstRow;
                if (firstRow != null)
                {
                    foreach (Cell cell in firstRow.Cells)
                        tableWidth += cell.CellFormat.Width;
                }
                if (tableWidth > maxTableWidth)
                    maxTableWidth = tableWidth;
            }

            // If a table is wider than the current page width, switch to landscape.
            if (maxTableWidth > section.PageSetup.PageWidth)
                section.PageSetup.Orientation = Orientation.Landscape;
        }

        // Save the document.
        string outputPath = Path.Combine(artifactsDir, "WideTableLandscape.docx");
        doc.Save(outputPath, SaveFormat.Docx);
    }
}

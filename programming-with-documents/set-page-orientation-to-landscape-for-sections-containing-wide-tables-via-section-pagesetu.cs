using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare output directory and file name.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "WideTableLandscape.docx");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -----------------------------------------------------------------
        // Section 1 – normal (portrait) orientation with a regular table.
        // -----------------------------------------------------------------
        builder.Writeln("Section 1: Portrait orientation with a normal table.");
        InsertSimpleTable(builder, columns: 3, rows: 3);

        // Start a new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // -----------------------------------------------------------------
        // Section 2 – will contain a wide table, set orientation to landscape.
        // -----------------------------------------------------------------
        builder.PageSetup.Orientation = Orientation.Landscape;
        builder.Writeln("Section 2: Landscape orientation for a wide table.");
        InsertWideTable(builder, columns: 10, rows: 3);

        // Save the document.
        doc.Save(outputPath);
    }

    // Inserts a simple table with the specified number of columns and rows.
    private static void InsertSimpleTable(DocumentBuilder builder, int columns, int rows)
    {
        Table table = builder.StartTable();
        for (int r = 0; r < rows; r++)
        {
            for (int c = 0; c < columns; c++)
            {
                builder.InsertCell();
                builder.Write($"R{r + 1}C{c + 1}");
            }
            builder.EndRow();
        }
        builder.EndTable();
    }

    // Inserts a wide table. Each cell is given a fixed width to ensure the table exceeds the page width.
    private static void InsertWideTable(DocumentBuilder builder, int columns, int rows)
    {
        Table table = builder.StartTable();

        // First row – set a fixed width for each cell.
        for (int c = 0; c < columns; c++)
        {
            builder.InsertCell();
            builder.CellFormat.Width = 100; // 100 points ≈ 1.39 inches.
            builder.Write($"Col {c + 1}");
        }
        builder.EndRow();

        // Remaining rows.
        for (int r = 1; r < rows; r++)
        {
            for (int c = 0; c < columns; c++)
            {
                builder.InsertCell();
                builder.Write($"R{r + 1}C{c + 1}");
            }
            builder.EndRow();
        }

        builder.EndTable();
    }
}

using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder to construct its contents.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        Table table = builder.StartTable();

        // ----- Header row -----
        builder.InsertCell();
        builder.Write("Product");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.InsertCell();
        builder.Write("Price");
        builder.EndRow();

        // ----- Data rows -----
        for (int i = 1; i <= 6; i++)
        {
            builder.InsertCell();
            builder.Write($"Item {i}");
            builder.InsertCell();
            builder.Write((i * 10).ToString());
            builder.InsertCell();
            builder.Write($"${i * 2}.00");
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Apply shading: header row gets a distinct color,
        // data rows alternate between two colors based on row index parity.
        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
        {
            Row row = table.Rows[rowIndex];
            Color shade = rowIndex == 0
                ? Color.LightGray                                   // Header row
                : (rowIndex % 2 == 0 ? Color.LightBlue : Color.LightCyan); // Alternating rows

            foreach (Cell cell in row.Cells)
            {
                cell.CellFormat.Shading.BackgroundPatternColor = shade;
            }
        }

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableAlternatingColors.docx");
        doc.Save(outputPath);
    }
}

using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a sample table with a header row and some numeric data.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Product");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.InsertCell();
        builder.Write("Price");
        builder.EndRow();

        // Data rows.
        AddDataRow(builder, "Apple", 10, 2.5);
        AddDataRow(builder, "Banana", 5, 1.2);
        AddDataRow(builder, "Carrot", 8, 0.8);

        // Finish the table.
        builder.EndTable();

        // Calculate column sums (skip the header row).
        int columnCount = table.Rows[0].Cells.Count;
        double[] sums = new double[columnCount];

        // Start from row index 1 to skip the header.
        for (int rowIdx = 1; rowIdx < table.Rows.Count; rowIdx++)
        {
            Row row = table.Rows[rowIdx];
            for (int colIdx = 0; colIdx < columnCount; colIdx++)
            {
                string cellText = row.Cells[colIdx].ToString(SaveFormat.Text).Trim();

                // Try to parse a double; if it fails, ignore (e.g., the first column with product names).
                if (double.TryParse(cellText, out double value))
                {
                    sums[colIdx] += value;
                }
            }
        }

        // Insert a new footer row with the totals.
        Row totalRow = new Row(doc);
        table.AppendChild(totalRow);

        for (int colIdx = 0; colIdx < columnCount; colIdx++)
        {
            Cell cell = new Cell(doc);
            totalRow.AppendChild(cell);
            cell.AppendChild(new Paragraph(doc));

            string text = colIdx == 0 ? "Total" : sums[colIdx].ToString("0.##");
            Run run = new Run(doc, text);
            run.Font.Bold = true; // Make totals bold.
            cell.FirstParagraph.AppendChild(run);

            // Apply a light gray shading to the footer row.
            cell.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
        }

        // Save the document.
        string outputPath = "TableWithFooterTotals.docx";
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }

    // Helper method to add a data row with product name, quantity and price.
    private static void AddDataRow(DocumentBuilder builder, string product, int quantity, double price)
    {
        builder.InsertCell();
        builder.Write(product);
        builder.InsertCell();
        builder.Write(quantity.ToString());
        builder.InsertCell();
        builder.Write(price.ToString("0.##"));
        builder.EndRow();
    }
}

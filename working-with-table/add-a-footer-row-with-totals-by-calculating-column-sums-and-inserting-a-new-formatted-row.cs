using System;
using System.IO;
using System.Linq;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a sample table with a header and some numeric data.
        builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.InsertCell();
        builder.Write("Price");
        builder.EndRow();

        // Data rows.
        for (int i = 1; i <= 5; i++)
        {
            builder.InsertCell();
            builder.Write($"Item {i}");
            builder.InsertCell();
            builder.Write((i * 2).ToString());                     // Quantity
            builder.InsertCell();
            builder.Write((i * 10.5).ToString("F2"));               // Price
            builder.EndRow();
        }

        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        int columnCount = table.Rows[0].Cells.Count;

        // Calculate sums for each numeric column (skip the first column which holds text).
        double[] sums = new double[columnCount];
        foreach (Row row in table.Rows.Skip(1)) // Skip header.
        {
            for (int col = 1; col < columnCount; col++) // Start from column 1 (Quantity).
            {
                string text = row.Cells[col].GetText().Trim();
                if (double.TryParse(text, out double value))
                {
                    sums[col] += value;
                }
            }
        }

        // Add a footer row with totals.
        Row footerRow = new Row(doc);
        table.Rows.Add(footerRow);

        for (int col = 0; col < columnCount; col++)
        {
            Cell cell = new Cell(doc);
            footerRow.Cells.Add(cell);

            Paragraph para = new Paragraph(doc);
            cell.AppendChild(para);

            Run run = new Run(doc);
            para.AppendChild(run);

            if (col == 0)
                run.Text = "Total";
            else
                run.Text = sums[col].ToString("F2");

            // Apply formatting: bold text and light gray shading.
            run.Font.Bold = true;
            cell.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
        }

        // Save the document.
        string outputPath = "TableWithFooter.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");
    }
}

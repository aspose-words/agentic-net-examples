using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Path for the output document.
        string outputPath = "TableWithFooter.docx";

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a sample table with numeric data (3 columns, 4 rows).
        builder.StartTable();

        // Header row
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.InsertCell();
        builder.Write("Price");
        builder.EndRow();

        // Data rows
        AddDataRow(builder, "Apple", 10, 0.5);
        AddDataRow(builder, "Banana", 5, 0.3);
        AddDataRow(builder, "Orange", 8, 0.4);
        AddDataRow(builder, "Grape", 12, 0.6);

        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Determine number of columns.
        int columnCount = table.Rows[0].Cells.Count;

        // Calculate column sums (skip the header row at index 0).
        double[] sums = new double[columnCount];
        for (int i = 1; i < table.Rows.Count; i++)
        {
            Row row = table.Rows[i];
            for (int col = 0; col < columnCount; col++)
            {
                string text = row.Cells[col].GetText().TrimEnd('\a'); // Remove cell end marker.
                if (double.TryParse(text, out double value))
                {
                    sums[col] += value;
                }
            }
        }

        // Create a new footer row.
        Row footerRow = new Row(doc);
        for (int col = 0; col < columnCount; col++)
        {
            // Create a new cell.
            Cell cell = new Cell(doc);

            // Create a paragraph that holds the sum value.
            Paragraph para = new Paragraph(doc);
            Run run = new Run(doc, sums[col].ToString());
            run.Font.Bold = true; // Make the total bold.
            para.AppendChild(run);
            cell.AppendChild(para);

            // Apply shading to highlight the footer row.
            cell.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;

            // Add the cell to the footer row.
            footerRow.Cells.Add(cell);
        }

        // Add the footer row to the table.
        table.Rows.Add(footerRow);

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        // Indicate success.
        Console.WriteLine("Document created successfully: " + Path.GetFullPath(outputPath));
    }

    // Helper method to add a data row with three cells.
    private static void AddDataRow(DocumentBuilder builder, string item, int quantity, double price)
    {
        builder.InsertCell();
        builder.Write(item);
        builder.InsertCell();
        builder.Write(quantity.ToString());
        builder.InsertCell();
        builder.Write(price.ToString());
        builder.EndRow();
    }
}

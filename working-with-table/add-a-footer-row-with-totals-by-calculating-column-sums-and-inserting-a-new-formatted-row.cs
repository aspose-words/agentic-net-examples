using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a sample table with a header row and some numeric data.
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
        string[,] data = {
            { "Apple", "10", "0.5" },
            { "Banana", "5", "0.3" },
            { "Orange", "8", "0.4" }
        };

        for (int i = 0; i < data.GetLength(0); i++)
        {
            for (int j = 0; j < data.GetLength(1); j++)
            {
                builder.InsertCell();
                builder.Write(data[i, j]);
            }
            builder.EndRow();
        }

        // End the table construction.
        builder.EndTable();

        // Retrieve the created table.
        Table table = doc.GetChildNodes(NodeType.Table, true)[0] as Table;
        if (table == null)
            throw new InvalidOperationException("Table was not created.");

        int columnCount = table.Rows[0].Cells.Count;
        double[] columnSums = new double[columnCount];

        // Calculate sums for numeric columns (skip the first column which is text).
        for (int rowIdx = 1; rowIdx < table.Rows.Count; rowIdx++)
        {
            Row row = table.Rows[rowIdx];
            for (int colIdx = 0; colIdx < columnCount; colIdx++)
            {
                string text = row.Cells[colIdx].GetText().TrimEnd('\a'); // Remove cell end marker.
                if (double.TryParse(text, out double value))
                {
                    columnSums[colIdx] += value;
                }
            }
        }

        // Insert a new footer row with totals.
        Row footerRow = new Row(doc);
        table.Rows.Add(footerRow);

        for (int colIdx = 0; colIdx < columnCount; colIdx++)
        {
            Cell cell = new Cell(doc);
            footerRow.Cells.Add(cell);

            // First column gets the label "Total".
            string cellText = colIdx == 0 ? "Total" : columnSums[colIdx].ToString();

            // Add the text to the cell.
            Paragraph para = new Paragraph(doc);
            Run run = new Run(doc, cellText);
            run.Font.Bold = true; // Make the totals bold.
            para.AppendChild(run);
            cell.AppendChild(para);

            // Apply shading to the footer cells.
            cell.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
        }

        // Save the document.
        string outputPath = "TableWithFooter.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);
    }
}

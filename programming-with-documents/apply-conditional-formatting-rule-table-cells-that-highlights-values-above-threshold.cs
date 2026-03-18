using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

class ConditionalFormattingExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple table with numeric values.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.EndRow();

        // Data rows.
        AddDataRow(builder, "Apples", 30);
        AddDataRow(builder, "Bananas", 75);
        AddDataRow(builder, "Carrots", 55);
        AddDataRow(builder, "Dates", 20);

        builder.EndTable();

        // Define the threshold above which cells will be highlighted.
        double threshold = 50.0;

        // Iterate through all rows except the header (row index 0).
        for (int rowIdx = 1; rowIdx < table.Rows.Count; rowIdx++)
        {
            Row row = table.Rows[rowIdx];
            // The quantity is in the second cell (index 1).
            Cell quantityCell = row.Cells[1];

            // Try to parse the cell text to a double.
            if (double.TryParse(quantityCell.GetText().Trim(), out double value))
            {
                if (value > threshold)
                {
                    // Highlight the cell with a yellow background.
                    quantityCell.CellFormat.Shading.BackgroundPatternColor = Color.Yellow;

                    // Make the text bold – CellFormat has no Font property, so we format the runs inside the cell.
                    foreach (Paragraph para in quantityCell.Paragraphs)
                    {
                        foreach (Run run in para.Runs)
                        {
                            run.Font.Bold = true;
                        }
                    }
                }
            }
        }

        // Save the document.
        doc.Save("ConditionalFormattingTable.docx");
    }

    // Helper method to add a data row to the table.
    private static void AddDataRow(DocumentBuilder builder, string item, double quantity)
    {
        builder.InsertCell();
        builder.Write(item);
        builder.InsertCell();
        builder.Write(quantity.ToString());
        builder.EndRow();
    }
}

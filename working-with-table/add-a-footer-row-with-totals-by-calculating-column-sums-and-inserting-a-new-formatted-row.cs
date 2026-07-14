using System;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple table with a header row and some numeric data.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.InsertCell();
        builder.Write("Price");
        builder.EndRow();

        // Data rows.
        AddDataRow(builder, "Apples", 10, 1.20);
        AddDataRow(builder, "Bananas", 5, 0.80);
        AddDataRow(builder, "Carrots", 8, 0.50);
        AddDataRow(builder, "Dates", 3, 2.00);

        // Finish the table and obtain the Table object.
        table = builder.EndTable();

        // Calculate column sums for the numeric columns (Quantity and Price).
        double quantitySum = 0;
        double priceSum = 0;

        // Skip the header row (index 0).
        for (int i = 1; i < table.Rows.Count; i++)
        {
            Row row = table.Rows[i];
            // Quantity column (index 1).
            double qty = ParseDouble(row.Cells[1].GetText());
            quantitySum += qty;

            // Price column (index 2).
            double price = ParseDouble(row.Cells[2].GetText());
            priceSum += price;
        }

        // Add a footer row with the totals.
        Row footerRow = new Row(doc);
        table.Rows.Add(footerRow);

        // First cell: label "Total".
        Cell labelCell = new Cell(doc);
        labelCell.AppendChild(new Paragraph(doc));
        Run labelRun = new Run(doc, "Total");
        labelRun.Font.Bold = true;
        labelCell.FirstParagraph.AppendChild(labelRun);
        // Apply shading to the footer row cells.
        labelCell.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
        footerRow.Cells.Add(labelCell);

        // Second cell: sum of Quantity.
        Cell qtyCell = new Cell(doc);
        qtyCell.AppendChild(new Paragraph(doc));
        Run qtyRun = new Run(doc, quantitySum.ToString());
        qtyRun.Font.Bold = true;
        qtyCell.FirstParagraph.AppendChild(qtyRun);
        qtyCell.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
        footerRow.Cells.Add(qtyCell);

        // Third cell: sum of Price.
        Cell priceCell = new Cell(doc);
        priceCell.AppendChild(new Paragraph(doc));
        Run priceRun = new Run(doc, priceSum.ToString("F2"));
        priceRun.Font.Bold = true;
        priceCell.FirstParagraph.AppendChild(priceRun);
        priceCell.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
        footerRow.Cells.Add(priceCell);

        // Save the document.
        doc.Save("TableWithFooter.docx");
    }

    // Helper method to add a data row to the table.
    private static void AddDataRow(DocumentBuilder builder, string item, int quantity, double price)
    {
        builder.InsertCell();
        builder.Write(item);
        builder.InsertCell();
        builder.Write(quantity.ToString());
        builder.InsertCell();
        builder.Write(price.ToString("F2"));
        builder.EndRow();
    }

    // Helper method to parse a double from cell text (removes any trailing control characters).
    private static double ParseDouble(string text)
    {
        // Cell.GetText() returns text ending with a cell marker (\\a). Trim it.
        string trimmed = text.Trim('\a', '\r', '\n', '\t', ' ');
        double result;
        double.TryParse(trimmed, out result);
        return result;
    }
}

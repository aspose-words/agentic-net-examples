using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Sample data to populate the table.
        string[,] data = new string[,]
        {
            { "Product", "Description", "Price", "Quantity" },
            { "Apple", "Fresh red apples", "$1.20", "50" },
            { "Banana", "Ripe yellow bananas", "$0.80", "120" },
            { "Cherry", "Sweet cherries from the valley", "$2.50", "30" }
        };

        // Start the table.
        Table table = builder.StartTable();

        // Populate the table with data.
        for (int row = 0; row < data.GetLength(0); row++)
        {
            for (int col = 0; col < data.GetLength(1); col++)
            {
                // Insert cell text.
                builder.InsertCell();
                builder.Write(data[row, col]);

                // Apply a simple border to each cell.
                Cell cell = (Cell)builder.CurrentParagraph.ParentNode;
                cell.CellFormat.Borders.LineWidth = 0.5;
                cell.CellFormat.Borders.Color = System.Drawing.Color.Black;
            }

            // End the current row.
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Adjust column widths proportionally to fit the content.
        // AutoFitToContents makes each column width fit its longest cell content.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document.
        string outputPath = "AdjustedTable.docx";
        doc.Save(outputPath);
    }
}

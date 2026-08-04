using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableFormulaUpdate
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table.
            Table table = builder.StartTable();

            // Header row.
            builder.InsertCell();
            builder.Write("Item");
            builder.InsertCell();
            builder.Write("Price");
            builder.EndRow();

            // First data row.
            builder.InsertCell();
            builder.Write("Apple");
            builder.InsertCell();
            builder.Write("10");
            builder.EndRow();

            // Second data row.
            builder.InsertCell();
            builder.Write("Orange");
            builder.InsertCell();
            builder.Write("15");
            builder.EndRow();

            // Total row with a formula field that sums the values above in the Price column.
            builder.InsertCell();
            builder.Write("Total");
            builder.InsertCell();
            // Insert a formula field =SUM(ABOVE). This overload updates the field automatically.
            builder.InsertField("=SUM(ABOVE)");
            builder.EndRow();

            // Finish building the table.
            builder.EndTable();

            // Insert a new data row before the total row.
            // The total row is currently the last row in the table.
            int totalRowIndex = table.Rows.Count - 1;

            // Create a new row.
            Row newRow = new Row(doc);
            // First cell of the new row.
            Cell cellItem = new Cell(doc);
            cellItem.AppendChild(new Paragraph(doc));
            cellItem.FirstParagraph.AppendChild(new Run(doc, "Banana"));
            newRow.AppendChild(cellItem);
            // Second cell of the new row.
            Cell cellPrice = new Cell(doc);
            cellPrice.AppendChild(new Paragraph(doc));
            cellPrice.FirstParagraph.AppendChild(new Run(doc, "20"));
            newRow.AppendChild(cellPrice);

            // Insert the new row before the total row.
            table.Rows.Insert(totalRowIndex, newRow);

            // Recalculate all fields in the document (including the formula field).
            doc.UpdateFields();

            // Save the document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "UpdatedTableFormulas.docx");
            doc.Save(outputPath);
        }
    }
}

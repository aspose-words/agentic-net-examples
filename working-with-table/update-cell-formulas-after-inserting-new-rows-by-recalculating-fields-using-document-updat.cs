using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableFormulaExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start building a table.
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
            builder.Write("2");
            builder.EndRow();

            // Second data row.
            builder.InsertCell();
            builder.Write("Banana");
            builder.InsertCell();
            builder.Write("3");
            builder.EndRow();

            // Total row with a formula field that sums the values above in the Price column.
            builder.InsertCell();
            builder.Write("Total");
            builder.InsertCell();
            // Insert a formula field. The result will be updated later.
            builder.InsertField("=SUM(ABOVE)", "");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // At this point the table has 4 rows (header, two data rows, total row).
            // Retrieve the total row so we can insert a new data row before it.
            Row totalRow = table.LastRow;

            // Create a new data row to be inserted before the total row.
            Row newRow = new Row(doc);
            // First cell: item name.
            Cell itemCell = new Cell(doc);
            itemCell.AppendChild(new Paragraph(doc));
            itemCell.FirstParagraph.AppendChild(new Run(doc, "Cherry"));
            newRow.AppendChild(itemCell);
            // Second cell: price value.
            Cell priceCell = new Cell(doc);
            priceCell.AppendChild(new Paragraph(doc));
            priceCell.FirstParagraph.AppendChild(new Run(doc, "5"));
            newRow.AppendChild(priceCell);

            // Insert the new row before the total row.
            table.InsertBefore(newRow, totalRow);

            // Recalculate all fields in the document (including the formula field).
            doc.UpdateFields();

            // Prepare output directory and file path.
            string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
            Directory.CreateDirectory(artifactsDir);
            string outputPath = Path.Combine(artifactsDir, "TableWithFormula.docx");

            // Save the document.
            doc.Save(outputPath);
        }
    }
}

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableUpdateFields
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple 4‑row table.
            // Row 1 – header.
            builder.StartTable();
            builder.InsertCell();
            builder.Write("Item");
            builder.InsertCell();
            builder.Write("Quantity");
            builder.EndRow();

            // Row 2 – first data row.
            builder.InsertCell();
            builder.Write("Apples");
            builder.InsertCell();
            builder.Write("10");
            builder.EndRow();

            // Row 3 – second data row.
            builder.InsertCell();
            builder.Write("Bananas");
            builder.InsertCell();
            builder.Write("20");
            builder.EndRow();

            // Row 4 – total row with a formula field that sums the values above it.
            builder.InsertCell();
            builder.Write("Total");
            builder.InsertCell();
            // Insert a formula field. Use the string overload which inserts the field code and updates it.
            builder.InsertField("=SUM(ABOVE)");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Save the initial document (optional, shows the state before modification).
            string initialPath = Path.Combine(Directory.GetCurrentDirectory(), "TableWithFormula_Initial.docx");
            doc.Save(initialPath);

            // -----------------------------------------------------------------
            // Modify the quantity of the first item from "10" to "15".
            // Locate the table, then the specific cell, and replace its text.
            // -----------------------------------------------------------------
            Table table = doc.FirstSection.Body.Tables[0];
            // Row index 1 corresponds to the second row (Apples).
            Cell quantityCell = table.Rows[1].Cells[1];
            // Clear existing runs and insert the new value.
            quantityCell.FirstParagraph.Runs.Clear();
            quantityCell.FirstParagraph.AppendChild(new Run(doc, "15"));

            // Recalculate all fields in the document so that the total reflects the change.
            doc.UpdateFields();

            // Save the updated document.
            string updatedPath = Path.Combine(Directory.GetCurrentDirectory(), "TableWithFormula_Updated.docx");
            doc.Save(updatedPath);
        }
    }
}

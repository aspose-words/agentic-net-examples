using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;

namespace AsposeWordsTableFieldUpdate
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple 2‑row table.
            // Row 1 – contains a numeric value.
            // Row 2 – contains a formula field that sums the values above it.
            builder.StartTable();

            // First row, first cell: write a number.
            builder.InsertCell();
            builder.Write("10");
            builder.EndRow();

            // Second row, first cell: insert a SUM(ABOVE) field.
            builder.InsertCell();
            // Use the string overload of InsertField which inserts the field and updates it automatically.
            builder.InsertField("=SUM(ABOVE)");
            builder.EndRow();

            builder.EndTable();

            // Locate the first cell and change its numeric value.
            Table table = doc.FirstSection.Body.Tables[0];
            Cell firstCell = table.Rows[0].Cells[0];

            // Replace the existing run text with a new value.
            // Ensure the cell has at least one paragraph and one run.
            if (firstCell.FirstParagraph.Runs.Count == 0)
            {
                firstCell.FirstParagraph.AppendChild(new Run(doc, "20"));
            }
            else
            {
                firstCell.FirstParagraph.Runs[0].Text = "20";
            }

            // Recalculate all fields in the document, including the formula field in the table.
            doc.UpdateFields();

            // Optional: verify the result of the formula field.
            // The document now contains a single field (the formula), accessible via the Range.Fields collection.
            Field formulaField = doc.Range.Fields[0];
            Console.WriteLine("Formula field result after update: " + formulaField?.Result);

            // Save the document to the current working directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "UpdatedTableFields.docx");
            doc.Save(outputPath);
        }
    }
}

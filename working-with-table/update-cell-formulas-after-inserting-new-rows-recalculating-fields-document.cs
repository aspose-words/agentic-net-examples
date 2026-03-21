using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsFieldUpdateExample
{
    class Program
    {
        static void Main()
        {
            // Create a new document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple table with a numeric cell and a formula field.
            Table table = builder.StartTable();

            // Header row.
            builder.InsertCell();
            builder.Write("Value");
            builder.EndRow();

            // Data row with a numeric value.
            builder.InsertCell();
            builder.Write("10");
            builder.EndRow();

            // Row with a formula field that sums the values above.
            builder.InsertCell();
            builder.InsertField("=SUM(ABOVE)", "0");
            builder.EndRow();

            builder.EndTable();

            // Get the first table in the document.
            table = doc.FirstSection.Body.Tables[0];

            // Insert a new row at the end of the table (clone the last row to keep formatting).
            Row newRow = (Row)table.LastRow.Clone(true);
            table.Rows.Add(newRow);

            // Fill the new row with a numeric value that the formula will include.
            newRow.Cells[0].FirstParagraph.AppendChild(new Run(doc, "42"));

            // Recalculate all fields, including the formula field.
            doc.UpdateFields();

            // Save the updated document to the current directory.
            string outputPath = "OutputWithUpdatedFields.docx";
            doc.Save(outputPath);
            Console.WriteLine($"Document saved to {outputPath}");
        }
    }
}

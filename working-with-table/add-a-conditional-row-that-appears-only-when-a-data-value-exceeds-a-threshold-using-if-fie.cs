using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsConditionalRow
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
            builder.Write("Quantity");
            builder.EndRow();

            // First data row.
            builder.InsertCell();
            builder.Write("Apples");
            builder.InsertCell();
            builder.Write("20");
            builder.EndRow();

            // Second data row.
            builder.InsertCell();
            builder.Write("Bananas");
            builder.InsertCell();
            builder.Write("45");
            builder.EndRow();

            // Conditional row: appears only if the quantity exceeds the threshold (30).
            // Insert an IF field that evaluates the expression; if true, the text is shown, otherwise the cell is empty.
            builder.InsertCell();
            builder.Write("Note");
            builder.InsertCell();
            // Use a proper string literal for the IF field code.
            builder.InsertField(@" IF 45 > 30 ""Quantity exceeds threshold"" """" ");
            builder.EndRow();

            // End the table.
            builder.EndTable();

            // Define output path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ConditionalRow.docx");

            // Save the document.
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The document was not saved correctly.");
        }
    }
}

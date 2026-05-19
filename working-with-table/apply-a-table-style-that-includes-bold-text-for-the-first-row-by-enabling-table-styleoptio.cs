using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableStyleExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table and add a few rows.
            Table table = builder.StartTable();

            // First row (header).
            builder.InsertCell();
            builder.Write("Item");
            builder.InsertCell();
            builder.Write("Quantity");
            builder.EndRow();

            // Second row.
            builder.InsertCell();
            builder.Write("Apples");
            builder.InsertCell();
            builder.Write("10");
            builder.EndRow();

            // Third row.
            builder.InsertCell();
            builder.Write("Bananas");
            builder.InsertCell();
            builder.Write("20");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Apply a built‑in table style.
            table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

            // Enable the first‑row conditional formatting.
            table.StyleOptions = TableStyleOptions.FirstRow;

            // Make the text in the first row bold via the conditional style.
            TableStyle tableStyle = (TableStyle)doc.Styles[StyleIdentifier.MediumShading1Accent1];
            tableStyle.ConditionalStyles[ConditionalStyleType.FirstRow].Font.Bold = true;

            // Adjust column widths to fit the content.
            table.AutoFit(AutoFitBehavior.AutoFitToContents);

            // Save the document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableWithBoldFirstRow.docx");
            doc.Save(outputPath);

            // Simple verification that the file was created.
            if (File.Exists(outputPath))
                Console.WriteLine($"Document saved successfully to: {outputPath}");
            else
                throw new InvalidOperationException("Failed to save the document.");
        }
    }
}

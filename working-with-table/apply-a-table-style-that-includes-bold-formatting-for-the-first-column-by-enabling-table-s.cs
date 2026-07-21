using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableStyleExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple 3‑row, 2‑column table.
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
            builder.Write("10");
            builder.EndRow();

            // Second data row.
            builder.InsertCell();
            builder.Write("Bananas");
            builder.InsertCell();
            builder.Write("20");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Create a custom table style.
            TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyTableStyle");

            // Make the first column bold via the conditional style.
            customStyle.ConditionalStyles[ConditionalStyleType.FirstColumn].Font.Bold = true;

            // Apply the custom style to the table.
            table.Style = customStyle;

            // Enable the FirstColumn option so the conditional style is applied.
            table.StyleOptions = TableStyleOptions.FirstColumn;

            // Optional: adjust column widths to fit the content.
            table.AutoFit(AutoFitBehavior.AutoFitToContents);

            // Save the document.
            const string outputPath = "TableStyleFirstColumnBold.docx";
            doc.Save(outputPath);

            // Simple verification that the file was created.
            if (System.IO.File.Exists(outputPath))
                Console.WriteLine($"Document saved successfully to '{outputPath}'.");
            else
                throw new InvalidOperationException("Failed to create the output document.");
        }
    }
}

using System;
using System.IO;
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

            // Start building the table.
            Table table = builder.StartTable();

            // ---------- Header row (will be styled bold) ----------
            builder.InsertCell();
            builder.Write("Header 1");
            builder.InsertCell();
            builder.Write("Header 2");
            builder.EndRow();

            // ---------- Data rows ----------
            for (int i = 1; i <= 3; i++)
            {
                builder.InsertCell();
                builder.Write($"Item {i}");
                builder.InsertCell();
                builder.Write($"Value {i}");
                builder.EndRow();
            }

            // ---------- Footer row (will be styled italic) ----------
            builder.InsertCell();
            builder.Write("Footer 1");
            builder.InsertCell();
            builder.Write("Footer 2");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Create a custom table style.
            TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyCustomStyle");

            // Apply bold formatting to the first (header) row.
            customStyle.ConditionalStyles[ConditionalStyleType.FirstRow].Font.Bold = true;

            // Apply italic formatting to the last (footer) row.
            customStyle.ConditionalStyles[ConditionalStyleType.LastRow].Font.Italic = true;

            // Assign the style to the table.
            table.Style = customStyle;

            // Enable the conditional formatting for first and last rows.
            table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.LastRow;

            // Save the document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "StyledTable.docx");
            doc.Save(outputPath);

            // Simple validation that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The document was not saved correctly.");
        }
    }
}

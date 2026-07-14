using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

namespace TableStyleExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // -----------------------------------------------------------------
            // Build a simple table with a header row, two data rows and a footer row.
            // -----------------------------------------------------------------
            Table table = builder.StartTable();

            // Header row (first row)
            builder.InsertCell();
            builder.Write("Header Column 1");
            builder.InsertCell();
            builder.Write("Header Column 2");
            builder.EndRow();

            // First data row
            builder.InsertCell();
            builder.Write("Data 1-1");
            builder.InsertCell();
            builder.Write("Data 1-2");
            builder.EndRow();

            // Second data row
            builder.InsertCell();
            builder.Write("Data 2-1");
            builder.InsertCell();
            builder.Write("Data 2-2");
            builder.EndRow();

            // Footer row (last row)
            builder.InsertCell();
            builder.Write("Footer Column 1");
            builder.InsertCell();
            builder.Write("Footer Column 2");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // -----------------------------------------------------------------
            // Create a custom table style.
            // The style will make the first row bold and the last row italic.
            // -----------------------------------------------------------------
            TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyCustomTableStyle");

            // Apply bold formatting to the first (header) row.
            customStyle.ConditionalStyles[ConditionalStyleType.FirstRow].Font.Bold = true;

            // Apply italic formatting to the last (footer) row.
            customStyle.ConditionalStyles[ConditionalStyleType.LastRow].Font.Italic = true;

            // Assign the custom style to the table.
            table.Style = customStyle;

            // Enable the conditional formatting for first and last rows.
            table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.LastRow;

            // -----------------------------------------------------------------
            // Save the document.
            // -----------------------------------------------------------------
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableStyleCustom.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create the output file: {outputPath}");

            // The program finishes automatically; no user interaction required.
        }
    }
}

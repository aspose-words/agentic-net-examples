using System;
using System.IO;
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

            // Build a simple table: header row, two data rows, footer row.
            builder.StartTable();

            // ----- Header row -----
            builder.InsertCell();
            builder.Write("Header 1");
            builder.InsertCell();
            builder.Write("Header 2");
            builder.EndRow();

            // ----- Data row 1 -----
            builder.InsertCell();
            builder.Write("Data 1A");
            builder.InsertCell();
            builder.Write("Data 1B");
            builder.EndRow();

            // ----- Data row 2 -----
            builder.InsertCell();
            builder.Write("Data 2A");
            builder.InsertCell();
            builder.Write("Data 2B");
            builder.EndRow();

            // ----- Footer row -----
            builder.InsertCell();
            builder.Write("Footer 1");
            builder.InsertCell();
            builder.Write("Footer 2");
            builder.EndRow();

            // Finish the table.
            Table table = builder.EndTable();

            // Create a custom table style.
            TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyCustomStyle");

            // Make the first row (header) bold.
            customStyle.ConditionalStyles[ConditionalStyleType.FirstRow].Font.Bold = true;

            // Make the last row (footer) italic.
            customStyle.ConditionalStyles[ConditionalStyleType.LastRow].Font.Italic = true;

            // Apply the custom style to the table.
            table.Style = customStyle;

            // Enable the conditional formatting for first and last rows.
            table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.LastRow;

            // Save the document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CustomTableStyle.docx");
            doc.Save(outputPath);
        }
    }
}

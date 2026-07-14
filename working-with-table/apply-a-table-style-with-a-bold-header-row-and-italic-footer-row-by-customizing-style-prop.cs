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

            // Build a table with a header row, two data rows, and a footer row.
            Table table = builder.StartTable();

            // Header row.
            builder.InsertCell();
            builder.Write("Header 1");
            builder.InsertCell();
            builder.Write("Header 2");
            builder.EndRow();

            // First data row.
            builder.InsertCell();
            builder.Write("Data 1");
            builder.InsertCell();
            builder.Write("Data 2");
            builder.EndRow();

            // Second data row.
            builder.InsertCell();
            builder.Write("Data 3");
            builder.InsertCell();
            builder.Write("Data 4");
            builder.EndRow();

            // Footer row.
            builder.InsertCell();
            builder.Write("Footer 1");
            builder.InsertCell();
            builder.Write("Footer 2");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Create a custom table style.
            TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyCustomStyle");

            // Make the first row (header) bold.
            customStyle.ConditionalStyles[ConditionalStyleType.FirstRow].Font.Bold = true;

            // Make the last row (footer) italic.
            customStyle.ConditionalStyles[ConditionalStyleType.LastRow].Font.Italic = true;

            // Apply the custom style to the table.
            table.Style = customStyle;

            // Enable the conditional formatting for the first and last rows.
            table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.LastRow;

            // Save the document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableWithCustomStyle.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The document was not saved correctly.");
        }
    }
}

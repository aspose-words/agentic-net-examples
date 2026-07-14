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

            // Build a simple 3x2 table.
            Table table = builder.StartTable();

            // First row (header).
            builder.InsertCell();
            builder.Write("Product");
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

            // Create a custom table style.
            TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyBoldFirstRowStyle");

            // Make the first row bold via the conditional style.
            customStyle.ConditionalStyles[ConditionalStyleType.FirstRow].Font.Bold = true;

            // Apply the custom style to the table.
            table.Style = customStyle;

            // Enable the FirstRow style option so that the conditional formatting is applied.
            table.StyleOptions = TableStyleOptions.FirstRow;

            // Save the document.
            doc.Save("TableStyleFirstRowBold.docx");
        }
    }
}

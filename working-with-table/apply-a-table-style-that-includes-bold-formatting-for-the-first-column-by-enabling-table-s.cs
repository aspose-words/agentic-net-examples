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

            // Start a table and add a few rows with two columns each.
            Table table = builder.StartTable();

            // Header row.
            builder.InsertCell();
            builder.Write("Item");
            builder.InsertCell();
            builder.Write("Quantity");
            builder.EndRow();

            // Data rows.
            builder.InsertCell();
            builder.Write("Apples");
            builder.InsertCell();
            builder.Write("10");
            builder.EndRow();

            builder.InsertCell();
            builder.Write("Bananas");
            builder.InsertCell();
            builder.Write("20");
            builder.EndRow();

            // Finish building the table.
            builder.EndTable();

            // Create a custom table style.
            TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyFirstColumnBoldStyle");

            // Make the first column bold via the conditional style.
            customStyle.ConditionalStyles[ConditionalStyleType.FirstColumn].Font.Bold = true;

            // Apply the custom style to the table.
            table.Style = customStyle;

            // Enable the FirstColumn option so the conditional style is applied.
            table.StyleOptions = TableStyleOptions.FirstColumn;

            // Save the document.
            doc.Save("TableStyleFirstColumnBold.docx");
        }
    }
}

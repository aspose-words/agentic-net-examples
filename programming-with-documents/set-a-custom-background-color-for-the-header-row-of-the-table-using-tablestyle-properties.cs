using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableHeaderStyle
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a new table.
            Table table = builder.StartTable();

            // ----- Header row -----
            // First header cell.
            builder.InsertCell();
            builder.Write("Product");
            // Second header cell.
            builder.InsertCell();
            builder.Write("Quantity");
            // End the header row.
            builder.EndRow();

            // ----- Data rows -----
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
            TableStyle headerStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "HeaderRowStyle");

            // Set the background color for the first row (header) using conditional style.
            headerStyle.ConditionalStyles[ConditionalStyleType.FirstRow].Shading.BackgroundPatternColor = Color.LightBlue;

            // Apply the style to the table.
            table.Style = headerStyle;

            // Ensure the style is applied to the first row.
            table.StyleOptions = TableStyleOptions.FirstRow;

            // Save the document to a file.
            doc.Save("HeaderRowStyle.docx");
        }
    }
}

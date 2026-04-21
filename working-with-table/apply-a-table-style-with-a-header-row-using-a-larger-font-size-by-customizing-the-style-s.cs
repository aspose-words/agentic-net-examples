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

            // Start a table.
            Table table = builder.StartTable();

            // ----- Header row (will use larger font via style) -----
            builder.InsertCell();
            builder.Write("Product");
            builder.InsertCell();
            builder.Write("Quantity");
            builder.EndRow();

            // ----- Data rows -----
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

            // Finish the table.
            builder.EndTable();

            // Create a custom table style.
            TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyCustomTableStyle");

            // Set a larger font size for the header row (first row conditional style).
            tableStyle.ConditionalStyles[ConditionalStyleType.FirstRow].Font.Size = 16; // larger font

            // Optionally set a default font size for the rest of the table.
            tableStyle.Font.Size = 12;

            // Apply the style to the table.
            table.Style = tableStyle;

            // Enable the FirstRow conditional formatting so the header style is applied.
            table.StyleOptions = TableStyleOptions.FirstRow;

            // Save the document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableWithHeaderStyle.docx");
            doc.Save(outputPath);
        }
    }
}

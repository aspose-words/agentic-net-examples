using Aspose.Words;
using Aspose.Words.Tables;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple table with a header row.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Product");
        builder.InsertCell();
        builder.Write("Price");
        builder.EndRow();

        // First data row.
        builder.InsertCell();
        builder.Write("Apple");
        builder.InsertCell();
        builder.Write("$1");
        builder.EndRow();

        // Second data row.
        builder.InsertCell();
        builder.Write("Banana");
        builder.InsertCell();
        builder.Write("$2");
        builder.EndRow();

        builder.EndTable();

        // Create a custom table style.
        TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "HeaderStyle");

        // Set a different background color for the header (first row) using the conditional style.
        tableStyle.ConditionalStyles[ConditionalStyleType.FirstRow].Shading.BackgroundPatternColor = Color.LightGray;

        // Optionally make header text bold.
        tableStyle.ConditionalStyles[ConditionalStyleType.FirstRow].Font.Bold = true;

        // Apply the style to the table.
        table.Style = tableStyle;

        // Enable the FirstRow conditional formatting.
        table.StyleOptions = TableStyleOptions.FirstRow;

        // Save the document.
        doc.Save("HeaderTableStyle.docx");
    }
}

using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2‑column table with a header row.
        builder.StartTable();

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

        builder.EndTable();

        // Create a custom table style.
        TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyHeaderStyle");

        // Set the background color for the header (first row) using a conditional style.
        tableStyle.ConditionalStyles[ConditionalStyleType.FirstRow].Shading.BackgroundPatternColor = Color.LightBlue;

        // Apply the style to the table.
        Table table = doc.FirstSection.Body.Tables[0];
        table.Style = tableStyle;

        // Enable the first‑row conditional formatting so the header color is applied.
        table.StyleOptions |= TableStyleOptions.FirstRow;

        // Save the document.
        doc.Save("HeaderRowStyle.docx");
    }
}

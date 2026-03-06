using Aspose.Words;
using Aspose.Words.Tables;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Load the WORDML template.
        Document doc = new Document("Template.docx");

        // Get the first table in the document (adjust index if needed).
        Table table = doc.FirstSection.Body.Tables[0];

        // Create a custom table style that will hold conditional formatting.
        TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyConditionalStyle");

        // ----- Conditional style for the last row -----
        // Set a background colour and make the text bold.
        ConditionalStyle lastRow = customStyle.ConditionalStyles[ConditionalStyleType.LastRow];
        lastRow.Shading.BackgroundPatternColor = Color.LightYellow;
        lastRow.Font.Bold = true;

        // ----- Conditional style for the first column -----
        // Add a dark blue right border to the first column cells.
        ConditionalStyle firstColumn = customStyle.ConditionalStyles[ConditionalStyleType.FirstColumn];
        firstColumn.Borders[BorderType.Right].Color = Color.DarkBlue;
        firstColumn.Borders[BorderType.Right].LineStyle = LineStyle.Single;

        // Apply the custom style to the table.
        table.Style = customStyle;

        // Enable the conditional styles we have defined.
        table.StyleOptions |= TableStyleOptions.LastRow | TableStyleOptions.FirstColumn;

        // Save the resulting document.
        doc.Save("Result.docx");
    }
}

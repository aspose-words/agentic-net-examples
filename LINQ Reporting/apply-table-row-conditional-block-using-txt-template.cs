using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

class TableConditionalExample
{
    static void Main()
    {
        // Load a plain‑text template. The TXT file is treated as a Word document.
        Document doc = new Document("Template.txt");

        // Use DocumentBuilder to construct a table.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ----- Build table header -----
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Product");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.EndRow();

        // ----- Build some data rows -----
        for (int i = 1; i <= 3; i++)
        {
            builder.InsertCell();
            builder.Write($"Item {i}");
            builder.InsertCell();
            builder.Write((i * 10).ToString());
            builder.EndRow();
        }
        builder.EndTable();

        // Retrieve the table we just created (first table in the document).
        Table table = doc.FirstSection.Body.Tables[0];

        // ----- Create a custom table style -----
        TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyConditionalStyle");

        // Apply conditional formatting to the first row (e.g., a light blue background).
        customStyle.ConditionalStyles[ConditionalStyleType.FirstRow].Shading.BackgroundPatternColor = Color.LightBlue;

        // Assign the style to the table.
        table.Style = customStyle;

        // Enable the FirstRow conditional style for this table.
        table.StyleOptions = TableStyleOptions.FirstRow;

        // Save the resulting document.
        doc.Save("Output.docx");
    }
}

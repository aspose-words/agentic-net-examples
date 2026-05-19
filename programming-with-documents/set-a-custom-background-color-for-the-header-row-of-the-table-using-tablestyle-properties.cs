using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        Table table = builder.StartTable();

        // ----- Header row -----
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow(); // end header row

        // ----- Data row -----
        builder.InsertCell();
        builder.Write("Data 1");
        builder.InsertCell();
        builder.Write("Data 2");
        builder.EndRow(); // end data row

        // Finish the table.
        builder.EndTable();

        // Create a custom table style.
        TableStyle headerStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyHeaderStyle");

        // Set a custom background color for the first row (header) conditional style.
        headerStyle.ConditionalStyles[ConditionalStyleType.FirstRow].Shading.BackgroundPatternColor = Color.LightBlue;

        // Apply the style to the table.
        table.Style = headerStyle;

        // Enable the FirstRow option so the conditional style is applied.
        table.StyleOptions |= TableStyleOptions.FirstRow;

        // Save the document.
        doc.Save("HeaderRowStyle.docx");
    }
}

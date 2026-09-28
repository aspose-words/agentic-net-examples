using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;   // Required for TextureIndex and Shading

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a table with a header row.
        Table table = builder.StartTable();

        // Header row cells.
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // Data row cells.
        builder.InsertCell();
        builder.Write("Data 1");
        builder.InsertCell();
        builder.Write("Data 2");
        builder.EndRow();

        builder.EndTable();

        // Create a custom table style.
        Style customStyle = doc.Styles.Add(StyleType.Table, "MyCustomTableStyle");

        // The Style object for a table can be cast to TableStyle to access shading properties.
        TableStyle tableStyle = (TableStyle)customStyle;
        tableStyle.Shading.BackgroundPatternColor = Color.LightBlue;
        tableStyle.Shading.Texture = TextureIndex.TextureNone;

        // Apply the style to the table and specify that only the first row (header) uses it.
        table.Style = customStyle;
        table.StyleOptions = TableStyleOptions.FirstRow;

        // Save the document.
        string outputPath = "TableWithHeaderStyle.docx";
        doc.Save(outputPath);
    }
}

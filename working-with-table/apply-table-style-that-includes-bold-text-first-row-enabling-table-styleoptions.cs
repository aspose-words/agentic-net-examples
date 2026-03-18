using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

class TableFirstRowBoldExample
{
    static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table.
        Table table = builder.StartTable();

        // First row (header).
        builder.InsertCell();
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.EndRow();

        // Second row (data).
        builder.InsertCell();
        builder.Writeln("Data 1");
        builder.InsertCell();
        builder.Writeln("Data 2");
        builder.EndRow();

        builder.EndTable();

        // Create a custom table style.
        TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyTableStyle");
        customStyle.Borders.Color = Color.Black;
        customStyle.Borders.LineStyle = LineStyle.Single;

        // Apply the custom style to the table.
        table.Style = customStyle;

        // Enable the FirstRow conditional formatting.
        table.StyleOptions = TableStyleOptions.FirstRow;

        // Make the text in the first row bold.
        ((TableStyle)table.Style).ConditionalStyles.FirstRow.Font.Bold = true;

        // Save the document.
        doc.Save("TableFirstRowBold.docx");
    }
}

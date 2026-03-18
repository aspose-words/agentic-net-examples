using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new document and a DocumentBuilder to construct its contents.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 3‑row table: header, data, footer.
        Table table = builder.StartTable();

        // ----- Header row -----
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // ----- Data row -----
        builder.InsertCell();
        builder.Write("Data 1");
        builder.InsertCell();
        builder.Write("Data 2");
        builder.EndRow();

        // ----- Footer row -----
        builder.InsertCell();
        builder.Write("Footer 1");
        builder.InsertCell();
        builder.Write("Footer 2");
        builder.EndRow();

        builder.EndTable();

        // Create a custom table style that will be applied to the table.
        TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyCustomTableStyle");

        // Enable conditional formatting for the first and last rows of the table.
        table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.LastRow;

        // Make the header (first row) bold.
        tableStyle.ConditionalStyles[ConditionalStyleType.FirstRow].Font.Bold = true;

        // Make the footer (last row) italic.
        tableStyle.ConditionalStyles[ConditionalStyleType.LastRow].Font.Italic = true;

        // Apply the custom style to the table.
        table.Style = tableStyle;

        // Save the document to a file.
        doc.Save("TableWithCustomStyle.docx");
    }
}

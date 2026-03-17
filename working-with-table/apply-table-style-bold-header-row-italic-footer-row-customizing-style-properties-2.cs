using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

class TableStyleExample
{
    static void Main()
    {
        // Create a new blank document and a DocumentBuilder for building its contents.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -----------------------------------------------------------------
        // Build a simple 3‑row table (header, data row, footer).
        // -----------------------------------------------------------------
        Table table = builder.StartTable();

        // Header row
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // Data row
        builder.InsertCell();
        builder.Write("Data 1");
        builder.InsertCell();
        builder.Write("Data 2");
        builder.EndRow();

        // Footer row
        builder.InsertCell();
        builder.Write("Footer 1");
        builder.InsertCell();
        builder.Write("Footer 2");
        builder.EndRow();

        builder.EndTable();

        // -----------------------------------------------------------------
        // Create a custom table style.
        // -----------------------------------------------------------------
        TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyCustomTableStyle");

        // Apply bold formatting to the first (header) row.
        customStyle.ConditionalStyles[ConditionalStyleType.FirstRow].Font.Bold = true;

        // Apply italic formatting to the last (footer) row.
        customStyle.ConditionalStyles[ConditionalStyleType.LastRow].Font.Italic = true;

        // Optional: set a base style to inherit default formatting (e.g., GridTable1Light).
        customStyle.BaseStyleName = StyleIdentifier.TableGrid1.ToString();

        // Assign the custom style to the table.
        table.Style = customStyle;

        // Enable the conditional formatting for first and last rows.
        table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.LastRow;

        // -----------------------------------------------------------------
        // Save the document.
        // -----------------------------------------------------------------
        doc.Save("TableStyleBoldHeaderItalicFooter.docx");
    }
}

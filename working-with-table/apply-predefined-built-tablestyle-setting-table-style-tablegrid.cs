using System;
using Aspose.Words;
using Aspose.Words.Tables;

class ApplyBuiltInTableStyle
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Retrieve the built‑in style named "TableGrid" and assign it to the table.
        // The style can be accessed via its StyleIdentifier.
        Style tableGridStyle = doc.Styles[StyleIdentifier.TableGrid];
        table.Style = tableGridStyle;   // Apply the built‑in style.

        // Save the document to disk.
        doc.Save("TableWithTableGridStyle.docx");
    }
}

using System;
using Aspose.Words;
using Aspose.Words.Tables;

class TableMarginExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        Table table = builder.StartTable();

        // Insert a few cells with sample text.
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndRow();

        // End the table construction.
        builder.EndTable();

        // Apply custom margins around the entire table.
        // LeftIndent adds space from the left page margin.
        // Right side margin can be simulated by reducing the table width.
        table.LeftIndent = 30; // 30 points left margin
        // Reduce the table width so that it does not touch the right page margin.
        // Here we set the preferred width to 100% of the page width minus the left indent.
        // Adjust the value as needed for the desired right margin.
        table.PreferredWidth = PreferredWidth.FromPercent(100);
        // If you are using a recent Aspose.Words version (23.5+), you can also set:
        // table.RightIndent = 30; // 30 points right margin

        // Save the document to disk.
        doc.Save("TableWithCustomMargins.docx");
    }
}

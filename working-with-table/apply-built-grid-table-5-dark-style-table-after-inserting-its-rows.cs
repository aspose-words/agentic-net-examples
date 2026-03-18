using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table and populate it with a few rows.
        Table table = builder.StartTable();

        // Row 1 – header cells.
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // Row 2 – data cells.
        builder.InsertCell();
        builder.Write("Data 1");
        builder.InsertCell();
        builder.Write("Data 2");
        builder.EndRow();

        // Row 3 – more data.
        builder.InsertCell();
        builder.Write("Data 3");
        builder.InsertCell();
        builder.Write("Data 4");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Apply the built‑in "Grid Table 5 Dark" style to the table.
        table.StyleIdentifier = StyleIdentifier.GridTable5Dark;

        // Enable the style's conditional formatting (first/last rows & columns, banding, etc.).
        table.StyleOptions = TableStyleOptions.FirstRow |
                             TableStyleOptions.LastRow |
                             TableStyleOptions.FirstColumn |
                             TableStyleOptions.LastColumn |
                             TableStyleOptions.RowBands |
                             TableStyleOptions.ColumnBands;

        // Resize the table so that column widths fit the cell contents.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document to disk.
        doc.Save("GridTable5Dark.docx");
    }
}

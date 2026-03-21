using System;
using Aspose.Words;
using Aspose.Words.Tables;

class TableBreakAcrossPagesExample
{
    static void Main()
    {
        // Create a new document and add a simple table.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a 2x2 table.
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Writeln("Cell 3");
        builder.InsertCell();
        builder.Writeln("Cell 4");
        builder.EndRow();
        builder.EndTable();

        // Get the first table in the document.
        Table table = doc.FirstSection.Body.Tables[0];

        // Prevent each row of the table from breaking across pages.
        foreach (Row row in table.Rows)
        {
            row.RowFormat.AllowBreakAcrossPages = false;
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}

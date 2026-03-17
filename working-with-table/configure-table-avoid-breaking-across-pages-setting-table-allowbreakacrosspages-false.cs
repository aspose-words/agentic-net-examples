using System;
using Aspose.Words;
using Aspose.Words.Tables;

class TableBreakAcrossPagesExample
{
    static void Main()
    {
        // Load an existing document that contains a table.
        Document doc = new Document("Input.docx");

        // Get the first table in the document (adjust as needed).
        Table table = doc.FirstSection.Body.Tables[0];

        // Prevent each row of the table from breaking across pages.
        // The Table class itself does not expose AllowBreakAcrossPages,
        // so we set the property on each Row's RowFormat.
        foreach (Row row in table.Rows)
        {
            row.RowFormat.AllowBreakAcrossPages = false;
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}

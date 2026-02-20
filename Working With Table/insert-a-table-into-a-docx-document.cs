using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder which will be used to construct the document content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table. The builder's cursor is now positioned inside the table.
        Table table = builder.StartTable();

        // ---- First row (header) ----
        builder.InsertCell();               // First cell of the first row.
        builder.Write("Header 1");          // Write text into the cell.
        builder.InsertCell();               // Second cell of the first row.
        builder.Write("Header 2");          // Write text into the cell.
        builder.EndRow();                   // Complete the first row.

        // ---- Second row (data) ----
        builder.InsertCell();               // First cell of the second row.
        builder.Write("Cell 1");            // Write text into the cell.
        builder.InsertCell();               // Second cell of the second row.
        builder.Write("Cell 2");            // Write text into the cell.
        builder.EndRow();                   // Complete the second row.

        // End the table construction.
        builder.EndTable();

        // Adjust the table to fit its contents.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document to a DOCX file.
        doc.Save("TableExample.docx");
    }
}

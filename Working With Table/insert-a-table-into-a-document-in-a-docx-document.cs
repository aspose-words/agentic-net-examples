using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document for building content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table. The builder returns the Table object for further configuration if needed.
        Table table = builder.StartTable();

        // ---- First row ----
        builder.InsertCell();               // First cell of the first row.
        builder.Write("Header 1");          // Write text into the first cell.
        builder.InsertCell();               // Second cell of the first row.
        builder.Write("Header 2");          // Write text into the second cell.
        builder.EndRow();                   // End the first row.

        // ---- Second row ----
        builder.InsertCell();               // First cell of the second row.
        builder.Write("Cell 1");            // Write text into the first cell.
        builder.InsertCell();               // Second cell of the second row.
        builder.Write("Cell 2");            // Write text into the second cell.
        builder.EndRow();                   // End the second row.

        // Finish the table.
        builder.EndTable();

        // Save the document as a DOCX file.
        doc.Save("TableInserted.docx");
    }
}

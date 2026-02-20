using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

class AddRowToPdf
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder which will be used to construct the table.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        Table table = builder.StartTable();

        // ----- First row (header) -----
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // ----- Second row (the row we are adding) -----
        builder.InsertCell();
        builder.Write("Value 1");
        builder.InsertCell();
        builder.Write("Value 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document as a PDF file.
        doc.Save("AddedRow.pdf", SaveFormat.Pdf);
    }
}

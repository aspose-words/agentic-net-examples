using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table and add two rows with two cells each.
        builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Row 1, cell 1");
        builder.InsertCell();
        builder.Write("Row 1, cell 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Row 2, cell 1");
        builder.InsertCell();
        builder.Write("Row 2, cell 2");
        builder.EndTable();

        // Configure TXT save options to preserve the table layout.
        TxtSaveOptions txtOptions = new TxtSaveOptions
        {
            PreserveTableLayout = true
        };

        // Save the document as a plain‑text file with the table layout preserved.
        doc.Save("TableOutput.txt", txtOptions);
    }
}

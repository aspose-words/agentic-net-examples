using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

class EnableTableOverlap
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to insert a floating table.
        DocumentBuilder builder = new DocumentBuilder(doc);
        Table table = builder.StartTable();

        // Add a simple cell so the table is valid.
        builder.InsertCell();
        builder.Write("Floating table cell");
        builder.EndRow();
        builder.EndTable();

        // Make the table a floating object by enabling text wrapping.
        table.TextWrapping = TextWrapping.Around;

        // The AllowOverlap property is read‑only and defaults to true,
        // which means the table already permits other floating objects to overlap it.
        // Verify the default value.
        if (!table.AllowOverlap)
        {
            // In practice this branch will never be hit because AllowOverlap is always true.
            Console.WriteLine("Table does not allow overlap – this should not happen.");
        }

        // Save the document to disk.
        doc.Save("TableWithOverlap.docx");
    }
}

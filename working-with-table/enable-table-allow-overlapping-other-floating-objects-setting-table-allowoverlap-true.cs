using System;
using Aspose.Words;
using Aspose.Words.Tables;

class EnableTableOverlap
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a simple table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.EndRow();
        builder.EndTable();

        // Get the first table in the document.
        Table table = doc.FirstSection?.Body?.Tables?[0];
        if (table == null)
        {
            Console.WriteLine("No table found in the document.");
            return;
        }

        // Ensure the table is a floating table (required for overlapping behavior).
        table.TextWrapping = TextWrapping.Around;

        // Overlapping of other floating objects is enabled by default.
        // The AllowOverlap property is read‑only in this version of Aspose.Words,
        // so we only read its value for debugging purposes.
        Console.WriteLine($"Table AllowOverlap (read‑only): {table.AllowOverlap}");

        // Save the modified document.
        doc.Save("Output.docx");
        Console.WriteLine("Document saved as Output.docx");
    }
}

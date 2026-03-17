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

        // Build a simple 2‑cell table.
        Table table = builder.StartTable();

        builder.InsertCell();
        builder.Write("Name");               // Left‑to‑right text.
        builder.InsertCell();
        builder.Write("مرحبا");               // Right‑to‑left Arabic text.

        builder.EndRow();
        builder.EndTable();

        // Enable right‑to‑left layout for the entire table.
        table.Bidi = true;

        // Save the document to a file.
        doc.Save("TableBidi.docx");
    }
}

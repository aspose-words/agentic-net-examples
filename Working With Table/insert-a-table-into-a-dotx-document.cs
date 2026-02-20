using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load an existing DOTX template.
        Document doc = new Document("Template.dotx");

        // Attach a DocumentBuilder to the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Position the builder where the table should be inserted.
        // Here we move to the end of the document, but any position can be used.
        builder.MoveToDocumentEnd();

        // Begin constructing a new table.
        Table table = builder.StartTable();

        // ---- First row (header) ----
        builder.InsertCell();                     // First cell of the header row.
        builder.Writeln("Header 1");              // Write header text.
        builder.InsertCell();                     // Second cell of the header row.
        builder.Writeln("Header 2");
        builder.EndRow();                         // Complete the header row.

        // ---- Second row (data) ----
        builder.InsertCell();                     // First cell of the data row.
        builder.Writeln("Value 1");
        builder.InsertCell();                     // Second cell of the data row.
        builder.Writeln("Value 2");
        builder.EndRow();                         // Complete the data row.

        // Finish the table.
        builder.EndTable();

        // Optional: apply a built‑in style and auto‑fit the table to its contents.
        table.StyleIdentifier = StyleIdentifier.LightShadingAccent1;
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the modified document. The output format can be .docx, .dotx, etc.
        doc.Save("Result.docx");
    }
}

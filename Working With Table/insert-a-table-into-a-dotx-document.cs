using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load the DOTX template file.
        Document doc = new Document("Template.dotx");

        // Create a DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Position the builder where the table should be inserted.
        // Here we move to the end of the main story (the document body).
        builder.MoveToDocumentEnd();

        // Start a new table. The method returns the created Table node.
        Table table = builder.StartTable();

        // ---- First row ----
        builder.InsertCell();                     // First cell of the first row.
        builder.Write("Cell 1,1");                // Add text to the cell.

        builder.InsertCell();                     // Second cell of the first row.
        builder.Write("Cell 1,2");
        builder.EndRow();                         // Finish the first row.

        // ---- Second row ----
        builder.InsertCell();                     // First cell of the second row.
        builder.Write("Cell 2,1");

        builder.InsertCell();                     // Second cell of the second row.
        builder.Write("Cell 2,2");
        builder.EndTable();                       // Close the table.

        // Optional: apply a built‑in style and auto‑fit the table to its contents.
        table.StyleIdentifier = StyleIdentifier.LightShadingAccent1;
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the modified document. The output format can be any supported type (e.g., DOCX).
        doc.Save("Result.docx");
    }
}

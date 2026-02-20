using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InsertTableIntoWordMl
{
    static void Main()
    {
        // Load an existing WORDML (WordprocessingML) document.
        // Replace the path with the actual location of your .xml file.
        Document doc = new Document("InputDocument.xml");

        // Create a DocumentBuilder attached to the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the end of the document (or any desired location).
        builder.MoveToDocumentEnd();

        // Start a new table.
        Table table = builder.StartTable();

        // ---- First row (header) ----
        builder.InsertCell();                     // First cell of the header row.
        builder.Write("Header 1");                // Write header text.
        builder.InsertCell();                     // Second cell of the header row.
        builder.Write("Header 2");
        builder.EndRow();                         // End the header row.

        // ---- Second row (data) ----
        builder.InsertCell();                     // First cell of the data row.
        builder.Write("Data 1");
        builder.InsertCell();                     // Second cell of the data row.
        builder.Write("Data 2");
        builder.EndRow();                         // End the data row.

        // Finish the table.
        builder.EndTable();

        // Optional: Apply a simple style and auto‑fit the table to its contents.
        table.StyleIdentifier = StyleIdentifier.LightShadingAccent1;
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the modified document back to WORDML format.
        // The output file will be a .xml WordprocessingML document.
        doc.Save("OutputDocument.xml");
    }
}

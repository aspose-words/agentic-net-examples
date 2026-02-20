using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

class InsertTableIntoRtf
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        Table table = builder.StartTable();

        // Insert first row (header).
        builder.InsertCell();
        builder.Writeln("Item");
        builder.InsertCell();
        builder.Writeln("Quantity");
        builder.EndRow();

        // Insert second row.
        builder.InsertCell();
        builder.Writeln("Apples");
        builder.InsertCell();
        builder.Writeln("20");
        builder.EndRow();

        // Insert third row.
        builder.InsertCell();
        builder.Writeln("Bananas");
        builder.InsertCell();
        builder.Writeln("40");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Optional: Auto‑fit the table to its contents.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document as RTF using RtfSaveOptions.
        RtfSaveOptions saveOptions = new RtfSaveOptions
        {
            // Example: enable pretty formatting for readability (optional).
            PrettyFormat = true
        };
        doc.Save("TableDocument.rtf", saveOptions);
    }
}

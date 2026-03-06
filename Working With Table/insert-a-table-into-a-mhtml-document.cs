using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();               // Cell (1,1)
        builder.Write("Cell 1,1");
        builder.InsertCell();               // Cell (1,2)
        builder.Write("Cell 1,2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();               // Cell (2,1)
        builder.Write("Cell 2,1");
        builder.InsertCell();               // Cell (2,2)
        builder.Write("Cell 2,2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Configure save options for MHTML output.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            // Export table widths as they are (default behavior).
            TableWidthOutputMode = HtmlElementSizeOutputMode.All
        };

        // Save the document as an MHTML file.
        doc.Save("TableInMhtml.mht", saveOptions);
    }
}

using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load an existing DOTM template.
        Document doc = new Document("Template.dotm");

        // Create a DocumentBuilder attached to the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        Table table = builder.StartTable();

        // ---- First row (header) ----
        builder.InsertCell();               // First cell
        builder.Write("Header 1");          // Write content
        builder.InsertCell();               // Second cell
        builder.Write("Header 2");
        builder.EndRow();                   // End the header row

        // ---- Second row (data) ----
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();                   // End the data row

        // Finish the table.
        builder.EndTable();

        // Adjust the table to fit its contents.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the modified document back to DOTM format.
        doc.Save("Result.dotm");
    }
}

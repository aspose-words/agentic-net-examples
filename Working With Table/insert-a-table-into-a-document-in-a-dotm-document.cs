using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the existing DOTM template.
        Document doc = new Document("Template.dotm");

        // Create a DocumentBuilder for the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Position the builder where the table should be inserted.
        builder.MoveToDocumentEnd();

        // Start a new table.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Cell 1,1");
        builder.InsertCell();
        builder.Write("Cell 1,2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Cell 2,1");
        builder.InsertCell();
        builder.Write("Cell 2,2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Optional: set title and description for accessibility.
        table.Title = "Sample Table";
        table.Description = "Demonstrates inserting a table into a DOTM document.";

        // Save the modified document.
        doc.Save("Output.dotm");
    }
}

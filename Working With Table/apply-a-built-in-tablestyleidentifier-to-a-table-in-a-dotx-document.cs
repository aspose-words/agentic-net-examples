using System;
using Aspose.Words;
using Aspose.Words.Tables;

class ApplyTableStyle
{
    static void Main()
    {
        // Load an existing DOTX template.
        Document doc = new Document("Template.dotx");

        // Create a DocumentBuilder to edit the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        Table table = builder.StartTable();

        // Insert a single cell with some text.
        builder.InsertCell();
        builder.Writeln("Sample cell");

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Apply a built‑in table style using its identifier.
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Optionally, specify which parts of the style to apply.
        table.StyleOptions = TableStyleOptions.FirstRow |
                             TableStyleOptions.FirstColumn |
                             TableStyleOptions.RowBands;

        // Save the modified document.
        doc.Save("StyledTable.docx");
    }
}

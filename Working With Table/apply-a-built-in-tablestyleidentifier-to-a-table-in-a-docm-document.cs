using System;
using Aspose.Words;
using Aspose.Words.Tables;

class ApplyTableStyle
{
    static void Main()
    {
        // Load an existing DOCM document.
        Document doc = new Document(@"C:\Docs\InputDocument.docm");

        // Create a DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        Table table = builder.StartTable();

        // Insert a single cell to satisfy the requirement of having at least one row before styling.
        builder.InsertCell();
        builder.Writeln("Sample cell");

        // End the first row and the table.
        builder.EndRow();
        builder.EndTable();

        // Apply a built‑in table style using its locale‑independent identifier.
        // For example, use the "TableGrid" style.
        table.StyleIdentifier = StyleIdentifier.TableGrid;

        // Optionally, specify which parts of the style should be applied.
        table.StyleOptions = TableStyleOptions.FirstRow |
                             TableStyleOptions.FirstColumn |
                             TableStyleOptions.RowBands;

        // Save the modified document as a DOCM file.
        doc.Save(@"C:\Docs\OutputDocument.docm");
    }
}

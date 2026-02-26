using System;
using Aspose.Words;
using Aspose.Words.Tables;

class ApplyTableStyle
{
    static void Main()
    {
        // Load the DOT template.
        Document doc = new Document("Template.dot");

        // Find the first table in the document.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table != null)
        {
            // Apply a built‑in table style using its identifier.
            table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

            // Optionally, specify which parts of the style are applied.
            table.StyleOptions = TableStyleOptions.FirstColumn |
                                 TableStyleOptions.RowBands |
                                 TableStyleOptions.FirstRow;

            // Adjust column widths to fit the content.
            table.AutoFit(AutoFitBehavior.AutoFitToContents);
        }

        // Save the modified document.
        doc.Save("StyledDocument.docx");
    }
}

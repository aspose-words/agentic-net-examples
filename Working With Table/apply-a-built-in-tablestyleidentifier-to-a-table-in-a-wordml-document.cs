using System;
using Aspose.Words;
using Aspose.Words.Tables;

class ApplyTableStyle
{
    static void Main()
    {
        // Load an existing Word document (WORDML or DOCX) that contains at least one table.
        Document doc = new Document("input.docx");

        // Retrieve the first table in the document.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
        {
            Console.WriteLine("No table found in the document.");
            return;
        }

        // Apply a built‑in table style using its style identifier.
        // Here we use the built‑in style "Medium Shading 1 Accent 1".
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Optionally specify which parts of the style should be applied.
        table.StyleOptions = TableStyleOptions.FirstRow |
                             TableStyleOptions.FirstColumn |
                             TableStyleOptions.RowBands;

        // Adjust the table width to fit its contents after applying the style.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the modified document.
        doc.Save("output.docx");
    }
}

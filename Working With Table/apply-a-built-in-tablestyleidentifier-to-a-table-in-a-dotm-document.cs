using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace ApplyTableStyleToDotm
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing DOTM template.
            // If the template contains macros, you may need to provide LoadOptions with LoadFormat.Dotm.
            Document doc = new Document("Template.dotm");

            // Find the first table in the document (adjust the index if needed).
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            // Apply a built‑in table style by its identifier.
            // Example: LightGrid style – you can replace with any other StyleIdentifier value.
            table.StyleIdentifier = StyleIdentifier.LightGrid;

            // Optionally specify which parts of the style should be applied.
            table.StyleOptions = TableStyleOptions.FirstRow |
                                 TableStyleOptions.FirstColumn |
                                 TableStyleOptions.RowBands;

            // Save the modified document.
            doc.Save("StyledDocument.docx");
        }
    }
}

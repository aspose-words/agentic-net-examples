using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableStyleToDirectFormattingDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple table and apply a built‑in table style.
            builder.StartTable();
            builder.InsertCell();
            builder.Writeln("Cell 1");
            builder.InsertCell();
            builder.Writeln("Cell 2");
            builder.EndRow();
            builder.EndTable();

            // Apply a built‑in style to the table (e.g., "Table Grid").
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
            table.Style = doc.Styles["Table Grid"];

            // Convert all formatting defined in the table styles into direct formatting.
            doc.ExpandTableStylesToDirectFormatting();

            // Save the document. The tables now have their formatting applied directly.
            doc.Save("TablesDirectFormatting.docx");
        }
    }
}

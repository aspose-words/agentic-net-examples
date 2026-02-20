using System;
using Aspose.Words;
using Aspose.Words.Tables;

class JoinTablesInPdf
{
    static void Main()
    {
        // Load the PDF document (Aspose.Words can load PDF files).
        Document doc = new Document("input.pdf");

        // Get the collection of tables in the first section.
        TableCollection tables = doc.FirstSection.Body.Tables;

        // Ensure there are at least two tables to join.
        if (tables.Count < 2)
        {
            Console.WriteLine("The document must contain at least two tables.");
            return;
        }

        // Retrieve the two tables we want to merge.
        Table table1 = tables[0];
        Table table2 = tables[1];

        // Create a new empty table that will hold the merged rows.
        Table mergedTable = new Table(doc);

        // Copy the formatting from the first table (optional, adjust as needed).
        mergedTable.Style = table1.Style;
        mergedTable.StyleIdentifier = table1.StyleIdentifier;
        mergedTable.StyleOptions = table1.StyleOptions;
        mergedTable.Alignment = table1.Alignment;
        mergedTable.PreferredWidth = table1.PreferredWidth;
        mergedTable.AllowAutoFit = table1.AllowAutoFit;
        mergedTable.CellSpacing = table1.CellSpacing;
        mergedTable.TextWrapping = table1.TextWrapping;

        // Helper method to clone rows from a source table into the merged table.
        void AppendRows(Table source)
        {
            foreach (Row srcRow in source.Rows)
            {
                // Clone the row (deep clone) and add it to the merged table.
                Row clonedRow = (Row)srcRow.Clone(true);
                mergedTable.Rows.Add(clonedRow);
            }
        }

        // Append rows from both tables.
        AppendRows(table1);
        AppendRows(table2);

        // Insert the merged table into the document at the position of the first table.
        // ParentNode is a Node, but InsertBefore is defined on CompositeNode, so cast it.
        CompositeNode parent = (CompositeNode)table1.ParentNode;
        parent.InsertBefore(mergedTable, table1);

        // Remove the original tables from the document.
        table1.Remove();
        table2.Remove();

        // Save the modified document as PDF.
        doc.Save("output.pdf");
    }
}

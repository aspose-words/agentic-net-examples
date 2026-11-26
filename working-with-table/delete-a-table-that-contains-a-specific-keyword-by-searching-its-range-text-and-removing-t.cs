using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build the first table (does NOT contain the keyword).
        builder.StartTable();
        builder.InsertCell();
        builder.Write("First table cell 1");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("First table cell 2");
        builder.EndRow();
        builder.EndTable();

        // Add a paragraph separator.
        builder.Writeln();

        // Build the second table (contains the keyword to trigger deletion).
        builder.StartTable();
        builder.InsertCell();
        builder.Write("This cell will be deleted: DELETE_ME");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Second table cell 2");
        builder.EndRow();
        builder.EndTable();

        // Save the original document (optional, for verification).
        doc.Save("Original.docx");

        // Keyword to search for inside tables.
        string keyword = "DELETE_ME";

        // Retrieve all tables in the document.
        NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);

        // Iterate backwards so removal does not affect the loop index.
        for (int i = tables.Count - 1; i >= 0; i--)
        {
            Table table = (Table)tables[i];
            if (table.Range.Text.Contains(keyword))
            {
                // Remove the table node from its parent.
                table.ParentNode.RemoveChild(table);
            }
        }

        // Save the modified document where the matching table has been removed.
        doc.Save("Modified.docx");
    }
}

using System;
using System.IO;
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
        builder.Write("First table, cell 1");
        builder.InsertCell();
        builder.Write("First table, cell 2");
        builder.EndRow();
        builder.EndTable();

        // Build the second table (contains the keyword "DeleteMe").
        builder.StartTable();
        builder.InsertCell();
        builder.Write("This table contains DeleteMe keyword");
        builder.InsertCell();
        builder.Write("Second table, cell 2");
        builder.EndRow();
        builder.EndTable();

        // Keyword to search for.
        const string keyword = "DeleteMe";

        // Find all tables in the document.
        NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);

        // Iterate backwards so that removal does not affect the collection indexing.
        for (int i = tables.Count - 1; i >= 0; i--)
        {
            Table table = (Table)tables[i];

            // If the table's text contains the keyword, remove the whole table node.
            if (table.Range.Text.Contains(keyword))
            {
                table.Remove();
            }
        }

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DeletedTable.docx");
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}

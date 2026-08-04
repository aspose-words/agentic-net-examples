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

        // Build the first table that contains the keyword "DeleteMe".
        Table tableWithKeyword = builder.StartTable();
        builder.InsertCell();
        builder.Write("This cell will be deleted because it contains the keyword DeleteMe.");
        builder.InsertCell();
        builder.Write("Another cell.");
        builder.EndRow();
        builder.EndTable();

        // Build a second table that does NOT contain the keyword.
        Table tableWithoutKeyword = builder.StartTable();
        builder.InsertCell();
        builder.Write("This table stays.");
        builder.InsertCell();
        builder.Write("More content.");
        builder.EndRow();
        builder.EndTable();

        // Search for tables whose full text contains the specific keyword and remove them.
        NodeCollection allTables = doc.GetChildNodes(NodeType.Table, true);
        for (int i = allTables.Count - 1; i >= 0; i--)
        {
            Table tbl = (Table)allTables[i];
            if (tbl.Range.Text.Contains("DeleteMe"))
            {
                tbl.Remove();
            }
        }

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DeletedTable.docx");
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("The output document was not saved correctly.");
        }
    }
}

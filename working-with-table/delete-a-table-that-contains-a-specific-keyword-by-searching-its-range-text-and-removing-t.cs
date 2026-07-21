using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class DeleteTableByKeyword
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build first table (does NOT contain the keyword).
        builder.StartTable();
        builder.InsertCell();
        builder.Write("First table, cell 1.");
        builder.InsertCell();
        builder.Write("First table, cell 2.");
        builder.EndRow();
        builder.EndTable();

        // Build second table (contains the keyword "DeleteMe").
        builder.StartTable();
        builder.InsertCell();
        builder.Write("This table will be deleted. Keyword: DeleteMe");
        builder.InsertCell();
        builder.Write("Another cell.");
        builder.EndRow();
        builder.EndTable();

        // Save the original document (optional, for inspection).
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Original.docx");
        doc.Save(outputPath);

        // Keyword to search for.
        const string keyword = "DeleteMe";

        // Find and delete tables that contain the keyword.
        NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
        // Iterate backwards because removing nodes changes the collection indexing.
        for (int i = tables.Count - 1; i >= 0; i--)
        {
            Table tbl = (Table)tables[i];
            if (tbl.Range.Text != null && tbl.Range.Text.Contains(keyword))
            {
                // Remove the entire table node from the document.
                tbl.Remove();
            }
        }

        // Save the modified document.
        string resultPath = Path.Combine(Directory.GetCurrentDirectory(), "DeletedTable.docx");
        doc.Save(resultPath);

        // Simple verification that the file was created.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}

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
        builder.EndRow();
        builder.InsertCell();
        builder.Write("First table, cell 2");
        builder.EndRow();
        builder.EndTable();

        // Add a paragraph between tables for clarity.
        builder.Writeln();

        // Build the second table (contains the keyword "DeleteMe").
        builder.StartTable();
        builder.InsertCell();
        builder.Write("This table will be deleted because it contains DeleteMe");
        builder.EndRow();
        builder.EndTable();

        // Save the original document (optional, just for reference).
        string originalPath = "Original.docx";
        doc.Save(originalPath);

        // Keyword to search for.
        const string keyword = "DeleteMe";

        // Find tables whose full text contains the keyword and remove the first match.
        NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
        foreach (Table table in tables)
        {
            if (table.Range.Text.Contains(keyword))
            {
                // Remove the table node from the document.
                table.Remove();
                break; // Assuming only one table needs to be removed.
            }
        }

        // Save the modified document.
        string outputPath = "Result.docx";
        doc.Save(outputPath);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }
    }
}

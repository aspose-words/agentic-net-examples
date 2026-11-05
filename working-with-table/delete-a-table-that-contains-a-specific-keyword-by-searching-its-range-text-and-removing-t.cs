using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build first table (does not contain the keyword).
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1A");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Cell 1B");
        builder.EndRow();
        builder.EndTable();

        // Add a paragraph separator.
        builder.Writeln();

        // Build second table (contains the keyword "DeleteMe").
        builder.StartTable();
        builder.InsertCell();
        builder.Write("DeleteMe");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Cell 2B");
        builder.EndRow();
        builder.EndTable();

        // Save the original document (optional, just to have an input file).
        string originalPath = "Original.docx";
        doc.Save(originalPath);

        // Load the document (simulating processing an existing file).
        Document loadedDoc = new Document(originalPath);

        // Search for tables whose full text contains the keyword and remove them.
        string keyword = "DeleteMe";
        NodeCollection tables = loadedDoc.GetChildNodes(NodeType.Table, true);
        foreach (Table table in tables)
        {
            if (table.Range.Text != null && table.Range.Text.Contains(keyword))
            {
                table.Remove();
                // Since we modified the collection while iterating, break after removal.
                break;
            }
        }

        // Save the modified document.
        string resultPath = "Result.docx";
        loadedDoc.Save(resultPath);

        // Verify that the result file was created.
        if (!File.Exists(resultPath))
        {
            throw new Exception("The result document was not saved correctly.");
        }

        // Optionally, indicate success (no interactive input required).
        Console.WriteLine("Table containing the keyword was removed and document saved as " + resultPath);
    }
}

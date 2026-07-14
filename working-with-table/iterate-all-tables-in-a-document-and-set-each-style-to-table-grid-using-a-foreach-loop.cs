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

        // Build the first sample table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("First table, Cell 1");
        builder.InsertCell();
        builder.Write("First table, Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Add a paragraph separator.
        builder.Writeln();

        // Build the second sample table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Second table, Cell 1");
        builder.InsertCell();
        builder.Write("Second table, Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Iterate over all tables in the document and set their style to "Table Grid".
        NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
        foreach (Table table in tables)
        {
            table.StyleIdentifier = StyleIdentifier.TableGrid;
        }

        // Save the modified document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The document was not saved correctly.");
    }
}

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

        // Build first sample table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Table 1 - Cell 1");
        builder.InsertCell();
        builder.Write("Table 1 - Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Build second sample table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Table 2 - Cell 1");
        builder.InsertCell();
        builder.Write("Table 2 - Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Iterate over all tables in the document and apply the "Table Grid" style.
        NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
        foreach (Table table in tables)
        {
            // Apply built‑in style "Table Grid".
            table.StyleIdentifier = StyleIdentifier.TableGrid;
        }

        // Save the modified document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.docx");
        doc.Save(outputPath);
    }
}

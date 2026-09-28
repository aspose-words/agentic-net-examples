using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some introductory text.
        builder.Writeln("Document with sample tables.");

        // Build a simple 2x2 table.
        builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Writeln("Cell 3");
        builder.InsertCell();
        builder.Writeln("Cell 4");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the initial document (optional, demonstrates creation).
        string initialPath = "Sample.docx";
        doc.Save(initialPath);

        // Iterate all tables in the document and set their style to "Table Grid".
        foreach (Table table in doc.GetChildNodes(NodeType.Table, true).OfType<Table>())
        {
            // Apply the built‑in Table Grid style.
            table.StyleIdentifier = StyleIdentifier.TableGrid;
        }

        // Save the modified document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }

        // Inform that the process completed successfully.
        Console.WriteLine("All tables styled as 'Table Grid' and saved to " + outputPath);
    }
}

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

        // Insert the first table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("R1C1");
        builder.InsertCell();
        builder.Write("R1C2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("R2C1");
        builder.InsertCell();
        builder.Write("R2C2");
        builder.EndRow();
        builder.EndTable();

        // Insert a paragraph break between tables.
        builder.Writeln();

        // Insert the second table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("A1");
        builder.InsertCell();
        builder.Write("A2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("B1");
        builder.InsertCell();
        builder.Write("B2");
        builder.EndRow();
        builder.EndTable();

        // Save the original document.
        string originalPath = "Original.docx";
        doc.Save(originalPath);

        // Iterate all tables and set each style to "Table Grid".
        foreach (Table table in doc.GetChildNodes(NodeType.Table, true).OfType<Table>())
        {
            table.StyleIdentifier = StyleIdentifier.TableGrid;
        }

        // Save the modified document.
        string modifiedPath = "Modified.docx";
        doc.Save(modifiedPath);

        // Verify that the modified file was saved.
        if (!File.Exists(modifiedPath))
        {
            throw new InvalidOperationException("The modified document was not saved correctly.");
        }
    }
}

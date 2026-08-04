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
        builder.Write("A1");
        builder.InsertCell();
        builder.Write("B1");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("A2");
        builder.InsertCell();
        builder.Write("B2");
        builder.EndTable();

        // Build the second sample table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("C1");
        builder.InsertCell();
        builder.Write("D1");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("C2");
        builder.InsertCell();
        builder.Write("D2");
        builder.EndTable();

        // Iterate all tables in the document and set each style to "Table Grid".
        foreach (Table table in doc.GetChildNodes(NodeType.Table, true))
        {
            table.StyleIdentifier = StyleIdentifier.TableGrid;
        }

        // Save the resulting document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "TablesWithGridStyle.docx");
        doc.Save(outputPath);
    }
}

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add first table.
        builder.Writeln("First Table:");
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
        builder.EndRow();
        builder.EndTable();

        // Add a paragraph between tables.
        builder.Writeln();
        builder.Writeln("Second Table:");

        // Add second table.
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
        builder.EndRow();
        builder.EndTable();

        // Save the sample document.
        string samplePath = "Sample.docx";
        doc.Save(samplePath);

        // Load the document for processing.
        Document loadedDoc = new Document(samplePath);

        // Iterate all tables and set each style to "Table Grid".
        foreach (Table table in loadedDoc.GetChildNodes(NodeType.Table, true))
        {
            table.StyleIdentifier = StyleIdentifier.TableGrid;
        }

        // Save the modified document.
        string resultPath = "Result.docx";
        loadedDoc.Save(resultPath);

        // Verify that the output file was created.
        if (!File.Exists(resultPath))
        {
            throw new InvalidOperationException($"Failed to create output file: {resultPath}");
        }
    }
}

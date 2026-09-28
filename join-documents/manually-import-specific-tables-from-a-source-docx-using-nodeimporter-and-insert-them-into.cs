using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Paths for sample documents
        string destPath = "Destination.docx";
        string srcPath = "Source.docx";
        string resultPath = "Result.docx";

        // Create destination document with a simple paragraph
        var destDoc = new Document();
        var destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("Destination Document");
        destDoc.Save(destPath);

        // Create source document containing two tables
        var srcDoc = new Document();
        var srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.Writeln("Source Document with Tables");

        // First table
        srcBuilder.StartTable();
        srcBuilder.InsertCell();
        srcBuilder.Write("A1");
        srcBuilder.InsertCell();
        srcBuilder.Write("B1");
        srcBuilder.EndRow();
        srcBuilder.EndTable();

        // Second table
        srcBuilder.StartTable();
        srcBuilder.InsertCell();
        srcBuilder.Write("C1");
        srcBuilder.InsertCell();
        srcBuilder.Write("D1");
        srcBuilder.EndRow();
        srcBuilder.EndTable();

        srcDoc.Save(srcPath);

        // Load the documents for processing
        var destination = new Document(destPath);
        var source = new Document(srcPath);

        // Prepare a NodeImporter to import nodes from source to destination
        var importer = new NodeImporter(source, destination, ImportFormatMode.KeepSourceFormatting);

        // Import each table from the source document into the destination document
        NodeCollection sourceTables = source.GetChildNodes(NodeType.Table, true);
        foreach (Table table in sourceTables)
        {
            Node importedTable = importer.ImportNode(table, true);
            // Append the imported table to the end of the destination body
            destination.FirstSection.Body.AppendChild(importedTable);
        }

        // Save the merged result
        destination.Save(resultPath);

        // Validation: ensure the result file exists
        if (!File.Exists(resultPath))
        {
            throw new InvalidOperationException($"The merged document was not saved to '{resultPath}'.");
        }

        // Validation: ensure the result contains the expected number of tables
        var resultDoc = new Document(resultPath);
        int expectedTableCount = sourceTables.Count;
        int actualTableCount = resultDoc.GetChildNodes(NodeType.Table, true).Count;
        if (actualTableCount != expectedTableCount)
        {
            throw new InvalidOperationException($"Table count mismatch. Expected: {expectedTableCount}, Actual: {actualTableCount}.");
        }

        // Program completed successfully
        Console.WriteLine("Tables imported and merged document created successfully.");
    }
}

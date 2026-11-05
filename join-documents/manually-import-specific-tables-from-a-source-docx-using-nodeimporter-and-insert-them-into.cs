using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Define file paths.
        string sourcePath = "Source.docx";
        string destinationPath = "Destination.docx";
        string outputPath = "Merged.docx";

        // Create a source document with a table.
        var sourceDoc = new Document();
        var sourceBuilder = new DocumentBuilder(sourceDoc);
        sourceBuilder.Writeln("Source Document");
        sourceBuilder.StartTable();
        sourceBuilder.InsertCell();
        sourceBuilder.Write("Header 1");
        sourceBuilder.InsertCell();
        sourceBuilder.Write("Header 2");
        sourceBuilder.EndRow();
        sourceBuilder.InsertCell();
        sourceBuilder.Write("Row1Col1");
        sourceBuilder.InsertCell();
        sourceBuilder.Write("Row1Col2");
        sourceBuilder.EndRow();
        sourceBuilder.EndTable();
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // Create a destination document with some initial content.
        var destDoc = new Document();
        var destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("Destination Document");
        destBuilder.Writeln("This is the original content of the destination document.");
        destDoc.Save(destinationPath, SaveFormat.Docx);

        // Load the documents.
        var source = new Document(sourcePath);
        var destination = new Document(destinationPath);

        // Prepare a NodeImporter to import nodes from source to destination.
        var importer = new NodeImporter(source, destination, ImportFormatMode.KeepSourceFormatting);

        // Import each table from the source document and append it to the destination document.
        var tables = source.GetChildNodes(NodeType.Table, true);
        foreach (Table table in tables)
        {
            Node importedTable = importer.ImportNode(table, true);
            // Append the imported table to the end of the destination document.
            destination.LastSection.Body.AppendChild(importedTable);
        }

        // Save the merged document.
        destination.Save(outputPath, SaveFormat.Docx);

        // Validate that the output file exists.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"The merged document was not created at path: {outputPath}");
        }

        // Validate that the merged document contains the expected number of tables.
        int expectedTableCount = destination.GetChildNodes(NodeType.Table, true).Count;
        if (expectedTableCount < 1)
        {
            throw new InvalidOperationException("The merged document does not contain any tables as expected.");
        }

        // Optional: Write a simple confirmation to the console.
        Console.WriteLine($"Merged document created successfully with {expectedTableCount} table(s).");
    }
}

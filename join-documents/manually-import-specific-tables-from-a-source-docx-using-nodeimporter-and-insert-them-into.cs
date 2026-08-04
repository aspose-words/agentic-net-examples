using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Paths for the sample documents
        string sourcePath = "Source.docx";
        string destinationPath = "Destination.docx";
        string mergedPath = "Merged.docx";

        // Create source document with two tables
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);

        // First table
        Table table1 = srcBuilder.StartTable();
        srcBuilder.InsertCell();
        srcBuilder.Write("Source Table 1 - Row 1, Col 1");
        srcBuilder.InsertCell();
        srcBuilder.Write("Source Table 1 - Row 1, Col 2");
        srcBuilder.EndRow();
        srcBuilder.InsertCell();
        srcBuilder.Write("Source Table 1 - Row 2, Col 1");
        srcBuilder.InsertCell();
        srcBuilder.Write("Source Table 1 - Row 2, Col 2");
        srcBuilder.EndRow();
        srcBuilder.EndTable();

        srcBuilder.Writeln(); // Add a paragraph between tables

        // Second table
        Table table2 = srcBuilder.StartTable();
        srcBuilder.InsertCell();
        srcBuilder.Write("Source Table 2 - Row 1, Col 1");
        srcBuilder.InsertCell();
        srcBuilder.Write("Source Table 2 - Row 1, Col 2");
        srcBuilder.EndRow();
        srcBuilder.InsertCell();
        srcBuilder.Write("Source Table 2 - Row 2, Col 1");
        srcBuilder.InsertCell();
        srcBuilder.Write("Source Table 2 - Row 2, Col 2");
        srcBuilder.EndRow();
        srcBuilder.EndTable();

        sourceDoc.Save(sourcePath);

        // Create destination document with some initial content
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("Destination Document Start");
        destDoc.Save(destinationPath);

        // Load documents (optional, they are already in memory)
        Document source = new Document(sourcePath);
        Document destination = new Document(destinationPath);

        // Find the first table in the source document
        NodeCollection sourceTables = source.GetChildNodes(NodeType.Table, true);
        if (sourceTables.Count == 0)
        {
            throw new InvalidOperationException("No tables found in the source document.");
        }

        Table tableToImport = (Table)sourceTables[0];

        // Import the table into the destination document
        NodeImporter importer = new NodeImporter(source, destination, ImportFormatMode.KeepSourceFormatting);
        Node importedTable = importer.ImportNode(tableToImport, true);

        // Insert the imported table at the end of the destination document body
        Body destBody = destination.FirstSection.Body;
        destBody.AppendChild(importedTable);

        // Save the merged document
        destination.Save(mergedPath);

        // Validation: check that the merged file exists
        if (!File.Exists(mergedPath))
        {
            throw new FileNotFoundException("Merged document was not created.", mergedPath);
        }

        // Validation: ensure the merged document contains at least one table
        Document mergedDoc = new Document(mergedPath);
        NodeCollection mergedTables = mergedDoc.GetChildNodes(NodeType.Table, true);
        if (mergedTables.Count == 0)
        {
            throw new InvalidOperationException("Merged document does not contain any tables.");
        }

        // Optional: indicate success (no interactive input)
        Console.WriteLine("Tables imported and merged document created successfully.");
    }
}

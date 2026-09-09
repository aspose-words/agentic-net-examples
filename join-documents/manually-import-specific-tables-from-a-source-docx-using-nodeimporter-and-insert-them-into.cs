using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Define file paths in the current directory.
        string sourcePath = Path.Combine(Directory.GetCurrentDirectory(), "Source.docx");
        string destinationPath = Path.Combine(Directory.GetCurrentDirectory(), "Destination.docx");
        string mergedPath = Path.Combine(Directory.GetCurrentDirectory(), "Merged.docx");

        // -----------------------------------------------------------------
        // 1. Create a source document containing three tables.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);

        srcBuilder.Writeln("Source Document Header");

        // Table 1
        srcBuilder.StartTable();
        srcBuilder.InsertCell();
        srcBuilder.Write("Table1 Row1 Col1");
        srcBuilder.InsertCell();
        srcBuilder.Write("Table1 Row1 Col2");
        srcBuilder.EndRow();
        srcBuilder.EndTable();
        srcBuilder.Writeln(); // Ensure the table is closed properly.

        // Table 2
        srcBuilder.StartTable();
        srcBuilder.InsertCell();
        srcBuilder.Write("Table2 Row1 Col1");
        srcBuilder.InsertCell();
        srcBuilder.Write("Table2 Row1 Col2");
        srcBuilder.EndRow();
        srcBuilder.EndTable();
        srcBuilder.Writeln(); // Ensure the table is closed properly.

        // Table 3
        srcBuilder.StartTable();
        srcBuilder.InsertCell();
        srcBuilder.Write("Table3 Row1 Col1");
        srcBuilder.InsertCell();
        srcBuilder.Write("Table3 Row1 Col2");
        srcBuilder.EndRow();
        srcBuilder.EndTable();
        srcBuilder.Writeln(); // Ensure the table is closed properly.

        // Save the source document.
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Create a destination document with some initial content.
        // -----------------------------------------------------------------
        Document destinationDoc = new Document();
        DocumentBuilder dstBuilder = new DocumentBuilder(destinationDoc);
        dstBuilder.Writeln("Destination Document Header");
        destinationDoc.Save(destinationPath);

        // -----------------------------------------------------------------
        // 3. Import specific tables (first and third) from source into destination.
        // -----------------------------------------------------------------
        // Retrieve all tables from the source document.
        NodeCollection sourceTables = sourceDoc.GetChildNodes(NodeType.Table, true);
        if (sourceTables.Count < 3)
            throw new InvalidOperationException("Source document does not contain the expected number of tables.");

        // Select the first and third tables.
        Table table1 = (Table)sourceTables[0];
        Table table3 = (Table)sourceTables[2];

        // Create a NodeImporter for efficient repeated imports.
        NodeImporter importer = new NodeImporter(sourceDoc, destinationDoc, ImportFormatMode.KeepSourceFormatting);

        // Import the selected tables.
        Node importedTable1 = importer.ImportNode(table1, true);
        Node importedTable3 = importer.ImportNode(table3, true);

        // Append the imported tables to the end of the destination document.
        CompositeNode dstBody = destinationDoc.FirstSection.Body;
        dstBody.AppendChild(importedTable1);
        dstBody.AppendChild(importedTable3);

        // -----------------------------------------------------------------
        // 4. Save the merged document.
        // -----------------------------------------------------------------
        destinationDoc.Save(mergedPath);

        // -----------------------------------------------------------------
        // 5. Validation: ensure the file exists and contains expected table text.
        // -----------------------------------------------------------------
        if (!File.Exists(mergedPath))
            throw new FileNotFoundException("Merged document was not created.", mergedPath);

        Document mergedDoc = new Document(mergedPath);
        string mergedText = mergedDoc.GetText();

        if (!mergedText.Contains("Table1 Row1 Col1") || !mergedText.Contains("Table3 Row1 Col1"))
            throw new InvalidOperationException("Merged document does not contain the expected table content.");

        // If execution reaches this point, the operation succeeded.
        Console.WriteLine("Tables imported and merged document created successfully at:");
        Console.WriteLine(mergedPath);
    }
}

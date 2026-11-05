using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a source document that contains an original table.
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("Document before the table.");
        srcBuilder.StartTable();
        srcBuilder.InsertCell();
        srcBuilder.Write("Original Cell 1");
        srcBuilder.InsertCell();
        srcBuilder.Write("Original Cell 2");
        srcBuilder.EndRow();
        srcBuilder.EndTable();
        srcBuilder.Writeln("Document after the table.");

        // Create a template document that contains the replacement table.
        Document templateDoc = new Document();
        DocumentBuilder tmplBuilder = new DocumentBuilder(templateDoc);
        tmplBuilder.StartTable();
        tmplBuilder.InsertCell();
        tmplBuilder.Write("Template Cell A");
        tmplBuilder.InsertCell();
        tmplBuilder.Write("Template Cell B");
        tmplBuilder.EndRow();
        tmplBuilder.EndTable();

        // Locate the original table in the source document.
        Table originalTable = sourceDoc.GetChild(NodeType.Table, 0, true) as Table;
        if (originalTable == null)
            throw new InvalidOperationException("Original table not found in the source document.");

        // Import the template table into the source document.
        NodeImporter importer = new NodeImporter(templateDoc, sourceDoc, ImportFormatMode.KeepSourceFormatting);
        Table templateTable = templateDoc.GetChild(NodeType.Table, 0, true) as Table;
        Table importedTable = importer.ImportNode(templateTable, true) as Table;
        if (importedTable == null)
            throw new InvalidOperationException("Failed to import the template table.");

        // Replace the original table with the imported template table.
        // The parent of a Table is a CompositeNode (e.g., Body). Cast to CompositeNode to use InsertAfter.
        CompositeNode parent = originalTable.ParentNode as CompositeNode;
        if (parent == null)
            throw new InvalidOperationException("Unable to locate a valid parent for the original table.");

        parent.InsertAfter(importedTable, originalTable);
        originalTable.Remove();

        // Save the modified document.
        string outputPath = "Output.docx";
        sourceDoc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }
}

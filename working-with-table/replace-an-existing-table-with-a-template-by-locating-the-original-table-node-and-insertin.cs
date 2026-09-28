using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a source document with an original table.
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("Document before the table.");
        srcBuilder.StartTable();
        srcBuilder.InsertCell();
        srcBuilder.Write("Original Cell 1");
        srcBuilder.EndRow();
        srcBuilder.EndTable();
        srcBuilder.Writeln("Document after the table.");
        sourceDoc.Save("original.docx");

        // Locate the original table in the source document.
        Table originalTable = sourceDoc.GetChild(NodeType.Table, 0, true) as Table;
        if (originalTable == null)
            throw new InvalidOperationException("Original table not found.");

        // Build a template table in a separate document.
        Document templateDoc = new Document();
        DocumentBuilder tmplBuilder = new DocumentBuilder(templateDoc);
        tmplBuilder.StartTable();
        tmplBuilder.InsertCell();
        tmplBuilder.Write("Template Cell A");
        tmplBuilder.InsertCell();
        tmplBuilder.Write("Template Cell B");
        tmplBuilder.EndRow();
        tmplBuilder.EndTable();

        // Import the template table into the source document.
        NodeImporter importer = new NodeImporter(templateDoc, sourceDoc, ImportFormatMode.KeepSourceFormatting);
        Table importedTable = importer.ImportNode(templateDoc.FirstSection.Body.Tables[0], true) as Table;
        if (importedTable == null)
            throw new InvalidOperationException("Failed to import template table.");

        // Replace the original table with the imported template table.
        // Insert the new table after the original one, then remove the original.
        originalTable.ParentNode.InsertAfter(importedTable, originalTable);
        originalTable.Remove();

        // Save the resulting document.
        sourceDoc.Save("result.docx");

        // Verify that the output file was created.
        if (!File.Exists("result.docx"))
            throw new InvalidOperationException("Result document was not saved.");
    }
}

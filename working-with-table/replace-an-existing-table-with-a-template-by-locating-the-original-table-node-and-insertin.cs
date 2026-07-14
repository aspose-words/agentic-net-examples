using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // ---------- Create a source document with an original table ----------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Build a simple 2x2 table that will be replaced later
        Table originalTable = builder.StartTable();
        builder.InsertCell();
        builder.Write("Original 1,1");
        builder.InsertCell();
        builder.Write("Original 1,2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Original 2,1");
        builder.InsertCell();
        builder.Write("Original 2,2");
        builder.EndRow();
        builder.EndTable();

        // Save the source document (optional, but demonstrates the file exists)
        sourceDoc.Save("Source.docx");

        // ---------- Create a template document containing the replacement table ----------
        Document templateDoc = new Document();
        DocumentBuilder tmplBuilder = new DocumentBuilder(templateDoc);

        // Build a different 2x2 table that will replace the original one
        Table templateTable = tmplBuilder.StartTable();
        tmplBuilder.InsertCell();
        tmplBuilder.Write("Template A");
        tmplBuilder.InsertCell();
        tmplBuilder.Write("Template B");
        tmplBuilder.EndRow();

        tmplBuilder.InsertCell();
        tmplBuilder.Write("Template C");
        tmplBuilder.InsertCell();
        tmplBuilder.Write("Template D");
        tmplBuilder.EndRow();
        tmplBuilder.EndTable();

        // ---------- Import the template table into the source document ----------
        NodeImporter importer = new NodeImporter(templateDoc, sourceDoc, ImportFormatMode.KeepSourceFormatting);
        Table importedTable = (Table)importer.ImportNode(templateTable, true);

        // ---------- Locate the original table and replace it ----------
        Table tableToReplace = (Table)sourceDoc.GetChild(NodeType.Table, 0, true);
        // Insert the new table after the original one
        tableToReplace.ParentNode.InsertAfter(importedTable, tableToReplace);
        // Remove the original table from the document
        tableToReplace.Remove();

        // ---------- Save the final document ----------
        sourceDoc.Save("Result.docx");
    }
}

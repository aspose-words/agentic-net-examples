using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableReplacementExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a source document that contains the original table.
            Document sourceDoc = new Document();
            DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
            srcBuilder.Writeln("Document with original table:");
            srcBuilder.StartTable();
            srcBuilder.InsertCell();
            srcBuilder.Write("Original Cell 1");
            srcBuilder.EndRow();
            srcBuilder.InsertCell();
            srcBuilder.Write("Original Cell 2");
            srcBuilder.EndRow();
            srcBuilder.EndTable();
            srcBuilder.Writeln("End of document.");
            sourceDoc.Save("Original.docx");

            // Create a template document that contains the replacement table.
            Document templateDoc = new Document();
            DocumentBuilder tmplBuilder = new DocumentBuilder(templateDoc);
            tmplBuilder.StartTable();
            tmplBuilder.InsertCell();
            tmplBuilder.Write("Template Cell A");
            tmplBuilder.EndRow();
            tmplBuilder.InsertCell();
            tmplBuilder.Write("Template Cell B");
            tmplBuilder.EndRow();
            tmplBuilder.EndTable();
            templateDoc.Save("Template.docx");

            // Locate the original table in the source document.
            Table originalTable = sourceDoc.GetChild(NodeType.Table, 0, true) as Table;
            if (originalTable == null)
                throw new InvalidOperationException("Original table not found.");

            // Import the template table into the source document.
            Table templateTable = templateDoc.GetChild(NodeType.Table, 0, true) as Table;
            if (templateTable == null)
                throw new InvalidOperationException("Template table not found.");

            NodeImporter importer = new NodeImporter(templateDoc, sourceDoc, ImportFormatMode.KeepSourceFormatting);
            Table importedTable = importer.ImportNode(templateTable, true) as Table;
            if (importedTable == null)
                throw new InvalidOperationException("Failed to import template table.");

            // Replace the original table with the imported template table.
            // Insert the imported table after the original, then remove the original.
            originalTable.ParentNode.InsertAfter(importedTable, originalTable);
            originalTable.Remove();

            // Save the resulting document.
            sourceDoc.Save("Result.docx");

            // Verify that the output file was created.
            if (!File.Exists("Result.docx"))
                throw new FileNotFoundException("Result.docx was not created.");

            // Optional: clean up intermediate files (comment out if you want to keep them).
            // File.Delete("Original.docx");
            // File.Delete("Template.docx");
        }
    }
}

using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

class ExportTables
{
    static void Main()
    {
        // Path to the source document that contains the tables.
        string sourcePath = "SourceDocument.docx";

        // Load the source document.
        Document sourceDoc = new Document(sourcePath);

        // Convert any table style formatting to direct formatting so that each table
        // retains its appearance when saved independently.
        sourceDoc.ExpandTableStylesToDirectFormatting();

        // Retrieve all tables in the document (including those inside sections, headers, etc.).
        NodeCollection tables = sourceDoc.GetChildNodes(NodeType.Table, true);

        // Export each table to its own DOCX file.
        for (int i = 0; i < tables.Count; i++)
        {
            Table table = (Table)tables[i];

            // Create a new blank document for the current table.
            Document tableDoc = new Document();

            // Remove the default empty paragraph that Aspose.Words adds to a new document.
            tableDoc.FirstSection.Body.RemoveAllChildren();

            // Import the table node from the source document into the new document.
            // KeepSourceFormatting preserves the original formatting of the table.
            NodeImporter importer = new NodeImporter(sourceDoc, tableDoc, ImportFormatMode.KeepSourceFormatting);
            Node importedTable = importer.ImportNode(table, true);

            // Add the imported table to the body of the new document.
            tableDoc.FirstSection.Body.AppendChild(importedTable);

            // Save the new document. The file name includes the table index.
            string outputPath = $"Table_{i}.docx";
            tableDoc.Save(outputPath);
        }
    }
}

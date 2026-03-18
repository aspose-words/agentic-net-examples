using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Replacing;

class ReplaceTableWithTemplate
{
    static void Main()
    {
        // Load the document that contains the table to be replaced.
        Document mainDoc = new Document("MainDocument.docx");

        // Load the template document that contains the replacement table.
        Document templateDoc = new Document("TemplateDocument.docx");

        // Locate the original table in the main document.
        // For this example we replace the first table found.
        Table originalTable = (Table)mainDoc.GetChild(NodeType.Table, 0, true);
        if (originalTable == null)
            throw new InvalidOperationException("No table found in the main document.");

        // Locate the replacement table in the template document.
        // For this example we use the first table in the template.
        Table templateTable = (Table)templateDoc.GetChild(NodeType.Table, 0, true);
        if (templateTable == null)
            throw new InvalidOperationException("No table found in the template document.");

        // Import the template table into the main document.
        NodeImporter importer = new NodeImporter(templateDoc, mainDoc, ImportFormatMode.KeepSourceFormatting);
        Table importedTable = (Table)importer.ImportNode(templateTable, true);

        // Insert the imported table after the original table and then remove the original.
        CompositeNode parent = originalTable.ParentNode;
        parent.InsertAfter(importedTable, originalTable);
        originalTable.Remove();

        // Save the modified document.
        mainDoc.Save("ResultDocument.docx");
    }
}

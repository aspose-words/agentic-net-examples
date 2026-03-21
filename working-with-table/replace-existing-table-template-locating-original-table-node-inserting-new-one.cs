using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

class ReplaceTableWithTemplate
{
    static void Main()
    {
        const string mainDocPath = "MainDocument.docx";
        const string templateDocPath = "TemplateDocument.docx";
        const string resultDocPath = "ResultDocument.docx";

        // Ensure the main document exists; if not, create a simple one with a table.
        if (!File.Exists(mainDocPath))
        {
            CreateDocumentWithTable(mainDocPath, new[,]
            {
                { "Header 1", "Header 2" },
                { "Main 1", "Main 2" }
            });
        }

        // Ensure the template document exists; if not, create a simple one with a different table.
        if (!File.Exists(templateDocPath))
        {
            CreateDocumentWithTable(templateDocPath, new[,]
            {
                { "Template Header 1", "Template Header 2" },
                { "Template 1", "Template 2" },
                { "Template 3", "Template 4" }
            });
        }

        // Load the document that contains the table to be replaced.
        Document mainDoc = new Document(mainDocPath);

        // Load the template document that contains the replacement table.
        Document templateDoc = new Document(templateDocPath);

        // Locate the original table in the main document (first table). If none, create one.
        Table originalTable = (Table)mainDoc.GetChild(NodeType.Table, 0, true);
        if (originalTable == null)
        {
            CreateDocumentWithTable(mainDocPath, new[,]
            {
                { "Fallback Header 1", "Fallback Header 2" },
                { "Fallback 1", "Fallback 2" }
            });
            mainDoc = new Document(mainDocPath);
            originalTable = (Table)mainDoc.GetChild(NodeType.Table, 0, true);
        }

        // Locate the replacement table in the template document (first table). If none, create one.
        Table templateTable = (Table)templateDoc.GetChild(NodeType.Table, 0, true);
        if (templateTable == null)
        {
            CreateDocumentWithTable(templateDocPath, new[,]
            {
                { "Fallback Template Header 1", "Fallback Template Header 2" },
                { "Fallback Template 1", "Fallback Template 2" }
            });
            templateDoc = new Document(templateDocPath);
            templateTable = (Table)templateDoc.GetChild(NodeType.Table, 0, true);
        }

        // Import the template table into the main document.
        NodeImporter importer = new NodeImporter(templateDoc, mainDoc, ImportFormatMode.KeepSourceFormatting);
        Table importedTable = (Table)importer.ImportNode(templateTable, true);

        // Insert the imported table after the original table and then remove the original.
        CompositeNode parent = originalTable.ParentNode;
        parent.InsertAfter(importedTable, originalTable);
        originalTable.Remove();

        // Save the modified document.
        mainDoc.Save(resultDocPath);

        Console.WriteLine($"Table replaced successfully. Result saved to '{resultDocPath}'.");
    }

    private static void CreateDocumentWithTable(string filePath, string[,] data)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.StartTable();
        for (int row = 0; row < data.GetLength(0); row++)
        {
            for (int col = 0; col < data.GetLength(1); col++)
            {
                builder.InsertCell();
                builder.Write(data[row, col]);
            }
            builder.EndRow();
        }
        builder.EndTable();

        doc.Save(filePath);
    }
}

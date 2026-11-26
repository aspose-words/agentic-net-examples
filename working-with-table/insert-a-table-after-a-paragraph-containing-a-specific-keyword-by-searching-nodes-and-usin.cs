using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and add some paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First paragraph.");
        builder.Writeln("Paragraph with keyword: InsertHere");
        builder.Writeln("Another paragraph.");

        // Search for the paragraph that contains the specific keyword.
        Paragraph targetParagraph = null;
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph para in paragraphs)
        {
            if (para.GetText().Contains("InsertHere"))
            {
                targetParagraph = para;
                break;
            }
        }

        if (targetParagraph == null)
            throw new InvalidOperationException("Keyword paragraph not found.");

        // Build a table in a temporary document.
        Document tempDoc = new Document();
        DocumentBuilder tblBuilder = new DocumentBuilder(tempDoc);
        tblBuilder.StartTable();

        tblBuilder.InsertCell();
        tblBuilder.Write("Cell 1");
        tblBuilder.EndRow();

        tblBuilder.InsertCell();
        tblBuilder.Write("Cell 2");
        tblBuilder.EndRow();

        tblBuilder.EndTable();

        // Retrieve the created table.
        Table table = (Table)tempDoc.GetChild(NodeType.Table, 0, true);

        // Import the table into the main document.
        NodeImporter importer = new NodeImporter(tempDoc, doc, ImportFormatMode.KeepSourceFormatting);
        Table importedTable = (Table)importer.ImportNode(table, true);

        // Insert the table after the target paragraph.
        CompositeNode parent = targetParagraph.ParentNode as CompositeNode;
        parent.InsertAfter(importedTable, targetParagraph);

        // Save the resulting document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to save the document.");
    }
}

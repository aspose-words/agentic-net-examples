using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a sample source document.
        const string sourcePath = "source.docx";
        CreateSourceDocument(sourcePath);

        // Load the source document.
        Document sourceDoc = new Document(sourcePath);

        // Locate the start cell (first cell in the document).
        Cell startCell = sourceDoc.GetChildNodes(NodeType.Cell, true)[0] as Cell;
        if (startCell == null)
            throw new InvalidOperationException("Start cell not found.");

        // Locate the end paragraph (the first paragraph that follows the table).
        Paragraph endParagraph = null;
        NodeCollection allParagraphs = sourceDoc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph para in allParagraphs)
        {
            // The paragraph whose previous sibling is a table is the first paragraph after the table.
            if (para.PreviousSibling != null && para.PreviousSibling.NodeType == NodeType.Table)
            {
                endParagraph = para;
                break;
            }
        }

        if (endParagraph == null)
            throw new InvalidOperationException("End paragraph not found.");

        // Build a new document that will contain the extracted range.
        Document resultDoc = new Document();
        resultDoc.RemoveAllChildren();

        // Create a new section and body for the result document.
        Section resultSection = new Section(resultDoc);
        resultDoc.AppendChild(resultSection);
        Body resultBody = new Body(resultDoc);
        resultSection.AppendChild(resultBody);

        // Import the table that contains the start cell.
        Table containingTable = startCell.GetAncestor(NodeType.Table) as Table;
        if (containingTable == null)
            throw new InvalidOperationException("Containing table not found.");

        NodeImporter importer = new NodeImporter(sourceDoc, resultDoc, ImportFormatMode.KeepSourceFormatting);
        Node importedTable = importer.ImportNode(containingTable, true);
        resultBody.AppendChild(importedTable);

        // Import the end paragraph.
        Node importedParagraph = importer.ImportNode(endParagraph, true);
        resultBody.AppendChild(importedParagraph);

        // Save the extracted content.
        const string resultPath = "extracted.docx";
        resultDoc.Save(resultPath);

        // Verify that the output file was created.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("Extraction output was not created.");
    }

    // Helper method to create a sample document with a table followed by paragraphs.
    private static void CreateSourceDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1,1");
        builder.InsertCell();
        builder.Write("Cell 1,2");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Cell 2,1");
        builder.InsertCell();
        builder.Write("Cell 2,2");
        builder.EndRow();
        builder.EndTable();

        // Add paragraphs after the table.
        builder.Writeln("Paragraph after table 1.");
        builder.Writeln("Paragraph after table 2.");

        doc.Save(filePath);
    }
}

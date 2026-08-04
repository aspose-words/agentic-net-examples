using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // -----------------------------
        // Create a source document with a nested table
        // -----------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Intro paragraph.
        builder.Writeln("Document with nested tables:");

        // Build the outer table (2x2).
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Outer Cell 1");
        builder.InsertCell();
        builder.Write("Outer Cell 2");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Outer Cell 3");
        builder.InsertCell();
        builder.Write("Outer Cell 4");
        builder.EndTable();

        // Retrieve the outer table node.
        Table outerTable = sourceDoc.GetChildNodes(NodeType.Table, true)[0] as Table;
        if (outerTable == null)
            throw new InvalidOperationException("Outer table was not created.");

        // Insert an inner table into the first cell of the outer table.
        Cell targetCell = outerTable.FirstRow.FirstCell;
        builder.MoveTo(targetCell.FirstParagraph);
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Inner Cell 1");
        builder.InsertCell();
        builder.Write("Inner Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Save the source document to a local file.
        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // -----------------------------
        // Load the document and extract the outer table (including its nested table)
        // -----------------------------
        Document loadedDoc = new Document(sourcePath);

        Table tableToExtract = loadedDoc.GetChildNodes(NodeType.Table, true)[0] as Table;
        if (tableToExtract == null)
            throw new InvalidOperationException("Table to extract was not found.");

        // Create a new empty document that will hold the extracted segment.
        Document resultDoc = new Document();
        resultDoc.RemoveAllChildren();

        // Build the minimal required structure: Section -> Body.
        Section resultSection = new Section(resultDoc);
        resultDoc.AppendChild(resultSection);

        Body resultBody = new Body(resultDoc);
        resultSection.AppendChild(resultBody);

        // Import the table (with its nested table) into the new document.
        NodeImporter importer = new NodeImporter(loadedDoc, resultDoc, ImportFormatMode.KeepSourceFormatting);
        Node importedTable = importer.ImportNode(tableToExtract, true);
        resultBody.AppendChild(importedTable);

        // Save the extracted segment.
        const string resultPath = "extracted.docx";
        resultDoc.Save(resultPath);

        // Verify that the output file was created.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The extracted document was not created.");
    }
}

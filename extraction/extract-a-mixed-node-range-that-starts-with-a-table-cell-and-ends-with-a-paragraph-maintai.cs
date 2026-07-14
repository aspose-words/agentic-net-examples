using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // -------------------------------------------------
        // 1. Create a sample source document with a table and a paragraph.
        // -------------------------------------------------
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);

        // Build a 2x2 table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndRow();
        builder.EndTable();

        // Paragraph after the table – this will be the end boundary.
        builder.Writeln("End paragraph.");

        // Save the source document locally.
        const string sourcePath = "source.docx";
        source.Save(sourcePath);

        // -------------------------------------------------
        // 2. Load the document for extraction.
        // -------------------------------------------------
        Document loaded = new Document(sourcePath);

        // Locate the start node – the first cell of the first table.
        Table table = loaded.GetChildNodes(NodeType.Table, true)[0] as Table;
        if (table == null)
            throw new InvalidOperationException("Table not found in the source document.");

        Cell startCell = table.FirstRow.FirstCell;
        if (startCell == null)
            throw new InvalidOperationException("Start cell not found.");

        // Locate the end node – the paragraph that follows the table.
        // The body contains the table (as a block node) and then the paragraph.
        Paragraph endParagraph = loaded.FirstSection.Body.Paragraphs[1];
        if (endParagraph == null)
            throw new InvalidOperationException("End paragraph not found.");

        // -------------------------------------------------
        // 3. Prepare the destination document.
        // -------------------------------------------------
        Document result = new Document();
        result.RemoveAllChildren();

        Section resultSection = new Section(result);
        result.AppendChild(resultSection);

        Body resultBody = new Body(result);
        resultSection.AppendChild(resultBody);

        // -------------------------------------------------
        // 4. Import the required nodes using NodeImporter.
        // -------------------------------------------------
        NodeImporter importer = new NodeImporter(loaded, result, ImportFormatMode.KeepSourceFormatting);

        // Import the whole table (which contains the start cell) into the result.
        Node importedTable = importer.ImportNode(table, true);
        resultBody.AppendChild(importedTable);

        // Import the end paragraph.
        Node importedParagraph = importer.ImportNode(endParagraph, true);
        resultBody.AppendChild(importedParagraph);

        // -------------------------------------------------
        // 5. Save the extracted mixed range.
        // -------------------------------------------------
        const string resultPath = "extracted.docx";
        result.Save(resultPath);

        // Verify that the output file was created.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The extracted document was not created.");
    }
}

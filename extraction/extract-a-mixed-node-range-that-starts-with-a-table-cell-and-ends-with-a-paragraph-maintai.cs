using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

public class ExtractMixedRange
{
    public static void Main()
    {
        // -------------------------------------------------
        // 1. Create a sample source document.
        // -------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Paragraph before the table.
        builder.Writeln("Paragraph before the table.");

        // Table with a single cell that will be the start of the extraction range.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Content inside the start cell.");
        builder.EndRow();
        builder.EndTable();

        // Paragraphs after the table.
        builder.Writeln("Paragraph after table - part 1.");
        builder.Writeln("Paragraph after table - part 2 - END TARGET.");

        // Save the source document.
        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // -------------------------------------------------
        // 2. Load the document for extraction.
        // -------------------------------------------------
        Document loadedDoc = new Document(sourcePath);
        Body body = loadedDoc.FirstSection.Body;

        // Locate the start cell (first cell of the first table).
        Cell startCell = loadedDoc.GetChildNodes(NodeType.Cell, true)
                                   .Cast<Cell>()
                                   .FirstOrDefault();
        if (startCell == null)
            throw new InvalidOperationException("Start cell not found.");

        // The table that contains the start cell.
        Table startTable = startCell.ParentRow?.ParentTable;
        if (startTable == null)
            throw new InvalidOperationException("Parent table of start cell not found.");

        // Locate the end paragraph by its unique text.
        Paragraph endParagraph = loadedDoc.GetChildNodes(NodeType.Paragraph, true)
                                          .Cast<Paragraph>()
                                          .FirstOrDefault(p => p.GetText().Contains("END TARGET"));
        if (endParagraph == null)
            throw new InvalidOperationException("End paragraph not found.");

        // Gather the direct child nodes of the body in document order.
        List<Node> bodyChildren = body.GetChildNodes(NodeType.Any, false)
                                      .Cast<Node>()
                                      .ToList();

        int startIndex = bodyChildren.IndexOf(startTable);
        int endIndex = bodyChildren.IndexOf(endParagraph);
        if (startIndex == -1 || endIndex == -1 || startIndex > endIndex)
            throw new InvalidOperationException("Invalid extraction boundaries.");

        // -------------------------------------------------
        // 3. Build a new document that will contain the extracted range.
        // -------------------------------------------------
        Document resultDoc = new Document();
        resultDoc.RemoveAllChildren(); // Clear the default empty section.

        // Create a minimal valid structure: Section -> Body.
        Section resultSection = new Section(resultDoc);
        resultDoc.AppendChild(resultSection);
        Body resultBody = new Body(resultDoc);
        resultSection.AppendChild(resultBody);

        // Use NodeImporter to import nodes from the source document into the result document.
        NodeImporter importer = new NodeImporter(loadedDoc, resultDoc, ImportFormatMode.KeepSourceFormatting);

        for (int i = startIndex; i <= endIndex; i++)
        {
            Node nodeToImport = bodyChildren[i];
            Node importedNode = importer.ImportNode(nodeToImport, true);
            resultBody.AppendChild(importedNode);
        }

        // -------------------------------------------------
        // 4. Save the extracted content.
        // -------------------------------------------------
        const string resultPath = "extracted.docx";
        resultDoc.Save(resultPath);

        // Validate that the output file was created.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("Extraction result file was not created.");
    }
}

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a source document with sample content.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("First paragraph in source.");
        srcBuilder.Writeln("Second paragraph in source.");
        srcBuilder.StartTable();
        srcBuilder.InsertCell();
        srcBuilder.Write("Cell 1");
        srcBuilder.InsertCell();
        srcBuilder.Write("Cell 2");
        srcBuilder.EndRow();
        srcBuilder.EndTable();
        srcBuilder.Writeln("Third paragraph in source.");

        // Save the source document (lifecycle rule: save).
        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Load the source document (lifecycle rule: load).
        // -----------------------------------------------------------------
        Document loadedSource = new Document(sourcePath);

        // -----------------------------------------------------------------
        // 3. Extract all block-level nodes (paragraphs and tables).
        // -----------------------------------------------------------------
        NodeCollection sourceBlocks = loadedSource.GetChildNodes(NodeType.Any, true);
        // Filter only Paragraph and Table nodes that are direct children of the body.
        var extractedBlocks = new System.Collections.Generic.List<Node>();
        foreach (Node node in sourceBlocks)
        {
            if (node.NodeType == NodeType.Paragraph || node.NodeType == NodeType.Table)
                extractedBlocks.Add(node);
        }

        // -----------------------------------------------------------------
        // 4. Create a new destination document.
        // -----------------------------------------------------------------
        Document destDoc = new Document();
        // Remove the default empty section/paragraph.
        destDoc.RemoveAllChildren();

        // Build a minimal document structure: Section -> Body.
        Section destSection = new Section(destDoc);
        destDoc.AppendChild(destSection);
        Body destBody = new Body(destDoc);
        destSection.AppendChild(destBody);

        // -----------------------------------------------------------------
        // 5. Import extracted nodes into the destination document and prepend them.
        // -----------------------------------------------------------------
        NodeImporter importer = new NodeImporter(loadedSource, destDoc, ImportFormatMode.KeepSourceFormatting);

        // Prepend in reverse order so the original order is preserved.
        for (int i = extractedBlocks.Count - 1; i >= 0; i--)
        {
            Node importedNode = importer.ImportNode(extractedBlocks[i], true);
            destBody.PrependChild(importedNode);
        }

        // -----------------------------------------------------------------
        // 6. Save the destination document.
        // -----------------------------------------------------------------
        const string resultPath = "result.docx";
        destDoc.Save(resultPath);

        // -----------------------------------------------------------------
        // 7. Validate that the output file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The result document was not created.");

        // Optional: indicate success (no console interaction required by the task).
    }
}

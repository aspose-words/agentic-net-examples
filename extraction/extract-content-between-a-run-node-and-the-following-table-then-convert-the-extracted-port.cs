using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;

public class ExtractRunToTableToXps
{
    public static void Main()
    {
        // Create a sample document with runs and a table.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // First paragraph with two runs.
        builder.Writeln("Paragraph before the run.");
        builder.Write("RunStart ");               // This run will be the start marker.
        builder.Write("Middle run text. ");       // Additional run(s) after the start marker.
        builder.Writeln();                       // End of paragraph.

        // Insert a table after the runs.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.EndRow();
        builder.EndTable();

        // Ensure the document has the expected structure.
        Node startRun = sourceDoc.GetChildNodes(NodeType.Run, true)[0];
        if (startRun == null)
            throw new InvalidOperationException("Run node not found.");

        // Locate the first table that follows the start run.
        Node endTable = FindNextNodeOfType(startRun, NodeType.Table);
        if (endTable == null)
            throw new InvalidOperationException("Following table not found.");

        // Create a new empty document that will hold the extracted content.
        Document extractedDoc = new Document();
        extractedDoc.RemoveAllChildren();

        // Build the minimal required structure: Section -> Body.
        Section section = new Section(extractedDoc);
        extractedDoc.AppendChild(section);
        Body body = new Body(extractedDoc);
        section.AppendChild(body);

        // Helper paragraph for inline runs.
        Paragraph inlineParagraph = null;

        // Traverse nodes between startRun (exclusive) and endTable (exclusive).
        Node current = startRun.NextPreOrder(sourceDoc);
        while (current != null && current != endTable)
        {
            // If we encounter a block node, import it directly to the body.
            if (current.NodeType == NodeType.Paragraph || current.NodeType == NodeType.Table)
            {
                Node importedBlock = extractedDoc.ImportNode(current, true);
                body.AppendChild(importedBlock);
                // Skip the subtree of the block node to avoid duplicate imports.
                current = current.NextSibling;
                continue;
            }

            // Inline runs are placed inside a paragraph.
            if (current.NodeType == NodeType.Run)
            {
                if (inlineParagraph == null)
                {
                    inlineParagraph = new Paragraph(extractedDoc);
                    body.AppendChild(inlineParagraph);
                }

                Node importedRun = extractedDoc.ImportNode(current, true);
                inlineParagraph.AppendChild(importedRun);
            }
            else
            {
                // For any other node types, import them as they appear.
                Node imported = extractedDoc.ImportNode(current, true);
                body.AppendChild(imported);
            }

            current = current.NextPreOrder(sourceDoc);
        }

        // Validate that something was extracted.
        if (body.Count == 0)
            throw new InvalidOperationException("No content was extracted between the run and the table.");

        // Save the extracted portion as XPS.
        string outputPath = "ExtractedContent.xps";
        extractedDoc.Save(outputPath, new XpsSaveOptions());

        // Indicate successful completion (no interactive output required).
        // The file "ExtractedContent.xps" will be created in the program's working directory.
    }

    // Helper method to find the next node of a specific type after a given node.
    private static Node FindNextNodeOfType(Node startNode, NodeType targetType)
    {
        Node node = startNode.NextPreOrder(startNode.Document);
        while (node != null)
        {
            if (node.NodeType == targetType)
                return node;
            node = node.NextPreOrder(startNode.Document);
        }
        return null;
    }
}

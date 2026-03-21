using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;

class ExtractAndConvert
{
    static void Main()
    {
        // -----------------------------
        // Create a sample source document.
        // -----------------------------
        Document srcDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(srcDoc);

        builder.Writeln("Intro paragraph before the marker.");

        // Paragraph that contains the start marker.
        builder.Writeln("StartMarker");

        // Content that should be extracted.
        builder.Writeln("First line to extract.");
        builder.Writeln("Second line to extract.");

        // Insert a table – extraction stops before this node.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.EndRow();
        builder.EndTable();

        // Content after the table (should NOT be extracted).
        builder.Writeln("Paragraph after the table.");

        // -------------------------------------------------
        // Locate the start Run containing the marker text.
        // -------------------------------------------------
        Run startRun = null;
        foreach (Run run in srcDoc.GetChildNodes(NodeType.Run, true))
        {
            if (run.Text.Contains("StartMarker"))
            {
                startRun = run;
                break;
            }
        }

        if (startRun == null)
            throw new InvalidOperationException("Start run not found.");

        // -------------------------------------------------
        // Collect block nodes (Paragraphs, etc.) after the start run up to the next table.
        // -------------------------------------------------
        List<Node> extractedNodes = new List<Node>();
        Node curNode = startRun.ParentNode?.NextSibling; // start from the paragraph after the marker paragraph
        while (curNode != null && curNode.NodeType != NodeType.Table)
        {
            extractedNodes.Add(curNode);
            curNode = curNode.NextSibling;
        }

        // -------------------------------------------------
        // Build a new document containing the extracted nodes.
        // -------------------------------------------------
        Document destDoc = new Document();
        destDoc.RemoveAllChildren();

        Section section = new Section(destDoc);
        destDoc.AppendChild(section);
        Body body = new Body(destDoc);
        section.AppendChild(body);

        NodeImporter importer = new NodeImporter(srcDoc, destDoc, ImportFormatMode.KeepSourceFormatting);

        foreach (Node node in extractedNodes)
        {
            Node importedNode = importer.ImportNode(node, true);
            body.AppendChild(importedNode);
        }

        // -------------------------------------------------
        // Save the extracted portion as XPS.
        // -------------------------------------------------
        XpsSaveOptions xpsOptions = new XpsSaveOptions();
        destDoc.Save("ExtractedContent.xps", xpsOptions);
    }
}

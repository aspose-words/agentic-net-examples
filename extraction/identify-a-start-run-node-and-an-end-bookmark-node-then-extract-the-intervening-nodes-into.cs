using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

public class ExtractBetweenRunAndBookmark
{
    public static void Main()
    {
        // ---------- Create a sample source document ----------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Intro paragraph (outside the extraction range).
        builder.Writeln("Intro paragraph.");

        // Start Run node – this will be the start marker.
        builder.InsertParagraph();
        builder.Write("StartRun");
        Run startRun = builder.CurrentParagraph.Runs[0];

        // Content that should be extracted.
        builder.Writeln(); // Move to next paragraph.
        builder.Writeln("Paragraph between 1");
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell1");
        builder.EndRow();
        builder.EndTable();

        // End Bookmark – this will be the end marker.
        builder.StartBookmark("EndBookmark");
        builder.Writeln("End bookmark paragraph");
        builder.EndBookmark("EndBookmark");

        // Save the source document.
        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // ---------- Load the document for processing ----------
        Document loadedDoc = new Document(sourcePath);

        // Locate the start Run node by its text.
        Run foundStartRun = loadedDoc.GetChildNodes(NodeType.Run, true)
            .OfType<Run>()
            .FirstOrDefault(r => r.Text == "StartRun");
        if (foundStartRun == null)
            throw new InvalidOperationException("Start Run node not found.");

        // Locate the end BookmarkStart node.
        Bookmark endBookmark = loadedDoc.Range.Bookmarks["EndBookmark"];
        if (endBookmark == null)
            throw new InvalidOperationException("End bookmark not found.");
        BookmarkStart endBookmarkStart = endBookmark.BookmarkStart;

        // Build a flat list of all nodes in document order.
        NodeCollection allNodes = loadedDoc.GetChildNodes(NodeType.Any, true);
        int startIndex = -1;
        int endIndex = -1;
        for (int i = 0; i < allNodes.Count; i++)
        {
            if (allNodes[i] == foundStartRun)
                startIndex = i;
            if (allNodes[i] == endBookmarkStart)
                endIndex = i;
        }

        if (startIndex == -1 || endIndex == -1)
            throw new InvalidOperationException("Boundary nodes not found in the node collection.");
        if (endIndex <= startIndex)
            throw new InvalidOperationException("End bookmark occurs before start run.");

        // ---------- Prepare the result document ----------
        Document resultDoc = new Document();
        resultDoc.RemoveAllChildren();
        Section resultSection = new Section(resultDoc);
        resultDoc.AppendChild(resultSection);
        Body resultBody = new Body(resultDoc);
        resultSection.AppendChild(resultBody);

        // NodeImporter will handle moving nodes between documents.
        NodeImporter importer = new NodeImporter(loadedDoc, resultDoc, ImportFormatMode.KeepSourceFormatting);

        // Clone and append nodes that lie between the start run and end bookmark.
        for (int i = startIndex + 1; i < endIndex; i++)
        {
            Node currentNode = allNodes[i];
            switch (currentNode.NodeType)
            {
                case NodeType.Paragraph:
                case NodeType.Table:
                    // Import the node into the result document and append.
                    Node importedBlock = importer.ImportNode(currentNode, true);
                    resultBody.AppendChild(importedBlock);
                    break;

                case NodeType.Run:
                    // Import the run, wrap it inside a new paragraph.
                    Node importedRun = importer.ImportNode(currentNode, true);
                    Paragraph para = new Paragraph(resultDoc);
                    para.AppendChild(importedRun);
                    resultBody.AppendChild(para);
                    break;

                default:
                    // Other node types are ignored for this example.
                    break;
            }
        }

        // ---------- Save the extracted content ----------
        const string resultPath = "extracted.docx";
        resultDoc.Save(resultPath);

        // Validate that the output file was created.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("Extraction failed: output file was not created.");
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample source document with three paragraphs.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("Source Paragraph 1");
        srcBuilder.Writeln("Source Paragraph 2");
        srcBuilder.Writeln("Source Paragraph 3");
        sourceDoc.Save("source.docx");

        // -----------------------------------------------------------------
        // 2. Load the source document and extract its paragraphs.
        // -----------------------------------------------------------------
        Document loadedSource = new Document("source.docx");
        List<Node> extractedNodes = new List<Node>();
        foreach (Paragraph para in loadedSource.FirstSection.Body.Paragraphs)
        {
            // Deep clone each paragraph so the original stays unchanged.
            extractedNodes.Add(para.Clone(true));
        }

        if (extractedNodes.Count == 0)
            throw new InvalidOperationException("No paragraphs were extracted from the source document.");

        // -----------------------------------------------------------------
        // 3. Create a destination document with a custom bookmark.
        // -----------------------------------------------------------------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("Destination Header");
        destBuilder.StartBookmark("InsertHere");
        destBuilder.EndBookmark("InsertHere");
        destBuilder.Writeln("Destination Footer");

        // -----------------------------------------------------------------
        // 4. Locate the bookmark and prepare to insert the extracted nodes.
        // -----------------------------------------------------------------
        Bookmark bookmark = destDoc.Range.Bookmarks["InsertHere"];
        if (bookmark == null)
            throw new InvalidOperationException("Bookmark 'InsertHere' was not found in the destination document.");

        // The bookmark resides inside a paragraph. We'll insert after that paragraph.
        Paragraph bookmarkParagraph = bookmark.BookmarkStart.ParentNode as Paragraph;
        if (bookmarkParagraph == null)
            throw new InvalidOperationException("Bookmark is not placed inside a paragraph.");

        // -----------------------------------------------------------------
        // 5. Import the extracted nodes into the destination document and insert them.
        // -----------------------------------------------------------------
        NodeImporter importer = new NodeImporter(loadedSource, destDoc, ImportFormatMode.KeepSourceFormatting);
        Node referenceNode = bookmarkParagraph; // start inserting after this node

        foreach (Node node in extractedNodes)
        {
            Node importedNode = importer.ImportNode(node, true);
            // Insert after the previously inserted node (or after the bookmark paragraph for the first node).
            referenceNode.ParentNode.InsertAfter(importedNode, referenceNode);
            referenceNode = importedNode; // update reference for the next insertion
        }

        // -----------------------------------------------------------------
        // 6. Save the result and verify the output file.
        // -----------------------------------------------------------------
        string resultPath = "result.docx";
        destDoc.Save(resultPath);

        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The result document was not created.");

        // Write a simple confirmation log.
        File.WriteAllText("execution.log", $"Extraction and insertion completed successfully. Output: {resultPath}");
    }
}

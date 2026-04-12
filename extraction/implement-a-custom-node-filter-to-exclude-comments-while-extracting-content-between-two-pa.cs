using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // -------------------------------------------------
        // Create a sample source document with a commented paragraph.
        // -------------------------------------------------
        Document sourceDoc = new Document();
        sourceDoc.RemoveAllChildren();

        Section srcSection = new Section(sourceDoc);
        sourceDoc.AppendChild(srcSection);

        Body srcBody = new Body(sourceDoc);
        srcSection.AppendChild(srcBody);

        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Paragraph 1 – plain text.
        builder.Writeln("First paragraph.");

        // Paragraph 2 – contains a comment.
        Paragraph paraWithComment = new Paragraph(sourceDoc);
        Comment comment = new Comment(sourceDoc)
        {
            Author = "Author",
            Initial = "AU",
            DateTime = DateTime.Now
        };
        comment.SetText("This is a comment.");

        // Link comment range start/end to the comment's Id.
        CommentRangeStart rangeStart = new CommentRangeStart(sourceDoc, comment.Id);
        CommentRangeEnd rangeEnd = new CommentRangeEnd(sourceDoc, comment.Id);

        // Assemble the paragraph: start marker, text run, end marker, comment node.
        paraWithComment.AppendChild(rangeStart);
        paraWithComment.AppendChild(new Run(sourceDoc, "Second paragraph with a comment."));
        paraWithComment.AppendChild(rangeEnd);
        paraWithComment.AppendChild(comment);
        srcBody.AppendChild(paraWithComment);

        // Paragraph 3 – plain text.
        builder.Writeln("Third paragraph.");

        // Paragraph 4 – plain text.
        builder.Writeln("Fourth paragraph.");

        // -------------------------------------------------
        // Define extraction boundaries (inclusive).
        // -------------------------------------------------
        Paragraph startParagraph = sourceDoc.FirstSection.Body.Paragraphs[1]; // second paragraph (with comment)
        Paragraph endParagraph = sourceDoc.FirstSection.Body.Paragraphs[2];   // third paragraph

        if (startParagraph == null || endParagraph == null)
            throw new InvalidOperationException("Start or end paragraph not found.");

        // -------------------------------------------------
        // Prepare destination document.
        // -------------------------------------------------
        Document destDoc = new Document();
        destDoc.RemoveAllChildren();

        Section destSection = new Section(destDoc);
        destDoc.AppendChild(destSection);

        Body destBody = new Body(destDoc);
        destSection.AppendChild(destBody);

        // -------------------------------------------------
        // Import the selected range using NodeImporter.
        // -------------------------------------------------
        int startIndex = sourceDoc.FirstSection.Body.Paragraphs.IndexOf(startParagraph);
        int endIndex = sourceDoc.FirstSection.Body.Paragraphs.IndexOf(endParagraph);

        if (startIndex < 0 || endIndex < 0 || startIndex > endIndex)
            throw new InvalidOperationException("Invalid paragraph range.");

        // NodeImporter handles style and list translation between documents.
        NodeImporter importer = new NodeImporter(sourceDoc, destDoc, ImportFormatMode.KeepSourceFormatting);

        for (int i = startIndex; i <= endIndex; i++)
        {
            Paragraph srcPara = sourceDoc.FirstSection.Body.Paragraphs[i];

            // Import the paragraph (deep clone) into the destination document.
            Node importedNode = importer.ImportNode(srcPara, true);
            Paragraph importedPara = (Paragraph)importedNode;

            // Remove comment-related nodes from the imported paragraph.
            RemoveNodesByType(importedPara, NodeType.Comment);
            RemoveNodesByType(importedPara, NodeType.CommentRangeStart);
            RemoveNodesByType(importedPara, NodeType.CommentRangeEnd);

            destBody.AppendChild(importedPara);
        }

        // -------------------------------------------------
        // Save and display the result.
        // -------------------------------------------------
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Extracted.docx");
        destDoc.Save(outputPath);

        Console.WriteLine("Extracted text (comments excluded):");
        Console.WriteLine(destDoc.GetText().Trim());
    }

    // Helper: removes all child nodes of a specific type from a composite node.
    private static void RemoveNodesByType(CompositeNode parent, NodeType type)
    {
        NodeCollection nodes = parent.GetChildNodes(type, true);
        List<Node> toRemove = new List<Node>();
        foreach (Node node in nodes)
            toRemove.Add(node);

        foreach (Node node in toRemove)
            node.Remove();
    }
}

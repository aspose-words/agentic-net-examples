using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // 1. Create a sample source document with comments.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);

        builder.Writeln("Paragraph 1 - introductory text.");

        // Paragraph 2 (start of extraction range) with a comment.
        builder.Writeln("Paragraph 2 - start of range.");
        Comment comment1 = new Comment(source, "Alice", "A", DateTime.Today);
        comment1.SetText("Comment on paragraph 2.");
        builder.CurrentParagraph.AppendChild(comment1);

        int commentId = comment1.Id;

        // Paragraph 3 (inside extraction range) with a commented text range.
        builder.Writeln("Paragraph 3 - middle of range.");
        builder.CurrentParagraph.AppendChild(new CommentRangeStart(source, commentId));
        builder.Writeln("Commented text inside paragraph 3.");
        builder.CurrentParagraph.AppendChild(new CommentRangeEnd(source, commentId));

        // Add a second comment (no range) to the same paragraph.
        Comment comment2 = new Comment(source, "Bob", "B", DateTime.Today);
        comment2.SetText("Another comment.");
        builder.CurrentParagraph.AppendChild(comment2);

        // Paragraph 4 (end of extraction range).
        builder.Writeln("Paragraph 4 - end of range.");

        // Paragraph 5 (after extraction range).
        builder.Writeln("Paragraph 5 - concluding text.");

        const string sourcePath = "source.docx";
        source.Save(sourcePath);

        // 2. Load the document for processing.
        Document loaded = new Document(sourcePath);

        Paragraph startParagraph = loaded.FirstSection.Body.Paragraphs[1]; // Paragraph 2
        Paragraph endParagraph = loaded.FirstSection.Body.Paragraphs[3];   // Paragraph 4

        if (startParagraph == null || endParagraph == null)
            throw new InvalidOperationException("Boundary paragraphs not found.");

        // 3. Prepare the result document (empty body).
        Document result = new Document();
        result.RemoveAllChildren();                     // clear the default section/paragraph
        Section resultSection = new Section(result);
        result.AppendChild(resultSection);
        Body resultBody = new Body(result);
        resultSection.AppendChild(resultBody);

        // 4. Clone paragraphs within the range, removing comments.
        int startIndex = loaded.FirstSection.Body.Paragraphs.IndexOf(startParagraph);
        int endIndex = loaded.FirstSection.Body.Paragraphs.IndexOf(endParagraph);
        if (startIndex < 0 || endIndex < 0 || startIndex > endIndex)
            throw new InvalidOperationException("Invalid paragraph boundaries.");

        // NodeImporter will handle importing nodes from the source document into the result document.
        NodeImporter importer = new NodeImporter(loaded, result, ImportFormatMode.KeepSourceFormatting);

        for (int i = startIndex; i <= endIndex; i++)
        {
            Paragraph srcPara = loaded.FirstSection.Body.Paragraphs[i];
            // Import the paragraph (deep clone) into the destination document.
            Paragraph importedPara = (Paragraph)importer.ImportNode(srcPara, true);
            // Remove comments from the imported paragraph.
            RemoveCommentsFromNode(importedPara);
            resultBody.AppendChild(importedPara);
        }

        // 5. Save the extracted content.
        const string resultPath = "extracted.docx";
        result.Save(resultPath);

        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The extracted document was not created.");
    }

    // Recursively removes comment nodes and comment range markers from a node tree.
    private static void RemoveCommentsFromNode(Node node)
    {
        if (node == null) return;
        if (!node.IsComposite) return;

        CompositeNode composite = (CompositeNode)node;

        // Take a snapshot of the children to avoid modifying the collection while iterating.
        Node[] children = new Node[composite.GetChildNodes(NodeType.Any, false).Count];
        int idx = 0;
        foreach (Node child in composite.GetChildNodes(NodeType.Any, false))
            children[idx++] = child;

        foreach (Node child in children)
        {
            if (child.NodeType == NodeType.Comment ||
                child.NodeType == NodeType.CommentRangeStart ||
                child.NodeType == NodeType.CommentRangeEnd)
            {
                child.Remove();
                continue;
            }

            if (child.IsComposite)
                RemoveCommentsFromNode(child);
        }
    }
}

using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a sample source document.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);

        builder.Writeln("Paragraph 1");
        builder.Writeln("Paragraph 2");

        // Add a comment anchored to the second paragraph.
        Comment comment = new Comment(source, "Author", "A", DateTime.Now);
        builder.CurrentParagraph.AppendChild(comment);
        // The comment text is placed inside a separate paragraph within the comment node.
        builder.MoveTo(comment.AppendChild(new Paragraph(source)));
        builder.Writeln("This is a comment.");

        // Continue building the document.
        builder.MoveToDocumentEnd();
        builder.Writeln("Paragraph 3");
        builder.Writeln("Paragraph 4");

        // Save the source document locally.
        const string sourcePath = "source.docx";
        source.Save(sourcePath);

        // Load the document for extraction.
        Document loaded = new Document(sourcePath);

        // Define the start and end paragraphs for the extraction range (inclusive).
        Paragraph startParagraph = loaded.FirstSection.Body.Paragraphs[1]; // "Paragraph 2"
        Paragraph endParagraph = loaded.FirstSection.Body.Paragraphs[3];   // "Paragraph 4"

        if (startParagraph == null || endParagraph == null)
            throw new InvalidOperationException("Boundary paragraphs not found.");

        // Prepare the result document.
        Document result = new Document();
        result.RemoveAllChildren();
        Section resultSection = new Section(result);
        result.AppendChild(resultSection);
        Body resultBody = new Body(result);
        resultSection.AppendChild(resultBody);

        // Iterate from the start paragraph to the end paragraph, cloning each paragraph,
        // removing comment nodes, and importing the cleaned node into the result document.
        bool withinRange = false;
        foreach (Paragraph para in loaded.FirstSection.Body.Paragraphs)
        {
            if (para == startParagraph)
                withinRange = true;

            if (withinRange)
            {
                // Deep clone the paragraph.
                Paragraph clonedPara = (Paragraph)para.Clone(true);
                // Remove comment nodes from the cloned paragraph.
                RemoveComments(clonedPara);
                // Import the cleaned paragraph into the result document.
                Node importedNode = result.ImportNode(clonedPara, true);
                resultBody.AppendChild(importedNode);
            }

            if (para == endParagraph)
                break;
        }

        // Save the extracted content.
        const string resultPath = "extracted.docx";
        result.Save(resultPath);

        // Verify that the output file was created.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The extracted document was not created.");
    }

    // Removes all comment nodes from the given paragraph (including nested comment structures).
    private static void RemoveComments(Paragraph paragraph)
    {
        // Collect comment nodes to avoid modifying the collection while iterating.
        var comments = paragraph.GetChildNodes(NodeType.Comment, true)
                                .Cast<Node>()
                                .ToList();

        foreach (Node commentNode in comments)
        {
            commentNode.Remove();
        }
    }
}

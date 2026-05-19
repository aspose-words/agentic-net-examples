using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with a paragraph and a comment.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Paragraph before the target range.
        builder.Writeln("Paragraph before target.");

        // Target paragraph – this will be the start boundary.
        builder.Writeln("Target paragraph.");

        // Insert a comment attached to a range of text.
        Comment comment = new Comment(doc, "Monitor", "M", DateTime.Now);
        Paragraph lastParagraph = doc.FirstSection.Body.LastParagraph;
        lastParagraph.AppendChild(new CommentRangeStart(doc, comment.Id));
        lastParagraph.AppendChild(new Run(doc, "Commented text."));
        lastParagraph.AppendChild(new CommentRangeEnd(doc, comment.Id));
        lastParagraph.AppendChild(comment);

        // Paragraph after the comment – this will be outside the extraction range.
        builder.Writeln("Paragraph after comment.");

        // Save the document locally.
        const string sourceFile = "sample.docx";
        doc.Save(sourceFile);

        // Load the document for extraction.
        Document loaded = new Document(sourceFile);

        // Locate the start paragraph (the second paragraph in the body).
        Paragraph startParagraph = loaded.FirstSection.Body.Paragraphs[1];
        if (startParagraph == null)
            throw new InvalidOperationException("Start paragraph not found.");

        // Locate the comment node.
        Comment commentNode = loaded.GetChildNodes(NodeType.Comment, true)[0] as Comment;
        if (commentNode == null)
            throw new InvalidOperationException("Comment node not found.");

        // Extract text between the start paragraph and the comment node.
        string extractedText = ExtractBetweenNodes(startParagraph, commentNode);

        // Log the extracted content to the monitoring system (console output).
        Console.WriteLine("Extracted content between paragraph and comment:");
        Console.WriteLine(extractedText.Trim());

        // Persist the extracted text to a file for verification.
        const string outputFile = "extracted.txt";
        File.WriteAllText(outputFile, extractedText);
        if (!File.Exists(outputFile))
            throw new InvalidOperationException("Extraction output file was not created.");
    }

    // Helper method to collect text from nodes that lie between two given nodes.
    private static string ExtractBetweenNodes(Node startNode, Node endNode)
    {
        StringBuilder sb = new StringBuilder();
        Node current = startNode.NextSibling;
        while (current != null && current != endNode)
        {
            sb.Append(current.GetText());
            current = current.NextSibling;
        }
        return sb.ToString();
    }
}

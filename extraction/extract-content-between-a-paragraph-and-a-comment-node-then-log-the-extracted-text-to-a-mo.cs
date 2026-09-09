using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample document with paragraphs and a comment.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Paragraph that will serve as the start marker.
        builder.Writeln("Paragraph A - Start");

        // Paragraph that lies between the start paragraph and the comment.
        builder.Writeln("Paragraph B - Between");

        // Create a comment node and attach it to the current paragraph (Paragraph B).
        Comment comment = new Comment(doc, "Monitor", "MT", DateTime.Now);
        comment.SetText("This is a sample comment.");
        builder.CurrentParagraph.AppendChild(comment);

        // Paragraph after the comment (not part of the extraction range).
        builder.Writeln("Paragraph C - After");

        // Save the source document (optional, for inspection).
        const string sourcePath = "source.docx";
        doc.Save(sourcePath);

        // Load the document to simulate a separate extraction step.
        Document loaded = new Document(sourcePath);

        // Locate the start paragraph (Paragraph A).
        Paragraph startParagraph = loaded.FirstSection.Body.Paragraphs[0];
        if (startParagraph == null)
            throw new InvalidOperationException("Start paragraph not found.");

        // Locate the comment node.
        Comment targetComment = loaded.GetChildNodes(NodeType.Comment, true)[0] as Comment;
        if (targetComment == null)
            throw new InvalidOperationException("Comment node not found.");

        // Collect all nodes that appear between the start paragraph and the comment node.
        List<Node> betweenNodes = new List<Node>();
        Node current = startParagraph.NextSibling;
        while (current != null && current != targetComment)
        {
            betweenNodes.Add(current);
            current = current.NextSibling;
        }

        // Extract text from the collected nodes.
        string extractedText = string.Empty;
        foreach (Node node in betweenNodes)
        {
            // Use GetText for block nodes; for inline nodes GetText also works.
            extractedText += node.GetText();
        }

        // Trim the result to remove trailing paragraph breaks.
        extractedText = extractedText.Trim();

        // Log the extracted text to the monitoring system (simulated via console output).
        Console.WriteLine("Extracted text between paragraph and comment:");
        Console.WriteLine(extractedText);

        // Additionally, write the extracted text to a deterministic file.
        const string outputPath = "extracted.txt";
        File.WriteAllText(outputPath, extractedText);

        // Validate that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the extracted text file.");
    }
}

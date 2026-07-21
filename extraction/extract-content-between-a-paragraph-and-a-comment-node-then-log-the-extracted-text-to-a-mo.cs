using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // -------------------------------------------------
        // Step 1: Create a sample document.
        // -------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Paragraph before the target range.
        builder.Writeln("Paragraph before target.");

        // Target paragraph – start boundary for extraction.
        builder.Writeln("Target paragraph.");

        // Paragraph that will be extracted (content between the target paragraph and the comment).
        builder.Writeln("Intermediate content to extract.");

        // Insert a comment after the intermediate paragraph.
        Comment comment = new Comment(sourceDoc)
        {
            Author = "Monitor",
            Initial = "M",
            DateTime = DateTime.Now
        };
        comment.SetText("This is a comment.");
        // The comment is attached to the current paragraph (the intermediate one).
        builder.CurrentParagraph.AppendChild(comment);

        // Paragraph after the comment.
        builder.Writeln("Paragraph after comment.");

        // Save the source document locally.
        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // -------------------------------------------------
        // Step 2: Load the document for extraction.
        // -------------------------------------------------
        Document loadedDoc = new Document(sourcePath);

        // Locate the target paragraph (the second paragraph in the body).
        Paragraph startParagraph = loadedDoc.FirstSection.Body.Paragraphs[1];
        if (startParagraph == null)
            throw new InvalidOperationException("Start paragraph not found.");

        // -------------------------------------------------
        // Step 3: Extract text between the start paragraph and the comment node.
        // -------------------------------------------------
        StringBuilder extractedBuilder = new StringBuilder();

        // Begin with the node immediately after the start paragraph.
        Node currentNode = startParagraph.NextSibling;

        // Walk forward until we encounter a paragraph that contains a comment.
        while (currentNode != null)
        {
            // If this node is a paragraph that holds a comment, stop.
            if (currentNode.NodeType == NodeType.Paragraph)
            {
                Paragraph para = (Paragraph)currentNode;
                if (para.GetChildNodes(NodeType.Comment, true).Count > 0)
                    break;
            }

            // Append the textual representation of the node.
            extractedBuilder.Append(currentNode.GetText());
            currentNode = currentNode.NextSibling;
        }

        string extractedText = extractedBuilder.ToString().Trim();
        if (string.IsNullOrEmpty(extractedText))
            throw new InvalidOperationException("No content was extracted between the paragraph and the comment.");

        // -------------------------------------------------
        // Step 4: Log the extracted text to a monitoring system.
        // -------------------------------------------------
        MonitoringSystem.Log(extractedText);
    }
}

// Simple monitoring system that records messages as JSON lines.
public static class MonitoringSystem
{
    private const string LogFile = "monitor.log";

    public static void Log(string message)
    {
        var entry = new
        {
            TimestampUtc = DateTime.UtcNow,
            Message = message
        };

        string json = JsonConvert.SerializeObject(entry);
        // Write to console for immediate visibility.
        Console.WriteLine(json);
        // Append to a log file.
        File.AppendAllText(LogFile, json + Environment.NewLine);
    }
}

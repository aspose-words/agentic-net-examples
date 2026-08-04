using System;
using System.IO;
using System.Text;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // ------------------------------------------------------------
        // Create a sample document containing paragraphs and a comment.
        // ------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Paragraph before target.");               // Paragraph 0
        builder.Writeln("Start paragraph.");                      // Paragraph 1 (start boundary)

        // Text that will be between the start paragraph and the comment.
        builder.Writeln("This is the first line of extracted text.");
        builder.Writeln("This is the second line of extracted text.");

        // Insert a comment anchored to a piece of text.
        Comment comment = new Comment(doc, "Alice", "A", DateTime.Now);
        comment.SetText("This is a comment.");

        // The comment range start/end are placed around the following run.
        builder.CurrentParagraph.AppendChild(new CommentRangeStart(doc, comment.Id));
        builder.Write("Commented text.");
        builder.CurrentParagraph.AppendChild(new CommentRangeEnd(doc, comment.Id));
        builder.CurrentParagraph.AppendChild(comment); // Append the comment node itself.

        // Additional paragraph after the comment.
        builder.Writeln("Paragraph after comment.");

        // Save the document locally.
        const string docPath = "sample.docx";
        doc.Save(docPath);

        // ------------------------------------------------------------
        // Load the document for extraction.
        // ------------------------------------------------------------
        Document loadedDoc = new Document(docPath);

        // Locate the start paragraph (index 1).
        Paragraph startParagraph = loadedDoc.FirstSection.Body.Paragraphs[1];
        if (startParagraph == null)
            throw new InvalidOperationException("Start paragraph not found.");

        // Locate the comment range start node to determine the end boundary.
        CommentRangeStart commentRangeStart = loadedDoc.GetChildNodes(NodeType.CommentRangeStart, true)[0] as CommentRangeStart;
        if (commentRangeStart == null)
            throw new InvalidOperationException("Comment range start not found.");

        Paragraph endParagraph = commentRangeStart.ParentNode as Paragraph;
        if (endParagraph == null)
            throw new InvalidOperationException("End paragraph (containing comment) not found.");

        // Extract text between the start and end paragraphs (exclusive).
        NodeCollection bodyParagraphs = loadedDoc.FirstSection.Body.Paragraphs;
        int startIndex = bodyParagraphs.IndexOf(startParagraph);
        int endIndex = bodyParagraphs.IndexOf(endParagraph);

        if (endIndex <= startIndex)
            throw new InvalidOperationException("Invalid paragraph boundaries for extraction.");

        StringBuilder extractedBuilder = new StringBuilder();
        for (int i = startIndex + 1; i < endIndex; i++)
        {
            Paragraph para = bodyParagraphs[i] as Paragraph;
            if (para != null)
                extractedBuilder.Append(para.GetText());
        }

        string extractedText = extractedBuilder.ToString().Trim();

        // ------------------------------------------------------------
        // Log the extracted text (simulated monitoring system).
        // ------------------------------------------------------------
        Console.WriteLine("=== Extracted Content ===");
        Console.WriteLine(extractedText);
        Console.WriteLine("=========================");

        const string logPath = "extracted.txt";
        File.WriteAllText(logPath, extractedText);

        // Validate that the log file was created.
        if (!File.Exists(logPath))
            throw new InvalidOperationException("Failed to create the extraction log file.");
    }
}

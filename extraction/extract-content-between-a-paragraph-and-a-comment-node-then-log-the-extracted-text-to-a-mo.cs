using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class ExtractBetweenParagraphAndComment
{
    public static void Main()
    {
        // Create a sample document with paragraphs and a comment.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("First paragraph.");
        builder.Writeln("Target paragraph."); // This will be the start boundary.

        // Create a comment attached to the current paragraph.
        Comment comment = new Comment(doc)
        {
            Author = "Tester",
            Initial = "T",
            DateTime = DateTime.Now
        };
        comment.SetText("This is a comment.");
        builder.CurrentParagraph.AppendChild(comment);

        builder.Writeln("Third paragraph."); // This will be after the comment.

        // Save the sample document.
        const string inputPath = "sample.docx";
        doc.Save(inputPath);

        // Load the document for extraction.
        Document loadedDoc = new Document(inputPath);

        // Locate the target paragraph (second paragraph in the body).
        Paragraph targetParagraph = loadedDoc.FirstSection.Body.Paragraphs[1];
        if (targetParagraph == null)
            throw new InvalidOperationException("Target paragraph not found.");

        // Locate the first comment node in the document.
        Comment commentNode = loadedDoc.GetChildNodes(NodeType.Comment, true)
                                      .OfType<Comment>()
                                      .FirstOrDefault();
        if (commentNode == null)
            throw new InvalidOperationException("Comment node not found.");

        // Extract text between the paragraph and the comment.
        // In this example the comment is attached to the target paragraph,
        // so we combine the paragraph text and the comment text.
        string paragraphText = targetParagraph.GetText().Trim();
        string commentText = commentNode.GetText().Trim();

        string extractedContent = $"Paragraph: \"{paragraphText}\" | Comment: \"{commentText}\"";

        // Log the extracted content to a monitoring system (simulated via console output).
        Console.WriteLine($"Monitoring Log: {extractedContent}");

        // Additionally, write the extracted content to a deterministic file for validation.
        const string outputPath = "extracted.txt";
        File.WriteAllText(outputPath, extractedContent);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the extracted output file.");
    }
}

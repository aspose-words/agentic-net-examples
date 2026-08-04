using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // -------------------------------------------------
        // Create a sample source document with a comment.
        // -------------------------------------------------
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);

        builder.Writeln("Paragraph 1");
        builder.Writeln("Paragraph 2"); // Start extraction here.

        // Add a comment attached to the next run.
        Comment comment = new Comment(source, "Alice", "A", DateTime.Today);
        builder.CurrentParagraph.AppendChild(comment);
        // The comment contains its own paragraph where we write the commented text.
        Paragraph commentParagraph = (Paragraph)comment.AppendChild(new Paragraph(source));
        builder.MoveTo(commentParagraph);
        builder.Write("Commented text.");
        builder.MoveToDocumentEnd(); // Return cursor after the comment.

        builder.Writeln("Paragraph 3");
        builder.Writeln("Paragraph 4"); // End extraction here.

        // Save the source document locally.
        const string sourcePath = "source.docx";
        source.Save(sourcePath);

        // -------------------------------------------------
        // Load the document for extraction.
        // -------------------------------------------------
        Document loaded = new Document(sourcePath);

        // Identify the start and end paragraphs (inclusive).
        // Paragraph indices are zero‑based.
        Paragraph startParagraph = loaded.FirstSection.Body.Paragraphs[1]; // "Paragraph 2"
        Paragraph endParagraph = loaded.FirstSection.Body.Paragraphs[3];   // "Paragraph 4"

        if (startParagraph == null || endParagraph == null)
            throw new InvalidOperationException("Boundary paragraphs not found.");

        // -------------------------------------------------
        // Prepare the result document.
        // -------------------------------------------------
        Document result = new Document();
        result.RemoveAllChildren();
        Section resultSection = new Section(result);
        result.AppendChild(resultSection);
        Body resultBody = new Body(result);
        resultSection.AppendChild(resultBody);

        // Use a NodeImporter to keep source formatting when cloning nodes.
        NodeImporter importer = new NodeImporter(loaded, result, ImportFormatMode.KeepSourceFormatting);

        // Walk from startParagraph to endParagraph, cloning block‑level nodes.
        Node current = startParagraph;
        while (current != null)
        {
            // Clone only Paragraph or Table nodes.
            if (current.NodeType == NodeType.Paragraph || current.NodeType == NodeType.Table)
            {
                Node importedNode = importer.ImportNode(current, true);

                // If the node is a paragraph, remove any comment‑related child nodes.
                if (importedNode is Paragraph para)
                {
                    foreach (Node commentNode in para.GetChildNodes(NodeType.Comment, true).ToList())
                        commentNode.Remove();

                    foreach (Node rangeStart in para.GetChildNodes(NodeType.CommentRangeStart, true).ToList())
                        rangeStart.Remove();

                    foreach (Node rangeEnd in para.GetChildNodes(NodeType.CommentRangeEnd, true).ToList())
                        rangeEnd.Remove();
                }

                resultBody.AppendChild(importedNode);
            }

            if (current == endParagraph)
                break;

            current = current.NextSibling;
        }

        // Save the extracted content.
        const string resultPath = "extracted.docx";
        result.Save(resultPath);

        // Validate that the output file was created.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("Extraction failed – output file not found.");

        Console.WriteLine("Extraction completed successfully.");
    }
}

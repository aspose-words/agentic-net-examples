using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // -------------------------------------------------
        // 1. Create a sample source document with a comment.
        // -------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        builder.Writeln("Paragraph 1");
        builder.Writeln("Paragraph 2");

        // Add a comment anchored to the second paragraph.
        Comment comment = new Comment(sourceDoc, "John Doe", "JD", DateTime.Today);
        builder.CurrentParagraph.AppendChild(comment);
        builder.MoveTo(comment.AppendChild(new Paragraph(sourceDoc)));
        builder.Writeln("This is a comment.");
        builder.MoveToDocumentEnd();

        builder.Writeln("Paragraph 3");

        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // -------------------------------------------------
        // 2. Load the document for extraction.
        // -------------------------------------------------
        Document loadedDoc = new Document(sourcePath);

        // Define the start and end paragraphs (inclusive).
        Paragraph startPara = loadedDoc.FirstSection.Body.Paragraphs[1]; // "Paragraph 2"
        Paragraph endPara = loadedDoc.FirstSection.Body.Paragraphs[2];   // "Paragraph 3"

        if (startPara == null || endPara == null)
            throw new InvalidOperationException("Boundary paragraphs not found.");

        // -------------------------------------------------
        // 3. Prepare the result document.
        // -------------------------------------------------
        Document resultDoc = new Document();
        resultDoc.RemoveAllChildren();

        Section resultSection = new Section(resultDoc);
        resultDoc.AppendChild(resultSection);

        Body resultBody = new Body(resultDoc);
        resultSection.AppendChild(resultBody);

        // -------------------------------------------------
        // 4. Import nodes between the boundaries, skipping comments.
        // -------------------------------------------------
        // NodeImporter handles cloning across documents while preserving formatting.
        NodeImporter importer = new NodeImporter(loadedDoc, resultDoc, ImportFormatMode.KeepSourceFormatting);

        Node currentNode = startPara;
        while (currentNode != null)
        {
            // Skip comment nodes and their range markers.
            if (currentNode.NodeType != NodeType.Comment &&
                currentNode.NodeType != NodeType.CommentRangeStart &&
                currentNode.NodeType != NodeType.CommentRangeEnd)
            {
                // Import the node into the destination document.
                Node importedNode = importer.ImportNode(currentNode, true);
                resultBody.AppendChild(importedNode);
            }

            if (currentNode == endPara)
                break;

            currentNode = currentNode.NextSibling;
        }

        // -------------------------------------------------
        // 5. Save and verify the extracted document.
        // -------------------------------------------------
        const string resultPath = "extracted.docx";
        resultDoc.Save(resultPath);

        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The extracted document was not created.");

        Console.WriteLine("Extraction completed successfully.");
    }
}

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare input and output folders.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputPdfs");

        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create two sample Word documents with a bookmark named "Extract".
        CreateSampleDocument(Path.Combine(inputDir, "Sample1.docx"));
        CreateSampleDocument(Path.Combine(inputDir, "Sample2.docx"));

        // Process each document in the input folder.
        foreach (string filePath in Directory.GetFiles(inputDir, "*.docx"))
        {
            // Load the source document.
            Document sourceDoc = new Document(filePath);

            // Locate the bookmark that defines the extraction range.
            Bookmark extractBookmark = sourceDoc.Range.Bookmarks["Extract"];
            if (extractBookmark == null)
                throw new InvalidOperationException($"Bookmark 'Extract' not found in {filePath}.");

            // Create a new document that will hold the extracted content.
            Document extractedDoc = new Document();
            extractedDoc.RemoveAllChildren();

            // Build a minimal document structure: Section -> Body.
            Section section = new Section(extractedDoc);
            extractedDoc.AppendChild(section);
            Body body = new Body(extractedDoc);
            section.AppendChild(body);

            // Preserve the original formatting by cloning the nodes inside the bookmark.
            // The bookmark's Text property contains the plain text; to keep formatting we clone its child nodes.
            Node startNode = extractBookmark.BookmarkStart;
            Node endNode = extractBookmark.BookmarkEnd;

            // Collect all nodes that are descendants of the bookmark.
            Node currentNode = startNode;
            while (currentNode != null && currentNode != endNode)
            {
                // Move to the next node in document order.
                Node nextNode = currentNode.NextPreOrder(sourceDoc);
                // If the node is a block-level node (Paragraph or Table), clone and add it.
                if (currentNode.NodeType == NodeType.Paragraph || currentNode.NodeType == NodeType.Table)
                {
                    Node imported = extractedDoc.ImportNode(currentNode, true);
                    body.AppendChild(imported);
                }
                currentNode = nextNode;
            }

            // Save the extracted content as a PDF.
            string pdfFileName = Path.GetFileNameWithoutExtension(filePath) + ".pdf";
            string pdfPath = Path.Combine(outputDir, pdfFileName);
            extractedDoc.Save(pdfPath, SaveFormat.Pdf);

            // Verify that the PDF was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create PDF: {pdfPath}");
        }

        // All done.
        Console.WriteLine("Batch extraction completed successfully.");
    }

    // Helper method to create a sample document with a bookmark named "Extract".
    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Document header.");

        // Define the range to be extracted.
        builder.StartBookmark("Extract");
        builder.Writeln("First line of extractable content.");
        builder.Writeln("Second line of extractable content.");
        builder.EndBookmark("Extract");

        builder.Writeln("Document footer.");

        doc.Save(filePath);
    }
}

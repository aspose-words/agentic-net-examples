using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document with a run node followed by some text and a bookmark.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        builder.Writeln("Paragraph before the run.");
        builder.Write("RunStart");                     // The run we will locate.
        builder.Write(" TextBetween ");                // Content that should be extracted.
        builder.StartBookmark("TargetBookmark");       // Bookmark that follows the run.
        builder.Writeln("Content inside the bookmark.");
        builder.EndBookmark("TargetBookmark");
        builder.Writeln("Paragraph after the bookmark.");

        // Save the source document.
        const string sourcePath = "sample.docx";
        sourceDoc.Save(sourcePath);

        // Load the document for extraction.
        Document doc = new Document(sourcePath);

        // Locate the run node with the exact text "RunStart".
        Run runNode = null;
        foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
        {
            if (run.Text == "RunStart")
            {
                runNode = run;
                break;
            }
        }

        if (runNode == null)
            throw new InvalidOperationException("Run node with text 'RunStart' was not found.");

        // Locate the next bookmark after the run node.
        BookmarkStart nextBookmarkStart = null;
        Node current = runNode.NextSibling;
        while (current != null && nextBookmarkStart == null)
        {
            if (current.NodeType == NodeType.BookmarkStart)
                nextBookmarkStart = (BookmarkStart)current;
            else
                current = current.NextSibling;
        }

        if (nextBookmarkStart == null)
            throw new InvalidOperationException("No bookmark found after the specified run node.");

        // Extract the text between the run node and the bookmark start.
        StringBuilder extractedText = new StringBuilder();
        Node extractor = runNode.NextSibling;
        while (extractor != null && extractor != nextBookmarkStart)
        {
            extractedText.Append(extractor.GetText());
            extractor = extractor.NextSibling;
        }

        // Build a temporary document containing the extracted text.
        Document tempDoc = new Document();
        DocumentBuilder tempBuilder = new DocumentBuilder(tempDoc);
        tempBuilder.Writeln(extractedText.ToString().Trim());

        // Convert the extracted segment to HTML.
        string html = tempDoc.FirstSection.Body.FirstParagraph.ToString(SaveFormat.Html);

        // Save the HTML output.
        const string htmlPath = "extracted.html";
        File.WriteAllText(htmlPath, html);

        // Verify that the HTML file was created.
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("HTML extraction output was not created.");

        // Optional: display a short confirmation (no interactive input required).
        Console.WriteLine("Extraction completed successfully.");
    }
}

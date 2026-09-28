using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files.
        string destPath = "Destination.docx";
        string sourcePath = "Source.docx";
        string outputPath = "Result.docx";

        // ---------- Create destination document with a bookmark ----------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("This is the beginning of the destination document.");
        destBuilder.StartBookmark("InsertHere");
        destBuilder.Writeln("Bookmark location.");
        destBuilder.EndBookmark("InsertHere");
        destBuilder.Writeln("This is the end of the destination document.");
        destDoc.Save(destPath, SaveFormat.Docx);

        // ---------- Create source document with several paragraphs ----------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("First imported paragraph.");
        srcBuilder.Writeln("Second imported paragraph.");
        srcBuilder.Writeln("Third imported paragraph.");
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // ---------- Load documents (simulating a typical import scenario) ----------
        Document destination = new Document(destPath);
        Document source = new Document(sourcePath);

        // Locate the bookmark in the destination document.
        Bookmark bookmark = destination.Range.Bookmarks["InsertHere"]
            ?? throw new InvalidOperationException("Bookmark 'InsertHere' not found in destination document.");

        // The paragraph that contains the bookmark start.
        Paragraph bookmarkParagraph = bookmark.BookmarkStart.ParentNode as Paragraph
            ?? throw new InvalidOperationException("Bookmark start is not inside a paragraph.");

        // Parent node where paragraphs can be inserted (normally the Body of the section).
        CompositeNode? parentComposite = bookmarkParagraph.ParentNode as CompositeNode
            ?? throw new InvalidOperationException("Unable to locate a valid composite parent for insertion.");

        // Prepare the importer.
        NodeImporter importer = new NodeImporter(source, destination, ImportFormatMode.KeepSourceFormatting);

        // Reference node for insertion – start with the paragraph that holds the bookmark.
        Node referenceNode = bookmarkParagraph;

        // Import each paragraph from the source document and insert after the bookmark.
        NodeCollection sourceParagraphs = source.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph para in sourceParagraphs)
        {
            // Import the paragraph node.
            Node importedNode = importer.ImportNode(para, true);

            // Insert the imported paragraph after the current reference node.
            parentComposite.InsertAfter(importedNode, referenceNode);

            // Update the reference node so the next paragraph is inserted after the newly added one.
            referenceNode = importedNode;
        }

        // Save the resulting document.
        destination.Save(outputPath, SaveFormat.Docx);

        // Validate that the output file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The merged document was not saved correctly.", outputPath);
    }
}

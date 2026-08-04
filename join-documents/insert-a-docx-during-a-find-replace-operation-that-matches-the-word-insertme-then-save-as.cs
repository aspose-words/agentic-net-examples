using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create the main document that contains the placeholder word "INSERTME".
        Document mainDoc = new Document();
        DocumentBuilder mainBuilder = new DocumentBuilder(mainDoc);
        mainBuilder.Writeln("This is the main document.");
        mainBuilder.Writeln("INSERTME"); // Placeholder to be replaced.
        mainBuilder.Writeln("End of the main document.");
        mainDoc.Save("MainDoc.docx", SaveFormat.Docx);

        // Create the document whose content will be inserted.
        Document insertDoc = new Document();
        DocumentBuilder insertBuilder = new DocumentBuilder(insertDoc);
        insertBuilder.Writeln("This is the inserted document content.");
        insertDoc.Save("InsertDoc.docx", SaveFormat.Docx);

        // Load the main document (optional, we can reuse the same instance).
        Document srcDoc = new Document("MainDoc.docx");

        // Set up find‑replace options with a custom callback that inserts the document.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new InsertDocumentHandler()
        };

        // Perform the replace operation. The matched word will be removed and the document inserted.
        srcDoc.Range.Replace(new Regex("INSERTME"), string.Empty, options);

        // Save the resulting document.
        srcDoc.Save("Result.docx", SaveFormat.Docx);
    }

    // Callback that is invoked for each match found during the replace operation.
    private class InsertDocumentHandler : IReplacingCallback
    {
        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // Load the document to be inserted.
            Document docToInsert = new Document("InsertDoc.docx");

            // The match is inside a paragraph; insert after that paragraph.
            Paragraph placeholderParagraph = (Paragraph)args.MatchNode.ParentNode;
            InsertDocument(placeholderParagraph, docToInsert);

            // Remove the placeholder paragraph.
            placeholderParagraph.Remove();

            // Skip further processing of this match.
            return ReplaceAction.Skip;
        }

        // Inserts all nodes of docToInsert after the specified insertion destination.
        private static void InsertDocument(Node insertionDestination, Document docToInsert)
        {
            if (insertionDestination.NodeType != NodeType.Paragraph && insertionDestination.NodeType != NodeType.Table)
                throw new ArgumentException("The destination node must be a paragraph or a table.");

            CompositeNode dstStory = insertionDestination.ParentNode;

            NodeImporter importer = new NodeImporter(
                docToInsert, insertionDestination.Document, ImportFormatMode.KeepSourceFormatting);

            foreach (Section srcSection in docToInsert.Sections)
            {
                foreach (Node srcNode in srcSection.Body)
                {
                    // Skip the last empty paragraph of a section.
                    if (srcNode.NodeType == NodeType.Paragraph)
                    {
                        Paragraph para = (Paragraph)srcNode;
                        if (para.IsEndOfSection && !para.HasChildNodes)
                            continue;
                    }

                    Node newNode = importer.ImportNode(srcNode, true);
                    dstStory.InsertAfter(newNode, insertionDestination);
                    insertionDestination = newNode;
                }
            }
        }
    }
}

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create the destination document with a placeholder tag.
        Document destination = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destination);
        destBuilder.Writeln("Header of the main document.");
        destBuilder.Writeln("[INSERT_DOC]"); // Placeholder to be replaced.
        destBuilder.Writeln("Footer of the main document.");

        // Create the document that will be inserted.
        Document toInsert = new Document();
        DocumentBuilder insertBuilder = new DocumentBuilder(toInsert);
        insertBuilder.Writeln("This paragraph comes from the inserted document.");
        insertBuilder.Writeln("Another line from the inserted document.");

        // Set up find‑replace with a callback that inserts the document.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new InsertDocumentCallback(toInsert)
        };

        // Perform the replace operation. The placeholder is removed and the document is inserted.
        destination.Range.Replace(new Regex(@"\[INSERT_DOC\]"), string.Empty, options);

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.docx");
        destination.Save(outputPath, SaveFormat.Docx);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The merged document was not saved correctly.");
    }

    // Callback that inserts a document at the location of the matched placeholder.
    private class InsertDocumentCallback : IReplacingCallback
    {
        private readonly Document _documentToInsert;

        public InsertDocumentCallback(Document documentToInsert)
        {
            _documentToInsert = documentToInsert ?? throw new ArgumentNullException(nameof(documentToInsert));
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // The match is inside a paragraph; insert the document after this paragraph.
            Paragraph placeholderParagraph = (Paragraph)args.MatchNode.ParentNode;
            InsertDocument(placeholderParagraph, _documentToInsert);

            // Remove the paragraph that contained the placeholder tag.
            placeholderParagraph.Remove();

            // Skip further processing for this match.
            return ReplaceAction.Skip;
        }

        // Inserts all nodes of another document after the specified paragraph or table.
        private static void InsertDocument(Node insertionDestination, Document docToInsert)
        {
            if (insertionDestination.NodeType != NodeType.Paragraph && insertionDestination.NodeType != NodeType.Table)
                throw new ArgumentException("Insertion destination must be a paragraph or a table.");

            CompositeNode destinationStory = insertionDestination.ParentNode;

            // NodeImporter lives in Aspose.Words namespace; no extra using required.
            NodeImporter importer = new NodeImporter(docToInsert, insertionDestination.Document, ImportFormatMode.KeepSourceFormatting);

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

                    Node importedNode = importer.ImportNode(srcNode, true);
                    destinationStory.InsertAfter(importedNode, insertionDestination);
                    insertionDestination = importedNode;
                }
            }
        }
    }
}

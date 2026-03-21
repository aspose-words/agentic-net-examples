using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Create the destination document with a placeholder.
        Document mainDoc = new Document();
        DocumentBuilder mainBuilder = new DocumentBuilder(mainDoc);
        mainBuilder.Writeln("This is the beginning of the document.");
        mainBuilder.Writeln("INSERT_HERE");
        mainBuilder.Writeln("This is the end of the document.");

        // Create the document that will be inserted.
        Document insertDoc = new Document();
        DocumentBuilder insertBuilder = new DocumentBuilder(insertDoc);
        insertBuilder.Writeln("=== Inserted Content Start ===");
        insertBuilder.Writeln("Here is the content that will be inserted.");
        insertBuilder.Writeln("=== Inserted Content End ===");

        // Configure find/replace options with a custom callback that will insert another document.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new InsertDocHandler(insertDoc)
        };

        // Replace the placeholder; the callback will handle the insertion.
        mainDoc.Range.Replace(new Regex("INSERT_HERE"), string.Empty, options);

        // Save the modified document.
        mainDoc.Save("Result.docx");
        Console.WriteLine("Result.docx created successfully.");
    }

    // Callback that inserts a separate DOCX at the location of each match.
    private class InsertDocHandler : IReplacingCallback
    {
        private readonly Document _docToInsert;

        public InsertDocHandler(Document docToInsert)
        {
            _docToInsert = docToInsert ?? throw new ArgumentNullException(nameof(docToInsert));
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // The match node is inside a Run; its parent paragraph is the insertion point.
            Paragraph placeholderParagraph = (Paragraph)args.MatchNode.ParentNode;

            // Insert the entire sub‑document after the paragraph containing the match.
            InsertDocument(placeholderParagraph, _docToInsert);

            // Remove the paragraph that held the placeholder text.
            placeholderParagraph.Remove();

            // Skip further processing for this match.
            return ReplaceAction.Skip;
        }

        // Helper method that inserts all nodes of another document after a paragraph or table.
        private static void InsertDocument(Node insertionDestination, Document docToInsert)
        {
            if (insertionDestination.NodeType != NodeType.Paragraph && insertionDestination.NodeType != NodeType.Table)
                throw new ArgumentException("Destination must be a paragraph or table.", nameof(insertionDestination));

            CompositeNode destinationStory = insertionDestination.ParentNode;
            NodeImporter importer = new NodeImporter(docToInsert, insertionDestination.Document, ImportFormatMode.KeepSourceFormatting);

            // Insert nodes from each section of the source document.
            foreach (Section srcSection in docToInsert.Sections)
            {
                foreach (Node srcNode in srcSection.Body)
                {
                    // Skip the last empty paragraph of a section.
                    if (srcNode.NodeType == NodeType.Paragraph)
                    {
                        Paragraph p = (Paragraph)srcNode;
                        if (p.IsEndOfSection && !p.HasChildNodes)
                            continue;
                    }

                    Node newNode = importer.ImportNode(srcNode, true);
                    destinationStory.InsertAfter(newNode, insertionDestination);
                    insertionDestination = newNode;
                }
            }
        }
    }
}

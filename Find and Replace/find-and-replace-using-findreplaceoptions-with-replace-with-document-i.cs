using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace FindReplaceWithDocumentExample
{
    class Program
    {
        static void Main()
        {
            // Paths to the source documents and the output file.
            string dataDir = @"C:\Data\";
            string mainDocPath = Path.Combine(dataDir, "Document insertion destination.docx");
            string subDocPath = Path.Combine(dataDir, "Document.docx");
            string outputPath = Path.Combine(dataDir, "InsertDocumentAtReplaceResult.docx");

            // Load the main document (the one that contains the placeholder).
            Document mainDoc = new Document(mainDocPath);

            // Configure find/replace options with a custom callback.
            FindReplaceOptions options = new FindReplaceOptions
            {
                ReplacingCallback = new InsertDocumentAtReplaceHandler(subDocPath)
            };

            // Replace the placeholder "[MY_DOCUMENT]" with the contents of the sub‑document.
            mainDoc.Range.Replace(new Regex(@"\[MY_DOCUMENT\]"), string.Empty, options);

            // Save the modified document.
            mainDoc.Save(outputPath);
        }
    }

    // Custom callback that inserts another document at the location of each match.
    class InsertDocumentAtReplaceHandler : IReplacingCallback
    {
        private readonly string _subDocumentPath;

        public InsertDocumentAtReplaceHandler(string subDocumentPath)
        {
            _subDocumentPath = subDocumentPath;
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // Load the document that will be inserted.
            Document subDoc = new Document(_subDocumentPath);

            // The match is inside a paragraph; insert the sub‑document after that paragraph.
            Paragraph placeholderParagraph = (Paragraph)args.MatchNode.ParentNode;
            InsertDocument(placeholderParagraph, subDoc);

            // Remove the paragraph that contained the placeholder text.
            placeholderParagraph.Remove();

            // Skip further processing for this match because we have already handled it.
            return ReplaceAction.Skip;
        }

        // Inserts all nodes of another document after a given paragraph or table.
        private static void InsertDocument(Node insertionDestination, Document docToInsert)
        {
            if (insertionDestination.NodeType != NodeType.Paragraph && insertionDestination.NodeType != NodeType.Table)
                throw new ArgumentException("The destination node must be either a paragraph or a table.");

            CompositeNode dstStory = insertionDestination.ParentNode;

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

                    Node newNode = importer.ImportNode(srcNode, true);
                    dstStory.InsertAfter(newNode, insertionDestination);
                    insertionDestination = newNode;
                }
            }
        }
    }
}

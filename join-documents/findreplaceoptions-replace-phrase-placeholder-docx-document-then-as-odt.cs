using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Saving;

namespace AsposeWordsReplaceWithDocument
{
    // Implements a callback that inserts another document at the location of the match.
    class InsertDocumentHandler : IReplacingCallback
    {
        private readonly Document _documentToInsert;

        public InsertDocumentHandler(Document documentToInsert)
        {
            _documentToInsert = documentToInsert;
        }

        // Called for each match found by the Replace method.
        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // The match is expected to be a whole placeholder token inside a paragraph.
            // Insert the document after that paragraph and then remove the placeholder paragraph.
            var placeholderParagraph = (Paragraph)args.MatchNode.ParentNode;
            InsertDocument(placeholderParagraph, _documentToInsert);
            placeholderParagraph.Remove();

            // Skip the default replacement because we have already handled it.
            return ReplaceAction.Skip;
        }

        // Inserts all nodes of the source document after the specified insertion destination.
        private static void InsertDocument(Node insertionDestination, Document docToInsert)
        {
            // The destination must be a paragraph or a table.
            if (insertionDestination.NodeType != NodeType.Paragraph && insertionDestination.NodeType != NodeType.Table)
                throw new ArgumentException("The destination node must be either a paragraph or a table.");

            // The story (body) that contains the destination node.
            CompositeNode dstStory = insertionDestination.ParentNode;

            // Importer handles style and formatting preservation.
            NodeImporter importer = new NodeImporter(docToInsert, insertionDestination.Document, ImportFormatMode.KeepSourceFormatting);

            // Iterate through all sections and their body nodes of the source document.
            foreach (Section srcSection in docToInsert.Sections)
            {
                foreach (Node srcNode in srcSection.Body)
                {
                    // Skip the final empty paragraph of a section.
                    if (srcNode.NodeType == NodeType.Paragraph)
                    {
                        Paragraph para = (Paragraph)srcNode;
                        if (para.IsEndOfSection && !para.HasChildNodes)
                            continue;
                    }

                    // Import the node into the destination document.
                    Node newNode = importer.ImportNode(srcNode, true);

                    // Insert the imported node after the current insertion point.
                    dstStory.InsertAfter(newNode, insertionDestination);
                    insertionDestination = newNode;
                }
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Paths to the source documents.
            string mainDocPath = "MainDocument.docx";      // Contains the placeholder token [PLACEHOLDER]
            string insertDocPath = "InsertDocument.docx"; // Document whose content will replace the placeholder

            // Load the main document (the one that contains the placeholder).
            Document mainDoc = new Document(mainDocPath);

            // Load the document that will be inserted.
            Document insertDoc = new Document(insertDocPath);

            // Configure the find/replace operation.
            FindReplaceOptions options = new FindReplaceOptions
            {
                // Use a callback to perform a custom replacement (inserting a whole document).
                ReplacingCallback = new InsertDocumentHandler(insertDoc)
            };

            // Perform the replace. The pattern looks for the literal token [PLACEHOLDER].
            // The replacement string is empty because the callback handles insertion.
            mainDoc.Range.Replace(new Regex(@"\[PLACEHOLDER\]"), string.Empty, options);

            // Save the resulting document as ODT using OdtSaveOptions.
            OdtSaveOptions odtOptions = new OdtSaveOptions();
            mainDoc.Save("ResultDocument.odt", odtOptions);
        }
    }
}

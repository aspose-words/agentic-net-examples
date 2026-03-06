using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Saving;

namespace AsposeWordsDynamicInsert
{
    // Callback that replaces a placeholder with the contents of another document.
    class InsertDocumentHandler : IReplacingCallback
    {
        private readonly Document _docToInsert;

        public InsertDocumentHandler(Document docToInsert)
        {
            _docToInsert = docToInsert;
        }

        // This method is called for each match of the placeholder.
        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // The placeholder is inside a Run node; its parent is a Paragraph.
            Paragraph paragraph = (Paragraph)args.MatchNode.ParentNode;

            // Insert the document after the paragraph that contains the placeholder.
            InsertDocumentAfterParagraph(paragraph, _docToInsert);

            // Remove the paragraph that held the placeholder.
            paragraph.Remove();

            // Skip the default replace operation because we already handled it.
            return ReplaceAction.Skip;
        }

        // Inserts all nodes of sourceDoc after the specified paragraph.
        private static void InsertDocumentAfterParagraph(Paragraph insertionPoint, Document sourceDoc)
        {
            // The story (body) that will receive the new nodes.
            CompositeNode dstStory = insertionPoint.ParentNode;

            // Importer handles style and formatting conflicts.
            NodeImporter importer = new NodeImporter(sourceDoc, insertionPoint.Document, ImportFormatMode.KeepSourceFormatting);

            // Walk through each section and each node in the source document.
            foreach (Section srcSection in sourceDoc.Sections)
            {
                foreach (Node srcNode in srcSection.Body)
                {
                    // Skip the final empty paragraph that Word adds to each section.
                    if (srcNode.NodeType == NodeType.Paragraph)
                    {
                        Paragraph para = (Paragraph)srcNode;
                        if (para.IsEndOfSection && !para.HasChildNodes)
                            continue;
                    }

                    // Import the node into the destination document.
                    Node newNode = importer.ImportNode(srcNode, true);

                    // Insert after the current insertion point and advance the pointer.
                    dstStory.InsertAfter(newNode, insertionPoint);
                    insertionPoint = (Paragraph)newNode;
                }
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Load the RTF template that contains a placeholder like [INSERT_DOC].
            Document mainDoc = new Document("Template.rtf");

            // Load the document that we want to insert dynamically.
            Document docToInsert = new Document("InsertContent.docx");

            // Set up find-and-replace options with our custom callback.
            FindReplaceOptions options = new FindReplaceOptions();
            options.ReplacingCallback = new InsertDocumentHandler(docToInsert);

            // Replace the placeholder text with the contents of docToInsert.
            // The placeholder can be any unique string; here we use "[INSERT_DOC]".
            mainDoc.Range.Replace("[INSERT_DOC]", string.Empty, options);

            // Save the result back to RTF using custom save options if needed.
            RtfSaveOptions saveOptions = new RtfSaveOptions
            {
                // Example: make the file smaller (optional).
                ExportCompactSize = true,
                // Keep the generator name (default true).
                ExportGeneratorName = true
            };

            mainDoc.Save("Result.rtf", saveOptions);
        }
    }
}

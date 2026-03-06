using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

// Load the WORDML (WordprocessingML) template, insert a sub‑document at a placeholder, and save the result.
class DynamicDocumentInsertion
{
    // Path to the main template (WORDML format, e.g., .xml or .docx that contains the placeholder).
    private const string TemplatePath = @"C:\Docs\Template.xml";

    // Path to the document that will be inserted dynamically.
    private const string SubDocumentPath = @"C:\Docs\SectionToInsert.docx";

    // Path where the final document will be saved.
    private const string OutputPath = @"C:\Docs\Result.docx";

    static void Main()
    {
        // 1. Load the template document.
        Document mainDoc = new Document(TemplatePath);

        // 2. Prepare a find‑replace operation that looks for a unique placeholder token.
        //    In this example the placeholder is [[INSERT_DOC]].
        FindReplaceOptions replaceOptions = new FindReplaceOptions
        {
            // The callback will perform the actual insertion.
            ReplacingCallback = new InsertDocumentCallback(SubDocumentPath)
        };

        // Perform the replace – the placeholder text itself is replaced with an empty string
        // because the callback inserts the sub‑document manually.
        mainDoc.Range.Replace(new Regex(@"\[\[INSERT_DOC\]\]"), string.Empty, replaceOptions);

        // 3. Save the populated document.
        mainDoc.Save(OutputPath);
    }

    // Callback that is invoked for each placeholder match.
    private class InsertDocumentCallback : IReplacingCallback
    {
        private readonly string _docToInsertPath;

        public InsertDocumentCallback(string docToInsertPath)
        {
            _docToInsertPath = docToInsertPath;
        }

        // This method is called for every match found by the Replace operation.
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Load the document that should be inserted.
            Document subDoc = new Document(_docToInsertPath);

            // The match node is a Run inside a Paragraph that contains the placeholder.
            // Insert the sub‑document after that paragraph.
            Paragraph placeholderParagraph = (Paragraph)args.MatchNode.ParentNode;

            InsertDocumentAfterParagraph(placeholderParagraph, subDoc);

            // Remove the paragraph that held the placeholder token.
            placeholderParagraph.Remove();

            // Skip the default replace action because we have already handled the insertion.
            return ReplaceAction.Skip;
        }

        // Inserts all nodes of 'docToInsert' after the specified paragraph.
        private static void InsertDocumentAfterParagraph(Paragraph insertionPoint, Document docToInsert)
        {
            // The destination must be a paragraph or a table.
            if (insertionPoint.NodeType != NodeType.Paragraph && insertionPoint.NodeType != NodeType.Table)
                throw new ArgumentException("Insertion point must be a paragraph or a table.");

            // Parent node where new content will be added.
            CompositeNode destinationParent = insertionPoint.ParentNode;

            // Use NodeImporter for efficient style and list handling.
            NodeImporter importer = new NodeImporter(docToInsert, insertionPoint.Document, ImportFormatMode.KeepSourceFormatting);

            // Keep a reference to the node after which we will insert new content.
            Node currentInsertAfter = insertionPoint;

            // Iterate through all block‑level nodes of the source document.
            foreach (Section srcSection in docToInsert.Sections)
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
                    Node importedNode = importer.ImportNode(srcNode, true);

                    // Insert the imported node after the current insertion point.
                    destinationParent.InsertAfter(importedNode, currentInsertAfter);
                    currentInsertAfter = importedNode; // Move the insertion point forward.
                }
            }
        }
    }
}

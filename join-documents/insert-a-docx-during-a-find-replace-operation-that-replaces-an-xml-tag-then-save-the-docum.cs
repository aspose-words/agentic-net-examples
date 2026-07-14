using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files.
        const string destPath = "Destination.docx";
        const string subPath = "SubDocument.docx";
        const string outputPath = "MergedOutput.docx";

        // Create the destination document with a placeholder tag.
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("This is the main document.");
        destBuilder.Writeln("[INSERT_DOC]"); // Placeholder to be replaced.
        destDoc.Save(destPath, SaveFormat.Docx);

        // Create the document that will be inserted.
        Document subDoc = new Document();
        DocumentBuilder subBuilder = new DocumentBuilder(subDoc);
        subBuilder.Writeln("This is the inserted document content.");
        subDoc.Save(subPath, SaveFormat.Docx);

        // Set up find‑replace with a callback that inserts the sub‑document.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new InsertDocumentHandler(subPath)
        };

        // Perform the replace; the placeholder will be removed and the sub‑document inserted.
        destDoc.Range.Replace(new Regex(@"\[INSERT_DOC\]"), string.Empty, options);

        // Save the merged result.
        destDoc.Save(outputPath, SaveFormat.Docx);

        // Simple validation to ensure the output file was created and contains expected text.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The merged output file was not created.");

        Document result = new Document(outputPath);
        string resultText = result.GetText();

        if (!resultText.Contains("This is the main document.") ||
            !resultText.Contains("This is the inserted document content."))
            throw new InvalidOperationException("The merged document does not contain expected content.");
    }

    // Callback that inserts a document at the location of the matched placeholder.
    private class InsertDocumentHandler : IReplacingCallback
    {
        private readonly string _subDocumentPath;

        public InsertDocumentHandler(string subDocumentPath)
        {
            _subDocumentPath = subDocumentPath;
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // Load the document to be inserted.
            Document subDoc = new Document(_subDocumentPath);

            // The placeholder resides inside a paragraph; insert after that paragraph.
            Paragraph placeholderParagraph = (Paragraph)args.MatchNode.ParentNode;
            InsertDocument(placeholderParagraph, subDoc);

            // Remove the placeholder paragraph.
            placeholderParagraph.Remove();

            // Skip further processing of this match.
            return ReplaceAction.Skip;
        }

        // Inserts all nodes of the source document after the specified paragraph.
        private static void InsertDocument(Node insertionDestination, Document docToInsert)
        {
            if (insertionDestination.NodeType != NodeType.Paragraph && insertionDestination.NodeType != NodeType.Table)
                throw new ArgumentException("Insertion destination must be a paragraph or table.");

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

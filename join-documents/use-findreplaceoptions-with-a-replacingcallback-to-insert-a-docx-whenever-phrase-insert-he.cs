using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create the main document that contains the placeholder text.
        Document mainDoc = new Document();
        DocumentBuilder mainBuilder = new DocumentBuilder(mainDoc);
        mainBuilder.Writeln("This is the main document.");
        mainBuilder.Writeln("INSERT_HERE"); // Placeholder to be replaced.
        mainBuilder.Writeln("End of the main document.");

        // Create the document that will be inserted at the placeholder.
        string insertPath = Path.Combine(Directory.GetCurrentDirectory(), "Insert.docx");
        Document insertDoc = new Document();
        DocumentBuilder insertBuilder = new DocumentBuilder(insertDoc);
        insertBuilder.Writeln("This is the inserted document content.");
        insertDoc.Save(insertPath);

        // Configure FindReplaceOptions with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new InsertDocumentAtReplaceHandler(insertPath)
        };

        // Perform the replace operation. The placeholder text is removed,
        // and the content of Insert.docx is inserted at its location.
        mainDoc.Range.Replace(new Regex("INSERT_HERE"), "", options);

        // Save the resulting document.
        string resultPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.docx");
        mainDoc.Save(resultPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The merged document was not saved correctly.");

        // Optional: verify that the inserted text is present.
        Document resultDoc = new Document(resultPath);
        string resultText = resultDoc.GetText();
        if (!resultText.Contains("This is the inserted document content."))
            throw new InvalidOperationException("The inserted document content was not found in the result.");
    }

    // Callback that inserts a document at each match of the placeholder.
    private class InsertDocumentAtReplaceHandler : IReplacingCallback
    {
        private readonly string _insertPath;

        public InsertDocumentAtReplaceHandler(string insertPath) => _insertPath = insertPath;

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // Load the document to be inserted.
            Document subDoc = new Document(_insertPath);

            // The placeholder resides in a paragraph; insert after that paragraph.
            Paragraph placeholderParagraph = (Paragraph)args.MatchNode.ParentNode;
            InsertDocument(placeholderParagraph, subDoc);

            // Remove the paragraph that contained the placeholder text.
            placeholderParagraph.Remove();

            // Skip further processing of this match.
            return ReplaceAction.Skip;
        }

        // Inserts all nodes of the source document after the specified paragraph.
        private static void InsertDocument(Node insertionDestination, Document docToInsert)
        {
            if (insertionDestination.NodeType != NodeType.Paragraph && insertionDestination.NodeType != NodeType.Table)
                throw new ArgumentException("The destination node must be a paragraph or a table.");

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

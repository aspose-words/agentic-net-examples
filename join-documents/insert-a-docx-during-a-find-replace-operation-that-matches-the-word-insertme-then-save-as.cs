using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Paths for the sample documents.
        const string mainDocPath = "MainDoc.docx";
        const string insertDocPath = "InsertDoc.docx";
        const string resultDocPath = "Result.docx";

        // Create the main document containing the placeholder word "INSERTME".
        var mainDoc = new Document();
        var mainBuilder = new DocumentBuilder(mainDoc);
        mainBuilder.Writeln("This is the start of the main document.");
        mainBuilder.Writeln("INSERTME"); // Placeholder to be replaced.
        mainBuilder.Writeln("This is the end of the main document.");
        mainDoc.Save(mainDocPath, SaveFormat.Docx);

        // Create the document that will be inserted.
        var insertDoc = new Document();
        var insertBuilder = new DocumentBuilder(insertDoc);
        insertBuilder.Writeln("Inserted content line 1.");
        insertBuilder.Writeln("Inserted content line 2.");
        insertDoc.Save(insertDocPath, SaveFormat.Docx);

        // Load the main document for processing.
        var doc = new Document(mainDocPath);

        // Set up find‑replace options with a custom callback.
        var options = new FindReplaceOptions
        {
            ReplacingCallback = new InsertDocumentAtReplaceHandler(insertDocPath)
        };

        // Perform the replace operation; the callback will insert the document.
        doc.Range.Replace(new Regex("INSERTME"), string.Empty, options);

        // Save the resulting document.
        doc.Save(resultDocPath, SaveFormat.Docx);

        // Validate that the output file was created.
        if (!File.Exists(resultDocPath))
            throw new InvalidOperationException("The result document was not created.");
    }

    // Callback that inserts a document at the location of each match.
    private class InsertDocumentAtReplaceHandler : IReplacingCallback
    {
        private readonly string _documentPath;

        public InsertDocumentAtReplaceHandler(string documentPath)
        {
            _documentPath = documentPath ?? throw new ArgumentNullException(nameof(documentPath));
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // Load the document to be inserted.
            var subDoc = new Document(_documentPath);

            // The match is inside a paragraph; insert after that paragraph.
            var paragraph = (Paragraph)args.MatchNode.ParentNode;
            InsertDocument(paragraph, subDoc);

            // Remove the placeholder paragraph.
            paragraph.Remove();

            // Skip further processing of this match.
            return ReplaceAction.Skip;
        }

        // Inserts all nodes of another document after the specified paragraph.
        private static void InsertDocument(Node insertionDestination, Document docToInsert)
        {
            if (insertionDestination.NodeType != NodeType.Paragraph && insertionDestination.NodeType != NodeType.Table)
                throw new ArgumentException("The destination node must be a paragraph or a table.");

            var destinationStory = insertionDestination.ParentNode;
            var importer = new NodeImporter(docToInsert, insertionDestination.Document, ImportFormatMode.KeepSourceFormatting);

            foreach (Section srcSection in docToInsert.Sections)
            {
                foreach (Node srcNode in srcSection.Body)
                {
                    // Skip the last empty paragraph of a section.
                    if (srcNode.NodeType == NodeType.Paragraph)
                    {
                        var para = (Paragraph)srcNode;
                        if (para.IsEndOfSection && !para.HasChildNodes)
                            continue;
                    }

                    var newNode = importer.ImportNode(srcNode, true);
                    destinationStory.InsertAfter(newNode, insertionDestination);
                    insertionDestination = newNode;
                }
            }
        }
    }
}

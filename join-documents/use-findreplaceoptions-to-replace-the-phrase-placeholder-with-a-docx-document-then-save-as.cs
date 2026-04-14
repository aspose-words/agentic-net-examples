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
        // Create the main document that contains the placeholder.
        Document mainDoc = new Document();
        DocumentBuilder mainBuilder = new DocumentBuilder(mainDoc);
        mainBuilder.Writeln("This is the beginning of the document.");
        mainBuilder.Writeln("PLACEHOLDER");
        mainBuilder.Writeln("This is the end of the document.");

        // Save the main document temporarily (optional, just to illustrate creation).
        const string mainPath = "MainDocument.docx";
        mainDoc.Save(mainPath, SaveFormat.Docx);

        // Create the document that will replace the placeholder.
        Document insertDoc = new Document();
        DocumentBuilder insertBuilder = new DocumentBuilder(insertDoc);
        insertBuilder.Writeln("=== Inserted Document Start ===");
        insertBuilder.Writeln("This content comes from the inserted document.");
        insertBuilder.Writeln("=== Inserted Document End ===");

        // Save the insert document temporarily (optional).
        const string insertPath = "InsertDocument.docx";
        insertDoc.Save(insertPath, SaveFormat.Docx);

        // Load the documents from the saved files to simulate real-world usage.
        Document main = new Document(mainPath);
        Document sub = new Document(insertPath);

        // Set up FindReplaceOptions with a custom callback that inserts the document.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new InsertDocumentAtReplaceHandler(sub);

        // Perform the replace operation. The placeholder text will be removed and the document inserted.
        main.Range.Replace(new Regex("PLACEHOLDER"), string.Empty, options);

        // Save the resulting document as ODT.
        const string outputPath = "ResultDocument.odt";
        main.Save(outputPath, SaveFormat.Odt);

        // Validate that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The ODT file was not created.");

        // Optional validation: ensure the inserted text is present.
        string resultText = main.GetText();
        if (!resultText.Contains("Inserted Document Start"))
            throw new InvalidOperationException("The inserted document content was not found in the result.");

        // Clean up temporary files (optional).
        File.Delete(mainPath);
        File.Delete(insertPath);
    }

    // Callback that inserts a document at the location of each match.
    private class InsertDocumentAtReplaceHandler : IReplacingCallback
    {
        private readonly Document _docToInsert;

        public InsertDocumentAtReplaceHandler(Document docToInsert)
        {
            _docToInsert = docToInsert ?? throw new ArgumentNullException(nameof(docToInsert));
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // The match is expected to be inside a paragraph.
            Paragraph placeholderParagraph = args.MatchNode.ParentNode as Paragraph;
            if (placeholderParagraph == null)
                return ReplaceAction.Skip;

            // Insert the document after the placeholder paragraph.
            InsertDocument(placeholderParagraph, _docToInsert);

            // Remove the placeholder paragraph.
            placeholderParagraph.Remove();

            // Skip further processing for this match.
            return ReplaceAction.Skip;
        }

        // Inserts all nodes from the source document after the specified paragraph.
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

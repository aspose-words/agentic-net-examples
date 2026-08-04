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
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the documents.
        string mainDocPath = Path.Combine(outputDir, "MainDocument.docx");
        string subDocPath = Path.Combine(outputDir, "SubDocument.docx");
        string resultPath = Path.Combine(outputDir, "Result.docx");

        // Create the main document containing a placeholder XML tag.
        Document mainDoc = new Document();
        DocumentBuilder mainBuilder = new DocumentBuilder(mainDoc);
        mainBuilder.Writeln("This is the main document.");
        // Placeholder tag that will be replaced.
        mainBuilder.Writeln("<InsertDoc/>");
        mainBuilder.Writeln("End of the main document.");
        mainDoc.Save(mainDocPath);

        // Create the document that will be inserted.
        Document subDoc = new Document();
        DocumentBuilder subBuilder = new DocumentBuilder(subDoc);
        subBuilder.Writeln("=== Inserted Document Start ===");
        subBuilder.Writeln("This content comes from the inserted document.");
        subBuilder.Writeln("=== Inserted Document End ===");
        subDoc.Save(subDocPath);

        // Load the main document for processing.
        Document processingDoc = new Document(mainDocPath);

        // Set up find‑replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new InsertDocumentAtReplaceHandler(subDocPath)
        };

        // Perform the replace operation on the placeholder tag.
        processingDoc.Range.Replace(new Regex("<InsertDoc/>"), "", options);

        // Save the final document.
        processingDoc.Save(resultPath, SaveFormat.Docx);

        // Validate that the result file was created.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The result document was not created.");

        // Verify that the inserted content exists.
        string resultText = new Document(resultPath).GetText();
        if (!resultText.Contains("Inserted Document Start"))
            throw new InvalidOperationException("The inserted document content was not found in the result.");
    }

    // Callback that inserts a document at the location of the matched placeholder.
    private class InsertDocumentAtReplaceHandler : IReplacingCallback
    {
        private readonly string _subDocPath;

        public InsertDocumentAtReplaceHandler(string subDocPath)
        {
            _subDocPath = subDocPath;
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // Load the document to be inserted.
            Document subDoc = new Document(_subDocPath);

            // The match is inside a paragraph; get that paragraph.
            Paragraph placeholderParagraph = args.MatchNode?.ParentNode as Paragraph;
            if (placeholderParagraph == null)
                return ReplaceAction.Skip;

            // Insert the sub‑document after the placeholder paragraph.
            InsertDocument(placeholderParagraph, subDoc);

            // Remove the placeholder paragraph.
            placeholderParagraph.Remove();

            // Skip further processing of this match.
            return ReplaceAction.Skip;
        }

        // Inserts all nodes of docToInsert after the specified paragraph or table.
        private static void InsertDocument(Node insertionDestination, Document docToInsert)
        {
            if (insertionDestination.NodeType != NodeType.Paragraph && insertionDestination.NodeType != NodeType.Table)
                throw new ArgumentException("Insertion destination must be a paragraph or table.");

            CompositeNode destinationStory = insertionDestination.ParentNode;

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
                    destinationStory.InsertAfter(newNode, insertionDestination);
                    insertionDestination = newNode;
                }
            }
        }
    }
}

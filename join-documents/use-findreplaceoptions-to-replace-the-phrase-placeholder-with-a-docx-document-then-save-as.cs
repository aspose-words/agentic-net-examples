using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // Create the main document that contains the placeholder.
        Document mainDoc = new Document();
        DocumentBuilder mainBuilder = new DocumentBuilder(mainDoc);
        mainBuilder.Writeln("This is the main document.");
        mainBuilder.Writeln("PLACEHOLDER"); // The text to be replaced.
        mainBuilder.Writeln("End of the main document.");

        // Create the document that will replace the placeholder.
        Document insertDoc = new Document();
        DocumentBuilder insertBuilder = new DocumentBuilder(insertDoc);
        insertBuilder.Writeln("This text comes from the inserted document.");
        insertBuilder.Writeln("It can contain multiple paragraphs.");

        // Set up FindReplaceOptions with a custom callback that inserts the document.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new InsertDocumentHandler(insertDoc);

        // Perform the replace operation. The placeholder text is removed and the document is inserted.
        mainDoc.Range.Replace(new Regex("PLACEHOLDER"), string.Empty, options);

        // Save the resulting document as ODT.
        string outputPath = Path.Combine(outputDir, "Result.odt");
        mainDoc.Save(outputPath, SaveFormat.Odt);

        // Simple validation that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output ODT file was not created.");

        // Optionally, you could load the saved ODT to verify its contents.
        Document verifyDoc = new Document(outputPath);
        Console.WriteLine("Resulting document text:");
        Console.WriteLine(verifyDoc.GetText());
    }

    // Callback that inserts a document at the location of each match.
    private class InsertDocumentHandler : IReplacingCallback
    {
        private readonly Document _docToInsert;

        public InsertDocumentHandler(Document docToInsert)
        {
            _docToInsert = docToInsert ?? throw new ArgumentNullException(nameof(docToInsert));
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // The match node is a Run; its parent is a Paragraph.
            Paragraph placeholderParagraph = (Paragraph)args.MatchNode.ParentNode;

            // Insert the document after the paragraph that contains the placeholder.
            InsertDocument(placeholderParagraph, _docToInsert);

            // Remove the original placeholder paragraph.
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

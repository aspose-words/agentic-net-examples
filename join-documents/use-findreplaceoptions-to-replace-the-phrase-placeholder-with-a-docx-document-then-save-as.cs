using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Prepare working directory.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        Directory.CreateDirectory(workDir);

        // Paths for the source documents and the final output.
        string mainDocPath = Path.Combine(workDir, "Main.docx");
        string subDocPath = Path.Combine(workDir, "Insert.docx");
        string outputPath = Path.Combine(workDir, "Result.odt");

        // Create the main document containing the placeholder.
        Document mainDoc = new Document();
        DocumentBuilder mainBuilder = new DocumentBuilder(mainDoc);
        mainBuilder.Writeln("This is the start of the main document.");
        mainBuilder.Writeln("PLACEHOLDER"); // The text to be replaced.
        mainBuilder.Writeln("This is the end of the main document.");
        mainDoc.Save(mainDocPath, SaveFormat.Docx);

        // Create the document that will replace the placeholder.
        Document subDoc = new Document();
        DocumentBuilder subBuilder = new DocumentBuilder(subDoc);
        subBuilder.Writeln("Inserted line 1.");
        subBuilder.Writeln("Inserted line 2.");
        subDoc.Save(subDocPath, SaveFormat.Docx);

        // Load the main document for processing.
        Document loadedMain = new Document(mainDocPath);

        // Configure find‑replace to use a custom callback that inserts a document.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new InsertDocumentHandler(subDocPath)
        };

        // Replace the placeholder (exact word match) with the content of subDoc.
        loadedMain.Range.Replace(new Regex(@"\bPLACEHOLDER\b"), string.Empty, options);

        // Save the resulting document as ODT.
        loadedMain.Save(outputPath, SaveFormat.Odt);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output ODT file was not created.");

        // Optional: indicate success.
        Console.WriteLine("Document processed and saved to: " + outputPath);
    }

    // Callback that inserts a document at the location of each match.
    private class InsertDocumentHandler : IReplacingCallback
    {
        private readonly string _documentPath;

        public InsertDocumentHandler(string documentPath)
        {
            _documentPath = documentPath;
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // Load the document to be inserted.
            Document insertDoc = new Document(_documentPath);

            // The match is inside a paragraph; insert after that paragraph.
            Paragraph placeholderParagraph = (Paragraph)args.MatchNode.ParentNode;
            InsertDocument(placeholderParagraph, insertDoc);

            // Remove the placeholder paragraph.
            placeholderParagraph.Remove();

            // Skip the default replacement since we already handled it.
            return ReplaceAction.Skip;
        }

        // Inserts all nodes of insertDoc after the specified paragraph.
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

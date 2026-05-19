using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace InsertDocViaFindReplace
{
    // Callback that inserts a document when the placeholder is found.
    class InsertDocumentAtReplaceHandler : IReplacingCallback
    {
        private readonly string _sourcePath;

        public InsertDocumentAtReplaceHandler(string sourcePath)
        {
            _sourcePath = sourcePath;
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // Load the document that will be inserted.
            Document subDoc = new Document(_sourcePath);

            // The placeholder resides in a paragraph; insert after it.
            Paragraph para = (Paragraph)args.MatchNode.ParentNode;
            InsertDocument(para, subDoc);

            // Remove the placeholder paragraph.
            para.Remove();

            // Skip further processing of this match.
            return ReplaceAction.Skip;
        }

        // Inserts all nodes of another document after a paragraph or table.
        private static void InsertDocument(Node insertionDestination, Document docToInsert)
        {
            if (insertionDestination.NodeType != NodeType.Paragraph && insertionDestination.NodeType != NodeType.Table)
                throw new ArgumentException("The destination node must be either a paragraph or table.");

            CompositeNode dstStory = insertionDestination.ParentNode;

            NodeImporter importer = new NodeImporter(
                docToInsert, insertionDestination.Document, ImportFormatMode.KeepSourceFormatting);

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

    class Program
    {
        static void Main()
        {
            // Prepare a folder for temporary files.
            string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
            Directory.CreateDirectory(workDir);

            // Paths for the documents.
            string destinationPath = Path.Combine(workDir, "Destination.docx");
            string sourcePath = Path.Combine(workDir, "Source.docx");
            string resultPath = Path.Combine(workDir, "Result.docx");

            // Create the destination document containing the placeholder.
            Document destDoc = new Document();
            DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
            destBuilder.Writeln("Document start.");
            destBuilder.Writeln("INSERT_HERE"); // Placeholder to be replaced.
            destBuilder.Writeln("Document end.");
            destDoc.Save(destinationPath);

            // Create the source document that will be inserted.
            Document srcDoc = new Document();
            DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
            srcBuilder.Writeln("This is the inserted content.");
            srcDoc.Save(sourcePath);

            // Load the destination document for processing.
            Document mainDoc = new Document(destinationPath);

            // Configure find‑replace with the custom callback.
            FindReplaceOptions options = new FindReplaceOptions
            {
                ReplacingCallback = new InsertDocumentAtReplaceHandler(sourcePath)
            };

            // Perform the replace; the placeholder text is removed and the source document is inserted.
            mainDoc.Range.Replace(new Regex("INSERT_HERE"), "", options);

            // Save the final merged document.
            mainDoc.Save(resultPath);

            // Simple validation to ensure the file was created.
            if (!File.Exists(resultPath))
                throw new InvalidOperationException("The result document was not created.");

            // Optionally, you could output the path for verification (no interactive prompts required).
            Console.WriteLine($"Result document saved to: {resultPath}");
        }
    }
}

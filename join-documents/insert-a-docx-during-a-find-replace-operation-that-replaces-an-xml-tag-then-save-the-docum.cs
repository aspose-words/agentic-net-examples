using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace AsposeWordsJoinExample
{
    public class Program
    {
        public static void Main()
        {
            // Prepare output directory.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Create the main document that contains the placeholder tag.
            Document mainDoc = new Document();
            DocumentBuilder mainBuilder = new DocumentBuilder(mainDoc);
            mainBuilder.Writeln("This is the main document.");
            mainBuilder.Writeln("[INSERT_DOC]"); // Placeholder to be replaced.
            mainBuilder.Writeln("End of the main document.");

            // Create the document that will be inserted.
            Document subDoc = new Document();
            DocumentBuilder subBuilder = new DocumentBuilder(subDoc);
            subBuilder.Writeln("This is the inserted document.");
            string subDocPath = Path.Combine(outputDir, "SubDoc.docx");
            subDoc.Save(subDocPath, SaveFormat.Docx);

            // Perform find‑replace with a callback that inserts the sub‑document.
            FindReplaceOptions options = new FindReplaceOptions
            {
                ReplacingCallback = new InsertDocumentAtReplaceHandler(subDocPath)
            };
            mainDoc.Range.Replace(new Regex(@"\[INSERT_DOC\]"), string.Empty, options);

            // Save the merged result.
            string mergedPath = Path.Combine(outputDir, "Merged.docx");
            mainDoc.Save(mergedPath, SaveFormat.Docx);

            // Validate that the file was created.
            if (!File.Exists(mergedPath))
                throw new InvalidOperationException("The merged document was not created.");
        }

        // Callback that inserts a document at the location of the matched tag.
        private class InsertDocumentAtReplaceHandler : IReplacingCallback
        {
            private readonly string _documentPath;

            public InsertDocumentAtReplaceHandler(string documentPath)
            {
                _documentPath = documentPath;
            }

            ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
            {
                // Load the document to be inserted.
                Document docToInsert = new Document(_documentPath);

                // The match is inside a paragraph; insert after that paragraph.
                Paragraph paragraph = (Paragraph)args.MatchNode.ParentNode;
                InsertDocument(paragraph, docToInsert);

                // Remove the placeholder paragraph.
                paragraph.Remove();

                // Skip further processing of this match.
                return ReplaceAction.Skip;
            }

            // Inserts all nodes of docToInsert after the specified insertion destination.
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
}

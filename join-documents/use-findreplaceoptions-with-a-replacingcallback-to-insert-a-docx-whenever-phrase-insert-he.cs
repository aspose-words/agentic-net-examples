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
        // Prepare a folder for temporary files.
        string dataDir = Path.Combine(Environment.CurrentDirectory, "Data");
        Directory.CreateDirectory(dataDir);

        // Create the source document that will be inserted.
        string sourcePath = Path.Combine(dataDir, "Source.docx");
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("This is the inserted content.");
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // Create the destination document containing the placeholder phrase.
        string destPath = Path.Combine(dataDir, "Destination.docx");
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("Document start.");
        destBuilder.Writeln("INSERT_HERE"); // Placeholder to be replaced.
        destBuilder.Writeln("Document end.");
        destDoc.Save(destPath, SaveFormat.Docx);

        // Load the destination document for processing.
        Document mainDoc = new Document(destPath);

        // Configure find/replace to use a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new InsertDocumentHandler(sourcePath)
        };

        // Replace the placeholder with the content of the source document.
        mainDoc.Range.Replace(new Regex("INSERT_HERE"), string.Empty, options);

        // Save the merged result.
        string resultPath = Path.Combine(dataDir, "Result.docx");
        mainDoc.Save(resultPath, SaveFormat.Docx);

        // Simple validation that the output file was created.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The merged document was not saved correctly.");
    }

    // Callback that inserts a document at the location of each match.
    private class InsertDocumentHandler : IReplacingCallback
    {
        private readonly string _docToInsertPath;

        public InsertDocumentHandler(string docToInsertPath)
        {
            _docToInsertPath = docToInsertPath;
        }

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // Load the document to be inserted.
            Document subDoc = new Document(_docToInsertPath);

            // The match is inside a Run; its parent paragraph is the insertion point.
            Paragraph para = (Paragraph)args.MatchNode.ParentNode;

            // Insert the document after the placeholder paragraph.
            InsertDocument(para, subDoc);

            // Remove the placeholder paragraph.
            para.Remove();

            // Skip further processing of this match.
            return ReplaceAction.Skip;
        }

        // Inserts all nodes of a document after a paragraph or table.
        private static void InsertDocument(Node insertionDestination, Document docToInsert)
        {
            if (insertionDestination.NodeType != NodeType.Paragraph && insertionDestination.NodeType != NodeType.Table)
                throw new ArgumentException("The destination node must be a paragraph or table.");

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
                        Paragraph p = (Paragraph)srcNode;
                        if (p.IsEndOfSection && !p.HasChildNodes)
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

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace ImportCommentsExample
{
    class Program
    {
        static void Main()
        {
            // Use the current directory for all files so the example works everywhere.
            string baseDir = AppContext.BaseDirectory;
            string sourcePath = Path.Combine(baseDir, "CommentsSource.docx");
            string destinationPath = Path.Combine(baseDir, "NewDocument.docx");
            string outputPath = Path.Combine(baseDir, "NewDocumentWithComments.docx");

            // -----------------------------------------------------------------
            // Create a source document that contains a comment.
            // -----------------------------------------------------------------
            Document srcDoc = new Document();
            DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
            srcBuilder.Writeln("Paragraph in source document.");

            // Insert a comment that covers the next line (manual creation because StartComment/EndComment may be unavailable).
            Comment comment = new Comment(srcDoc, "John Doe", "JD", DateTime.Now);
            CommentRangeStart rangeStart = new CommentRangeStart(srcDoc, comment.Id);
            CommentRangeEnd rangeEnd = new CommentRangeEnd(srcDoc, comment.Id);

            srcBuilder.InsertNode(rangeStart);
            srcBuilder.Writeln("This text is commented.");
            srcBuilder.InsertNode(rangeEnd);
            srcBuilder.InsertNode(comment);

            srcDoc.Save(sourcePath);

            // -----------------------------------------------------------------
            // Create a destination document (blank or with some content).
            // -----------------------------------------------------------------
            Document dstDoc = new Document();
            DocumentBuilder dstBuilder = new DocumentBuilder(dstDoc);
            dstBuilder.Writeln("Paragraph in destination document.");
            dstDoc.Save(destinationPath);

            // -----------------------------------------------------------------
            // Load the documents (could also reuse the objects created above).
            // -----------------------------------------------------------------
            srcDoc = new Document(sourcePath);
            dstDoc = new Document(destinationPath);

            // Prepare the importer.
            NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);

            // Get all top‑level comments from the source document.
            NodeCollection srcComments = srcDoc.GetChildNodes(NodeType.Comment, true);

            // Choose where to attach the imported comments – first paragraph of the destination.
            Paragraph targetParagraph = dstDoc.FirstSection.Body.FirstParagraph;

            foreach (Comment srcComment in srcComments)
            {
                // The CommentRangeStart node is the sibling that appears before the comment node.
                CommentRangeStart srcRangeStart = srcComment.PreviousSibling as CommentRangeStart;
                // The CommentRangeEnd node appears after the comment node.
                CommentRangeEnd srcRangeEnd = srcComment.NextSibling as CommentRangeEnd;

                // Import and insert the range start, if it exists.
                if (srcRangeStart != null)
                {
                    Node importedStart = importer.ImportNode(srcRangeStart, true);
                    targetParagraph.AppendChild(importedStart);
                }

                // Import the commented text runs (if any) that belong to the comment range.
                // In this simple example we just copy the runs between the range start and end.
                if (srcRangeStart != null && srcRangeEnd != null)
                {
                    Node current = srcRangeStart.NextSibling;
                    while (current != null && current != srcRangeEnd)
                    {
                        Node importedRun = importer.ImportNode(current, true);
                        targetParagraph.AppendChild(importedRun);
                        current = current.NextSibling;
                    }
                }
                else
                {
                    // Fallback placeholder if the range nodes are missing.
                    Run placeholder = new Run(dstDoc, "Commented text.");
                    targetParagraph.AppendChild(placeholder);
                }

                // Import and insert the range end, if it exists.
                if (srcRangeEnd != null)
                {
                    Node importedEnd = importer.ImportNode(srcRangeEnd, true);
                    targetParagraph.AppendChild(importedEnd);
                }

                // Import the comment itself and attach it to the paragraph.
                Node importedComment = importer.ImportNode(srcComment, true);
                targetParagraph.AppendChild(importedComment);
            }

            // Save the resulting document.
            dstDoc.Save(outputPath);

            Console.WriteLine($"Comment import completed. Output saved to: {outputPath}");
        }
    }
}

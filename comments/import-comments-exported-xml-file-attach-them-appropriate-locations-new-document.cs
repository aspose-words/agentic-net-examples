using System;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace ImportCommentsExample
{
    class Program
    {
        static void Main()
        {
            // Paths to the source (exported comments) and destination documents.
            string sourcePath = @"C:\Docs\CommentsSource.docx";   // document with exported comments
            string destinationPath = @"C:\Docs\NewDocument.docx"; // blank or existing document
            string outputPath = @"C:\Docs\NewDocumentWithComments.docx";

            // Load the documents.
            Document srcDoc = new Document(sourcePath);
            Document dstDoc = new Document(destinationPath);

            // NodeImporter will handle importing nodes while preserving formatting.
            NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);

            // Get all top‑level comments from the source document.
            NodeCollection srcComments = srcDoc.GetChildNodes(NodeType.Comment, true);

            // Choose where to attach the imported comments – first paragraph of the destination.
            Paragraph targetParagraph = dstDoc.FirstSection.Body.FirstParagraph;

            foreach (Comment srcComment in srcComments)
            {
                // The CommentRangeStart node is the sibling that appears *before* the comment node.
                CommentRangeStart srcRangeStart = srcComment.PreviousSibling as CommentRangeStart;
                // The CommentRangeEnd node appears *after* the comment node.
                CommentRangeEnd srcRangeEnd = srcComment.NextSibling as CommentRangeEnd;

                // Import and insert the range start, if it exists.
                if (srcRangeStart != null)
                {
                    Node importedStart = importer.ImportNode(srcRangeStart, true);
                    targetParagraph.AppendChild(importedStart);
                }

                // Insert a placeholder run that will be the text covered by the comment.
                Run placeholder = new Run(dstDoc, "Commented text.");
                targetParagraph.AppendChild(placeholder);

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
        }
    }
}

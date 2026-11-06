using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace BatchCommentAdder
{
    public class Program
    {
        public static void Main()
        {
            // Define folders for input and output documents.
            string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
            string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputDocs");

            // Ensure the folders exist.
            Directory.CreateDirectory(inputFolder);
            Directory.CreateDirectory(outputFolder);

            // Create a few sample documents to demonstrate the batch process.
            CreateSampleDocuments(inputFolder);

            // Standardized disclaimer text to be added as a comment.
            const string disclaimerText = "This document is confidential and intended for the designated recipient only.";

            // Process each DOCX file in the input folder.
            foreach (string filePath in Directory.GetFiles(inputFolder, "*.docx"))
            {
                // Load the document.
                Document doc = new Document(filePath);

                // Use DocumentBuilder to navigate to the start of the document.
                DocumentBuilder builder = new DocumentBuilder(doc);
                builder.MoveToDocumentStart();

                // Insert a paragraph that will host the disclaimer comment.
                builder.Writeln("Disclaimer:");

                // Ensure we have a paragraph to attach the comment to.
                Paragraph? targetParagraph = builder.CurrentParagraph;
                if (targetParagraph == null)
                {
                    // If for some reason there is no current paragraph, skip this document.
                    continue;
                }

                // Create the comment node.
                Comment disclaimerComment = new Comment(doc)
                {
                    Author = "System",
                    Initial = "SYS",
                    DateTime = DateTime.Now
                };

                // Add a paragraph and run inside the comment so it contains visible text.
                Paragraph commentParagraph = new Paragraph(doc);
                Run commentRun = new Run(doc, disclaimerText);
                commentParagraph.AppendChild(commentRun);
                disclaimerComment.AppendChild(commentParagraph);

                // Attach the comment to the target paragraph.
                targetParagraph.AppendChild(disclaimerComment);

                // Save the modified document to the output folder, preserving the original file name.
                string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
                doc.Save(outputPath);
            }
        }

        // Helper method to generate sample DOCX files.
        private static void CreateSampleDocuments(string folderPath)
        {
            for (int i = 1; i <= 3; i++)
            {
                Document sampleDoc = new Document();
                DocumentBuilder builder = new DocumentBuilder(sampleDoc);
                builder.Writeln($"This is sample document {i}.");
                string fileName = $"Sample{i}.docx";
                string fullPath = Path.Combine(folderPath, fileName);
                sampleDoc.Save(fullPath);
            }
        }
    }
}

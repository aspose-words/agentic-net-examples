using System;
using System.IO;
using Aspose.Words;

class BatchDisclaimer
{
    static void Main()
    {
        // Folder containing the source documents
        string inputFolder = @"C:\Docs\Input";

        // Folder where the processed documents will be saved
        string outputFolder = @"C:\Docs\Output";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Retrieve all Word documents in the input folder (adjust the pattern if needed)
        string[] docFiles = Directory.GetFiles(inputFolder, "*.docx");

        foreach (string sourcePath in docFiles)
        {
            // Load the document using the constructor that accepts a file name
            Document doc = new Document(sourcePath);

            // Ensure the document has at least one paragraph to attach the comment to
            doc.EnsureMinimum();

            // Create a standardized disclaimer comment
            Comment disclaimer = new Comment(doc, "System", "SYS", DateTime.Now);
            disclaimer.SetText("Disclaimer: This document is confidential and intended solely for the recipient.");

            // Insert the comment into the first paragraph
            Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;

            // Add a comment range (optional placeholder text inside the range)
            firstParagraph.AppendChild(new CommentRangeStart(doc, disclaimer.Id));
            firstParagraph.AppendChild(new Run(doc, " ")); // placeholder run
            firstParagraph.AppendChild(new CommentRangeEnd(doc, disclaimer.Id));

            // Append the comment node itself
            firstParagraph.AppendChild(disclaimer);

            // Save the modified document, preserving the original file name
            string fileName = Path.GetFileName(sourcePath);
            string destinationPath = Path.Combine(outputFolder, fileName);
            doc.Save(destinationPath);
        }
    }
}

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

class BatchDisclaimer
{
    static void Main()
    {
        // Use folders relative to the executable location
        string baseDir = AppContext.BaseDirectory;
        string inputFolder = Path.Combine(baseDir, "Input");
        string outputFolder = Path.Combine(baseDir, "Output");

        // Ensure both directories exist
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // If there are no .docx files, create a sample document to demonstrate functionality
        string[] docFiles = Directory.GetFiles(inputFolder, "*.docx");
        if (docFiles.Length == 0)
        {
            string samplePath = Path.Combine(inputFolder, "Sample.docx");
            var sampleDoc = new Document();
            var builder = new DocumentBuilder(sampleDoc);
            builder.Writeln("This is a sample document.");
            sampleDoc.Save(samplePath);
            docFiles = new[] { samplePath };
        }

        foreach (string sourcePath in docFiles)
        {
            // Load the document
            Document doc = new Document(sourcePath);
            doc.EnsureMinimum();

            // Create a standardized disclaimer comment
            Comment disclaimer = new Comment(doc, "System", "SYS", DateTime.Now);
            disclaimer.SetText("Disclaimer: This document is confidential and intended solely for the recipient.");

            // Insert the comment into the first paragraph
            Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;

            // Add comment range start, a placeholder run, and comment range end
            firstParagraph.AppendChild(new CommentRangeStart(doc, disclaimer.Id));
            firstParagraph.AppendChild(new Run(doc, " ")); // placeholder run
            firstParagraph.AppendChild(new CommentRangeEnd(doc, disclaimer.Id));

            // Append the comment node itself
            firstParagraph.AppendChild(disclaimer);

            // Save the modified document preserving the original file name
            string fileName = Path.GetFileName(sourcePath);
            string destinationPath = Path.Combine(outputFolder, fileName);
            doc.Save(destinationPath);
        }

        Console.WriteLine("Processing complete. Check the 'Output' folder.");
    }
}

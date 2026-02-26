using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class DocumentSplitter
{
    /// <summary>
    /// Splits a DOCX document into multiple HTML files at heading paragraphs.
    /// Each heading up to <paramref name="headingLevel"/> creates a new part.
    /// </summary>
    /// <param name="inputPath">Full path to the source DOCX file.</param>
    /// <param name="outputFolder">Folder where split HTML files will be written.</param>
    /// <param name="headingLevel">Maximum heading level to split on (1‑9).</param>
    public static void SplitByHeadings(string inputPath, string outputFolder, int headingLevel)
    {
        // Load the source document (load rule).
        Document doc = new Document(inputPath);

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Configure HTML save options to split at heading paragraphs.
        HtmlSaveOptions options = new HtmlSaveOptions
        {
            // Split the document wherever a heading paragraph is encountered.
            DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,
            // Define up to which heading level the split should occur.
            DocumentSplitHeadingLevel = headingLevel,
            // Optional: specify a folder for any images extracted during HTML conversion.
            ImagesFolder = Path.Combine(outputFolder, "Images")
        };

        // Build the base file name for the first part.
        string baseFileName = Path.Combine(outputFolder,
            Path.GetFileNameWithoutExtension(inputPath) + ".html");

        // Save the document using the configured options (save rule).
        // Aspose.Words will generate multiple files:
        //   baseFileName               – first part,
        //   baseFileName-01.html, etc. – subsequent parts.
        doc.Save(baseFileName, options);
    }

    // Example usage.
    static void Main()
    {
        string inputDocx = @"C:\Docs\SourceDocument.docx";
        string outputDir = @"C:\Docs\SplitOutput";
        int maxHeadingLevel = 2; // Split at Heading 1 and Heading 2.

        SplitByHeadings(inputDocx, outputDir, maxHeadingLevel);

        Console.WriteLine("Document split completed.");
    }
}

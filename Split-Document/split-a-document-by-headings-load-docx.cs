using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputFile = @"C:\Docs\SourceDocument.docx";

        // Base name for the split HTML output files.
        // The main file will be "SplitDocument.html",
        // subsequent parts will be "SplitDocument-01.html", "SplitDocument-02.html", etc.
        string outputFile = @"C:\Docs\SplitDocument.html";

        // Load the existing DOCX document.
        Document doc = new Document(inputFile);

        // Configure HTML save options to split the document at heading paragraphs.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            // Split at headings (Heading 1, Heading 2, ...).
            DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,

            // Maximum heading level to trigger a split (1 = Heading 1 only, 2 = Heading 1 & 2, etc.).
            // Adjust as needed; here we split at Heading 1 and Heading 2.
            DocumentSplitHeadingLevel = 2
        };

        // Save the document. Aspose.Words will create multiple HTML files according to the split criteria.
        doc.Save(outputFile, saveOptions);
    }
}

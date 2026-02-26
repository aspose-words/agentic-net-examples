using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = @"C:\Docs\SourceDocument.docx";

        // Base path for the output HTML files.
        // Aspose.Words will create additional files (e.g., SourceDocument-01.html, SourceDocument-02.html) for each chapter.
        string outputPath = @"C:\Docs\SourceDocument.html";

        // Load the DOCX document.
        Document doc = new Document(inputPath);

        // Set up HTML save options to split the document at heading paragraphs.
        // This will treat each heading (e.g., Heading 1) as the start of a new chapter.
        HtmlSaveOptions options = new HtmlSaveOptions();
        options.DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph;
        // Assuming chapters are marked with Heading 1 style; adjust if needed.
        options.DocumentSplitHeadingLevel = 1;

        // Save the document; multiple HTML files will be generated, one per chapter.
        doc.Save(outputPath, options);
    }
}

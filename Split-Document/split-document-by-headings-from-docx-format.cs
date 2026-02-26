using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class SplitDocumentByHeadings
{
    static void Main()
    {
        // Path to the source DOCX file.
        string sourcePath = @"C:\Docs\SourceDocument.docx";

        // Directory where the split HTML files will be written.
        string outputDir = @"C:\Docs\SplitOutput";

        // Ensure the output directory exists.
        if (!Directory.Exists(outputDir))
            Directory.CreateDirectory(outputDir);

        // The base name for the first HTML file. Additional parts will be named
        // automatically as <base>-01.html, <base>-02.html, etc.
        string outputFile = Path.Combine(outputDir, "Split.html");

        // Load the DOCX document.
        Document doc = new Document(sourcePath);

        // Configure HTML save options to split the document at heading paragraphs.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            // Split at paragraphs that use heading styles.
            DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,

            // Split at heading levels 1 and 2 (set to 0 to disable heading splits).
            DocumentSplitHeadingLevel = 2
        };

        // Save the document. Aspose.Words will create multiple HTML files according
        // to the heading split criteria.
        doc.Save(outputFile, saveOptions);

        Console.WriteLine("Document split completed. Files are located in:");
        Console.WriteLine(outputDir);
    }
}

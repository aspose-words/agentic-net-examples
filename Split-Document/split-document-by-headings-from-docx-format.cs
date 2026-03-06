using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class SplitByHeadings
{
    static void Main()
    {
        // Path to the source DOCX file.
        string sourcePath = @"C:\Docs\SourceDocument.docx";

        // Folder where the split HTML files will be written.
        string outputFolder = @"C:\Docs\SplitOutput";

        // Ensure the output folder exists.
        Directory.CreateDirectory(outputFolder);

        // Load the DOCX document.
        Document doc = new Document(sourcePath);

        // Configure HTML save options to split the document at heading paragraphs.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            // Split at paragraphs that use heading styles.
            DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,

            // Split at heading levels 1 and 2 (set to 0 to disable heading splitting).
            DocumentSplitHeadingLevel = 2
        };

        // Base file name for the first part; subsequent parts will receive suffixes like -01.html, -02.html, etc.
        string baseFileName = "SplitDocument.html";

        // Save the document; Aspose.Words will automatically create multiple HTML files according to the headings.
        doc.Save(Path.Combine(outputFolder, baseFileName), saveOptions);
    }
}

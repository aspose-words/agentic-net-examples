using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsSplitExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source Word document.
            string sourcePath = @"C:\Docs\SourceDocument.docx";

            // Load the document using the Document(string) constructor.
            Document doc = new Document(sourcePath);

            // Configure save options to split the document.
            // Here we split at each section break; other criteria can be used (PageBreak, HeadingParagraph, etc.).
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak
            };

            // Destination folder for the split parts.
            string outputFolder = @"C:\Docs\SplitOutput";
            Directory.CreateDirectory(outputFolder);

            // Base file name for the first part; Aspose.Words will generate additional part files.
            string baseFileName = Path.Combine(outputFolder, "SplitPart.html");

            // Save the document. Because DocumentSplitCriteria is set, the document will be split
            // into multiple HTML files (one per section) in the same folder.
            doc.Save(baseFileName, saveOptions);
        }
    }
}

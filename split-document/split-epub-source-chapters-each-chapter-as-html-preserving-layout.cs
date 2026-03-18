using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class EpubSplitter
{
    static void Main()
    {
        // Path to the source EPUB file.
        string inputEpubPath = @"C:\Input\book.epub";

        // Folder where the split HTML chapters will be written.
        string outputFolder = @"C:\Output\Chapters";

        // Clean the output folder if it already exists and recreate it.
        if (Directory.Exists(outputFolder))
            Directory.Delete(outputFolder, true);
        Directory.CreateDirectory(outputFolder);

        // Load the EPUB document.
        Document doc = new Document(inputEpubPath);

        // Configure HTML save options to split the document at heading paragraphs (chapters).
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            // Split at headings (e.g., Heading 1, Heading 2, …).
            DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,

            // Define up to which heading level the split should occur.
            // Adjust this value according to the heading styles used in the EPUB.
            DocumentSplitHeadingLevel = 2,

            // Preserve page setup information (margins, size) in the generated HTML.
            ExportPageSetup = true,

            // Export headers and footers per section so they appear in the HTML where appropriate.
            ExportHeadersFootersMode = ExportHeadersFootersMode.PerSection,

            // Use UTF‑8 encoding for the output files.
            Encoding = System.Text.Encoding.UTF8
        };

        // Save the document. Aspose.Words will create the main HTML file and additional
        // part files (e.g., book-01.html, book-02.html, …) in the same directory.
        string mainHtmlPath = Path.Combine(outputFolder, "book.html");
        doc.Save(mainHtmlPath, htmlOptions);
    }
}

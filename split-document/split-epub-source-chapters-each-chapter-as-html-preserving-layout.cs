using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class EpubSplitter
{
    static void Main()
    {
        // Determine output folder relative to the current directory.
        string outputFolder = Path.Combine(Environment.CurrentDirectory, "Output", "Chapters");

        // Clean the output folder if it already exists and recreate it.
        if (Directory.Exists(outputFolder))
            Directory.Delete(outputFolder, true);
        Directory.CreateDirectory(outputFolder);

        // Create a sample document with headings to simulate an EPUB structure.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Title (not a heading)
        builder.Writeln("Sample Book Title");

        // Chapter 1
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is the content of chapter 1.");

        // Chapter 2
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is the content of chapter 2.");

        // Configure HTML save options to split the document at heading paragraphs (chapters).
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,
            DocumentSplitHeadingLevel = 1,
            ExportPageSetup = true,
            ExportHeadersFootersMode = ExportHeadersFootersMode.PerSection,
            Encoding = System.Text.Encoding.UTF8
        };

        // Save the document. Aspose.Words will create the main HTML file and additional
        // part files (e.g., book-01.html, book-02.html, …) in the same directory.
        string mainHtmlPath = Path.Combine(outputFolder, "book.html");
        doc.Save(mainHtmlPath, htmlOptions);
    }
}

using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with headings that start on new pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First heading – will cause a split because of the heading level and page break.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.ParagraphFormat.PageBreakBefore = true; // Ensure the heading starts on a new page.
        builder.Writeln("Chapter 1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of chapter 1.");
        builder.Writeln("More content of chapter 1.");

        // Second heading.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.ParagraphFormat.PageBreakBefore = true;
        builder.Writeln("Chapter 2");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of chapter 2.");

        // Third heading.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.ParagraphFormat.PageBreakBefore = true;
        builder.Writeln("Chapter 3");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of chapter 3.");

        // Configure HtmlSaveOptions to split by heading paragraphs and page breaks.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph | DocumentSplitCriteria.PageBreak,
            DocumentSplitHeadingLevel = 1 // Split at Heading 1 paragraphs.
        };

        // Optional: customize the filenames of the split parts.
        saveOptions.DocumentPartSavingCallback = new PartRenamer("CombinedSplit");

        // Save the document; Aspose.Words will create multiple HTML files.
        string mainFilePath = Path.Combine(outputDir, "CombinedSplit.html");
        doc.Save(mainFilePath, saveOptions);

        // Verify that split parts were created.
        string[] splitFiles = Directory.GetFiles(outputDir, "CombinedSplit_Part*.html");
        if (splitFiles.Length == 0)
            throw new Exception("No split parts were generated.");

        // (Optional) Output the list of generated files.
        foreach (string file in splitFiles)
            Console.WriteLine($"Generated: {Path.GetFileName(file)}");
    }

    // Callback to rename each document part produced by the split operation.
    private class PartRenamer : IDocumentPartSavingCallback
    {
        private readonly string _baseName;
        private int _counter;

        public PartRenamer(string baseName)
        {
            _baseName = baseName;
            _counter = 0;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            string extension = Path.GetExtension(args.DocumentPartFileName);
            args.DocumentPartFileName = $"{_baseName}_Part{++_counter}{extension}";
        }
    }
}

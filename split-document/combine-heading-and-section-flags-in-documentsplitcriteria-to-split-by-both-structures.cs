using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class SplitDocumentExample
{
    public static void Main()
    {
        // Define a folder for all output files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with headings and explicit section breaks.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Heading 1 – will be a split point because of HeadingParagraph flag.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 1");

        // Insert a section break – will be a split point because of SectionBreak flag.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Heading 2 – another split point.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Heading 2");

        // Insert another section break.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Heading 3 – split point (level 3 is allowed by DocumentSplitHeadingLevel).
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("Heading 3");

        // Normal paragraph – will belong to the last split part.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is a normal paragraph that follows Heading 3.");

        // Configure HTML save options to split by both headings and sections.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            // Combine the two criteria using a bitwise OR.
            DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph | DocumentSplitCriteria.SectionBreak,
            // Split at heading levels 1 through 3.
            DocumentSplitHeadingLevel = 3,
            // Use a callback to give each split part a deterministic file name.
            DocumentPartSavingCallback = new SavedDocumentPartRename("CombinedSplit", DocumentSplitCriteria.HeadingParagraph | DocumentSplitCriteria.SectionBreak)
        };

        // Save the document; the callback will create separate files for each part.
        string mainFileName = Path.Combine(outputDir, "CombinedSplit.html");
        doc.Save(mainFileName, saveOptions);

        // Verify that at least one split part was created.
        string[] splitFiles = Directory.GetFiles(outputDir, "CombinedSplit_part_*.html");
        if (splitFiles.Length == 0)
            throw new InvalidOperationException("No split output files were generated.");

        // Output the names of the generated files (optional, for visual confirmation).
        Console.WriteLine("Generated split files:");
        foreach (string file in splitFiles)
            Console.WriteLine(" - " + Path.GetFileName(file));
    }

    // Callback that renames each document part created during the split operation.
    private class SavedDocumentPartRename : IDocumentPartSavingCallback
    {
        private readonly string _baseName;
        private readonly DocumentSplitCriteria _criteria;
        private int _partIndex = 0;

        public SavedDocumentPartRename(string baseName, DocumentSplitCriteria criteria)
        {
            _baseName = baseName;
            _criteria = criteria;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Generate a simple deterministic file name for each part.
            string newFileName = $"{_baseName}_part_{++_partIndex}{Path.GetExtension(args.DocumentPartFileName)}";
            args.DocumentPartFileName = newFileName;
            // Let Aspose.Words handle the file stream automatically.
        }
    }
}

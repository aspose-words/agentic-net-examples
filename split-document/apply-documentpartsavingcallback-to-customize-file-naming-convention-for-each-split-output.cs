using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace SplitDocumentExample
{
    public class Program
    {
        public static void Main()
        {
            // Prepare output directory.
            string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
            Directory.CreateDirectory(artifactsDir);

            // Create a sample document with three sections.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Content of Section 1");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Writeln("Content of Section 2");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Writeln("Content of Section 3");

            // Configure HTML save options to split by section break and use a custom naming callback.
            HtmlSaveOptions options = new HtmlSaveOptions
            {
                DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
                DocumentPartSavingCallback = new SavedDocumentPartRename(artifactsDir, "SplitDocument.html",
                    DocumentSplitCriteria.SectionBreak)
            };

            // Save the main document; the callback will rename each part.
            string mainOutputPath = Path.Combine(artifactsDir, "SplitDocument.html");
            doc.Save(mainOutputPath, options);

            // Validate that the expected split files were created.
            // The callback creates files named like:
            // "SplitDocument.html part 1, of type Section.html"
            string[] splitFiles = Directory.GetFiles(artifactsDir, "SplitDocument.html part *.html");
            if (splitFiles.Length != doc.Sections.Count)
                throw new InvalidOperationException(
                    $"Expected {doc.Sections.Count} split files, but found {splitFiles.Length}.");

            // Display the generated file names.
            foreach (string file in splitFiles)
                Console.WriteLine($"Created split part: {Path.GetFileName(file)}");
        }
    }

    // Callback that customizes the file name for each document part.
    internal class SavedDocumentPartRename : IDocumentPartSavingCallback
    {
        private readonly string _outputDir;
        private readonly string _baseFileName;
        private readonly DocumentSplitCriteria _criteria;
        private int _count;

        public SavedDocumentPartRename(string outputDir, string baseFileName, DocumentSplitCriteria criteria)
        {
            _outputDir = outputDir;
            _baseFileName = baseFileName;
            _criteria = criteria;
            _count = 0;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Determine the type of split part for naming.
            string partType = _criteria switch
            {
                DocumentSplitCriteria.PageBreak => "Page",
                DocumentSplitCriteria.ColumnBreak => "Column",
                DocumentSplitCriteria.SectionBreak => "Section",
                DocumentSplitCriteria.HeadingParagraph => "Heading",
                _ => "Part"
            };

            // Build a unique file name for the part.
            string extension = Path.GetExtension(args.DocumentPartFileName);
            string partFileName = $"{_baseFileName} part {++_count}, of type {partType}{extension}";

            // Set the file name and stream where Aspose.Words will write the part.
            args.DocumentPartFileName = partFileName;
            args.DocumentPartStream = new FileStream(Path.Combine(_outputDir, partFileName), FileMode.Create);
        }
    }
}

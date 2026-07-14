using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace SplitDocumentExample
{
    // Callback that assigns custom filenames for each split part.
    public class SplitDocumentPartCallback : IDocumentPartSavingCallback
    {
        private readonly string _baseName;
        private readonly DocumentSplitCriteria _criteria;
        private int _partIndex = 0;
        private readonly string _outputFolder;

        public SplitDocumentPartCallback(string baseName, DocumentSplitCriteria criteria, string outputFolder)
        {
            _baseName = baseName;
            _criteria = criteria;
            _outputFolder = outputFolder;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Determine a readable part type name.
            string partType = _criteria switch
            {
                DocumentSplitCriteria.PageBreak => "Page",
                DocumentSplitCriteria.ColumnBreak => "Column",
                DocumentSplitCriteria.SectionBreak => "Section",
                DocumentSplitCriteria.HeadingParagraph => "Heading",
                _ => "Part"
            };

            // Build a unique filename for the part.
            string partFileName = $"{_baseName} part {++_partIndex}, of type {partType}{Path.GetExtension(args.DocumentPartFileName)}";

            // Set the filename and stream where Aspose.Words will write this part.
            args.DocumentPartFileName = partFileName;
            string fullPath = Path.Combine(_outputFolder, partFileName);
            args.DocumentPartStream = new FileStream(fullPath, FileMode.Create);
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Define an output directory for all generated files.
            string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(artifactsDir);

            // -----------------------------------------------------------------
            // 1. Create a sample document with multiple sections.
            // -----------------------------------------------------------------
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("This is content of Section 1.");
            // Insert a section break (new page) to start Section 2.
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Writeln("This is content of Section 2.");

            // Save the original document (optional, for reference).
            string sourcePath = Path.Combine(artifactsDir, "SourceDocument.docx");
            doc.Save(sourcePath);

            // -----------------------------------------------------------------
            // 2. Configure HtmlSaveOptions to split the document by sections.
            // -----------------------------------------------------------------
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                // Create a DocumentSplitCriteria object and set it to split at section breaks.
                DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak
            };

            // Assign the custom callback that will name each split part.
            saveOptions.DocumentPartSavingCallback = new SplitDocumentPartCallback(
                baseName: "SectionPart",
                criteria: saveOptions.DocumentSplitCriteria,
                outputFolder: artifactsDir);

            // -----------------------------------------------------------------
            // 3. Save the document; Aspose.Words will generate separate files.
            // -----------------------------------------------------------------
            string mainHtmlPath = Path.Combine(artifactsDir, "Combined.html");
            doc.Save(mainHtmlPath, saveOptions);

            // -----------------------------------------------------------------
            // 4. Validate that the expected split files were created.
            // -----------------------------------------------------------------
            // Expected filenames based on the callback logic.
            string[] expectedFiles =
            {
                Path.Combine(artifactsDir, "SectionPart part 1, of type Section.html"),
                Path.Combine(artifactsDir, "SectionPart part 2, of type Section.html")
            };

            foreach (string filePath in expectedFiles)
            {
                if (!File.Exists(filePath))
                {
                    throw new FileNotFoundException($"Expected split file not found: {filePath}");
                }
            }

            // If execution reaches this point, the split operation succeeded.
            Console.WriteLine("Document split by sections completed successfully.");
        }
    }
}

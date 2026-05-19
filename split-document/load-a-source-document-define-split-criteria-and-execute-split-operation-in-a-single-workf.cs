using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace SplitDocumentExample
{
    // Custom callback to control the filenames of split document parts.
    class SavedDocumentPartRename : IDocumentPartSavingCallback
    {
        private readonly string _baseFileName;
        private readonly DocumentSplitCriteria _splitCriteria;
        private int _partCount = 0;

        public SavedDocumentPartRename(string baseFileName, DocumentSplitCriteria splitCriteria)
        {
            _baseFileName = baseFileName;
            _splitCriteria = splitCriteria;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Determine a readable part type name.
            string partType = _splitCriteria switch
            {
                DocumentSplitCriteria.PageBreak => "Page",
                DocumentSplitCriteria.ColumnBreak => "Column",
                DocumentSplitCriteria.SectionBreak => "Section",
                DocumentSplitCriteria.HeadingParagraph => "Heading",
                _ => "Part"
            };

            // Build a unique filename for each part.
            string partFileName = $"{_baseFileName}_part{++_partCount}_{partType}{Path.GetExtension(args.DocumentPartFileName)}";

            // Set the filename; Aspose.Words will place the file in the same folder as the main output.
            args.DocumentPartFileName = partFileName;
        }
    }

    class Program
    {
        static void Main()
        {
            // Define folders for input and output.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // -----------------------------------------------------------------
            // 1. Create a sample source document with two sections.
            // -----------------------------------------------------------------
            Document sourceDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sourceDoc);

            // First section content.
            builder.Writeln("Section 1 - Paragraph 1");
            builder.Writeln("Section 1 - Paragraph 2");
            // Insert a section break to start a new section.
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            // Second section content.
            builder.Writeln("Section 2 - Paragraph 1");
            builder.Writeln("Section 2 - Paragraph 2");

            // Save the source document to disk.
            string sourcePath = Path.Combine(outputDir, "SourceDocument.docx");
            sourceDoc.Save(sourcePath);

            // -----------------------------------------------------------------
            // 2. Load the source document.
            // -----------------------------------------------------------------
            Document loadedDoc = new Document(sourcePath);

            // -----------------------------------------------------------------
            // 3. Define split criteria (split by section break) and save.
            // -----------------------------------------------------------------
            string baseOutputName = "SplitDocument.html";
            string baseOutputPath = Path.Combine(outputDir, baseOutputName);

            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
                // Optional: keep original formatting.
                // Keep source formatting when splitting.
                // This is the default behavior for HTML saving.
                DocumentPartSavingCallback = new SavedDocumentPartRename(
                    Path.GetFileNameWithoutExtension(baseOutputName), DocumentSplitCriteria.SectionBreak)
            };

            // Save the document; this will generate multiple HTML files.
            loadedDoc.Save(baseOutputPath, saveOptions);

            // -----------------------------------------------------------------
            // 4. Validate that split parts were created.
            // -----------------------------------------------------------------
            // The main file and the split parts are placed in the same folder.
            string[] expectedFiles =
            {
                baseOutputPath,
                Path.Combine(outputDir, "SplitDocument_part1_Section.html"),
                Path.Combine(outputDir, "SplitDocument_part2_Section.html")
            };

            bool allExist = true;
            foreach (string file in expectedFiles)
            {
                if (!File.Exists(file))
                {
                    allExist = false;
                    Console.WriteLine($"Expected file not found: {file}");
                }
            }

            if (allExist)
                Console.WriteLine("Document split successfully. All expected files are present.");
            else
                Console.WriteLine("Document split failed. Some expected files are missing.");
        }
    }
}

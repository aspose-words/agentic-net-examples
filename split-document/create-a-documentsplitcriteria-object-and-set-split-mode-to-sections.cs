using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace SplitDocumentExample
{
    // Callback that assigns deterministic filenames to each split part.
    class SectionPartSavingCallback : IDocumentPartSavingCallback
    {
        private readonly string _baseFileName;
        private int _partIndex = 0;

        public SectionPartSavingCallback(string baseFileName)
        {
            _baseFileName = baseFileName;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Generate names like "Document_Part1.html", "Document_Part2.html", etc.
            string extension = Path.GetExtension(args.DocumentPartFileName);
            args.DocumentPartFileName = $"{_baseFileName}_Part{++_partIndex}{extension}";
        }
    }

    class Program
    {
        static void Main()
        {
            // Folder for all generated files.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // -----------------------------------------------------------------
            // 1. Create a sample document with multiple sections.
            // -----------------------------------------------------------------
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // First section.
            builder.Writeln("Section 1 - First paragraph.");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Writeln("Section 1 - Second paragraph.");

            // Second section.
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Writeln("Section 2 - Only paragraph.");

            // -----------------------------------------------------------------
            // 2. Create a DocumentSplitCriteria object and set it to split by sections.
            // -----------------------------------------------------------------
            DocumentSplitCriteria splitCriteria = DocumentSplitCriteria.SectionBreak;

            // -----------------------------------------------------------------
            // 3. Configure HtmlSaveOptions to use the split criteria.
            // -----------------------------------------------------------------
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                DocumentSplitCriteria = splitCriteria,
                DocumentPartSavingCallback = new SectionPartSavingCallback("Document")
            };

            // -----------------------------------------------------------------
            // 4. Save the document. The save operation will produce one HTML file
            //    per section because of the split criteria.
            // -----------------------------------------------------------------
            string mainFilePath = Path.Combine(outputDir, "Document.html");
            doc.Save(mainFilePath, saveOptions);

            // -----------------------------------------------------------------
            // 5. Validate that the expected split files exist.
            // -----------------------------------------------------------------
            for (int i = 1; i <= doc.Sections.Count; i++)
            {
                string partPath = Path.Combine(outputDir, $"Document_Part{i}.html");
                if (!File.Exists(partPath))
                    throw new FileNotFoundException($"Expected split file not found: {partPath}");
            }

            // Indicate successful completion.
            Console.WriteLine("Document split into sections and saved successfully.");
        }
    }
}

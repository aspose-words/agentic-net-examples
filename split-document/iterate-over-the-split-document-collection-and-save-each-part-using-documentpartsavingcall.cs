using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace SplitDocumentExample
{
    // Custom callback to control how each document part is saved.
    public class CustomDocumentPartSavingCallback : IDocumentPartSavingCallback
    {
        private readonly string _outputDirectory;
        private readonly string _baseFileName;
        private readonly DocumentSplitCriteria _splitCriteria;
        private int _partCount = 0;

        public CustomDocumentPartSavingCallback(string outputDirectory, string baseFileName, DocumentSplitCriteria splitCriteria)
        {
            _outputDirectory = outputDirectory;
            _baseFileName = baseFileName;
            _splitCriteria = splitCriteria;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Determine the type of part being saved (section, page, etc.).
            string partType = _splitCriteria switch
            {
                DocumentSplitCriteria.PageBreak => "Page",
                DocumentSplitCriteria.ColumnBreak => "Column",
                DocumentSplitCriteria.SectionBreak => "Section",
                DocumentSplitCriteria.HeadingParagraph => "Heading",
                _ => "Part"
            };

            // Create a unique file name for the part.
            string partFileName = $"{_baseFileName}_part{++_partCount}_{partType}{Path.GetExtension(args.DocumentPartFileName)}";

            // Set the file name (without path) and the stream where Aspose.Words will write the part.
            args.DocumentPartFileName = partFileName;
            string fullPath = Path.Combine(_outputDirectory, partFileName);
            args.DocumentPartStream = new FileStream(fullPath, FileMode.Create);

            // Ensure the stream will be closed by Aspose.Words after saving.
            args.KeepDocumentPartStreamOpen = false;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Define output folder.
            string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
            Directory.CreateDirectory(artifactsDir);

            // Create a sample document with multiple sections.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Section 1
            builder.Writeln("Section 1 - Introduction");
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            // Section 2
            builder.Writeln("Section 2 - Body");
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            // Section 3
            builder.Writeln("Section 3 - Conclusion");

            // Prepare HTML save options with splitting by section.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak
            };

            // Base file name for the main output (used only for naming parts).
            string baseFileName = "SplitDocument";

            // Assign the custom callback.
            saveOptions.DocumentPartSavingCallback = new CustomDocumentPartSavingCallback(
                artifactsDir,
                baseFileName,
                saveOptions.DocumentSplitCriteria);

            // Save the document; Aspose.Words will invoke the callback for each part.
            string mainOutputPath = Path.Combine(artifactsDir, $"{baseFileName}.html");
            doc.Save(mainOutputPath, saveOptions);

            // Verify that the expected part files were created.
            string[] partFiles = Directory.GetFiles(artifactsDir, $"{baseFileName}_part*_*.html");
            if (partFiles.Length == 0)
                throw new InvalidOperationException("No document parts were saved.");

            // Output the list of generated files (optional, for demonstration).
            Console.WriteLine("Generated document parts:");
            foreach (string file in partFiles)
                Console.WriteLine(Path.GetFileName(file));
        }
    }
}

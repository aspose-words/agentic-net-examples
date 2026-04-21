using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace DocumentPartSavingDemo
{
    // Custom callback that renames each document part generated during saving.
    public class CustomPartNamingCallback : IDocumentPartSavingCallback
    {
        private readonly string _outputFolder;
        private readonly string _baseFileName;
        private readonly DocumentSplitCriteria _splitCriteria;
        private int _partCounter = 0;

        public CustomPartNamingCallback(string outputFolder, string baseFileName, DocumentSplitCriteria splitCriteria)
        {
            _outputFolder = outputFolder;
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

            // Build a unique file name for the part.
            string partFileName = $"{_baseFileName}_part_{++_partCounter}_of_{partType}{Path.GetExtension(args.DocumentPartFileName)}";

            // Set the file name (Aspose will combine it with the folder of the main output file).
            args.DocumentPartFileName = partFileName;

            // Alternatively, provide a stream directly.
            string fullPath = Path.Combine(_outputFolder, partFileName);
            args.DocumentPartStream = new FileStream(fullPath, FileMode.Create);
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare output directory.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Base name for the main HTML file (without extension).
            string baseFileName = "SplitDocument";

            // Create a sample document with three sections.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Content of Section 1.");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Writeln("Content of Section 2.");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Writeln("Content of Section 3.");

            // Configure HTML save options to split by section.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
                // Assign the custom callback to control part file names.
                DocumentPartSavingCallback = new CustomPartNamingCallback(outputDir, baseFileName, DocumentSplitCriteria.SectionBreak)
            };

            // Save the document; Aspose will invoke the callback for each part.
            string mainOutputPath = Path.Combine(outputDir, $"{baseFileName}.html");
            doc.Save(mainOutputPath, saveOptions);

            // Verify that the expected split files were created (three sections => three parts).
            string[] partFiles = Directory.GetFiles(outputDir, $"{baseFileName}_part_*_of_Section.html");
            if (partFiles.Length != 3)
                throw new InvalidOperationException($"Expected 3 split parts, but found {partFiles.Length}.");

            // Optional: inform the user (no interactive input required).
            Console.WriteLine($"Document split completed. Main file: {mainOutputPath}");
            Console.WriteLine("Generated parts:");
            foreach (string file in partFiles)
                Console.WriteLine($" - {Path.GetFileName(file)}");
        }
    }
}

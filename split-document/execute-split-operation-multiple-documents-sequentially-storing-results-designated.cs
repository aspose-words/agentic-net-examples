using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsSplitExample
{
    // Callback that controls how each document part is saved when the document is split.
    internal class SavedDocumentPartRename : IDocumentPartSavingCallback
    {
        private readonly string _outputDirectory;
        private readonly string _baseFileName;
        private readonly DocumentSplitCriteria _splitCriteria;
        private int _partCount;

        public SavedDocumentPartRename(string outputDirectory, string baseFileName, DocumentSplitCriteria splitCriteria)
        {
            _outputDirectory = outputDirectory;
            _baseFileName = baseFileName;
            _splitCriteria = splitCriteria;
            _partCount = 0;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            string partType = _splitCriteria switch
            {
                DocumentSplitCriteria.PageBreak => "Page",
                DocumentSplitCriteria.ColumnBreak => "Column",
                DocumentSplitCriteria.SectionBreak => "Section",
                DocumentSplitCriteria.HeadingParagraph => "Heading",
                _ => "Part"
            };

            string partFileName = $"{_baseFileName}_part{++_partCount}_{partType}{Path.GetExtension(args.DocumentPartFileName)}";

            args.DocumentPartFileName = partFileName;
            args.DocumentPartStream = new FileStream(Path.Combine(_outputDirectory, partFileName), FileMode.Create);
            args.KeepDocumentPartStreamOpen = false;
        }
    }

    public static class DocumentSplitter
    {
        // Splits each document in the inputFiles array according to the specified criteria
        // and stores the resulting parts in the corresponding outputDirectories.
        public static void SplitDocuments(string[] inputFiles, string[] outputDirectories)
        {
            if (inputFiles == null) throw new ArgumentNullException(nameof(inputFiles));
            if (outputDirectories == null) throw new ArgumentNullException(nameof(outputDirectories));
            if (inputFiles.Length != outputDirectories.Length)
                throw new ArgumentException("The number of input files must match the number of output directories.");

            for (int i = 0; i < inputFiles.Length; i++)
            {
                string inputPath = inputFiles[i];
                string outputDir = outputDirectories[i];

                Directory.CreateDirectory(outputDir);

                Document doc = new Document(inputPath);

                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak
                };

                string baseFileName = Path.GetFileNameWithoutExtension(inputPath) + ".html";

                saveOptions.DocumentPartSavingCallback = new SavedDocumentPartRename(
                    outputDir,
                    baseFileName,
                    saveOptions.DocumentSplitCriteria);

                string mainOutputPath = Path.Combine(outputDir, baseFileName);
                doc.Save(mainOutputPath, saveOptions);
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Create temporary directories for source documents and split results.
            string tempRoot = Path.Combine(Path.GetTempPath(), "AsposeSplitDemo");
            string sourceDir = Path.Combine(tempRoot, "Sources");
            string outputRoot = Path.Combine(tempRoot, "Outputs");
            Directory.CreateDirectory(sourceDir);
            Directory.CreateDirectory(outputRoot);

            // Prepare source document paths.
            string[] sources = new[]
            {
                Path.Combine(sourceDir, "Report1.docx"),
                Path.Combine(sourceDir, "Report2.docx")
            };

            // Generate simple documents with a section break to allow splitting.
            for (int i = 0; i < sources.Length; i++)
            {
                Document doc = new Document();
                DocumentBuilder builder = new DocumentBuilder(doc);
                builder.Writeln($"Document {i + 1} - First section");
                builder.InsertBreak(BreakType.SectionBreakNewPage);
                builder.Writeln($"Document {i + 1} - Second section");
                doc.Save(sources[i]);
            }

            // Define output directories for each source document.
            string[] destinations = new[]
            {
                Path.Combine(outputRoot, "Report1"),
                Path.Combine(outputRoot, "Report2")
            };

            // Perform the split operation.
            DocumentSplitter.SplitDocuments(sources, destinations);

            Console.WriteLine("Splitting completed. Check the folder:");
            Console.WriteLine(outputRoot);
        }
    }
}

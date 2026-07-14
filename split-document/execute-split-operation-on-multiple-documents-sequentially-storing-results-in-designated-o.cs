using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Saving;

namespace SplitDocumentExample
{
    // Custom callback to control how each split part is saved.
    internal class SavedDocumentPartRename : IDocumentPartSavingCallback
    {
        private readonly string _outputDirectory;
        private readonly DocumentSplitCriteria _criteria;
        private int _partIndex = 0;

        public SavedDocumentPartRename(string outputDirectory, DocumentSplitCriteria criteria)
        {
            _outputDirectory = outputDirectory;
            _criteria = criteria;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Increment part counter.
            _partIndex++;

            // Determine a simple part type name for readability (optional).
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
            string partFileName = $"Part{_partIndex}_{partType}{extension}";
            string fullPath = Path.Combine(_outputDirectory, partFileName);

            // Provide a stream for Aspose.Words to write the part.
            args.DocumentPartStream = new FileStream(fullPath, FileMode.Create);
            // Also set the file name (without path) for completeness.
            args.DocumentPartFileName = partFileName;
        }
    }

    public class Program
    {
        static void Main()
        {
            // Base working directory.
            string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
            string sourceDir = Path.Combine(baseDir, "Source");
            string outputRoot = Path.Combine(baseDir, "Output");

            // Ensure clean environment.
            if (Directory.Exists(baseDir))
                Directory.Delete(baseDir, true);
            Directory.CreateDirectory(sourceDir);
            Directory.CreateDirectory(outputRoot);

            // Create sample source documents.
            CreateSampleDocuments(sourceDir);

            // Process each document: split by section and save parts.
            foreach (string sourcePath in Directory.GetFiles(sourceDir, "*.docx"))
            {
                // Load the source document.
                Document doc = new Document(sourcePath);

                // Prepare output folder for this document.
                string docNameWithoutExt = Path.GetFileNameWithoutExtension(sourcePath);
                string docOutputDir = Path.Combine(outputRoot, docNameWithoutExt);
                Directory.CreateDirectory(docOutputDir);

                // Configure HTML save options to split at each section break.
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
                    DocumentPartSavingCallback = new SavedDocumentPartRename(docOutputDir, DocumentSplitCriteria.SectionBreak)
                };

                // Save the document; the callback will create separate files for each part.
                string masterFilePath = Path.Combine(docOutputDir, $"{docNameWithoutExt}_master.html");
                doc.Save(masterFilePath, saveOptions);

                // Validate that split parts were created (at least two sections expected).
                string[] partFiles = Directory.GetFiles(docOutputDir, "*.html")
                                              .Where(f => !f.EndsWith("_master.html", StringComparison.OrdinalIgnoreCase))
                                              .ToArray();

                if (partFiles.Length < 2)
                    throw new InvalidOperationException($"Expected at least 2 split parts for '{sourcePath}', but found {partFiles.Length}.");

                // Optional: output result summary.
                Console.WriteLine($"Document '{Path.GetFileName(sourcePath)}' split into {partFiles.Length} parts in folder:");
                Console.WriteLine($"  {docOutputDir}");
            }
        }

        // Helper to create a few sample documents with multiple sections.
        private static void CreateSampleDocuments(string folder)
        {
            for (int i = 1; i <= 2; i++)
            {
                Document doc = new Document();
                DocumentBuilder builder = new DocumentBuilder(doc);

                // First section.
                builder.Writeln($"Document {i} - Section 1");
                builder.InsertBreak(BreakType.SectionBreakNewPage);

                // Second section.
                builder.Writeln($"Document {i} - Section 2");
                builder.InsertBreak(BreakType.SectionBreakNewPage);

                // Third section.
                builder.Writeln($"Document {i} - Section 3");

                string filePath = Path.Combine(folder, $"Doc{i}.docx");
                doc.Save(filePath);
            }
        }
    }
}

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace DocumentPartSavingExample
{
    // Callback that decides the file name and format for each document part.
    public class PartSavingCallback : IDocumentPartSavingCallback
    {
        private int _partIndex;
        private readonly string _outputFolder;

        public PartSavingCallback(string outputFolder)
        {
            _outputFolder = outputFolder;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            _partIndex++;

            // Even parts -> DOCX, Odd parts -> PDF
            string extension = (_partIndex % 2 == 0) ? ".docx" : ".pdf";
            string fileName = $"Part{_partIndex}{extension}";

            // Set the file name (without path) and provide a stream where Aspose.Words will write the part.
            args.DocumentPartFileName = fileName;
            args.DocumentPartStream = new FileStream(Path.Combine(_outputFolder, fileName), FileMode.Create);
            args.KeepDocumentPartStreamOpen = false; // Let Aspose.Words close the stream after writing.
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare output directory.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Create a sample document with several sections.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            for (int i = 1; i <= 4; i++)
            {
                builder.Writeln($"This is content of section {i}.");
                // Insert a section break after each section except the last one.
                if (i < 4)
                    builder.InsertBreak(BreakType.SectionBreakNewPage);
            }

            // Configure HTML save options to split the document by section.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
                DocumentPartSavingCallback = new PartSavingCallback(outputDir)
            };

            // Save the document; the callback will create separate files for each part.
            string mainFilePath = Path.Combine(outputDir, "Combined.html");
            doc.Save(mainFilePath, saveOptions);

            // Simple validation: ensure that the expected part files were created.
            for (int i = 1; i <= 4; i++)
            {
                string expectedExtension = (i % 2 == 0) ? ".docx" : ".pdf";
                string expectedPath = Path.Combine(outputDir, $"Part{i}{expectedExtension}");
                if (!File.Exists(expectedPath))
                    throw new FileNotFoundException($"Expected part file not found: {expectedPath}");
            }

            // All done.
        }
    }
}

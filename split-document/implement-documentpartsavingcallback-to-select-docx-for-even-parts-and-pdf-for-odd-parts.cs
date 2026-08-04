using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace DocumentPartSavingDemo
{
    // Callback that assigns a different file extension based on the part index:
    // even parts -> .docx, odd parts -> .pdf
    public class PartFormatCallback : IDocumentPartSavingCallback
    {
        private readonly string _outputFolder;
        private int _partIndex;

        public PartFormatCallback(string outputFolder)
        {
            _outputFolder = outputFolder;
            _partIndex = 0;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Determine the desired extension.
            string extension = (_partIndex % 2 == 0) ? ".docx" : ".pdf";

            // Build a unique file name for the part.
            string partFileName = $"Part_{_partIndex + 1}{extension}";

            // Set the file name (without path) that Aspose.Words will use.
            args.DocumentPartFileName = partFileName;

            // Create a stream that writes the part to the output folder.
            string fullPath = Path.Combine(_outputFolder, partFileName);
            args.DocumentPartStream = new FileStream(fullPath, FileMode.Create);

            // Increment the counter for the next part.
            _partIndex++;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare output directory.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Create a sample document with several sections to trigger splitting.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            for (int i = 1; i <= 4; i++)
            {
                builder.Writeln($"Section {i}");
                // Insert a section break after each section except the last.
                if (i < 4)
                    builder.InsertBreak(BreakType.SectionBreakNewPage);
            }

            // Configure HTML save options to split by section.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
                DocumentPartSavingCallback = new PartFormatCallback(outputDir)
            };

            // Save the document; the callback will create separate files.
            string mainFilePath = Path.Combine(outputDir, "Combined.html");
            doc.Save(mainFilePath, saveOptions);

            // Simple verification that the expected files were created.
            for (int i = 1; i <= 4; i++)
            {
                string expectedExtension = (i % 2 == 1) ? ".docx" : ".pdf"; // 1st part is even index (0) -> .docx
                string expectedPath = Path.Combine(outputDir, $"Part_{i}{expectedExtension}");
                if (!File.Exists(expectedPath))
                    throw new FileNotFoundException($"Expected part file not found: {expectedPath}");
            }

            // Program ends without waiting for user input.
        }
    }
}

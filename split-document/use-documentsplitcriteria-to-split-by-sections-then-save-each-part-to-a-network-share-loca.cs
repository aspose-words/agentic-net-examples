using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace SplitDocumentBySection
{
    public class Program
    {
        public static void Main()
        {
            // Create a sample document with three sections.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Content of Section 1");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Writeln("Content of Section 2");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Writeln("Content of Section 3");

            // Simulate a network share by using a folder inside the current directory.
            string networkSharePath = Path.Combine(Environment.CurrentDirectory, "NetworkShare");
            Directory.CreateDirectory(networkSharePath);

            // Base file name for the main HTML file.
            string baseFileName = "SplitDocument.html";

            // Configure HTML save options to split the document at each section break.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
                DocumentPartSavingCallback = new SavedDocumentPartRename(networkSharePath, baseFileName, DocumentSplitCriteria.SectionBreak)
            };

            // Save the document; the callback will create separate files for each section.
            doc.Save(Path.Combine(networkSharePath, baseFileName), saveOptions);

            // Verify that the expected number of split files were created.
            int expectedParts = doc.Sections.Count;
            int actualParts = Directory.GetFiles(networkSharePath, $"{Path.GetFileNameWithoutExtension(baseFileName)}_Part*{Path.GetExtension(baseFileName)}").Length;

            if (actualParts != expectedParts)
                throw new InvalidOperationException($"Expected {expectedParts} split parts, but found {actualParts}.");

            // Program ends automatically.
        }

        // Callback that redirects each document part to a custom file in the network share.
        private class SavedDocumentPartRename : IDocumentPartSavingCallback
        {
            private readonly string _outputFolder;
            private readonly string _baseFileName;
            private readonly DocumentSplitCriteria _criteria;
            private int _partIndex = 0;

            public SavedDocumentPartRename(string outputFolder, string baseFileName, DocumentSplitCriteria criteria)
            {
                _outputFolder = outputFolder;
                _baseFileName = Path.GetFileNameWithoutExtension(baseFileName);
                _criteria = criteria;
            }

            void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
            {
                // Generate a unique file name for each part.
                string partFileName = $"{_baseFileName}_Part{++_partIndex}{Path.GetExtension(args.DocumentPartFileName)}";

                // Set the file name (without path) that Aspose.Words will use.
                args.DocumentPartFileName = partFileName;

                // Provide a stream that writes directly to the desired location.
                string fullPath = Path.Combine(_outputFolder, partFileName);
                args.DocumentPartStream = new FileStream(fullPath, FileMode.Create);
                // Ensure the stream is writable.
                if (!args.DocumentPartStream.CanWrite)
                    throw new InvalidOperationException("Unable to write to the document part stream.");
                // Keep the stream closed after saving (default behavior).
                args.KeepDocumentPartStreamOpen = false;
            }
        }
    }
}

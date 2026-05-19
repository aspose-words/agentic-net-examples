using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace SplitDocumentExample
{
    public class Program
    {
        public static void Main()
        {
            // Define a folder that simulates a network share.
            string networkShareFolder = Path.Combine(Environment.CurrentDirectory, "NetworkShare");
            Directory.CreateDirectory(networkShareFolder);

            // Create a sample document with three sections.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Content of Section 1");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Writeln("Content of Section 2");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Writeln("Content of Section 3");

            // Configure HTML save options to split the document by section breaks.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
                DocumentPartSavingCallback = new SavedDocumentPartRename(networkShareFolder, DocumentSplitCriteria.SectionBreak)
            };

            // Save the document; each section will be saved as a separate HTML file in the network share folder.
            string mainOutputPath = Path.Combine(networkShareFolder, "SplitDocument.html");
            doc.Save(mainOutputPath, saveOptions);

            // Validate that the expected number of split files were created.
            string[] splitFiles = Directory.GetFiles(networkShareFolder, "SplitDocument part *.html");
            if (splitFiles.Length != doc.Sections.Count)
                throw new InvalidOperationException($"Expected {doc.Sections.Count} split files, but found {splitFiles.Length}.");

            Console.WriteLine($"Document split into {splitFiles.Length} parts and saved to '{networkShareFolder}'.");
        }

        // Callback that controls how each document part is saved.
        private class SavedDocumentPartRename : IDocumentPartSavingCallback
        {
            private readonly string _outputFolder;
            private readonly DocumentSplitCriteria _criteria;
            private int _partIndex;

            public SavedDocumentPartRename(string outputFolder, DocumentSplitCriteria criteria)
            {
                _outputFolder = outputFolder;
                _criteria = criteria;
                _partIndex = 0;
            }

            void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
            {
                // Determine a simple part type name for readability (always "Section" in this scenario).
                string partType = "Section";

                // Build a unique file name for the part.
                string partFileName = $"SplitDocument part {++_partIndex} of type {partType}{Path.GetExtension(args.DocumentPartFileName)}";

                // Set the file name (without path) and the stream where Aspose.Words will write the part.
                args.DocumentPartFileName = partFileName;
                args.DocumentPartStream = new FileStream(Path.Combine(_outputFolder, partFileName), FileMode.Create);
                args.KeepDocumentPartStreamOpen = false;
            }
        }
    }
}

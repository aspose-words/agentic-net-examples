using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace SplitDocumentWithCallback
{
    // Callback implementation that prefixes part file names with the section index.
    public class SectionPartSavingCallback : IDocumentPartSavingCallback
    {
        private readonly int _sectionIndex;

        public SectionPartSavingCallback(int sectionIndex)
        {
            _sectionIndex = sectionIndex;
        }

        // The interface method signature does not include a sender parameter.
        public void DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Prefix each part file name with the section identifier.
            args.DocumentPartFileName = $"Section{_sectionIndex}_{args.DocumentPartFileName}";
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare output directory.
            string outputDir = "SplitOutputs";
            Directory.CreateDirectory(outputDir);

            // Create a sample source document with three sections.
            Document sourceDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sourceDoc);

            for (int sec = 1; sec <= 3; sec++)
            {
                // Add a new section (except for the first one which already exists).
                if (sec > 1)
                    builder.InsertBreak(BreakType.SectionBreakNewPage);

                // Write some text indicating the section number.
                builder.Writeln($"This is content of Section {sec}.");
            }

            // Split the document by sections.
            for (int i = 0; i < sourceDoc.Sections.Count; i++)
            {
                // Create a new empty document to hold the split section.
                Document splitDoc = new Document();
                // Remove the default empty section that a new Document contains.
                splitDoc.RemoveAllChildren();

                // Import the current section from the source document.
                NodeImporter importer = new NodeImporter(sourceDoc, splitDoc, ImportFormatMode.KeepSourceFormatting);
                Section importedSection = (Section)importer.ImportNode(sourceDoc.Sections[i], true);
                splitDoc.AppendChild(importedSection);

                // Configure HTML save options with a DocumentPartSavingCallback to customize part file names.
                HtmlSaveOptions saveOptions = new HtmlSaveOptions();
                int sectionIndex = i + 1; // Capture the current section number for the callback.
                saveOptions.DocumentPartSavingCallback = new SectionPartSavingCallback(sectionIndex);

                // Save the split document as HTML.
                string outFile = Path.Combine(outputDir, $"Section{sectionIndex}.html");
                splitDoc.Save(outFile, saveOptions);
            }

            // Verify that the expected split files were created.
            for (int i = 1; i <= 3; i++)
            {
                string expectedFile = Path.Combine(outputDir, $"Section{i}.html");
                if (!File.Exists(expectedFile))
                    throw new FileNotFoundException($"Expected split file not found: {expectedFile}");
            }

            // Execution completed.
        }
    }
}

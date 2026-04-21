using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsSplitBySection
{
    // Custom callback to control where each split part is saved.
    public class NetworkSharePartSaver : IDocumentPartSavingCallback
    {
        private readonly string _shareFolder;
        private readonly DocumentSplitCriteria _criteria;
        private int _partIndex = 0;

        public NetworkSharePartSaver(string shareFolder, DocumentSplitCriteria criteria)
        {
            _shareFolder = shareFolder;
            _criteria = criteria;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Determine a readable part type name.
            string partType = _criteria switch
            {
                DocumentSplitCriteria.PageBreak => "Page",
                DocumentSplitCriteria.ColumnBreak => "Column",
                DocumentSplitCriteria.SectionBreak => "Section",
                DocumentSplitCriteria.HeadingParagraph => "Heading",
                _ => "Part"
            };

            // Build a unique file name for the part.
            string partFileName = $"Document_{partType}_Part_{++_partIndex}.html";

            // Save the part directly to the network share folder.
            string fullPath = Path.Combine(_shareFolder, partFileName);
            args.DocumentPartStream = new FileStream(fullPath, FileMode.Create);
            args.DocumentPartFileName = partFileName; // Not strictly required when using a stream.
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Define a folder that simulates a network share.
            // In a real scenario this could be a UNC path like @"\\Server\Share\Docs".
            string networkSharePath = Path.Combine(Path.GetTempPath(), "AsposeNetworkShare");
            Directory.CreateDirectory(networkSharePath);

            // Create a sample document with three sections.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Section 1 - First paragraph.");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Writeln("Section 2 - First paragraph.");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Writeln("Section 3 - First paragraph.");

            // Configure HTML save options to split by section.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
                DocumentPartSavingCallback = new NetworkSharePartSaver(networkSharePath, DocumentSplitCriteria.SectionBreak)
            };

            // The main file name is required; its location determines the base path for parts.
            string mainFilePath = Path.Combine(networkSharePath, "CombinedDocument.html");
            doc.Save(mainFilePath, saveOptions);

            // Validation: ensure a part file exists for each section.
            int expectedParts = doc.Sections.Count;
            int actualParts = Directory.GetFiles(networkSharePath, "Document_Section_Part_*.html").Length;

            if (actualParts != expectedParts)
                throw new InvalidOperationException($"Expected {expectedParts} split parts, but found {actualParts}.");

            // Optional: clean up the main combined file if not needed.
            // File.Delete(mainFilePath);
        }
    }
}

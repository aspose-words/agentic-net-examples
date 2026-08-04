using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

namespace SplitDocumentExample
{
    public class Program
    {
        public static void Main()
        {
            // Prepare output directory.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Create a sample document with two sections.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Content of the first section.");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Writeln("Content of the second section.");

            // Configure HTML save options to split by section.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
                DocumentPartSavingCallback = new SavedDocumentPartRename("SplitDocument", DocumentSplitCriteria.SectionBreak, outputDir)
            };

            // Save the document; the callback will create separate files for each part.
            string mainFilePath = Path.Combine(outputDir, "SplitDocument.html");
            doc.Save(mainFilePath, saveOptions);

            // Verify that split files were created.
            var splitFiles = Directory.GetFiles(outputDir, "SplitDocument_part*");
            if (!splitFiles.Any())
                throw new Exception("No split document parts were generated.");
        }

        // Callback that customizes the file name and stream for each document part.
        private class SavedDocumentPartRename : IDocumentPartSavingCallback
        {
            private readonly string _baseName;
            private readonly DocumentSplitCriteria _criteria;
            private readonly string _outputDir;
            private int _count;

            public SavedDocumentPartRename(string baseName, DocumentSplitCriteria criteria, string outputDir)
            {
                _baseName = baseName;
                _criteria = criteria;
                _outputDir = outputDir;
                _count = 0;
            }

            void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
            {
                string partType = _criteria switch
                {
                    DocumentSplitCriteria.PageBreak => "Page",
                    DocumentSplitCriteria.ColumnBreak => "Column",
                    DocumentSplitCriteria.SectionBreak => "Section",
                    DocumentSplitCriteria.HeadingParagraph => "Heading",
                    _ => "Part"
                };

                string partFileName = $"{_baseName}_part{++_count}_{partType}{Path.GetExtension(args.DocumentPartFileName)}";
                args.DocumentPartFileName = partFileName;

                string fullPath = Path.Combine(_outputDir, partFileName);
                args.DocumentPartStream = new FileStream(fullPath, FileMode.Create);
                args.KeepDocumentPartStreamOpen = false;
            }
        }
    }
}

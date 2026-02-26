using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace DocumentSplitExample
{
    public class DocumentSplitter
    {
        // Splits the input DOCX document into separate HTML parts using different split criteria.
        public void SplitByCriteria(string inputPath, string outputFolder)
        {
            // Ensure the output directory exists.
            Directory.CreateDirectory(outputFolder);

            // Load the source document.
            Document doc = new Document(inputPath);

            // Split by section breaks.
            SaveSplit(doc, outputFolder, "Section", DocumentSplitCriteria.SectionBreak);

            // Split by page breaks.
            SaveSplit(doc, outputFolder, "Page", DocumentSplitCriteria.PageBreak);

            // Split by column breaks.
            SaveSplit(doc, outputFolder, "Column", DocumentSplitCriteria.ColumnBreak);

            // Split by heading paragraphs (default heading level = 2).
            SaveSplit(doc, outputFolder, "Heading", DocumentSplitCriteria.HeadingParagraph);
        }

        // Helper method that configures HtmlSaveOptions and saves the document using a custom callback.
        private void SaveSplit(Document doc, string outputFolder, string prefix, DocumentSplitCriteria criteria)
        {
            HtmlSaveOptions options = new HtmlSaveOptions
            {
                DocumentSplitCriteria = criteria,
                // When splitting by headings we can limit the heading level (optional).
                DocumentSplitHeadingLevel = criteria.HasFlag(DocumentSplitCriteria.HeadingParagraph) ? 2 : 0,
                // Use a callback to give each part a meaningful file name.
                DocumentPartSavingCallback = new PartRenamer(prefix, criteria, outputFolder)
            };

            // The main file name is required but will contain only the first part.
            string mainFile = Path.Combine(outputFolder, $"{prefix}_Full.html");
            doc.Save(mainFile, options);
        }

        // Implements IDocumentPartSavingCallback to control the naming of each split part.
        private class PartRenamer : IDocumentPartSavingCallback
        {
            private readonly string _prefix;
            private readonly DocumentSplitCriteria _criteria;
            private readonly string _outputFolder;
            private int _partIndex = 0;

            public PartRenamer(string prefix, DocumentSplitCriteria criteria, string outputFolder)
            {
                _prefix = prefix;
                _criteria = criteria;
                _outputFolder = outputFolder;
            }

            void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
            {
                // Determine a readable part type based on the split criteria.
                string partType = _criteria switch
                {
                    DocumentSplitCriteria.PageBreak => "Page",
                    DocumentSplitCriteria.ColumnBreak => "Column",
                    DocumentSplitCriteria.SectionBreak => "Section",
                    DocumentSplitCriteria.HeadingParagraph => "Heading",
                    _ => "Part"
                };

                // Build a unique file name for the part.
                string partFileName = $"{_prefix}_Part{++_partIndex}_{partType}{Path.GetExtension(args.DocumentPartFileName)}";

                // Set the file name.
                args.DocumentPartFileName = partFileName;

                // Write to a custom stream (demonstrates both approaches).
                string fullPath = Path.Combine(_outputFolder, partFileName);
                args.DocumentPartStream = new FileStream(fullPath, FileMode.Create);
                // KeepDocumentPartStreamOpen remains false (default) – Aspose.Words will close the stream after writing.
            }
        }
    }

    // Entry point required for a console application.
    public static class Program
    {
        public static void Main(string[] args)
        {
            // Simple argument handling – you can replace these paths with your own.
            string inputPath = args.Length > 0 ? args[0] : @"C:\Docs\Input.docx";
            string outputFolder = args.Length > 1 ? args[1] : @"C:\Docs\Output";

            var splitter = new DocumentSplitter();
            splitter.SplitByCriteria(inputPath, outputFolder);

            Console.WriteLine($"Document split completed. Output folder: {outputFolder}");
        }
    }
}

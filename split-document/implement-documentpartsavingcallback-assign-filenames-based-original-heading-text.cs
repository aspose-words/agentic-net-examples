using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

namespace DocumentPartSavingExample
{
    // Callback that renames each document part using the original heading text.
    public class HeadingBasedDocumentPartRename : IDocumentPartSavingCallback
    {
        private readonly List<string> _headings;
        private int _partIndex;
        private readonly string _outputDir;

        public HeadingBasedDocumentPartRename(Document sourceDocument, string outputDir)
        {
            // Collect heading texts in the order they appear in the source document.
            _headings = sourceDocument
                .GetChildNodes(NodeType.Paragraph, true)
                .Cast<Paragraph>()
                .Where(p => p.ParagraphFormat.IsHeading)
                .Select(p => p.GetText().Trim())
                .ToList();

            _partIndex = 0;
            _outputDir = outputDir;
            Directory.CreateDirectory(_outputDir);
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            string headingText;
            if (_partIndex < _headings.Count)
            {
                headingText = _headings[_partIndex];
            }
            else
            {
                // Fallback for parts without a corresponding heading (e.g., content before the first heading).
                headingText = $"Part{_partIndex + 1}";
            }

            _partIndex++;

            // Preserve the original extension (e.g., .html).
            string extension = Path.GetExtension(args.DocumentPartFileName);
            string safeHeading = MakeFileNameSafe(headingText);
            string partFileName = $"{safeHeading}{extension}";

            // Assign the new file name.
            args.DocumentPartFileName = partFileName;

            // Ensure the directory exists; Aspose will handle the stream creation.
            // No need to set DocumentPartStream manually.
        }

        // Removes characters that are invalid in file names.
        private static string MakeFileNameSafe(string name)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
                name = name.Replace(c, '_');
            return string.IsNullOrWhiteSpace(name) ? "Untitled" : name;
        }
    }

    class Program
    {
        static void Main()
        {
            // Load the source document.
            Document doc = new Document("InputDocument.docx");

            // Output directory for the split parts.
            string outputDir = Path.Combine(Environment.CurrentDirectory, "Artifacts");
            Directory.CreateDirectory(outputDir);

            // Configure HTML save options to split by heading paragraphs.
            HtmlSaveOptions options = new HtmlSaveOptions
            {
                DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,
                DocumentPartSavingCallback = new HeadingBasedDocumentPartRename(doc, outputDir)
            };

            // Save the document; each part will be named after its heading (or a fallback name).
            doc.Save(Path.Combine(outputDir, "CombinedOutput.html"), options);
        }
    }
}

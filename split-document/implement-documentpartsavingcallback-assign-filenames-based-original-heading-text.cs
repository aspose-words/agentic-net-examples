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

        public HeadingBasedDocumentPartRename(Document sourceDocument, DocumentSplitCriteria splitCriteria)
        {
            // Collect heading texts in the order they appear in the source document.
            // This order matches the order of parts when splitting by HeadingParagraph.
            _headings = sourceDocument
                .GetChildNodes(NodeType.Paragraph, true)
                .Cast<Paragraph>()
                .Where(p => p.ParagraphFormat.IsHeading)
                .Select(p => p.GetText().Trim())
                .ToList();

            _partIndex = 0;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Guard against mismatched part count.
            if (_partIndex >= _headings.Count)
                throw new InvalidOperationException("More document parts than headings were generated.");

            // Use the heading text as the base file name.
            string headingText = _headings[_partIndex];
            _partIndex++;

            // Preserve the original extension (e.g., .html).
            string extension = Path.GetExtension(args.DocumentPartFileName);
            string safeHeading = MakeFileNameSafe(headingText);
            string partFileName = $"{safeHeading}{extension}";

            // Assign the new file name.
            args.DocumentPartFileName = partFileName;

            // Optionally write to a custom stream (demonstrates both ways).
            args.DocumentPartStream = new FileStream(Path.Combine(ArtifactsDir, partFileName), FileMode.Create);
            args.KeepDocumentPartStreamOpen = false;
        }

        // Removes characters that are invalid in file names.
        private static string MakeFileNameSafe(string name)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
                name = name.Replace(c, '_');
            return string.IsNullOrWhiteSpace(name) ? "Untitled" : name;
        }

        // Adjust this path as needed for your environment.
        private static readonly string ArtifactsDir = Path.Combine(Environment.CurrentDirectory, "Artifacts");
    }

    class Program
    {
        static void Main()
        {
            // Load the source document.
            Document doc = new Document("InputDocument.docx");

            // Configure HTML save options to split by heading paragraphs.
            HtmlSaveOptions options = new HtmlSaveOptions
            {
                DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,
                DocumentPartSavingCallback = new HeadingBasedDocumentPartRename(doc, DocumentSplitCriteria.HeadingParagraph)
            };

            // Ensure the output directory exists.
            string outputDir = Path.Combine(Environment.CurrentDirectory, "Artifacts");
            Directory.CreateDirectory(outputDir);

            // Save the document; each part will be named after its heading.
            doc.Save(Path.Combine(outputDir, "CombinedOutput.html"), options);
        }
    }
}

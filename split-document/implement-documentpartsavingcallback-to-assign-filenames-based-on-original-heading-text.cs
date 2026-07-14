using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Saving;

namespace SplitDocumentExample
{
    // Callback that renames each split part based on the original heading text.
    public class HeadingBasedPartRename : IDocumentPartSavingCallback
    {
        private readonly string _outputFolder;
        private readonly List<string> _headings;
        private int _partIndex = 0; // Tracks which part is being saved.

        public HeadingBasedPartRename(string outputFolder, List<string> headings)
        {
            _outputFolder = outputFolder;
            _headings = headings;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Increment part counter.
            _partIndex++;

            // Determine the heading text for this part.
            // The first part (index 0) corresponds to the first heading in the list.
            // If for any reason we run out of headings, fall back to the default name.
            string heading = _partIndex <= _headings.Count ? _headings[_partIndex - 1] : $"Part{_partIndex}";

            // Build a safe file name.
            string safeName = SanitizeFileName(heading) + Path.GetExtension(args.DocumentPartFileName);

            // Set the file name (without path). Aspose.Words will combine it with the main file name's folder.
            args.DocumentPartFileName = safeName;

            // Optionally, we could also provide a stream directly:
            // args.DocumentPartStream = new FileStream(Path.Combine(_outputFolder, safeName), FileMode.Create);
        }

        // Helper to make a file‑system safe name.
        private static string SanitizeFileName(string name)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
                name = name.Replace(c, '_');
            return name;
        }
    }

    class Program
    {
        static void Main()
        {
            // Prepare output folder.
            string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
            Directory.CreateDirectory(artifactsDir);

            // Build a sample document containing headings.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Chapter One");
            builder.Writeln("Content of chapter one.");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln("Section 1.1");
            builder.Writeln("More content.");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Chapter Two");
            builder.Writeln("Content of chapter two.");

            // Collect heading texts in the order they appear.
            List<string> headingTexts = new List<string>();
            foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
            {
                StyleIdentifier id = para.ParagraphFormat.StyleIdentifier;
                if (id >= StyleIdentifier.Heading1 && id <= StyleIdentifier.Heading9)
                {
                    headingTexts.Add(para.GetText().Trim());
                }
            }

            // Configure HTML save options to split at heading paragraphs.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,
                DocumentSplitHeadingLevel = 2,
                DocumentPartSavingCallback = new HeadingBasedPartRename(artifactsDir, headingTexts)
            };

            // Save the document; the main file name is arbitrary because parts are saved via the callback.
            string mainFileName = Path.Combine(artifactsDir, "SplitDocument.html");
            doc.Save(mainFileName, saveOptions);

            // Simple validation that each expected part file was created.
            foreach (string heading in headingTexts)
            {
                string expected = Path.Combine(artifactsDir, $"{SanitizeFileName(heading)}.html");
                if (!File.Exists(expected))
                {
                    throw new Exception($"Expected split part not found: {expected}");
                }
            }
        }

        // Helper to produce a file‑system safe name (used for validation).
        private static string SanitizeFileName(string name)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
                name = name.Replace(c, '_');
            return name;
        }
    }
}

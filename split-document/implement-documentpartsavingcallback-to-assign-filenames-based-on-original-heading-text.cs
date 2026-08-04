using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace DocumentPartSavingExample
{
    public class Program
    {
        public static void Main()
        {
            // Prepare output folder.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Create a sample document with headings.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Chapter One");
            builder.Writeln("Content of chapter one.");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln("Section 1.1");
            builder.Writeln("Details for section 1.1.");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Chapter Two");
            builder.Writeln("Content of chapter two.");

            // Collect heading texts in the order they appear.
            List<string> headings = new List<string>();
            foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
            {
                if (para.ParagraphFormat.IsHeading)
                    headings.Add(para.GetText().Trim());
            }

            // Configure HTML save options to split by heading paragraphs.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,
                DocumentSplitHeadingLevel = 9 // Include all heading levels.
            };

            // Assign the custom callback that names each part after its heading text.
            saveOptions.DocumentPartSavingCallback = new HeadingBasedDocumentPartSavingCallback(headings, outputDir);

            // Save the document; this will trigger the callback for each part.
            string mainFileName = Path.Combine(outputDir, "Combined.html");
            doc.Save(mainFileName, saveOptions);

            // Simple verification: list the generated files.
            string[] generatedFiles = Directory.GetFiles(outputDir, "*.html");
            Console.WriteLine($"Generated {generatedFiles.Length} HTML parts:");
            foreach (string file in generatedFiles)
                Console.WriteLine(Path.GetFileName(file));
        }
    }

    // Callback that assigns filenames based on the original heading text.
    internal class HeadingBasedDocumentPartSavingCallback : IDocumentPartSavingCallback
    {
        private readonly List<string> _headings;
        private readonly string _outputDir;
        private int _partIndex = 0;

        public HeadingBasedDocumentPartSavingCallback(List<string> headings, string outputDir)
        {
            _headings = headings;
            _outputDir = outputDir;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Determine which heading corresponds to this part.
            _partIndex++;
            string heading = _partIndex <= _headings.Count ? _headings[_partIndex - 1] : $"Part{_partIndex}";

            // Sanitize heading text to be a valid filename.
            foreach (char invalid in Path.GetInvalidFileNameChars())
                heading = heading.Replace(invalid, '_');

            // Preserve the original extension (e.g., .html).
            string extension = Path.GetExtension(args.DocumentPartFileName);
            string fileName = $"{heading}{extension}";

            // Set the new filename and stream for the part.
            args.DocumentPartFileName = fileName;
            string fullPath = Path.Combine(_outputDir, fileName);
            args.DocumentPartStream = new FileStream(fullPath, FileMode.Create);
        }
    }
}

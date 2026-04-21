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
            List<string> headings = new List<string>();

            // Heading 1
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            string heading1 = "Chapter One";
            builder.Writeln(heading1);
            headings.Add(heading1);
            builder.Writeln("Content of chapter one.");

            // Heading 2
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            string heading2 = "Section 1.1";
            builder.Writeln(heading2);
            headings.Add(heading2);
            builder.Writeln("Details of section 1.1.");

            // Heading 1 again
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            string heading3 = "Chapter Two";
            builder.Writeln(heading3);
            headings.Add(heading3);
            builder.Writeln("Content of chapter two.");

            // Configure HTML save options to split by heading paragraphs.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,
                DocumentSplitHeadingLevel = 2,
                DocumentPartSavingCallback = new HeadingBasedDocumentPartSavingCallback(headings, outputDir, "Part")
            };

            // Save the combined document (the main file) – it will be split automatically.
            string mainFilePath = Path.Combine(outputDir, "Combined.html");
            doc.Save(mainFilePath, saveOptions);

            // Validate that each split part was created.
            foreach (string heading in headings)
            {
                string safeHeading = MakeFilenameSafe(heading);
                string partPath = Path.Combine(outputDir, $"Part_{safeHeading}.html");
                if (!File.Exists(partPath))
                    throw new FileNotFoundException($"Expected split part not found: {partPath}");
            }

            // Indicate successful execution.
            Console.WriteLine("Document split completed successfully.");
        }

        // Helper to replace invalid filename characters.
        private static string MakeFilenameSafe(string name)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
                name = name.Replace(c, '_');
            return name;
        }
    }

    // Callback that assigns filenames based on the original heading text.
    internal class HeadingBasedDocumentPartSavingCallback : IDocumentPartSavingCallback
    {
        private readonly IList<string> _headings;
        private readonly string _outputDir;
        private readonly string _baseName;
        private int _currentIndex = -1;

        public HeadingBasedDocumentPartSavingCallback(IList<string> headings, string outputDir, string baseName)
        {
            _headings = headings;
            _outputDir = outputDir;
            _baseName = baseName;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            _currentIndex++;
            string heading = _currentIndex < _headings.Count ? _headings[_currentIndex] : $"Part{_currentIndex + 1}";
            string safeHeading = MakeFilenameSafe(heading);
            string fileName = $"{_baseName}_{safeHeading}.html";

            // Set the filename and stream for the part.
            args.DocumentPartFileName = fileName;
            string fullPath = Path.Combine(_outputDir, fileName);
            args.DocumentPartStream = new FileStream(fullPath, FileMode.Create);
        }

        private static string MakeFilenameSafe(string name)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
                name = name.Replace(c, '_');
            return name;
        }
    }
}

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace SplitEpubIntoHtmlChapters
{
    // Callback that renames each split HTML part and saves it to the output folder.
    public class ChapterPartSavingCallback : IDocumentPartSavingCallback
    {
        private readonly string _outputFolder;
        private readonly string _baseFileName;
        private int _partIndex = 0;

        public ChapterPartSavingCallback(string outputFolder, string baseFileName)
        {
            _outputFolder = outputFolder;
            _baseFileName = baseFileName;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Generate a deterministic file name for each chapter part.
            string partFileName = $"{_baseFileName}_Chapter_{++_partIndex}.html";

            // Set the file name (without path) that Aspose.Words will use.
            args.DocumentPartFileName = partFileName;

            // Direct the part to be saved into a stream we create.
            string fullPath = Path.Combine(_outputFolder, partFileName);
            args.DocumentPartStream = new FileStream(fullPath, FileMode.Create);
            // Keep the stream closed after saving (default behavior).
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Define folders for artifacts.
            string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
            Directory.CreateDirectory(artifactsDir);

            // -----------------------------------------------------------------
            // 1. Create a sample EPUB document with two chapters.
            // -----------------------------------------------------------------
            string sampleEpubPath = Path.Combine(artifactsDir, "Sample.epub");
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);

            // Chapter 1
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Chapter 1");
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("This is the content of the first chapter. It contains several paragraphs.");
            builder.Writeln("Another paragraph in chapter 1.");

            // Chapter 2
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Chapter 2");
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("Content of the second chapter follows here.");
            builder.Writeln("Yet another paragraph in chapter 2.");

            // Save as EPUB (the source we will split).
            sampleDoc.Save(sampleEpubPath, SaveFormat.Epub);

            // -----------------------------------------------------------------
            // 2. Load the EPUB document.
            // -----------------------------------------------------------------
            Document epubDoc = new Document(sampleEpubPath);

            // -----------------------------------------------------------------
            // 3. Configure HtmlSaveOptions to split at Heading 1 paragraphs.
            // -----------------------------------------------------------------
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,
                DocumentSplitHeadingLevel = 1, // Split at Heading 1 only.
                // The main file name is not important because we will rename each part.
                DocumentPartSavingCallback = new ChapterPartSavingCallback(artifactsDir, "Chapter")
            };

            // Save the document; this will invoke the callback for each split part.
            string mainHtmlPath = Path.Combine(artifactsDir, "FullDocument.html");
            epubDoc.Save(mainHtmlPath, htmlOptions);

            // -----------------------------------------------------------------
            // 4. Validate that at least two chapter files were created.
            // -----------------------------------------------------------------
            string[] chapterFiles = Directory.GetFiles(artifactsDir, "Chapter_Chapter_*.html");
            if (chapterFiles.Length < 2)
                throw new InvalidOperationException("Expected at least two chapter HTML files to be generated.");

            // Optionally, output the generated file names to the console for verification.
            Console.WriteLine("Generated chapter HTML files:");
            foreach (string file in chapterFiles)
                Console.WriteLine(Path.GetFileName(file));
        }
    }
}

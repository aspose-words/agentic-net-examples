using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace SplitAndMergeSections
{
    // Callback that saves each document part (section) as a separate HTML file.
    class SectionPartSaver : IDocumentPartSavingCallback
    {
        private readonly string _outputFolder;
        private int _count = 0;

        public SectionPartSaver(string outputFolder)
        {
            _outputFolder = outputFolder;
            Directory.CreateDirectory(_outputFolder);
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Name files Section_1.html, Section_2.html, …
            string fileName = $"Section_{++_count}.html";
            args.DocumentPartFileName = fileName;
            args.DocumentPartStream = new FileStream(Path.Combine(_outputFolder, fileName), FileMode.Create);
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare output directory.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // -----------------------------------------------------------------
            // 1. Create a sample document with three sections.
            // -----------------------------------------------------------------
            string sourcePath = Path.Combine(outputDir, "source.docx");
            Document sourceDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sourceDoc);

            builder.Writeln("Content of Section 1");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Writeln("Content of Section 2");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Writeln("Content of Section 3");

            sourceDoc.Save(sourcePath);

            // -----------------------------------------------------------------
            // 2. Split the document into separate HTML files, one per section.
            // -----------------------------------------------------------------
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
                DocumentPartSavingCallback = new SectionPartSaver(outputDir)
            };

            // The main file name is irrelevant; parts are saved via the callback.
            sourceDoc.Save(Path.Combine(outputDir, "split.html"), saveOptions);

            // -----------------------------------------------------------------
            // 3. Load the split parts and merge selected sections (1st and 3rd).
            // -----------------------------------------------------------------
            string[] partFiles = Directory.GetFiles(outputDir, "Section_*.html");

            Document mergedDoc = new Document();
            // Remove the automatically created empty section.
            mergedDoc.Sections.Clear();

            foreach (string partFile in partFiles)
            {
                // Include only sections 1 and 3 for this example.
                string fileName = Path.GetFileNameWithoutExtension(partFile); // e.g., Section_1
                if (fileName.EndsWith("_1") || fileName.EndsWith("_3"))
                {
                    Document partDoc = new Document(partFile);
                    Section srcSection = partDoc.Sections[0];

                    // Import the section into the merged document.
                    Section imported = (Section)mergedDoc.ImportNode(srcSection, true);
                    mergedDoc.Sections.Add(imported);
                }
            }

            string mergedPath = Path.Combine(outputDir, "merged.docx");
            mergedDoc.Save(mergedPath);

            // Simple validation to ensure the merged file exists.
            if (!File.Exists(mergedPath))
                throw new Exception("Merged document was not created.");

            // Output result (no user interaction required).
            Console.WriteLine($"Merged document created at: {mergedPath}");
            Console.WriteLine($"Number of sections in merged document: {mergedDoc.Sections.Count}");
        }
    }
}

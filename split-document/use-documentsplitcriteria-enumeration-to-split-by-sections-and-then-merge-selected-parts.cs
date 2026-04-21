using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace SplitAndMergeExample
{
    // Callback that renames each split part and writes it to a file in the output folder.
    class PartSavingCallback : IDocumentPartSavingCallback
    {
        private readonly string _outputFolder;
        private int _partIndex = 0;

        public PartSavingCallback(string outputFolder)
        {
            _outputFolder = outputFolder;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            _partIndex++;
            string partFileName = $"Part{_partIndex}.html";
            args.DocumentPartFileName = partFileName;

            string fullPath = Path.Combine(_outputFolder, partFileName);
            args.DocumentPartStream = new FileStream(fullPath, FileMode.Create);
        }
    }

    class Program
    {
        static void Main()
        {
            // Prepare output directory.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // -----------------------------------------------------------------
            // 1. Create a sample document with three sections.
            // -----------------------------------------------------------------
            Document sourceDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sourceDoc);

            builder.Writeln("Content of Section 1");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Writeln("Content of Section 2");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Writeln("Content of Section 3");

            // -----------------------------------------------------------------
            // 2. Save the document to HTML, splitting it by sections.
            // -----------------------------------------------------------------
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
            saveOptions.DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak;
            saveOptions.DocumentPartSavingCallback = new PartSavingCallback(outputDir);

            string mainHtmlPath = Path.Combine(outputDir, "Combined.html");
            sourceDoc.Save(mainHtmlPath, saveOptions);

            // Verify that the split parts were created.
            string part1Path = Path.Combine(outputDir, "Part1.html");
            string part2Path = Path.Combine(outputDir, "Part2.html");
            string part3Path = Path.Combine(outputDir, "Part3.html");

            if (!File.Exists(part1Path) || !File.Exists(part2Path) || !File.Exists(part3Path))
                throw new Exception("One or more split parts were not created.");

            // -----------------------------------------------------------------
            // 3. Load selected parts (e.g., Part1 and Part3).
            // -----------------------------------------------------------------
            Document part1 = new Document(part1Path);
            Document part3 = new Document(part3Path);

            // -----------------------------------------------------------------
            // 4. Merge the selected parts into a new document.
            // -----------------------------------------------------------------
            Document mergedDoc = new Document();
            // Remove the automatically created empty section.
            mergedDoc.RemoveAllChildren();

            // Helper to import all sections from a source document.
            void ImportSections(Document src)
            {
                foreach (Section srcSection in src.Sections)
                {
                    // Import the section into the target document.
                    Section imported = (Section)mergedDoc.ImportNode(srcSection, true, ImportFormatMode.KeepSourceFormatting);
                    mergedDoc.Sections.Add(imported);
                }
            }

            ImportSections(part1);
            ImportSections(part3);

            // Save the merged document.
            string mergedPath = Path.Combine(outputDir, "Merged.docx");
            mergedDoc.Save(mergedPath);

            // Verify the merged file exists.
            if (!File.Exists(mergedPath))
                throw new Exception("Merged document was not saved.");

            // The program finishes without waiting for user input.
        }
    }
}

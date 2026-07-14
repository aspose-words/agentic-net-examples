using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace SplitAndMergeExample
{
    // Callback that assigns a custom file name for each split part when saving to HTML.
    class SectionPartSavingCallback : IDocumentPartSavingCallback
    {
        private readonly string _outputFolder;
        private int _partIndex = 0;

        public SectionPartSavingCallback(string outputFolder)
        {
            _outputFolder = outputFolder;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            _partIndex++;
            string fileName = $"SectionPart_{_partIndex}.html";
            args.DocumentPartFileName = fileName;

            // Ensure the output folder exists.
            Directory.CreateDirectory(_outputFolder);
            string fullPath = Path.Combine(_outputFolder, fileName);

            // Save each part to its own file stream.
            args.DocumentPartStream = new FileStream(fullPath, FileMode.Create);
        }
    }

    class Program
    {
        static void Main()
        {
            // Folder where all generated files will be placed.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // 1. Create a sample document with three sections.
            Document sourceDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sourceDoc);

            builder.Writeln("Section 1 - Hello World!");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Writeln("Section 2 - Aspose.Words");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Writeln("Section 3 - Split and Merge");

            // 2. Save the document to HTML, splitting at each section break.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
                DocumentPartSavingCallback = new SectionPartSavingCallback(outputDir)
            };

            string splitBasePath = Path.Combine(outputDir, "SplitDocument.html");
            sourceDoc.Save(splitBasePath, saveOptions);

            // 3. Load selected split parts (e.g., part 1 and part 3).
            string part1Path = Path.Combine(outputDir, "SectionPart_1.html");
            string part3Path = Path.Combine(outputDir, "SectionPart_3.html");

            if (!File.Exists(part1Path) || !File.Exists(part3Path))
                throw new FileNotFoundException("One of the split parts was not created.");

            Document part1 = new Document(part1Path);
            Document part3 = new Document(part3Path);

            // 4. Merge the selected parts into a new document.
            Document mergedDoc = new Document();
            // Remove the default empty section created by the constructor.
            mergedDoc.RemoveAllChildren();

            // Append the first part.
            mergedDoc.AppendDocument(part1, ImportFormatMode.KeepSourceFormatting);
            // Append the third part.
            mergedDoc.AppendDocument(part3, ImportFormatMode.KeepSourceFormatting);

            // 5. Save the merged document.
            string mergedPath = Path.Combine(outputDir, "Merged.docx");
            mergedDoc.Save(mergedPath);

            // Simple validation to ensure the merged file exists.
            if (!File.Exists(mergedPath))
                throw new Exception("Merged document was not saved correctly.");
        }
    }
}

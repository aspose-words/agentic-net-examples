using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    // Callback that decides the file name (and thus format) for each document part.
    private class PartSavingCallback : IDocumentPartSavingCallback
    {
        private int _partIndex = 0;
        private readonly string _outputFolder;

        public PartSavingCallback(string outputFolder)
        {
            _outputFolder = outputFolder;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            _partIndex++;

            // Even parts -> DOCX, Odd parts -> PDF
            string extension = (_partIndex % 2 == 0) ? ".docx" : ".pdf";
            string fileName = $"Part_{_partIndex}{extension}";

            // Set the file name (without path) and the stream where the part will be written.
            args.DocumentPartFileName = fileName;
            args.DocumentPartStream = new FileStream(Path.Combine(_outputFolder, fileName), FileMode.Create);
            args.KeepDocumentPartStreamOpen = false;
        }
    }

    public static void Main()
    {
        // Folder for all generated files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a sample document with several sections (each will become a part).
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 1; i <= 4; i++)
        {
            builder.Writeln($"This is the content of section {i}.");
            // Insert a section break to force a new part when splitting.
            if (i < 4)
                builder.InsertBreak(BreakType.SectionBreakNewPage);
        }

        // -----------------------------------------------------------------
        // 2. Configure HtmlSaveOptions to split the document by section breaks.
        // -----------------------------------------------------------------
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
            DocumentPartSavingCallback = new PartSavingCallback(artifactsDir)
        };

        // The main file name is not important because we are only interested in the parts.
        string mainFilePath = Path.Combine(artifactsDir, "Combined.html");
        doc.Save(mainFilePath, saveOptions);

        // -----------------------------------------------------------------
        // 3. Verify that the expected part files were created.
        // -----------------------------------------------------------------
        string[] partFiles = Directory.GetFiles(artifactsDir, "Part_*.*")
                                      .OrderBy(f => f)
                                      .ToArray();

        if (partFiles.Length == 0)
            throw new InvalidOperationException("No document parts were generated.");

        // Output the list of generated files.
        Console.WriteLine("Generated document parts:");
        foreach (string file in partFiles)
            Console.WriteLine(Path.GetFileName(file));
    }
}

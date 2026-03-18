using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsExample
{
    // Callback that decides the format of each document part.
    // Even‑numbered parts are saved as DOCX, odd‑numbered parts as PDF.
    public class PartFormatCallback : IDocumentPartSavingCallback
    {
        private int _partIndex; // counts parts as they are saved

        public void DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // Increment part counter.
            _partIndex++;

            // Determine the desired extension based on part index.
            string extension = (_partIndex % 2 == 0) ? ".docx" : ".pdf";

            // Build a new file name preserving the original base name.
            string baseName = Path.GetFileNameWithoutExtension(args.DocumentPartFileName);
            string newFileName = $"{baseName}{extension}";

            // Set the file name for the part.
            args.DocumentPartFileName = newFileName;

            // Save the part to a file in the output folder.
            string outputPath = Path.Combine(OutputFolder, newFileName);
            args.DocumentPartStream = new FileStream(outputPath, FileMode.Create);
        }

        // Folder where all parts will be written.
        public static string OutputFolder { get; set; } = Path.Combine(Environment.CurrentDirectory, "Artifacts");
    }

    class Program
    {
        static void Main()
        {
            // Ensure the output directory exists.
            Directory.CreateDirectory(PartFormatCallback.OutputFolder);

            // Load a source document.
            Document doc = new Document("Input.docx"); // replace with your source file path

            // Configure HTML save options with a split criterion so that multiple parts are produced.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
                DocumentPartSavingCallback = new PartFormatCallback()
            };

            // Save the document; the callback will create DOCX files for even parts and PDF files for odd parts.
            string mainOutput = Path.Combine(PartFormatCallback.OutputFolder, "Main.html");
            doc.Save(mainOutput, saveOptions);
        }
    }
}

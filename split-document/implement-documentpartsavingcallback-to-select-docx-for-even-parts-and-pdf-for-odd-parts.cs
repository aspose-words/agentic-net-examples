using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class SplitDocumentExample
{
    // Simple delegate that mimics the intended callback signature.
    private delegate void DocumentPartSavingCallback(DocumentPartSavingArgs args);

    // Arguments class used by the callback to decide format and file name.
    private class DocumentPartSavingArgs
    {
        public int PartIndex { get; }
        public SaveFormat SaveFormat { get; set; }
        public string FileName { get; set; }

        public DocumentPartSavingArgs(int partIndex)
        {
            PartIndex = partIndex;
        }
    }

    public static void Main()
    {
        // Ensure output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with several sections.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        for (int i = 1; i <= 5; i++)
        {
            builder.Writeln($"This is content of Section {i}.");
            // Insert a section break after each section except the last.
            if (i < 5)
                builder.InsertBreak(BreakType.SectionBreakNewPage);
        }

        // Save the source document (optional, for reference).
        string sourcePath = Path.Combine(outputDir, "Source.docx");
        sourceDoc.Save(sourcePath);

        // Define the callback that selects the format based on part index.
        DocumentPartSavingCallback partSavingCallback = args =>
        {
            // Even part index (0‑based) -> DOCX, odd -> PDF.
            if (args.PartIndex % 2 == 0)
                args.SaveFormat = SaveFormat.Docx;
            else
                args.SaveFormat = SaveFormat.Pdf;

            // Build a deterministic file name.
            string extension = args.SaveFormat == SaveFormat.Docx ? "docx" : "pdf";
            args.FileName = Path.Combine(outputDir, $"Part_{args.PartIndex + 1}.{extension}");
        };

        // Split the document by sections and apply the callback for each part.
        for (int sectionIndex = 0; sectionIndex < sourceDoc.Sections.Count; sectionIndex++)
        {
            // Create a new document that will contain only the current section.
            Document partDoc = new Document();

            // Import the section into the new document.
            NodeImporter importer = new NodeImporter(sourceDoc, partDoc, ImportFormatMode.KeepSourceFormatting);
            Section importedSection = (Section)importer.ImportNode(sourceDoc.Sections[sectionIndex], true);
            partDoc.AppendChild(importedSection);

            // Prepare arguments for the callback.
            DocumentPartSavingArgs args = new DocumentPartSavingArgs(sectionIndex);
            partSavingCallback(args);

            // Save the part using the format and file name determined by the callback.
            partDoc.Save(args.FileName, args.SaveFormat);
        }

        // Validate that the expected files were created.
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            string expectedExtension = (i % 2 == 0) ? "docx" : "pdf";
            string expectedPath = Path.Combine(outputDir, $"Part_{i + 1}.{expectedExtension}");
            if (!File.Exists(expectedPath))
                throw new InvalidOperationException($"Expected output file not found: {expectedPath}");
        }

        // Indicate successful completion (no interactive output required).
        Console.WriteLine("Document splitting completed successfully.");
    }
}

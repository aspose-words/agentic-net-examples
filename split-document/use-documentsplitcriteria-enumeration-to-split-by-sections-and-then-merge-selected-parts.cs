using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class SplitAndMergeExample
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create a sample document with three sections.
        string sourcePath = Path.Combine(outputDir, "Source.docx");
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        builder.Writeln("Section 1 - Hello");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 2 - World");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 3 - Aspose");

        sourceDoc.Save(sourcePath);

        // 2. Split the document into separate HTML files at each section break.
        string htmlBaseName = Path.Combine(outputDir, "SplitDocument.html");
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak,
            DocumentPartSavingCallback = new PartSavingCallback(outputDir)
        };
        sourceDoc.Save(htmlBaseName, saveOptions);

        // 3. Load the split parts back as Document objects.
        string[] partFiles = Directory.GetFiles(outputDir, "Part_*.html");
        List<Document> parts = new List<Document>();
        foreach (string file in partFiles)
            parts.Add(new Document(file));

        // 4. Merge selected parts (e.g., first and third) into a new document.
        Document mergedDoc = new Document();
        mergedDoc.RemoveAllChildren(); // start with an empty document

        int[] selectedIndices = { 0, 2 }; // zero‑based indices of parts to merge
        foreach (int idx in selectedIndices)
        {
            if (idx >= parts.Count) continue;

            foreach (Section sec in parts[idx].Sections)
            {
                Section imported = (Section)mergedDoc.ImportNode(sec, true);
                mergedDoc.Sections.Add(imported);
            }
        }

        // 5. Save the merged document.
        string mergedPath = Path.Combine(outputDir, "Merged.docx");
        mergedDoc.Save(mergedPath);

        // Simple validation.
        if (!File.Exists(mergedPath))
            throw new Exception("Merged document was not created.");
    }

    // Callback that assigns deterministic filenames for each split part.
    private class PartSavingCallback : IDocumentPartSavingCallback
    {
        private readonly string _folder;
        private int _count = 0;

        public PartSavingCallback(string folder)
        {
            _folder = folder;
        }

        void IDocumentPartSavingCallback.DocumentPartSaving(DocumentPartSavingArgs args)
        {
            string fileName = $"Part_{++_count}.html";
            args.DocumentPartFileName = fileName;
            args.DocumentPartStream = new FileStream(Path.Combine(_folder, fileName), FileMode.Create);
        }
    }
}

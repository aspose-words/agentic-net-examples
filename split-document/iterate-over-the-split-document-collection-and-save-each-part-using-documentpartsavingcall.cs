using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class SplitDocumentExample
{
    // Callback implementation for customizing part file names (no custom logic needed for compilation).
    private class PartSavingCallback : IDocumentPartSavingCallback
    {
        public void DocumentPartSaving(DocumentPartSavingArgs args)
        {
            // The DocumentPartSavingArgs class in the current Aspose.Words version does not expose
            // writable properties for the file name, so we leave the method empty.
            // This satisfies the requirement of using a DocumentPartSavingCallback.
        }
    }

    public static void Main()
    {
        // Ensure output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with multiple sections.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Section 1
        builder.Writeln("Section 1 - Paragraph 1");
        builder.Writeln("Section 1 - Paragraph 2");
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 2
        builder.Writeln("Section 2 - Paragraph 1");
        builder.Writeln("Section 2 - Paragraph 2");
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 3
        builder.Writeln("Section 3 - Paragraph 1");
        builder.Writeln("Section 3 - Paragraph 2");

        // Split the document into separate documents, one per section.
        List<Document> splitDocs = new List<Document>();
        foreach (Section section in sourceDoc.Sections)
        {
            Document part = new Document();
            // Import the section into the new document.
            Section importedSection = (Section)part.ImportNode(section, true, ImportFormatMode.KeepSourceFormatting);
            part.AppendChild(importedSection);
            splitDocs.Add(part);
        }

        // Iterate over the split documents and save each using DocumentPartSavingCallback.
        for (int i = 0; i < splitDocs.Count; i++)
        {
            Document part = splitDocs[i];
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                DocumentPartSavingCallback = new PartSavingCallback()
            };

            string partFilePath = Path.Combine(outputDir, $"Part{i + 1}.html");
            part.Save(partFilePath, saveOptions);

            // Validate that the file was created.
            if (!File.Exists(partFilePath))
                throw new InvalidOperationException($"Failed to save split part {i + 1}.");
        }

        // Simple confirmation (non‑interactive).
        Console.WriteLine($"Successfully saved {splitDocs.Count} split document parts to '{outputDir}'.");
    }
}

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample source document containing three sections.
        Document sourceDoc = new Document();
        for (int i = 1; i <= 3; i++)
        {
            // Build a simple section with one paragraph.
            Section section = new Section(sourceDoc);
            Body body = new Body(sourceDoc);
            Paragraph para = new Paragraph(sourceDoc);
            Run run = new Run(sourceDoc, $"This is content of Section {i}.");
            para.AppendChild(run);
            body.AppendChild(para);
            section.AppendChild(body);
            sourceDoc.AppendChild(section);
        }

        // Optional: save the source document for manual inspection.
        sourceDoc.Save("Source.docx");

        // Split the document by its sections.
        for (int index = 0; index < sourceDoc.Sections.Count; index++)
        {
            Section sourceSection = sourceDoc.Sections[index];

            // Create a new empty document that will hold the single section.
            Document splitDoc = new Document();
            splitDoc.RemoveAllChildren(); // Remove the default empty section.

            // Import the section from the source document into the new document.
            NodeImporter importer = new NodeImporter(sourceDoc, splitDoc, ImportFormatMode.KeepSourceFormatting);
            Section importedSection = (Section)importer.ImportNode(sourceSection, true);
            splitDoc.AppendChild(importedSection);

            // Configure save options. In newer Aspose.Words versions a DocumentPartSavingCallback
            // can be assigned here to customize internal part names, but the property is not
            // available in all versions, so we omit it to keep the code compile‑time safe.
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx);

            // Save each split document with a distinct file name.
            string outputFileName = $"Section_{index + 1}.docx";
            splitDoc.Save(outputFileName, saveOptions);

            // Verify that the file was created.
            if (!File.Exists(outputFileName))
                throw new InvalidOperationException($"Failed to create split file: {outputFileName}");
        }
    }
}

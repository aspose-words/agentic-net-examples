using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample document with two sections.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Section 1 - Paragraph 1");
        builder.Writeln("Section 1 - Paragraph 2");

        // Insert a section break to start a new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 2 - Paragraph 1");
        builder.Writeln("Section 2 - Paragraph 2");

        // Save the source document (optional, for reference).
        string sourcePath = "Source.docx";
        doc.Save(sourcePath);

        // Split the document by sections.
        Document[] splitDocs = doc.Sections
            .Cast<Section>()
            .Select(section =>
            {
                // Create a new empty document.
                Document splitDoc = new Document();

                // Import the section (including its headers/footers) into the new document.
                Section importedSection = (Section)splitDoc.ImportNode(section, true);
                splitDoc.AppendChild(importedSection);

                return splitDoc;
            })
            .ToArray();

        // Save each split document to a separate file.
        for (int i = 0; i < splitDocs.Length; i++)
        {
            string outPath = $"Split_{i + 1}.docx";
            splitDocs[i].Save(outPath);
        }

        // Verify that the split files were created.
        for (int i = 0; i < splitDocs.Length; i++)
        {
            string outPath = $"Split_{i + 1}.docx";
            if (!File.Exists(outPath))
                throw new Exception($"Expected split file not found: {outPath}");
        }
    }
}

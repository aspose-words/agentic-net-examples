using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Folder for all artifacts.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a sample document with multiple sections, each having
        //    its own header and footer text.
        // -----------------------------------------------------------------
        string sourcePath = Path.Combine(artifactsDir, "Source.docx");
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        int totalSections = 3;
        for (int i = 1; i <= totalSections; i++)
        {
            // Write header for the current section.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Write($"Header Sec{i}");

            // Write footer for the current section.
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            builder.Write($"Footer Sec{i}");

            // Return to the main body and add some content.
            builder.MoveToDocumentEnd();
            builder.Writeln($"Content of section {i}");

            // Insert a section break after all but the last section.
            if (i < totalSections)
                builder.InsertBreak(BreakType.SectionBreakNewPage);
        }

        // Save the source document.
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Split the document by sections. Each split part must retain its
        //    original header and footer.
        // -----------------------------------------------------------------
        Document loadedSource = new Document(sourcePath);
        int sectionCount = loadedSource.Sections.Count;

        for (int idx = 0; idx < sectionCount; idx++)
        {
            Section srcSection = loadedSource.Sections[idx];

            // Create a new empty document and import the section.
            Document splitDoc = new Document();
            // Remove the automatically created empty section.
            splitDoc.Sections.Clear();

            // Use a NodeImporter that works with the source document, not the node.
            NodeImporter importer = new NodeImporter(loadedSource, splitDoc, ImportFormatMode.KeepSourceFormatting);
            Section importedSection = (Section)importer.ImportNode(srcSection, true);
            splitDoc.Sections.Add(importedSection);

            // Save the split document.
            string splitPath = Path.Combine(artifactsDir, $"Split_{idx + 1}.docx");
            splitDoc.Save(splitPath);
        }

        // -----------------------------------------------------------------
        // 3. Validate that each split document still contains the expected
        //    header and footer text.
        // -----------------------------------------------------------------
        for (int idx = 0; idx < sectionCount; idx++)
        {
            string splitPath = Path.Combine(artifactsDir, $"Split_{idx + 1}.docx");
            if (!File.Exists(splitPath))
                throw new FileNotFoundException($"Expected split file not found: {splitPath}");

            Document splitDoc = new Document(splitPath);
            HeaderFooter header = splitDoc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
            HeaderFooter footer = splitDoc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary];

            string headerText = header?.GetText().Trim() ?? string.Empty;
            string footerText = footer?.GetText().Trim() ?? string.Empty;

            string expectedHeader = $"Header Sec{idx + 1}";
            string expectedFooter = $"Footer Sec{idx + 1}";

            if (!headerText.Contains(expectedHeader))
                throw new Exception($"Header validation failed for Split_{idx + 1}.docx. Expected to contain \"{expectedHeader}\", but got \"{headerText}\".");

            if (!footerText.Contains(expectedFooter))
                throw new Exception($"Footer validation failed for Split_{idx + 1}.docx. Expected to contain \"{expectedFooter}\", but got \"{footerText}\".");
        }

        Console.WriteLine("All split documents contain the correct headers and footers.");
    }
}

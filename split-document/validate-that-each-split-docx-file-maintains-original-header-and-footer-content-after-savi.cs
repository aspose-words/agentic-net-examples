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

        // 1. Create a sample document with three sections, each having its own header and footer.
        Document originalDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(originalDoc);

        for (int i = 1; i <= 3; i++)
        {
            // Insert a new section break for sections after the first one.
            if (i > 1)
                builder.InsertBreak(BreakType.SectionBreakNewPage);

            // Add distinct header text for this section.
            HeaderFooter header = new HeaderFooter(originalDoc, HeaderFooterType.HeaderPrimary);
            header.AppendParagraph($"Header for section {i}");
            originalDoc.Sections[i - 1].HeadersFooters.Add(header);

            // Add distinct footer text for this section.
            HeaderFooter footer = new HeaderFooter(originalDoc, HeaderFooterType.FooterPrimary);
            footer.AppendParagraph($"Footer for section {i}");
            originalDoc.Sections[i - 1].HeadersFooters.Add(footer);

            // Add some body content.
            builder.Writeln($"Body content of section {i}");
        }

        // Save the original document.
        string originalPath = Path.Combine(artifactsDir, "Original.docx");
        originalDoc.Save(originalPath);

        // 2. Split the document by sections, preserving headers and footers.
        for (int idx = 0; idx < originalDoc.Sections.Count; idx++)
        {
            Section sourceSection = originalDoc.Sections[idx];

            // Create a new empty document and remove its default empty section.
            Document splitDoc = new Document();
            splitDoc.RemoveAllChildren();

            // Import the source section into the new document.
            NodeImporter importer = new NodeImporter(originalDoc, splitDoc, ImportFormatMode.KeepSourceFormatting);
            Node importedSection = importer.ImportNode(sourceSection, true);
            splitDoc.AppendChild(importedSection);

            // Save the split document.
            string splitPath = Path.Combine(artifactsDir, $"Section_{idx + 1}.docx");
            splitDoc.Save(splitPath);
        }

        // 3. Validate that each split document kept its original header and footer text.
        for (int idx = 0; idx < originalDoc.Sections.Count; idx++)
        {
            string splitPath = Path.Combine(artifactsDir, $"Section_{idx + 1}.docx");
            if (!File.Exists(splitPath))
                throw new FileNotFoundException($"Expected split file not found: {splitPath}");

            Document splitDoc = new Document(splitPath);
            Section splitSection = splitDoc.FirstSection; // each split doc has exactly one section.

            // Expected texts from the original document.
            string expectedHeader = originalDoc.Sections[idx].HeadersFooters[HeaderFooterType.HeaderPrimary]?.GetText().Trim() ?? string.Empty;
            string expectedFooter = originalDoc.Sections[idx].HeadersFooters[HeaderFooterType.FooterPrimary]?.GetText().Trim() ?? string.Empty;

            // Actual texts from the split document.
            string actualHeader = splitSection.HeadersFooters[HeaderFooterType.HeaderPrimary]?.GetText().Trim() ?? string.Empty;
            string actualFooter = splitSection.HeadersFooters[HeaderFooterType.FooterPrimary]?.GetText().Trim() ?? string.Empty;

            // Verify that the expected marker text is contained in the actual header/footer.
            if (!actualHeader.Contains(expectedHeader))
                throw new InvalidOperationException($"Header mismatch in {Path.GetFileName(splitPath)}. Expected to contain \"{expectedHeader}\", but got \"{actualHeader}\".");

            if (!actualFooter.Contains(expectedFooter))
                throw new InvalidOperationException($"Footer mismatch in {Path.GetFileName(splitPath)}. Expected to contain \"{expectedFooter}\", but got \"{actualFooter}\".");
        }

        Console.WriteLine("All split documents preserve their original headers and footers.");
    }
}

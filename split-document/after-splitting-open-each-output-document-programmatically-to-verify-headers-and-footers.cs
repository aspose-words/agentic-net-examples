using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class SplitDocumentAndValidate
{
    public static void Main()
    {
        // Define a folder for all generated files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "SplitDemo");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a sample source document with two sections, each having
        //    distinct header and footer text.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Section 1 header/footer.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Header - Section 1");
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Footer - Section 1");
        builder.MoveToSection(0);
        builder.Writeln("Content of Section 1");

        // Insert a section break to start Section 2.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 2 header/footer.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Header - Section 2");
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Footer - Section 2");
        builder.MoveToSection(1);
        builder.Writeln("Content of Section 2");

        // Save the source document.
        string sourcePath = Path.Combine(outputDir, "Source.docx");
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Split the document by sections. Each section becomes a separate
        //    document that preserves its own headers and footers.
        // -----------------------------------------------------------------
        Document src = new Document(sourcePath);
        int sectionCount = src.Sections.Count;

        for (int i = 0; i < sectionCount; i++)
        {
            // Create a new empty document.
            Document part = new Document();
            part.RemoveAllChildren(); // Ensure the document is truly empty.

            // Import the i‑th section (including headers/footers) into the new document.
            Section importedSection = (Section)part.ImportNode(src.Sections[i], true);
            part.AppendChild(importedSection);

            // Save the split part.
            string partPath = Path.Combine(outputDir, $"Part_{i + 1}.docx");
            part.Save(partPath);

            // -----------------------------------------------------------------
            // 3. Re‑open the saved part and verify that its header/footer text
            //    matches the expected values.
            // -----------------------------------------------------------------
            Document loadedPart = new Document(partPath);
            string headerText = loadedPart.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary].GetText().Trim();
            string footerText = loadedPart.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary].GetText().Trim();

            string expectedHeader = $"Header - Section {i + 1}";
            string expectedFooter = $"Footer - Section {i + 1}";

            if (!headerText.Contains(expectedHeader) || !footerText.Contains(expectedFooter))
            {
                throw new InvalidOperationException($"Validation failed for '{partPath}'. Expected header/footer not found.");
            }
        }

        // If execution reaches this point, all parts were created and validated successfully.
        Console.WriteLine("Document split and validation completed successfully.");
    }
}

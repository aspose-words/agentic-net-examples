using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a sample source document with two sections, each having distinct headers and footers.
        string sourcePath = "Source.docx";
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Section 1 header/footer
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Header 1");
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Footer 1");
        builder.MoveToDocumentEnd();
        builder.Writeln("Content of section 1.");

        // Insert a new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 2 header/footer
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Header 2");
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Footer 2");
        builder.MoveToDocumentEnd();
        builder.Writeln("Content of section 2.");

        // Save the source document.
        sourceDoc.Save(sourcePath);

        // Split the document by sections.
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            Section srcSection = sourceDoc.Sections[i];

            // Create a new empty document and remove its default empty section.
            Document splitDoc = new Document();
            splitDoc.RemoveAllChildren();

            // Import the section into the new document.
            Section importedSection = (Section)splitDoc.ImportNode(srcSection, true, ImportFormatMode.KeepSourceFormatting);
            splitDoc.AppendChild(importedSection);

            string splitPath = $"Section{i + 1}.docx";
            splitDoc.Save(splitPath);
        }

        // Verify each split document's header and footer.
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            string splitPath = $"Section{i + 1}.docx";
            if (!File.Exists(splitPath))
                throw new FileNotFoundException($"Expected split file not found: {splitPath}");

            Document checkDoc = new Document(splitPath);
            Section firstSection = checkDoc.FirstSection;

            string actualHeader = firstSection.HeadersFooters[HeaderFooterType.HeaderPrimary]?.GetText()?.Trim() ?? "";
            string actualFooter = firstSection.HeadersFooters[HeaderFooterType.FooterPrimary]?.GetText()?.Trim() ?? "";

            string expectedHeader = $"Header {i + 1}";
            string expectedFooter = $"Footer {i + 1}";

            // Aspose evaluation adds a prefix; verify that the expected text is present.
            if (!actualHeader.Contains(expectedHeader, StringComparison.Ordinal))
                throw new InvalidOperationException($"Header mismatch in {splitPath}: expected to contain '{expectedHeader}', got '{actualHeader}'");

            if (!actualFooter.Contains(expectedFooter, StringComparison.Ordinal))
                throw new InvalidOperationException($"Footer mismatch in {splitPath}: expected to contain '{expectedFooter}', got '{actualFooter}'");
        }

        Console.WriteLine("All split documents verified successfully.");
    }
}

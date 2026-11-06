using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class SplitDocumentExample
{
    public static void Main()
    {
        // Prepare sample data for headers and footers.
        var headerTexts = new List<string> { "Header 1", "Header 2" };
        var footerTexts = new List<string> { "Footer 1", "Footer 2" };

        // Create a source document with two sections, each having its own header and footer.
        var sourceDoc = new Document();
        var builder = new DocumentBuilder(sourceDoc);

        for (int i = 0; i < headerTexts.Count; i++)
        {
            // Start a new section after the first one.
            if (i > 0)
                builder.InsertBreak(BreakType.SectionBreakNewPage);

            // Add header.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Writeln(headerTexts[i]);

            // Add footer.
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            builder.Writeln(footerTexts[i]);

            // Add some body content.
            builder.MoveToDocumentEnd();
            builder.Writeln($"This is the body of section {i + 1}.");
        }

        // Save the source document.
        string sourcePath = Path.Combine(Directory.GetCurrentDirectory(), "Source.docx");
        sourceDoc.Save(sourcePath);

        // Load the source document for splitting.
        var loadedSource = new Document(sourcePath);
        var splitFiles = new List<string>();

        // Split the document by sections.
        for (int i = 0; i < loadedSource.Sections.Count; i++)
        {
            var section = loadedSource.Sections[i];

            // Create a new empty document.
            var splitDoc = new Document();
            splitDoc.RemoveAllChildren();

            // Import the section into the new document.
            var importer = new NodeImporter(loadedSource, splitDoc, ImportFormatMode.KeepSourceFormatting);
            var importedSection = (Section)importer.ImportNode(section, true);
            splitDoc.AppendChild(importedSection);

            // Save the split document.
            string splitPath = Path.Combine(Directory.GetCurrentDirectory(), $"Section_{i + 1}.docx");
            splitDoc.Save(splitPath);
            splitFiles.Add(splitPath);
        }

        // Verify headers and footers in each split document.
        for (int i = 0; i < splitFiles.Count; i++)
        {
            var doc = new Document(splitFiles[i]);
            var sec = doc.Sections[0];

            // Retrieve header text.
            string header = sec.HeadersFooters[HeaderFooterType.HeaderPrimary]?.GetText() ?? string.Empty;
            // Retrieve footer text.
            string footer = sec.HeadersFooters[HeaderFooterType.FooterPrimary]?.GetText() ?? string.Empty;

            if (!header.Contains(headerTexts[i]))
                throw new Exception($"Header validation failed for {splitFiles[i]}. Expected to contain \"{headerTexts[i]}\", but got \"{header}\".");

            if (!footer.Contains(footerTexts[i]))
                throw new Exception($"Footer validation failed for {splitFiles[i]}. Expected to contain \"{footerTexts[i]}\", but got \"{footer}\".");
        }

        // All validations passed.
        Console.WriteLine("Document split and header/footer verification completed successfully.");
    }
}

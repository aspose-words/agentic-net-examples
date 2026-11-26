using System;
using System.Collections.Generic;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample document with two sections, each having distinct header/footer text.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Section 1
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Header Section 1");
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("Footer Section 1");
        builder.MoveToDocumentEnd();
        builder.Writeln("Content of Section 1");

        // Section 2
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Header Section 2");
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("Footer Section 2");
        builder.MoveToDocumentEnd();
        builder.Writeln("Content of Section 2");

        // Save the source document (optional, just for reference).
        string sourcePath = "Source.docx";
        sourceDoc.Save(sourcePath);

        // Expected header/footer texts for validation after splitting.
        var expectedHeadersFooters = new List<(string Header, string Footer)>
        {
            ("Header Section 1", "Footer Section 1"),
            ("Header Section 2", "Footer Section 2")
        };

        // Split the document by sections.
        int index = 1;
        foreach (Section srcSection in sourceDoc.Sections)
        {
            // Create a new empty document.
            Document splitDoc = new Document();
            splitDoc.RemoveAllChildren(); // Remove the default empty section.

            // Import the current section into the new document.
            NodeImporter importer = new NodeImporter(sourceDoc, splitDoc, ImportFormatMode.KeepSourceFormatting);
            Section importedSection = (Section)importer.ImportNode(srcSection, true);
            splitDoc.AppendChild(importedSection);

            // Save the split document.
            string splitPath = $"Split_{index}.docx";
            splitDoc.Save(splitPath);

            // Reopen the saved document to verify header and footer.
            Document loadedDoc = new Document(splitPath);
            string headerText = loadedDoc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary].GetText();
            string footerText = loadedDoc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary].GetText();

            var expected = expectedHeadersFooters[index - 1];
            if (!headerText.Contains(expected.Header) || !footerText.Contains(expected.Footer))
            {
                throw new Exception($"Header/footer validation failed for {splitPath}");
            }

            index++;
        }

        // If execution reaches this point, all validations passed.
        Console.WriteLine("Document split and header/footer verification completed successfully.");
    }
}

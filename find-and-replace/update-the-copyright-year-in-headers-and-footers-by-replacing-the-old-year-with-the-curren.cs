using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with a header and a footer containing the old year.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Header with old year.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("(C) 2022 My Company");

        // Footer with old year.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("(C) 2022 My Company");

        // Body content (optional).
        builder.MoveToDocumentEnd();
        builder.Writeln("Sample body text.");

        // Save the sample document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loadedDoc = new Document(inputPath);

        // Prepare find-and-replace options.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false,
            FindWholeWordsOnly = false
        };

        // Define the old and new text.
        int currentYear = DateTime.Now.Year;
        string oldText = "(C) 2022 My Company";
        string newText = $"(C) {currentYear} My Company";

        // Replace in the header.
        HeaderFooter header = loadedDoc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
        int headerReplacements = header.Range.Replace(oldText, newText, options);

        // Replace in the footer.
        HeaderFooter footer = loadedDoc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary];
        int footerReplacements = footer.Range.Replace(oldText, newText, options);

        // Validate that at least one replacement occurred.
        if (headerReplacements == 0 && footerReplacements == 0)
            throw new InvalidOperationException("Expected at least one replacement in header or footer.");

        // Save the updated document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);
    }
}

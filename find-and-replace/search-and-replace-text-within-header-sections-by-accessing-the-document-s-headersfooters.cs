using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class HeaderFindReplaceExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a primary header with placeholder text.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Company: _CompanyName_");

        // Save the document to a local file.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document from the file.
        Document loadedDoc = new Document(inputPath);

        // Access the primary header of the first section.
        HeaderFooterCollection headersFooters = loadedDoc.FirstSection.HeadersFooters;
        HeaderFooter? header = headersFooters[HeaderFooterType.HeaderPrimary];

        if (header == null)
            throw new InvalidOperationException("The document does not contain a primary header.");

        // Perform find-and-replace within the header's range.
        FindReplaceOptions options = new FindReplaceOptions();
        int replacedCount = header.Range.Replace("_CompanyName_", "Acme Corp", options);

        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement in the header.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);
    }
}

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Define file names in the current working directory.
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "sample.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");

        // -----------------------------------------------------------------
        // 1. Create a sample document with a header that contains placeholder text.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the primary header of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Company: _CompanyName_");
        builder.Writeln(); // Add a line break in the header.

        // Add some body content so the document is not empty.
        builder.MoveToDocumentEnd();
        builder.Writeln("This is the body of the document.");

        // Save the initial document.
        doc.Save(inputPath);

        // -----------------------------------------------------------------
        // 2. Load the document and perform a find-and-replace inside the header.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(inputPath);

        // Retrieve the primary header of the first section.
        HeaderFooter header = loadedDoc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
        if (header == null)
            throw new InvalidOperationException("The document does not contain a primary header.");

        // Replace the placeholder with the actual company name.
        FindReplaceOptions options = new FindReplaceOptions();
        int replacedCount = header.Range.Replace("_CompanyName_", "Acme Corp", options);

        // Validate that a replacement actually occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement in the header.");

        // -----------------------------------------------------------------
        // 3. Save the modified document.
        // -----------------------------------------------------------------
        loadedDoc.Save(outputPath);
    }
}

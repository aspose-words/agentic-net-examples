using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a blank document and add a primary header with placeholder text.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the primary header of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Company: _CompanyName_");
        builder.Writeln("Date: _Date_");

        // Save the source document locally.
        const string inputPath = "HeaderInput.docx";
        doc.Save(inputPath);

        // -----------------------------------------------------------------
        // 2. Load the document and obtain the header's range.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(inputPath);

        // Retrieve the primary header from the first section.
        HeaderFooter header = loadedDoc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
        if (header == null)
            throw new InvalidOperationException("Primary header not found in the document.");

        // -----------------------------------------------------------------
        // 3. Perform find-and-replace operations inside the header.
        // -----------------------------------------------------------------
        FindReplaceOptions options = new FindReplaceOptions();

        int totalReplacements = 0;
        totalReplacements += header.Range.Replace("_CompanyName_", "Acme Corp", options);
        totalReplacements += header.Range.Replace("_Date_", DateTime.Today.ToString("yyyy-MM-dd"), options);

        if (totalReplacements == 0)
            throw new InvalidOperationException("Expected at least one replacement in the header.");

        // -----------------------------------------------------------------
        // 4. Save the modified document.
        // -----------------------------------------------------------------
        const string outputPath = "HeaderOutput.docx";
        loadedDoc.Save(outputPath);
    }
}

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Paths for the sample input and output documents.
        string inputPath = "sample.docx";
        string outputPath = "output.docx";

        // -----------------------------------------------------------------
        // 1. Create a sample document with a primary header containing text.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the primary header of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Company: _CompanyName_");

        // Save the document so it can be re‑loaded for the replace operation.
        doc.Save(inputPath);

        // ---------------------------------------------------------------
        // 2. Load the document and replace text inside the header section.
        // ---------------------------------------------------------------
        Document loadedDoc = new Document(inputPath);

        // Retrieve the primary header from the first section.
        HeaderFooter header = loadedDoc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
        if (header == null)
            throw new InvalidOperationException("The document does not contain a primary header.");

        // Perform a find‑and‑replace on the header's range.
        FindReplaceOptions options = new FindReplaceOptions();
        int replacedCount = header.Range.Replace("_CompanyName_", "Aspose Ltd.", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement in the header.");

        // ---------------------------------------------------------------
        // 3. Save the modified document.
        // ---------------------------------------------------------------
        loadedDoc.Save(outputPath);

        // Optional: output the result count to the console.
        Console.WriteLine($"Replacements performed in header: {replacedCount}");
    }
}

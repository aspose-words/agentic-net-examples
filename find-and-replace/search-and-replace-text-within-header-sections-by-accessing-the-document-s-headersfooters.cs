using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a primary header and write some placeholder text.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Company: Acme Corp");
        builder.Writeln();
        builder.Write("Report Date: 2023-01-01");

        // Access the primary header through the HeadersFooters collection.
        HeaderFooter primaryHeader = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
        if (primaryHeader == null)
            throw new InvalidOperationException("Primary header not found.");

        // Perform a find‑and‑replace operation inside the header.
        FindReplaceOptions options = new FindReplaceOptions();
        int replaceCount = primaryHeader.Range.Replace("Acme Corp", "Contoso Ltd", options);

        // Validate that a replacement actually occurred.
        if (replaceCount == 0)
            throw new InvalidOperationException("No occurrences of the search text were found in the header.");

        // Save the modified document to the local file system.
        const string outputFile = "HeaderReplaceResult.docx";
        doc.Save(outputFile);
    }
}

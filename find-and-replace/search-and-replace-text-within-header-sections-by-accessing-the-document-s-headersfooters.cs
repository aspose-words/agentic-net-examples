using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing; // Required by Aspose.Words for drawing types
using Newtonsoft.Json; // Included as per package requirement

public class HeaderFindReplaceExample
{
    public static void Main()
    {
        // Create a new document and add a header with sample text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a primary header and write text that will be replaced.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("This is the old header text.");

        // Add some body content that also contains the word "old" to demonstrate that only the header is affected.
        builder.MoveToDocumentEnd();
        builder.Writeln();
        builder.Writeln("Body contains old text as well.");

        // Save the initial document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document from disk.
        Document loadedDoc = new Document(inputPath);

        int totalReplacements = 0;

        // Iterate through each section's primary header and perform the replacement.
        foreach (Section section in loadedDoc.Sections)
        {
            HeaderFooter header = section.HeadersFooters[HeaderFooterType.HeaderPrimary];
            if (header != null)
            {
                int replaced = header.Range.Replace("old", "new", new FindReplaceOptions());
                totalReplacements += replaced;
            }
        }

        // Validate that at least one replacement occurred in the header.
        if (totalReplacements == 0)
            throw new InvalidOperationException("Expected at least one replacement in the header, but none were made.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);
    }
}

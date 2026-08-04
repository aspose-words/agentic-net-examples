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

        // Insert a primary header and write placeholder text.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Company: OldName");

        // Add a simple body paragraph.
        builder.MoveToDocumentEnd();
        builder.Writeln("Body content.");

        // Save the initial document to the local file system.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document from the saved file.
        Document loadedDoc = new Document(inputPath);

        // Retrieve the primary header from the first section.
        HeaderFooter header = loadedDoc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
        if (header == null)
            throw new InvalidOperationException("Header not found.");

        // Perform a find-and-replace operation within the header's range.
        FindReplaceOptions options = new FindReplaceOptions();
        int replaced = header.Range.Replace("OldName", "NewName", options);
        if (replaced == 0)
            throw new InvalidOperationException("No replacements were made in the header.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);
    }
}

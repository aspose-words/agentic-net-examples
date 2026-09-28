using System;
using Aspose.Words;
using Aspose.Words.Replacing;
using System.IO;

public class Program
{
    public static void Main()
    {
        // Create a sample document with a header that contains the text to be replaced.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a primary header and write placeholder text.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Header placeholder: OldValue");

        // Return to the main body and add some body text (should remain unchanged).
        builder.MoveToDocumentEnd();
        builder.Writeln("Body text with OldValue that should not be changed.");

        // Save the initial document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Perform replacement only within header sections.
        int totalReplacements = 0;
        foreach (Section section in loaded.Sections)
        {
            HeaderFooter header = section.HeadersFooters[HeaderFooterType.HeaderPrimary];
            if (header != null)
            {
                int replaced = header.Range.Replace("OldValue", "NewValue", new FindReplaceOptions());
                totalReplacements += replaced;
            }
        }

        // Validate that at least one replacement occurred.
        if (totalReplacements == 0)
            throw new InvalidOperationException("Expected at least one replacement in header.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }
}

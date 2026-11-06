using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with a footer that contains page numbers and placeholder text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some body content.
        builder.Writeln("This is the main body of the document.");

        // Move to the primary footer of the first (and only) section.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        // Write static text, insert page number fields, and a placeholder to be replaced.
        builder.Write("Confidential - ");
        builder.InsertField("PAGE", "1");
        builder.Write(" of ");
        builder.InsertField("NUMPAGES", "1");
        builder.Write(" - CompanyName");

        // Save the initial document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document from the saved file.
        Document loadedDoc = new Document(inputPath);

        // Prepare find-replace options (default options are sufficient for this case).
        FindReplaceOptions options = new FindReplaceOptions();

        // Perform replacement only within footer sections, preserving page number fields.
        int totalReplacements = 0;
        foreach (Section section in loadedDoc.Sections)
        {
            HeaderFooter footer = section.HeadersFooters[HeaderFooterType.FooterPrimary];
            if (footer != null)
            {
                int replaced = footer.Range.Replace("CompanyName", "Acme Corp", options);
                totalReplacements += replaced;
            }
        }

        // Validate that at least one replacement occurred.
        if (totalReplacements == 0)
            throw new InvalidOperationException("Expected at least one replacement in the footer.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);
    }
}

using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some body content.
        builder.Writeln("This is the main body of the document.");

        // Add a primary footer to the first (and only) section.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("Sample Footer Text");

        // Retrieve the first section.
        Section firstSection = doc.Sections[0];

        // Remove all footers from the section.
        // Aspose.Words stores headers and footers together in the HeadersFooters collection.
        // Iterate backwards and remove items whose type is a footer.
        for (int i = firstSection.HeadersFooters.Count - 1; i >= 0; i--)
        {
            HeaderFooter hf = firstSection.HeadersFooters[i];
            if (hf.HeaderFooterType == HeaderFooterType.FooterPrimary ||
                hf.HeaderFooterType == HeaderFooterType.FooterFirst ||
                hf.HeaderFooterType == HeaderFooterType.FooterEven)
            {
                firstSection.HeadersFooters.RemoveAt(i);
            }
        }

        // Save the resulting document.
        doc.Save("Output.docx");
    }
}

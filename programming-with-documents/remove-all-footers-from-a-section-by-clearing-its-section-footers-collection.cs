using System;
using System.Collections.Generic;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some body content.
        builder.Writeln("Main content of the document.");

        // Add a primary footer with sample text.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("This footer will be removed.");

        // Return to the main body.
        builder.MoveToDocumentEnd();

        // Remove all footers from each section.
        foreach (Section section in doc.Sections)
        {
            // Collect footers in the current section.
            List<HeaderFooter> footersToRemove = new List<HeaderFooter>();
            foreach (HeaderFooter hf in section.HeadersFooters)
            {
                if (hf.HeaderFooterType == HeaderFooterType.FooterPrimary ||
                    hf.HeaderFooterType == HeaderFooterType.FooterFirst ||
                    hf.HeaderFooterType == HeaderFooterType.FooterEven)
                {
                    footersToRemove.Add(hf);
                }
            }

            // Remove the collected footers.
            foreach (HeaderFooter hf in footersToRemove)
            {
                section.HeadersFooters.Remove(hf);
            }
        }

        // Save the modified document.
        const string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}

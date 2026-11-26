using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace RangeHeaderFooterExtraction
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Get the first (and only) section of the document.
            Section section = doc.FirstSection;

            // Ensure the section has a header/footer collection.
            HeaderFooterCollection hfCollection = section.HeadersFooters;

            // Add a primary header with sample text.
            HeaderFooter header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
            Paragraph headerPara = new Paragraph(doc);
            headerPara.AppendChild(new Run(doc, "Sample Header Text"));
            header.AppendChild(headerPara);
            hfCollection.Add(header);

            // Add a primary footer with sample text.
            HeaderFooter footer = new HeaderFooter(doc, HeaderFooterType.FooterPrimary);
            Paragraph footerPara = new Paragraph(doc);
            footerPara.AppendChild(new Run(doc, "Sample Footer Text"));
            footer.AppendChild(footerPara);
            hfCollection.Add(footer);

            // Save the document to disk.
            string filePath = "Sample.docx";
            doc.Save(filePath);

            // Load the document back (demonstrates loading rule).
            Document loadedDoc = new Document(filePath);

            // Extract plain text from all headers.
            foreach (HeaderFooter hf in loadedDoc.FirstSection.HeadersFooters)
            {
                if (hf.IsHeader)
                {
                    string headerText = hf.Range.Text.Trim();
                    Console.WriteLine("Header text extracted: " + headerText);
                }
            }

            // Extract plain text from all footers.
            foreach (HeaderFooter hf in loadedDoc.FirstSection.HeadersFooters)
            {
                // HeaderFooter does not have IsFooter; check the type instead.
                if (hf.HeaderFooterType == HeaderFooterType.FooterPrimary ||
                    hf.HeaderFooterType == HeaderFooterType.FooterFirst ||
                    hf.HeaderFooterType == HeaderFooterType.FooterEven)
                {
                    string footerText = hf.Range.Text.Trim();
                    Console.WriteLine("Footer text extracted: " + footerText);
                }
            }

            // The program ends here without waiting for user input.
        }
    }
}

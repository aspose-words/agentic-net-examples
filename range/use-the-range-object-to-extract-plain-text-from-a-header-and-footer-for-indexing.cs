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

            // Add a new section to the document.
            Section section = new Section(doc);
            doc.Sections.Add(section);

            // Create a header and add a paragraph with sample text.
            HeaderFooter header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
            Paragraph headerPara = new Paragraph(doc);
            headerPara.AppendChild(new Run(doc, "This is the header text for indexing."));
            header.AppendChild(headerPara);
            section.HeadersFooters.Add(header);

            // Create a footer and add a paragraph with sample text.
            HeaderFooter footer = new HeaderFooter(doc, HeaderFooterType.FooterPrimary);
            Paragraph footerPara = new Paragraph(doc);
            footerPara.AppendChild(new Run(doc, "This is the footer text for indexing."));
            footer.AppendChild(footerPara);
            section.HeadersFooters.Add(footer);

            // Save the document to a temporary file.
            string filePath = "Sample.docx";
            doc.Save(filePath);

            // Load the document back from the file.
            Document loadedDoc = new Document(filePath);

            // Iterate through each section and extract plain text from its header and footer using Range.
            foreach (Section sec in loadedDoc.Sections)
            {
                // Header extraction.
                HeaderFooter primaryHeader = sec.HeadersFooters[HeaderFooterType.HeaderPrimary];
                if (primaryHeader != null)
                {
                    string headerText = primaryHeader.Range.Text;
                    Console.WriteLine("Header Text: " + headerText);
                }

                // Footer extraction.
                HeaderFooter primaryFooter = sec.HeadersFooters[HeaderFooterType.FooterPrimary];
                if (primaryFooter != null)
                {
                    string footerText = primaryFooter.Range.Text;
                    Console.WriteLine("Footer Text: " + footerText);
                }
            }
        }
    }
}

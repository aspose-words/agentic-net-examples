using System;
using System.IO;
using Aspose.Words;

namespace AsposeWordsExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            int totalSections = 3;

            for (int i = 1; i <= totalSections; i++)
            {
                // Create a distinct primary header for the current section.
                builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
                builder.Write($"Header for Section {i}");

                // Create a distinct primary footer for the current section.
                builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
                builder.Write($"Footer for Section {i}");

                // Return to the body of the current section and add some text.
                builder.MoveToSection(i - 1);
                builder.Writeln($"This is the content of section {i}.");

                // Insert a section break to start a new section, except after the last one.
                if (i < totalSections)
                {
                    builder.InsertBreak(BreakType.SectionBreakNewPage);
                }
            }

            // Save the document to the current directory.
            string outputFile = Path.Combine(Directory.GetCurrentDirectory(), "MultiSectionHeadersFooters.docx");
            doc.Save(outputFile);
        }
    }
}

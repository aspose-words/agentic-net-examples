using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace HeaderFooterToPdfExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Attach a DocumentBuilder to the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Enable different headers/footers for the first page (optional).
            builder.PageSetup.DifferentFirstPageHeaderFooter = true;

            // Move the cursor to the primary header of the current section and add text.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Write("This is the primary header");

            // Move the cursor to the primary footer of the current section and add text.
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            builder.Write("Page ");
            // Insert a field that will display the current page number.
            builder.InsertField("PAGE", "");

            // Return the cursor to the main document body.
            builder.MoveToSection(0);

            // Add some sample content to the body.
            builder.Writeln("Hello World!");
            builder.Writeln("This document contains a header and a footer.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Second page content.");

            // Save the document as PDF.
            doc.Save("OutputWithHeaderFooter.pdf", SaveFormat.Pdf);
        }
    }
}

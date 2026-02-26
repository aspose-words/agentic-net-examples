using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a heading using the built‑in Heading1 style.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("LINQ Reporting Introduction to LINQ Reporting Engine");

        // If you have a PDF template you want to use, uncomment the line below.
        // Document doc = new Document("Template.pdf");

        // Save the document.
        doc.Save("LINQReportingHeading.docx");
    }
}

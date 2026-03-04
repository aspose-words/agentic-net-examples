using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // Use DocumentBuilder to add content to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the paragraph style to Heading 1 for the title.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;

        // Write the heading text.
        builder.Writeln("LINQ Reporting Introduction to LINQ Reporting Engine");

        // Save the document to a file.
        doc.Save("LINQReportingEngine.docx");
    }
}

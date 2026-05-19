using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add a primary footer and a first-page footer.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("Primary footer text.");
        builder.MoveToHeaderFooter(HeaderFooterType.FooterFirst);
        builder.Writeln("First page footer text.");

        // At this point the section contains footers.
        // Remove all footers by clearing the HeadersFooters collection of the first section.
        // This removes both header and footer nodes; for the purpose of the task we only needed to remove footers.
        doc.FirstSection.HeadersFooters.Clear();

        // Save the resulting document.
        doc.Save("Output.docx");
    }
}

using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder for simple paragraph insertion.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add an initial paragraph.
        builder.Writeln("First paragraph.");

        // Insert the target paragraph.
        builder.Writeln("Paragraph with page break before.");

        // Retrieve the paragraph that was just added.
        Paragraph targetParagraph = (Paragraph)doc.LastSection.Body.LastParagraph;

        // Set the PageBreakBefore property to true.
        targetParagraph.ParagraphFormat.PageBreakBefore = true;

        // Save the document.
        doc.Save("Output.docx");
    }
}

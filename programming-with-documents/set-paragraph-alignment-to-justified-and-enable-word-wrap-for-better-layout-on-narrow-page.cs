using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Sample long text to illustrate justification and word wrap.
        string sampleText = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. " +
                            "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " +
                            "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.";

        // Insert a new paragraph with the sample text.
        Paragraph paragraph = new Paragraph(doc);
        paragraph.AppendChild(new Run(doc, sampleText));

        // Set paragraph alignment to justified.
        paragraph.ParagraphFormat.Alignment = ParagraphAlignment.Justify;

        // Enable word wrap for the paragraph (useful on narrow pages).
        paragraph.ParagraphFormat.WordWrap = true;

        // Add the paragraph to the document body.
        doc.FirstSection.Body.AppendChild(paragraph);

        // Save the document to a file.
        string outputPath = "JustifiedParagraph.docx";
        doc.Save(outputPath);
    }
}

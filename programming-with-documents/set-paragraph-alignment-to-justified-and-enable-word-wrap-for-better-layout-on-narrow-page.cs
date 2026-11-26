using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add a paragraph with sample text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample paragraph that will be justified and have word wrap enabled. " +
                        "It contains enough text to demonstrate wrapping on narrow pages, ensuring the layout " +
                        "remains readable even when the page width is limited.");

        // Retrieve the first paragraph that was just added.
        Paragraph paragraph = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);

        // Set paragraph alignment to justified.
        paragraph.ParagraphFormat.Alignment = ParagraphAlignment.Justify;

        // Enable word wrap for the paragraph.
        paragraph.ParagraphFormat.WordWrap = true;

        // Save the document to a file.
        const string outputPath = "JustifiedParagraph.docx";
        doc.Save(outputPath);
    }
}

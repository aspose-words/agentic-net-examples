using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Tables;

class ParagraphManipulation
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document for easy content insertion.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the first paragraph.
        builder.Writeln("Hello Aspose.Words!");

        // Insert a second paragraph that will later be modified.
        builder.Writeln("This is the original paragraph that will be replaced.");

        // Insert a third paragraph with custom formatting (centered and extra space after).
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.ParagraphFormat.SpaceAfter = 12; // points
        builder.Writeln("Centered paragraph with extra space after.");

        // Reset paragraph formatting to defaults for subsequent inserts.
        builder.ParagraphFormat.ClearFormatting();

        // Perform a find‑and‑replace on the entire document.
        // Replace the word "original" with "updated".
        int replaceCount = doc.Range.Replace("original", "updated");
        // Example of regex replace (uncomment if needed):
        // int replaceCount = doc.Range.Replace(new Regex(@"Aspose\.\w+"), "Aspose.Words");

        // Access the second paragraph (index 1) and change its style.
        Paragraph secondParagraph = doc.FirstSection.Body.Paragraphs[1];
        secondParagraph.ParagraphFormat.StyleName = "Intense Quote";

        // Move the cursor to the end of the document and add a final paragraph.
        builder.MoveToDocumentEnd();
        builder.Writeln("Added at the end of the document.");

        // Save the modified document as a DOCX file.
        doc.Save("ParagraphManipulation.docx");
    }
}

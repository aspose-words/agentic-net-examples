using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Notes;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ---------- Header ----------
        // Move the cursor to the primary header of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        // Write some text and a PAGE field.
        builder.Writeln("Sample Header - Page ");
        builder.InsertField("PAGE", "");

        // ---------- Footer ----------
        // Move the cursor to the primary footer of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("Sample Footer - Confidential");

        // ---------- Body ----------
        // Return to the main story (body) of the document.
        builder.MoveToDocumentStart();

        // First paragraph.
        builder.Writeln("First paragraph. This is the start of the document.");

        // Insert a footnote after the first paragraph.
        Footnote footnote = builder.InsertFootnote(FootnoteType.Footnote, "This is a footnote.");

        // Second paragraph.
        builder.Writeln("Second paragraph. More content follows.");

        // Third paragraph.
        builder.Writeln("Third paragraph. End of sample content.");

        // ---------- Extract content between first and third paragraph ----------
        // Get references to the first and third paragraphs.
        Paragraph firstPara = doc.FirstSection.Body.Paragraphs[0];
        Paragraph thirdPara = doc.FirstSection.Body.Paragraphs[2];

        // Collect the text of all paragraphs from the first up to (and including) the third.
        string extractedRangeText = "";
        for (int i = 0; i <= 2; i++)
        {
            extractedRangeText += doc.FirstSection.Body.Paragraphs[i].GetText();
        }

        Console.WriteLine("Extracted text between first and third paragraph:");
        Console.WriteLine(extractedRangeText.Trim());

        // ---------- Replace text in the header ----------
        HeaderFooter header = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
        // Replace the placeholder "Sample Header" with "Updated Header".
        header.Range.Replace("Sample Header", "Updated Header");

        // ---------- Update footnote reference numbers and read footnote text ----------
        // Ensure that fields (including footnote references) are up‑to‑date.
        doc.UpdateFields();
        doc.UpdatePageLayout();

        // Retrieve the first footnote node in the document.
        Footnote firstFootnote = (Footnote)doc.GetChildNodes(NodeType.Footnote, true)[0];
        Console.WriteLine("Footnote text: " + firstFootnote.GetText().Trim());

        // ---------- Save the document ----------
        doc.Save("Output.docx");
    }
}

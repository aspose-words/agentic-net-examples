using System;
using Aspose.Words;
using Aspose.Words.Tables;

class HyphenationExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;
        // Optional: fine‑tune hyphenation behaviour.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch

        // Add a paragraph with long text that will be hyphenated.
        builder.Font.Size = 24;
        builder.Writeln(
            "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

        // Insert a table.
        builder.StartTable();

        // First cell – hyphenation will be suppressed.
        builder.InsertCell();
        builder.Writeln("This cell has hyphenation disabled. The quick brown fox jumps over the lazy dog.");
        // Retrieve the cell we just created.
        Cell cellWithSuppressedHyphens = doc.FirstSection.Body.Tables[0].Rows[0].Cells[0];
        // Suppress hyphenation for all paragraphs in this cell.
        foreach (Paragraph para in cellWithSuppressedHyphens.Paragraphs)
        {
            para.ParagraphFormat.SuppressAutoHyphens = true;
        }

        // Second cell – hyphenation follows the document settings.
        builder.InsertCell();
        builder.Writeln(
            "This cell follows the document hyphenation settings. The quick brown fox jumps over the lazy dog.");

        // End the table.
        builder.EndTable();

        // Save the document.
        doc.Save("HyphenationOverride.docx");
    }
}

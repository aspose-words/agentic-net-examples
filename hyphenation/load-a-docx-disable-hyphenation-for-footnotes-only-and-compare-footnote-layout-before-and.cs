using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Notes;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Define file names in the local folder.
        string originalPath = "Original.docx";
        string hyphenatedPath = "Hyphenated.docx";
        string footnoteNoHyphenPath = "FootnoteNoHyphen.docx";

        // -------------------------------------------------
        // 1. Create a sample document with a long paragraph and a footnote.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a narrow page width to force line wrapping.
        builder.PageSetup.PageWidth = 300; // points

        // Write a long paragraph that will need hyphenation.
        builder.Font.Size = 12;
        builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " +
                        "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.");

        // Insert a footnote with similarly long text.
        builder.InsertFootnote(FootnoteType.Footnote,
            "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " +
            "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.");

        // Save the original document (no hyphenation applied yet).
        doc.Save(originalPath);

        // -------------------------------------------------
        // 2. Load the document and enable automatic hyphenation.
        // -------------------------------------------------
        Document hyphenatedDoc = new Document(originalPath);
        hyphenatedDoc.HyphenationOptions.AutoHyphenation = true;
        hyphenatedDoc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch
        hyphenatedDoc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        hyphenatedDoc.HyphenationOptions.HyphenateCaps = true;

        // Save the document with hyphenation applied to the whole body, including footnotes.
        hyphenatedDoc.Save(hyphenatedPath);

        // -------------------------------------------------
        // 3. Load the hyphenated document and disable hyphenation for footnotes only.
        // -------------------------------------------------
        Document footnoteNoHyphenDoc = new Document(hyphenatedPath);

        // Find all footnote paragraphs and suppress automatic hyphens.
        NodeCollection footnotes = footnoteNoHyphenDoc.GetChildNodes(NodeType.Footnote, true);
        foreach (Footnote footnote in footnotes)
        {
            // Record the original state (should be false before we change it).
            bool originalSuppress = footnote.FirstParagraph?.ParagraphFormat?.SuppressAutoHyphens ?? false;

            // Disable hyphenation for this footnote paragraph.
            if (footnote.FirstParagraph != null)
                footnote.FirstParagraph.ParagraphFormat.SuppressAutoHyphens = true;

            // Output comparison info to the console.
            Console.WriteLine($"Footnote ID {footnote.FootnoteType} - SuppressAutoHyphens changed from {originalSuppress} to true.");
        }

        // Save the document where footnotes are not hyphenated.
        footnoteNoHyphenDoc.Save(footnoteNoHyphenPath);

        // -------------------------------------------------
        // 4. Report completion.
        // -------------------------------------------------
        Console.WriteLine($"Documents created:");
        Console.WriteLine($"  Original document: {Path.GetFullPath(originalPath)}");
        Console.WriteLine($"  Hyphenated document (footnotes hyphenated): {Path.GetFullPath(hyphenatedPath)}");
        Console.WriteLine($"  Footnote-no-hyphen document (footnotes hyphenation disabled): {Path.GetFullPath(footnoteNoHyphenPath)}");
    }
}

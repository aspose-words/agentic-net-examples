using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Saving;

class ChapterNumberRomanExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ------------------------------------------------------------
        // 1. Prepare headings that will act as chapter titles.
        //    The headings must have a numbered outline format.
        // ------------------------------------------------------------
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1 – Introduction");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2 – Getting Started");

        // ------------------------------------------------------------
        // 2. Configure page numbering to include chapter numbers.
        //    - Use HeadingLevelForChapter = 1 (the level of our headings).
        //    - Use a colon as the separator between chapter and page numbers.
        //    - Page numbers themselves will be Arabic (1,2,3…).
        // ------------------------------------------------------------
        PageSetup pageSetup = doc.FirstSection.PageSetup;
        pageSetup.HeadingLevelForChapter = 1;                     // Chapter titles are Heading 1.
        pageSetup.ChapterPageSeparator = ChapterPageSeparator.Colon;
        pageSetup.PageNumberStyle = NumberStyle.Arabic;           // Arabic page numbers.

        // ------------------------------------------------------------
        // 3. Insert a header that displays the chapter number in lower‑case Roman numerals.
        //    The field code uses the expression: = ChapterNumber \* roman
        //    (equivalent to the {=ChapterNumber:roman} syntax in Word).
        // ------------------------------------------------------------
        builder.MoveToSection(0);
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Font.Size = 12;
        builder.Write("Chapter ");
        // Insert the field that evaluates the current chapter number and formats it as lower‑case Roman.
        builder.InsertField("= ChapterNumber \\* roman", "");

        // Add a separator and the page number.
        builder.Write(" – Page ");
        builder.InsertField("PAGE", "");

        // ------------------------------------------------------------
        // 4. Save the document.
        // ------------------------------------------------------------
        doc.Save("ChapterNumberRoman.docx", SaveFormat.Docx);
    }
}

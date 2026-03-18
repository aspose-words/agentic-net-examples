using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add some heading paragraphs.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Heading 1
        builder.ParagraphFormat.Style = doc.Styles["Heading 1"];
        builder.Writeln("Heading 1");

        // Heading 2
        builder.ParagraphFormat.Style = doc.Styles["Heading 2"];
        builder.Writeln("Heading 2");

        // Heading 3
        builder.ParagraphFormat.Style = doc.Styles["Heading 3"];
        builder.Writeln("Heading 3");

        // Create HtmlSaveOptions and set the split criteria to headings.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph;

        // Optional: limit the heading level that triggers a split (e.g., up to level 2).
        saveOptions.DocumentSplitHeadingLevel = 2;

        // Save the document; it will be split into separate HTML files at heading paragraphs.
        doc.Save("SplitByHeadings.html", saveOptions);
    }
}

using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add some content so the split operation has something to process.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Section 1");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 2");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 3");

        // Create HtmlSaveOptions and set the split criteria to split by sections.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            // DocumentSplitCriteria is an enum; SectionBreak splits the document at each section break.
            DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak
        };

        // Save the document. Because DocumentSplitCriteria is set, Aspose.Words will generate
        // separate HTML files for each section (e.g., output.html, output-01.html, output-02.html).
        doc.Save("output.html", saveOptions);
    }
}

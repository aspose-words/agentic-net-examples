using System;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Drawing;

class RestartListNumberingExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a numbered list to the document and configure it to restart at each section.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);
        list.IsRestartAtEachSection = true; // <-- key setting

        // Apply the list to the following paragraphs.
        builder.ListFormat.List = list;
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.Writeln("Item 3");

        // Insert a heading paragraph. This will start a new section.
        builder.ListFormat.RemoveNumbers(); // stop list formatting for the heading
        builder.ParagraphFormat.Style = doc.Styles["Heading 1"];
        builder.Writeln("New Section Heading");

        // Insert a section break (new page) so the next list belongs to a new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Continue the list; numbering will restart at 1 because IsRestartAtEachSection is true.
        builder.ListFormat.List = list;
        builder.Writeln("Item 1 (new section)");
        builder.Writeln("Item 2 (new section)");
        builder.Writeln("Item 3 (new section)");

        // Clean up list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document.
        doc.Save("RestartListAtEachSection.docx");
    }
}

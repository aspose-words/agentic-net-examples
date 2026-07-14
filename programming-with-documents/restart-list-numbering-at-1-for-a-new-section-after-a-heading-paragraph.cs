using System;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First section heading.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Section 1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;

        // Create a numbered list and configure it to restart at each new section.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);
        list.IsRestartAtEachSection = true;

        // Apply the list to the following paragraphs.
        builder.ListFormat.List = list;
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");

        // Insert a section break (new page) to start a new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Second section heading.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Section 2");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;

        // Continue using the same list; numbering will restart at 1 because of IsRestartAtEachSection.
        builder.ListFormat.List = list;
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");

        // Save the document. Use a compliance level that supports the restart property.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
        {
            Compliance = OoxmlCompliance.Iso29500_2008_Transitional
        };
        doc.Save("RestartList.docx", saveOptions);
    }
}

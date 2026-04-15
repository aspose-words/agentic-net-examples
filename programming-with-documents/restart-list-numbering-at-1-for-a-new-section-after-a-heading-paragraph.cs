using System;
using System.IO;
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

        // Add a heading for the first section.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("First Section");

        // Create a numbered list and configure it to restart at each section.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);
        list.IsRestartAtEachSection = true; // Restart numbering for each new section.

        // Apply the list to the following paragraphs.
        builder.ListFormat.List = list;
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.Writeln("Item 3");

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Insert a section break (new page) after the heading.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Add a heading for the second section.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Second Section");

        // Apply the same list again; numbering will restart at 1 because of IsRestartAtEachSection.
        builder.ListFormat.List = list;
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.Writeln("Item 3");

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Save the document with a compliance level that supports the restart property.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "RestartListAtEachSection.docx");
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
        {
            Compliance = OoxmlCompliance.Iso29500_2008_Transitional
        };
        doc.Save(outputPath, saveOptions);
    }
}

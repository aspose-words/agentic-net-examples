using System;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Saving;

class InsertRestartNumTag
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a numbered list to the document.
        // The list will be of the default numbered template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Enable the restart numbering tag for the list.
        // In DOC format this causes a <restartNum> tag to be written for each list level.
        list.IsRestartAtEachSection = true;

        // Apply the list to the builder so that subsequent paragraphs become list items.
        builder.ListFormat.List = list;

        // Add some list items.
        builder.Writeln("First item");
        builder.Writeln("Second item");
        builder.Writeln("Third item");

        // Insert a section break to demonstrate that numbering restarts after the break.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Continue the list after the section break; numbering will restart because of the tag.
        builder.Writeln("First item after restart");
        builder.Writeln("Second item after restart");

        // Save the document in DOC format (the restartNum tag will be present).
        doc.Save("RestartNumTag.doc", SaveFormat.Doc);
    }
}

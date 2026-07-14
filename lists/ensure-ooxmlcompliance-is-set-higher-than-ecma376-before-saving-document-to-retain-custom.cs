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

        // Add a numbered list to the document.
        doc.Lists.Add(ListTemplate.NumberDefault);
        List docList = doc.Lists[0];

        // Enable restarting the list at each new section.
        docList.IsRestartAtEachSection = true;

        // Apply the list to the builder.
        builder.ListFormat.List = docList;

        // First section items.
        builder.Writeln("Item 1 in first section");
        builder.Writeln("Item 2 in first section");

        // Insert a section break to start a new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Second section items – numbering will restart because of the setting above.
        builder.Writeln("Item 1 in second section");
        builder.Writeln("Item 2 in second section");

        // Remove list formatting from the builder.
        builder.ListFormat.RemoveNumbers();

        // Prepare OOXML save options with a compliance level higher than Ecma376.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.Compliance = OoxmlCompliance.Iso29500_2008_Transitional;

        // Define the output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CustomList.docx");

        // Save the document using the specified save options.
        doc.Save(outputPath, saveOptions);
    }
}

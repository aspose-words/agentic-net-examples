using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Lists;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Define output directory and file name.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string filePath = Path.Combine(outputDir, "CustomList.docx");

        // Create a new blank document and a DocumentBuilder to add content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a numbered list to the document.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Customize the first level of the list (e.g., red font, larger size, start at 5).
        ListLevel level0 = list.ListLevels[0];
        level0.Font.Color = Color.Red;
        level0.Font.Size = 24;
        level0.StartAt = 5;

        // Enable restarting the list at each new section.
        // This property only takes effect when the OOXML compliance level is newer than Ecma376.
        list.IsRestartAtEachSection = true;

        // Apply the list to a few paragraphs, insert a section break, and continue the list.
        builder.ListFormat.List = list;
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Item 3");
        builder.Writeln("Item 4");
        builder.ListFormat.RemoveNumbers();

        // Create OoxmlSaveOptions and set compliance higher than Ecma376.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.Compliance = OoxmlCompliance.Iso29500_2008_Transitional;

        // Save the document using the specified save options.
        doc.Save(filePath, saveOptions);
    }
}

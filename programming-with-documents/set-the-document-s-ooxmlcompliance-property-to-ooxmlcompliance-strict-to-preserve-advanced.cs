using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a simple numbered list.
        doc.Lists.Add(ListTemplate.NumberDefault);
        List list = doc.Lists[0];
        // Enable advanced list setting: restart numbering at each section.
        list.IsRestartAtEachSection = true;

        // Populate the list with a few items.
        builder.ListFormat.List = list;
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Item 3");
        builder.Writeln("Item 4");
        builder.ListFormat.RemoveNumbers();

        // Configure OOXML save options to use strict compliance.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            Compliance = OoxmlCompliance.Iso29500_2008_Strict
        };

        // Define output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AdvancedListStrict.docx");

        // Save the document with the specified compliance level.
        doc.Save(outputPath, saveOptions);

        // Indicate completion.
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}

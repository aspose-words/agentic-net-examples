using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Lists;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Create output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a numbered list and enable restart at each section.
        doc.Lists.Add(ListTemplate.NumberDefault);
        List list = doc.Lists[0];
        list.IsRestartAtEachSection = true;

        // Write list items in the first section.
        builder.ListFormat.List = list;
        builder.Writeln("Item 1 in first section");
        builder.Writeln("Item 2 in first section");

        // Insert a section break and add items in the second section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Item 1 in second section");
        builder.Writeln("Item 2 in second section");
        builder.ListFormat.RemoveNumbers();

        // Set OOXML compliance higher than Ecma376 to retain the custom list setting.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.Compliance = OoxmlCompliance.Iso29500_2008_Transitional;

        // Save the document.
        string outFile = Path.Combine(outputDir, "ListRestart.docx");
        doc.Save(outFile, saveOptions);

        // Load the saved document to verify the list setting persisted.
        Document loaded = new Document(outFile);
        bool restart = loaded.Lists[0].IsRestartAtEachSection;

        // Output verification result.
        Console.WriteLine($"IsRestartAtEachSection persisted: {restart}");
    }
}

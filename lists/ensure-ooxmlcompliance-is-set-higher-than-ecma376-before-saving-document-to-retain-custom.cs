using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Saving;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a numbered list to the document.
        doc.Lists.Add(ListTemplate.NumberDefault);
        List docList = doc.Lists[0];

        // Enable restarting the list at each new section.
        // This property only takes effect when the OOXML compliance level is higher than Ecma376.
        docList.IsRestartAtEachSection = true;

        // Apply the list to the builder.
        builder.ListFormat.List = docList;

        // Write some list items, then insert a section break and write more items.
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Item 3");
        builder.Writeln("Item 4");

        // Configure OOXML save options with a compliance level higher than Ecma376.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.Compliance = OoxmlCompliance.Iso29500_2008_Transitional;

        // Save the document using the specified save options.
        string outPath = Path.Combine(outputDir, "CustomList.docx");
        doc.Save(outPath, saveOptions);

        // Load the saved document to verify that the list restart setting was retained.
        Document loadedDoc = new Document(outPath);
        bool isRestartEnabled = loadedDoc.Lists[0].IsRestartAtEachSection;

        // Output the verification result.
        Console.WriteLine($"IsRestartAtEachSection retained after save: {isRestartEnabled}");
    }
}

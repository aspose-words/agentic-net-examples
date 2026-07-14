using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Define output file path in the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AdvancedList.docx");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a numbered list to the document.
        doc.Lists.Add(ListTemplate.NumberDefault);
        List docList = doc.Lists[0];

        // Enable advanced list setting: restart numbering at each new section.
        docList.IsRestartAtEachSection = true;

        // Build list items before a section break.
        builder.ListFormat.List = docList;
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");

        // Insert a new section break.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Continue the list in the new section.
        builder.Writeln("Item 3");
        builder.Writeln("Item 4");
        builder.ListFormat.RemoveNumbers();

        // Save the document with strict OOXML compliance to preserve the advanced list setting.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
        {
            Compliance = OoxmlCompliance.Iso29500_2008_Strict,
            SaveFormat = SaveFormat.Docx
        };
        doc.Save(outputPath, saveOptions);

        // Reload the saved document to verify the setting persisted.
        Document loadedDoc = new Document(outputPath);
        bool isRestartEnabled = loadedDoc.Lists[0].IsRestartAtEachSection;

        // Output verification result.
        Console.WriteLine($"List restart at each section enabled: {isRestartEnabled}");
    }
}

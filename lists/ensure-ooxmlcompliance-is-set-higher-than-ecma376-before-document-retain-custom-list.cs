using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Lists;
using Aspose.Words.Settings;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a simple numbered list to demonstrate custom list settings.
        // The list will later be verified after saving with higher OOXML compliance.
        doc.Lists.Add(ListTemplate.NumberDefault);
        List list = doc.Lists[0];
        list.IsRestartAtEachSection = true; // Custom setting we want to preserve.

        // Apply the list to a few paragraphs.
        builder.ListFormat.List = list;
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.ListFormat.RemoveNumbers();

        // Configure OOXML save options with a compliance level newer than Ecma376.
        // Iso29500_2008_Transitional is higher than Ecma376_2006.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.Compliance = OoxmlCompliance.Iso29500_2008_Transitional;

        // Save the document using the configured options.
        doc.Save("CustomList_HigherCompliance.docx", saveOptions);

        // Reload the document to verify that the custom list setting persisted.
        Document loaded = new Document("CustomList_HigherCompliance.docx");
        bool isRestart = loaded.Lists[0].IsRestartAtEachSection;
        Console.WriteLine($"IsRestartAtEachSection preserved: {isRestart}");
    }
}

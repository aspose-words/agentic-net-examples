using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Lists;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a numbered list to the document.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Enable restarting the list at each new section.
        list.IsRestartAtEachSection = true;

        // Apply the list to some paragraphs.
        builder.ListFormat.List = list;
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");

        // Insert a section break and continue the list.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Item 3");
        builder.Writeln("Item 4");

        // Prepare save options with a compliance level higher than Ecma376.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.Compliance = OoxmlCompliance.Iso29500_2008_Transitional;

        // Save the document.
        string filePath = "CustomList.docx";
        doc.Save(filePath, saveOptions);

        // Load the saved document to verify that the list setting persisted.
        Document loadedDoc = new Document(filePath);
        bool isRestart = loadedDoc.Lists[0].IsRestartAtEachSection;

        // Output the verification result.
        Console.WriteLine($"IsRestartAtEachSection persisted: {isRestart}");
    }
}

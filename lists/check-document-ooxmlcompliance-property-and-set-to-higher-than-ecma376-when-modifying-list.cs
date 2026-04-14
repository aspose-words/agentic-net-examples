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
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string filePath = Path.Combine(outputDir, "ListCompliance.docx");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a numbered list and enable restart-at-each-section (requires higher OOXML compliance).
        List list = doc.Lists.Add(ListTemplate.NumberDefault);
        list.IsRestartAtEachSection = true;

        // Apply the list to some paragraphs.
        builder.ListFormat.List = list;
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Item 3");
        builder.Writeln("Item 4");
        builder.ListFormat.RemoveNumbers();

        // Determine the compliance level needed for the list feature.
        OoxmlCompliance compliance = doc.Compliance;
        if (compliance == OoxmlCompliance.Ecma376_2006)
        {
            // Upgrade to a version that supports IsRestartAtEachSection.
            compliance = OoxmlCompliance.Iso29500_2008_Transitional;
        }

        // Configure save options with the required compliance.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx);
        saveOptions.Compliance = compliance;

        // Save the document.
        doc.Save(filePath, saveOptions);

        // Load the saved document and output its compliance level.
        Document loadedDoc = new Document(filePath);
        Console.WriteLine($"Saved document compliance: {loadedDoc.Compliance}");
    }
}

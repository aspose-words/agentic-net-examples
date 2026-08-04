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

        // Add a simple numbered list.
        doc.Lists.Add(ListTemplate.NumberDefault);
        List list = doc.Lists[0];
        // Enable restarting the list at each new section (advanced list setting).
        list.IsRestartAtEachSection = true;

        // Apply the list to a few paragraphs.
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

        // Determine an output path relative to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AdvancedListStrict.docx");

        // Save the document with the specified compliance level.
        doc.Save(outputPath, saveOptions);

        // Reload the document to verify that the list setting is preserved.
        Document loadedDoc = new Document(outputPath);
        bool isRestart = loadedDoc.Lists[0].IsRestartAtEachSection;

        // Output the verification result.
        Console.WriteLine($"List restart at each section preserved: {isRestart}");
    }
}

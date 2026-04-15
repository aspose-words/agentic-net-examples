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
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a numbered list and enable an advanced setting:
        // restart numbering at each new section.
        doc.Lists.Add(ListTemplate.NumberDefault);
        List list = doc.Lists[0];
        list.IsRestartAtEachSection = true;

        // Apply the list to some paragraphs.
        builder.ListFormat.List = list;
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Item 3");
        builder.Writeln("Item 4");
        builder.ListFormat.RemoveNumbers();

        // Configure OOXML save options to use strict compliance.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx);
        saveOptions.Compliance = OoxmlCompliance.Iso29500_2008_Strict;

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "StrictCompliance.docx");

        // Save the document with the specified compliance level.
        doc.Save(outputPath, saveOptions);
    }
}

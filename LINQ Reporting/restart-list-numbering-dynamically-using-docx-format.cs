using System;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Saving;

class RestartListNumbering
{
    static void Main()
    {
        // Folder where the document will be saved.
        string artifactsDir = @"C:\Artifacts\";

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a numbered list and enable restart at each section.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);
        list.IsRestartAtEachSection = true; // <-- important for DOCX restart behavior

        // Apply the list to the first section.
        builder.ListFormat.List = list;
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");

        // Insert a section break (new page) – the list will restart after this.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Continue the list in the second section; numbering starts again from 1.
        builder.Writeln("Item 1 in second section");
        builder.Writeln("Item 2 in second section");

        // Stop list formatting for any following paragraphs.
        builder.ListFormat.RemoveNumbers();

        // Save the document with a compliance level higher than Ecma376_2006
        // so that the IsRestartAtEachSection flag is written to the DOCX file.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
        {
            Compliance = OoxmlCompliance.Iso29500_2008_Transitional
        };
        doc.Save(artifactsDir + "RestartListAtEachSection.docx", saveOptions);

        // Load the saved document to verify the flag.
        Document loaded = new Document(artifactsDir + "RestartListAtEachSection.docx");
        bool isRestartEnabled = loaded.Lists[0].IsRestartAtEachSection;
        Console.WriteLine($"IsRestartAtEachSection = {isRestartEnabled}");
    }
}

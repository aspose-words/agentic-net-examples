using System;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a multilevel list to demonstrate advanced list settings.
        builder.ListFormat.ApplyNumberDefault();
        builder.Writeln("Item 1");
        builder.ListFormat.ListIndent();
        builder.Writeln("Subitem 1.1");
        builder.ListFormat.ListOutdent();
        builder.Writeln("Item 2");

        // Save the document with strict OOXML compliance to preserve advanced list settings.
        string outputPath = "OutputStrictCompliance.docx";
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx);
        saveOptions.Compliance = OoxmlCompliance.Iso29500_2008_Strict;
        doc.Save(outputPath, saveOptions);
    }
}

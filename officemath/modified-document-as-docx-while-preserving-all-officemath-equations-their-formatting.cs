using System;
using Aspose.Words;
using Aspose.Words.Saving;

class PreserveOfficeMath
{
    static void Main()
    {
        // Load an existing document that contains OfficeMath equations.
        // The Document constructor follows the provided load rule.
        Document doc = new Document("InputWithMath.docx");

        // Example modification: add a paragraph at the end of the document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Additional content added while preserving equations.");

        // Prepare save options for DOCX format.
        // OoxmlSaveOptions follows the provided save rule and allows us to specify the format explicitly.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
        {
            SaveFormat = SaveFormat.Docx   // Ensure the output is DOCX.
        };

        // Save the modified document. This uses the Save(string, SaveOptions) rule.
        doc.Save("OutputPreservedMath.docx", saveOptions);
    }
}

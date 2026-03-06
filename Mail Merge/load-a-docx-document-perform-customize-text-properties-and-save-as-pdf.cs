using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the source DOCX file.
        Document doc = new Document("Input.docx");

        // Example of customizing text properties:
        // Replace the placeholder "[Title]" with the text "Custom Title"
        // and apply a specific font style to the replacement.
        FindReplaceOptions replaceOptions = new FindReplaceOptions();
        replaceOptions.ApplyFont.Name = "Arial";
        replaceOptions.ApplyFont.Size = 24;
        replaceOptions.ApplyFont.Bold = true;

        doc.Range.Replace("[Title]", "Custom Title", replaceOptions);

        // Save the modified document as PDF.
        doc.Save("Output.pdf");
    }
}

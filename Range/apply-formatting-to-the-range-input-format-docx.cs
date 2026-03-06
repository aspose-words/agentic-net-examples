using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class ApplyFormattingToRange
{
    static void Main()
    {
        // Load the source DOCX document.
        string inputPath = "input.docx";
        Document doc = new Document(inputPath);

        // Define find-and-replace options to apply formatting to the replaced content.
        FindReplaceOptions options = new FindReplaceOptions();

        // Example: center-align paragraphs that contain the replacement.
        options.ApplyParagraphFormat.Alignment = ParagraphAlignment.Center;

        // Example: make the replacement text bold.
        // The ReplacementFormat property can be used to specify that the replacement is treated as a field,
        // but for simple text formatting we modify the paragraph format as above.
        // Replace the placeholder "{placeholder}" with the desired text.
        string placeholder = "{placeholder}";
        string replacement = "Formatted Text";

        // Perform the replace operation with the specified formatting.
        doc.Range.Replace(placeholder, replacement, options);

        // Save the modified document.
        string outputPath = "output.docx";
        doc.Save(outputPath, SaveFormat.Docx);
    }
}

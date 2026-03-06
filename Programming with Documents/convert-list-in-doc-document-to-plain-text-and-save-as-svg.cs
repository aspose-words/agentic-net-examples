using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertListToPlainTextSvg
{
    static void Main()
    {
        // Load the source DOC document that contains a list.
        // The constructor with a file name follows the provided load rule.
        Document sourceDoc = new Document("InputDocument.doc");

        // Prepare text save options to get a clean plain‑text representation of the list.
        // SimplifyListLabels makes list symbols easier to read in plain text.
        TxtSaveOptions txtOptions = new TxtSaveOptions
        {
            SimplifyListLabels = true
        };

        // Export the whole document to plain text using the options above.
        // The ToString method with SaveOptions follows the provided save rule.
        string plainText = sourceDoc.ToString(txtOptions);

        // Create a new blank document that will hold the extracted plain text.
        // Using the parameterless constructor follows the provided create rule.
        Document textDoc = new Document();

        // Insert the plain‑text string into the new document.
        // DocumentBuilder is used to add a paragraph containing the text.
        DocumentBuilder builder = new DocumentBuilder(textDoc);
        builder.Writeln(plainText);

        // Configure SVG save options.
        // UsePlacedGlyphs renders text as curves, ensuring the SVG contains the text visually.
        SvgSaveOptions svgOptions = new SvgSaveOptions
        {
            TextOutputMode = SvgTextOutputMode.UsePlacedGlyphs,
            FitToViewPort = true,
            ShowPageBorder = false
        };

        // Save the document that now contains only the plain‑text list as an SVG file.
        // The Save method with a file name and SaveOptions follows the provided save rule.
        textDoc.Save("ListPlainText.svg", svgOptions);
    }
}

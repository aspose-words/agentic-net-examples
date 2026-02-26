using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("Input.doc");

        // Optional: extract the whole document as plain text (lists are included).
        // This demonstrates that the list content is converted to plain text.
        string plainText = doc.ToString(SaveFormat.Text);
        Console.WriteLine(plainText);

        // Configure SVG save options.
        // UsePlacedGlyphs renders all text as vector curves, which is suitable for SVG output.
        SvgSaveOptions svgOptions = new SvgSaveOptions
        {
            TextOutputMode = SvgTextOutputMode.UsePlacedGlyphs,
            ShowPageBorder = false,
            FitToViewPort = true
        };

        // Save the document as an SVG file.
        doc.Save("Output.svg", svgOptions);
    }
}

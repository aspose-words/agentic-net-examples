using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fields; // Added for Field and FieldType

class RemoveTocAndSaveSvg
{
    static void Main()
    {
        // Path to the source DOC document.
        string inputPath = "input.doc";

        // Path where the resulting SVG will be saved.
        string outputPath = "output.svg";

        // Load the document from the file system.
        Document doc = new Document(inputPath);

        // Remove all Table of Contents (TOC) fields from the document.
        // Collect fields to remove first to avoid modifying the collection while iterating.
        var tocFields = new System.Collections.Generic.List<Field>();
        foreach (Field field in doc.Range.Fields)
        {
            if (field.Type == FieldType.FieldTOC)
                tocFields.Add(field);
        }
        foreach (Field field in tocFields)
        {
            field.Remove();
        }

        // Configure SVG save options.
        SvgSaveOptions options = new SvgSaveOptions
        {
            // Make the SVG fill the viewport (optional, can be omitted).
            FitToViewPort = true,
            // Do not draw a page border around the SVG.
            ShowPageBorder = false,
            // Render text as placed glyphs (curves) to avoid font issues.
            TextOutputMode = SvgTextOutputMode.UsePlacedGlyphs
        };

        // Save the modified document as an SVG file using the configured options.
        doc.Save(outputPath, options);
    }
}

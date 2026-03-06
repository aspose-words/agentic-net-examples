using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Saving;

class RemoveTocAndConvertToPdf
{
    static void Main()
    {
        // Path to the source DOC document.
        string inputPath = @"C:\Docs\SourceDocument.doc";

        // Path where the resulting PDF will be saved.
        string outputPath = @"C:\Docs\ResultDocument.pdf";

        // Load the existing Word document.
        Document doc = new Document(inputPath);

        // Remove all Table of Contents (TOC) fields from the document.
        // Iterate backwards to safely modify the collection while removing items.
        for (int i = doc.Range.Fields.Count - 1; i >= 0; i--)
        {
            Field field = doc.Range.Fields[i];
            if (field.Type == FieldType.FieldTOC)
            {
                // Optionally remove the whole paragraph that contains the TOC field.
                // field.Remove(); // This removes only the field code.
                Node parentParagraph = field.Start.ParentNode;
                if (parentParagraph != null && parentParagraph.NodeType == NodeType.Paragraph)
                {
                    parentParagraph.Remove();
                }
                else
                {
                    // Fallback: just remove the field if the paragraph cannot be located.
                    field.Remove();
                }
            }
        }

        // Save the modified document as PDF.
        doc.Save(outputPath, SaveFormat.Pdf);
    }
}

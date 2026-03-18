using System;
using System.Drawing;
using Aspose.Words;

namespace CopyFontFormattingExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new document
            Document doc = new Document();

            // Create a source run with some font formatting
            Run sourceRun = new Run(doc, "Source text");
            sourceRun.Font.Name = "Arial";
            sourceRun.Font.Size = 24;
            sourceRun.Font.Bold = true;
            sourceRun.Font.Color = Color.Blue;

            // Create a target run with default formatting
            Run targetRun = new Run(doc, "Target text");

            // Copy the font formatting from the source run to the target run.
            // The Font property of Inline (base class of Run) is read‑only, so we cannot assign a new Font object.
            // Instead we copy the individual font attributes.
            targetRun.Font.Name = sourceRun.Font.Name;
            targetRun.Font.Size = sourceRun.Font.Size;
            targetRun.Font.Bold = sourceRun.Font.Bold;
            targetRun.Font.Italic = sourceRun.Font.Italic;
            targetRun.Font.Underline = sourceRun.Font.Underline;
            targetRun.Font.Color = sourceRun.Font.Color;
            targetRun.Font.HighlightColor = sourceRun.Font.HighlightColor;
            // Add any other font properties you need to copy.

            // Append both runs to the first paragraph of the document
            Paragraph para = doc.FirstSection.Body.FirstParagraph;
            para.AppendChild(sourceRun);
            para.AppendChild(targetRun);

            // Save the document
            doc.Save("CopyFontFormatting.docx");
        }
    }
}

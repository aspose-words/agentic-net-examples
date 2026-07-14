using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.BuildingBlocks;
using Aspose.Words.Themes;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Ensure the document has a glossary document (required for building blocks).
        if (doc.GlossaryDocument == null)
            doc.GlossaryDocument = new GlossaryDocument();

        // Access the document's theme and pick a theme color (Accent1 in this example).
        Color themeAccentColor = doc.Theme.Colors.Accent1;

        // Create a building block that will serve as the placeholder text.
        GlossaryDocument glossary = doc.GlossaryDocument;
        BuildingBlock placeholderBlock = new BuildingBlock(glossary)
        {
            Name = "MyPlaceholder"
        };

        // The building block must contain a section with a body and a paragraph.
        Section blockSection = new Section(glossary);
        placeholderBlock.AppendChild(blockSection);
        blockSection.EnsureMinimum(); // Creates Body and first Paragraph.

        // Add the placeholder run with the theme color.
        Run placeholderRun = new Run(glossary, "Placeholder text matching theme color");
        placeholderRun.Font.Color = themeAccentColor;
        blockSection.Body.FirstParagraph.AppendChild(placeholderRun);

        // Add the building block to the glossary document.
        glossary.AppendChild(placeholderBlock);

        // Insert an inline plain‑text content control (structured document tag).
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "SampleControl",
            Tag = "SampleTag",
            PlaceholderName = "MyPlaceholder"
        };

        // Insert the content control into the first paragraph of the document.
        Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;
        firstParagraph.AppendChild(sdt);

        // Save the resulting document.
        doc.Save("PlaceholderThemeColor.docx");
    }
}

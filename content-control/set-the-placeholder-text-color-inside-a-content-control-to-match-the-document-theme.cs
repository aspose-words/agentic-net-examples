using System;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.BuildingBlocks;
using Aspose.Words.Themes;

namespace ContentControlPlaceholderTheme
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Ensure the document has a glossary (required for building blocks).
            if (doc.GlossaryDocument == null)
                doc.GlossaryDocument = new GlossaryDocument();

            // Access the document theme and set a known accent color.
            Theme theme = doc.Theme;
            ThemeColors themeColors = theme.Colors;
            themeColors.Accent1 = System.Drawing.Color.CornflowerBlue; // Example theme accent color.

            // -----------------------------------------------------------------
            // Create a building block that will serve as the placeholder text.
            // -----------------------------------------------------------------
            GlossaryDocument glossary = doc.GlossaryDocument;

            BuildingBlock placeholderBlock = new BuildingBlock(glossary)
            {
                Name = "MyPlaceholder",
                // Mark the block as a placeholder for a StructuredDocumentTag.
                Type = BuildingBlockType.StructuredDocumentTagPlaceholderText
            };

            // Build the placeholder block: Section -> Body -> Paragraph -> Run.
            Section blockSection = new Section(glossary);
            // Ensure the section has a body and a paragraph.
            blockSection.EnsureMinimum();

            // Create the run with placeholder text and apply the theme accent color.
            Run placeholderRun = new Run(glossary, "Enter name here");
            placeholderRun.Font.Color = themeColors.Accent1;

            // Insert the run into the first paragraph of the section.
            blockSection.Body.FirstParagraph.AppendChild(placeholderRun);

            // Attach the section to the building block.
            placeholderBlock.AppendChild(blockSection);

            // Add the building block to the document glossary.
            glossary.AppendChild(placeholderBlock);

            // ---------------------------------------------------------------
            // Insert an inline plain‑text content control linked to the placeholder.
            // ---------------------------------------------------------------
            StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
            {
                Title = "Name",
                PlaceholderName = "MyPlaceholder"
            };

            // Place the content control into the first paragraph of the document.
            Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;
            firstParagraph.AppendChild(sdt);

            // Save the resulting document.
            doc.Save("PlaceholderTheme.docx");
        }
    }
}

using System;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Themes;
using Aspose.Words.BuildingBlocks;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Ensure the document has a glossary document (used for building blocks).
        GlossaryDocument glossary = doc.GlossaryDocument;
        if (glossary == null)
        {
            glossary = new GlossaryDocument();
            doc.GlossaryDocument = glossary;
        }

        // -----------------------------------------------------------------
        // Create a building block that will serve as the placeholder text.
        // -----------------------------------------------------------------
        BuildingBlock placeholderBlock = new BuildingBlock(glossary)
        {
            Name = "MyPlaceholder"
        };

        // Build the block structure: Section -> Body -> Paragraph -> Run.
        Section blockSection = new Section(glossary);
        Body blockBody = new Body(glossary);
        Paragraph blockParagraph = new Paragraph(glossary);
        Run blockRun = new Run(glossary, "Placeholder text shown when the content control is empty.");

        // Apply a theme color (Accent1) to the placeholder run.
        blockRun.Font.ThemeColor = ThemeColor.Accent1;

        // Assemble the block.
        blockParagraph.AppendChild(blockRun);
        blockBody.AppendChild(blockParagraph);
        blockSection.AppendChild(blockBody);
        placeholderBlock.AppendChild(blockSection);
        glossary.AppendChild(placeholderBlock);

        // ---------------------------------------------------------------
        // Insert a plain‑text content control and link it to the placeholder.
        // ---------------------------------------------------------------
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "Sample Content Control",
            PlaceholderName = "MyPlaceholder",
            IsShowingPlaceholderText = true
        };

        // Insert the content control into the document at the builder's current position.
        builder.InsertNode(sdt);

        // Save the resulting document.
        doc.Save("PlaceholderThemeColor.docx");
    }
}

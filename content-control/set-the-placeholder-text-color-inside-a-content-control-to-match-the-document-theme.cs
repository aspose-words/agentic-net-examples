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

        // Set a theme accent color that we will use for the placeholder text.
        Theme theme = doc.Theme;
        theme.Colors.Accent1 = Color.Blue;

        // Ensure the document has a glossary document (container for building blocks).
        GlossaryDocument glossary = doc.GlossaryDocument;
        if (glossary == null)
        {
            glossary = new GlossaryDocument();
            doc.GlossaryDocument = glossary;
        }

        // Create a building block that will serve as the placeholder text.
        BuildingBlock placeholderBlock = new BuildingBlock(glossary)
        {
            Name = "MyPlaceholder"
        };

        // Build the placeholder content: Section -> Body -> Paragraph -> Run.
        Section placeholderSection = new Section(glossary);
        placeholderBlock.AppendChild(placeholderSection);

        Body placeholderBody = new Body(glossary);
        placeholderSection.AppendChild(placeholderBody);

        Paragraph placeholderParagraph = new Paragraph(glossary);
        placeholderBody.AppendChild(placeholderParagraph);

        Run placeholderRun = new Run(glossary, "Placeholder text");
        // Apply the theme accent color to the placeholder font.
        placeholderRun.Font.Color = theme.Colors.Accent1;
        placeholderParagraph.AppendChild(placeholderRun);

        // Add the building block to the document's glossary.
        glossary.AppendChild(placeholderBlock);

        // Insert a plain‑text content control and link it to the placeholder building block.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Before the content control:");

        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "SampleControl",
            PlaceholderName = "MyPlaceholder"
        };
        builder.InsertNode(sdt);

        builder.Writeln("After the content control.");

        // Save the resulting document.
        doc.Save("PlaceholderTheme.docx");
    }
}

using System;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Themes; // <-- required for ThemeColor enum

class Program
{
    static void Main()
    {
        // Create a new blank document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a plain‑text content control (structured document tag)
        StructuredDocumentTag tag = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);

        // Make the control display its placeholder text
        tag.IsShowingPlaceholderText = true;

        // Set the placeholder text color to a theme color (e.g., Accent2)
        // Font.ThemeColor expects a value from the ThemeColor enum which lives in Aspose.Words.Themes
        tag.ContentsFont.ThemeColor = ThemeColor.Accent2;

        // Insert the content control into the document
        builder.InsertNode(tag);

        // Save the document
        doc.Save("PlaceholderThemeColor.docx");
    }
}

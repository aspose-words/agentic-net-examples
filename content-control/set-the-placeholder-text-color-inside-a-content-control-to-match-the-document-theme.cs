using System;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Themes;

namespace PlaceholderColorExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a plain‑text content control (structured document tag) at the current cursor position.
            StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
            {
                Title = "SamplePlaceholder",
                Tag = "SampleTag",
                // Ensure the placeholder text is shown when the control is empty.
                IsShowingPlaceholderText = true
            };
            builder.InsertNode(sdt);

            // Retrieve a theme color (e.g., Accent1) from the document's theme.
            // This color will be used for the placeholder frame of the content control.
            System.Drawing.Color themeAccent = doc.Theme.Colors.Accent1;

            // Apply the theme color to the content control's frame.
            sdt.Color = themeAccent;

            // Save the resulting document.
            doc.Save("PlaceholderColor.docx");
        }
    }
}

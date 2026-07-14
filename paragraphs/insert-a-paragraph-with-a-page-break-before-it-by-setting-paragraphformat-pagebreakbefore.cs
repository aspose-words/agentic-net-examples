using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder which will be used to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the PageBreakBefore flag so that the next paragraph starts on a new page.
        builder.ParagraphFormat.PageBreakBefore = true;

        // Insert the paragraph text. The builder will apply the current ParagraphFormat.
        builder.Writeln("This paragraph starts on a new page because of PageBreakBefore.");

        // Save the document to a file in the current directory.
        doc.Save("ParagraphWithPageBreakBefore.docx");
    }
}

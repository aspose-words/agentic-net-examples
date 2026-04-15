using System;
using Aspose.Words;

namespace AsposeWordsExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Force a page break before the next paragraph.
            builder.ParagraphFormat.PageBreakBefore = true;

            // Insert the paragraph text.
            builder.Writeln("This paragraph starts on a new page.");

            // Save the document.
            doc.Save("ParagraphWithPageBreakBefore.docx");
        }
    }
}

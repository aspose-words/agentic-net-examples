using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTabStopExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Attach a DocumentBuilder to the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a custom tab stop at 2 inches (2 * 72 points = 144 points).
            // The tab stop is left-aligned with no leader.
            builder.ParagraphFormat.TabStops.Add(144.0, TabAlignment.Left, TabLeader.None);

            // Write some text, insert a tab character, then write more text.
            // The text after the tab will align to the custom tab stop at 2 inches.
            builder.Write("First part");
            builder.Write(ControlChar.Tab);
            builder.Write("Aligned at 2 inches");

            // Finish the paragraph.
            builder.Writeln();

            // Save the document to a file in the current directory.
            doc.Save("CustomTabStop.docx");
        }
    }
}

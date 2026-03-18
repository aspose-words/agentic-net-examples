using System;
using Aspose.Words;
using Aspose.Words.Tables;

class TabStopExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // 1 inch = 72 points. Add a custom tab stop at 2 inches (144 points).
        // Align text to the left of the tab stop, no leader line.
        builder.ParagraphFormat.TabStops.Add(144.0, TabAlignment.Left, TabLeader.None);

        // Insert text with a tab character. The text after the tab will align to the custom tab stop.
        builder.Writeln("Label:" + ControlChar.Tab + "Value aligned at 2 inches");

        // Save the document.
        doc.Save("TabStopExample.docx");
    }
}

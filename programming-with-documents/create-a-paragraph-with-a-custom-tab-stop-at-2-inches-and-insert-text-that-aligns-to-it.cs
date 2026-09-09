using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder attached to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a custom tab stop at 2 inches (2 * 72 points = 144 points).
        // Align text to the left at this tab stop.
        builder.ParagraphFormat.TabStops.Add(144.0, TabAlignment.Left, TabLeader.None);

        // Insert a paragraph with text that uses the tab character to align to the custom tab stop.
        builder.Writeln("First part" + ControlChar.Tab + "Second part aligned at 2 inches");

        // Save the document to the local file system.
        doc.Save("CustomTabStop.docx");
    }
}

using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a custom tab stop at 2 inches (144 points) aligned to the left.
        builder.ParagraphFormat.TabStops.Add(144.0, TabAlignment.Left, TabLeader.None);

        // Insert a paragraph with text before and after the tab character.
        // The text after the tab will align to the custom tab stop.
        builder.Writeln("First part" + ControlChar.Tab + "Second part aligned at 2 inches");

        // Save the document to the current working directory.
        doc.Save("CustomTabStop.docx");
    }
}

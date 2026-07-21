using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder attached to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a custom tab stop at 2 inches (144 points) with left alignment.
        // TabLeader.None means no leader line will be displayed.
        builder.ParagraphFormat.TabStops.Add(144.0, TabAlignment.Left, TabLeader.None);

        // Insert text that uses the custom tab stop.
        // The text before the tab will be left-aligned, and the text after the tab
        // will start at the 2‑inch position.
        builder.Writeln("Item" + ControlChar.Tab + "Aligned at 2 inches");

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "CustomTabStop.docx");
        doc.Save(outputPath);
    }
}

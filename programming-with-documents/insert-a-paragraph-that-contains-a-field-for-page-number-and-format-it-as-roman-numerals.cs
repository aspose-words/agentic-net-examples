using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with the text "Page " followed by a PAGE field.
        builder.Write("Page ");
        builder.InsertField("PAGE", ""); // Inserts the PAGE field.
        builder.Writeln(); // Ends the paragraph.

        // Configure the document to display page numbers as uppercase Roman numerals.
        doc.FirstSection.PageSetup.PageNumberStyle = NumberStyle.UppercaseRoman;
        // Ensure numbering starts from 1.
        doc.FirstSection.PageSetup.RestartPageNumbering = true;
        doc.FirstSection.PageSetup.PageStartingNumber = 1;

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "PageNumberRoman.docx");
        doc.Save(outputPath);
    }
}

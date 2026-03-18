using System;
using Aspose.Words;

public static class RunExtensions
{
    // Reusable method that sets the font name and size for any Run instance.
    public static void ApplyFont(this Run run, string fontName, double fontSize)
    {
        if (run == null) throw new ArgumentNullException(nameof(run));
        if (string.IsNullOrEmpty(fontName)) throw new ArgumentException("Font name cannot be null or empty.", nameof(fontName));
        if (fontSize <= 0) throw new ArgumentOutOfRangeException(nameof(fontSize), "Font size must be greater than zero.");

        // Apply the desired font name. This updates NameAscii, NameBi, NameFarEast, and NameOther.
        run.Font.Name = fontName;

        // Apply the desired font size (in points).
        run.Font.Size = fontSize;
    }
}

public class Example
{
    public void CreateDocument()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Obtain the first paragraph (creates one automatically if the document is empty).
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

        // Create a run with some sample text.
        Run run = new Run(doc, "Hello Aspose.Words!");

        // Apply custom font formatting using the reusable extension method.
        run.ApplyFont("Calibri", 18);

        // Insert the run into the paragraph.
        paragraph.AppendChild(run);

        // Save the document.
        doc.Save("FormattedRun.docx");
    }
}

public static class Program
{
    // Entry point required for a console application.
    public static void Main(string[] args)
    {
        // Instantiate Example and run the document creation logic.
        new Example().CreateDocument();
        Console.WriteLine("Document created successfully.");
    }
}

using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a primary footer.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("Primary footer text.");

        // Add a first-page footer.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterFirst);
        builder.Writeln("First page footer text.");

        // Add an even-page footer.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterEven);
        builder.Writeln("Even page footer text.");

        // Remove all footers (and any headers) from the first section.
        doc.FirstSection.HeadersFooters.Clear();

        // Save the resulting document.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Artifacts");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "DocumentWithoutFooters.docx");
        doc.Save(outputPath);
    }
}

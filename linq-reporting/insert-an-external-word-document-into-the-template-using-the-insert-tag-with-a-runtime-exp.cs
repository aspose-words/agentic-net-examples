using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Wrapper
{
    public Document Document { get; set; }

    public Wrapper(Document doc)
    {
        Document = doc ?? throw new ArgumentNullException(nameof(doc));
    }
}

public class Program
{
    public static void Main()
    {
        // Create the external document that will be inserted.
        var sourceDoc = new Document();
        var srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("This is the content of the external document.");
        const string sourcePath = "Source.docx";
        sourceDoc.Save(sourcePath);

        // Create the template document containing the insert tag.
        var templateDoc = new Document();
        var tmplBuilder = new DocumentBuilder(templateDoc);
        tmplBuilder.Writeln("Before inserted document:");
        tmplBuilder.Writeln("<<doc [src.Document]>>");
        tmplBuilder.Writeln("After inserted document.");

        // Wrap the external document for the reporting engine.
        var wrapper = new Wrapper(new Document(sourcePath));

        // Build the report using the wrapper as the data source.
        var engine = new ReportingEngine();
        engine.BuildReport(templateDoc, wrapper, "src");

        // Save the final document.
        const string outputPath = "Result.docx";
        templateDoc.Save(outputPath);
    }
}

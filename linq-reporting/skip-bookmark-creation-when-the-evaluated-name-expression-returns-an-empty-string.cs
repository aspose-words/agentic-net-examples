using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    public string BookmarkName { get; set; } = "";
    public string Title { get; set; } = "";
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some environments)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create a template document with a conditional bookmark.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Conditional block: create bookmark only if BookmarkName is not empty.
        builder.Writeln("<<if [model.BookmarkName != \"\"]>>");
        builder.Writeln("<<bookmark [model.BookmarkName]>>");
        builder.Writeln("<<[model.Title]>>");
        builder.Writeln("<</bookmark>>");
        builder.Writeln("<</if>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for report generation.
        Document docWithBookmark = new Document(templatePath);
        Document docWithoutBookmark = new Document(templatePath);

        // Prepare data models.
        ReportModel modelWith = new ReportModel
        {
            BookmarkName = "MyBookmark",
            Title = "Report with bookmark"
        };

        ReportModel modelWithout = new ReportModel
        {
            BookmarkName = "", // Empty name – bookmark should be skipped.
            Title = "Report without bookmark"
        };

        // Build reports.
        ReportingEngine engine = new ReportingEngine();

        engine.BuildReport(docWithBookmark, modelWith, "model");
        engine.BuildReport(docWithoutBookmark, modelWithout, "model");

        // Save the generated reports.
        docWithBookmark.Save("ReportWithBookmark.docx");
        docWithoutBookmark.Save("ReportWithoutBookmark.docx");
    }
}

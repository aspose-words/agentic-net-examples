using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare paths
        string templatePath = "Template.docx";
        string resultPath = "Result.docx";

        // 1. Create the template document programmatically
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Paragraph with a required field
        builder.Writeln("Name: <<[model.Name]>>");

        // Paragraph that may become empty after rendering (optional field)
        builder.Writeln("<<[model.Optional]>>");

        // Save the template
        templateDoc.Save(templatePath);

        // 2. Load the template for reporting
        Document doc = new Document(templatePath);

        // 3. Prepare the data model
        ReportModel model = new ReportModel
        {
            Name = "John Doe",
            Optional = null // This will cause the second paragraph to be empty after rendering
        };

        // 4. Build the report using LINQ Reporting Engine
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // 5. Remove empty paragraphs that resulted from empty placeholders
        List<Paragraph> emptyParagraphs = new();
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            // GetText includes the paragraph break; Trim removes whitespace and the break
            if (string.IsNullOrWhiteSpace(para.GetText()))
                emptyParagraphs.Add(para);
        }

        foreach (Paragraph para in emptyParagraphs)
            para.Remove();

        // 6. Save the final document
        doc.Save(resultPath);
    }

    // Public data model class required by the reporting engine
    public class ReportModel
    {
        public string Name { get; set; } = string.Empty;
        public string? Optional { get; set; }
    }
}

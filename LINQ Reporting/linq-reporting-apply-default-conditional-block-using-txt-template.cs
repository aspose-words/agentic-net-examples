using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Prepare the data source that will be referenced from the template.
        // -----------------------------------------------------------------
        var model = new ReportModel
        {
            Name = "John Doe",
            IsMember = true
        };

        // -----------------------------------------------------------------
        // 2. Load the TXT template that contains LINQ Reporting tags.
        //    Example template content (Template.txt):
        //    <<if [model.IsMember]>>
        //    Hello <<[model.Name]>>, you are a member.
        //    <<else>>
        //    Hello <<[model.Name]>>, you are not a member.
        //    <<endif>>
        // -----------------------------------------------------------------
        const string templatePath = "Template.txt";
        Document doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Configure the ReportingEngine.
        //    RemoveEmptyParagraphs ensures that any paragraph left empty after
        //    conditional processing is removed from the final output.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.RemoveEmptyParagraphs
        };

        // -----------------------------------------------------------------
        // 4. Build the report. The data source is passed together with a name
        //    ("model") that is used inside the template tags.
        // -----------------------------------------------------------------
        engine.BuildReport(doc, model, "model");

        // -----------------------------------------------------------------
        // 5. Save the generated document as plain text.
        //    TxtSaveOptions allows us to customize how the text is written.
        // -----------------------------------------------------------------
        TxtSaveOptions saveOptions = new TxtSaveOptions
        {
            // Custom paragraph break – can be any string you need.
            ParagraphBreak = "\r\n---\r\n"
        };

        doc.Save("Report.txt", saveOptions);
    }

    // -----------------------------------------------------------------
    // Simple POCO class that represents the data used in the template.
    // -----------------------------------------------------------------
    public class ReportModel
    {
        public string Name { get; set; }
        public bool IsMember { get; set; }
    }
}

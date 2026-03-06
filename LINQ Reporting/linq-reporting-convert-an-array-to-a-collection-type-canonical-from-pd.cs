using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the source PDF document (lifecycle: load rule)
        Document pdfDoc = new Document("Input.pdf");

        // Convert the SectionCollection to an array using the provided ToArray method (rule)
        Section[] sectionsArray = pdfDoc.Sections.ToArray();

        // Convert the array to a canonical collection type (List<Section>)
        List<Section> sectionsList = sectionsArray.ToList();

        // Create a simple template document in memory (lifecycle: create rule)
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Placeholder to display the total number of sections
        builder.Writeln("Number of sections: <<[ds.Count]>>");

        // Loop through each section and output its body text
        builder.Writeln("<<foreach [ds]>><<[Body]>>\n<</foreach>>");

        // Initialize the ReportingEngine
        ReportingEngine engine = new ReportingEngine();

        // Build the report using the list as the data source; expose it as "ds" in the template
        engine.BuildReport(template, sectionsList, "ds");

        // Save the generated report as PDF (lifecycle: save rule)
        template.Save("ReportFromPdf.pdf", SaveFormat.Pdf);
    }
}

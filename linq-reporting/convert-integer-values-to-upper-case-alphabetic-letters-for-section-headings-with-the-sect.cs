using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for any encoding needs.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for the template and the generated report.
        string templatePath = "Template.docx";
        string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Begin a foreach loop over the Sections collection.
        builder.Writeln("<<foreach [sec in Sections]>>");
        // Write a heading that shows the numeric value and its corresponding letter.
        builder.Writeln("Section <<[sec.Number]>>: <<[sec.Letter]>>");
        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Prepare the data model.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel
        {
            Sections = new List<Section>
            {
                new Section { Number = 1 },
                new Section { Number = 2 },
                new Section { Number = 3 },
                new Section { Number = 4 },
                new Section { Number = 5 }
            }
        };

        // -----------------------------------------------------------------
        // 3. Load the template and build the report.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        reportDoc.Save(reportPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes.
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<Section> Sections { get; set; } = new();
}

public class Section
{
    public int Number { get; set; }

    // Converts the integer to an uppercase alphabetic representation (A, B, ..., Z, AA, AB, ...).
    public string Letter => ConvertToLetters(Number);

    private static string ConvertToLetters(int number)
    {
        if (number <= 0) return string.Empty;

        string result = string.Empty;
        int n = number;
        while (n > 0)
        {
            n--; // Adjust because 'A' corresponds to 1.
            char ch = (char)('A' + (n % 26));
            result = ch + result;
            n /= 26;
        }
        return result;
    }
}

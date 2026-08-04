using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Aspose.Words.Tables;

public class OfficeMathReportGenerator
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a few sample equations using the deterministic EQ‑field bootstrap workflow.
        InsertOfficeMath(builder, @"\f(1,2)");          // Fraction 1/2
        InsertOfficeMath(builder, @"\r(3,x)");          // Cube root of x
        InsertOfficeMath(builder, @"\i \su(n=1,5,n)"); // Integral with summation

        // Save the sample document.
        string docPath = "Sample.docx";
        doc.Save(docPath);

        // Enumerate all OfficeMath nodes in the document.
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        NodeCollection allParagraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        StringBuilder reportBuilder = new StringBuilder();
        for (int i = 0; i < officeMathNodes.Count; i++)
        {
            OfficeMath om = (OfficeMath)officeMathNodes[i];
            string mathObjectType = om.MathObjectType.ToString();

            // Determine the global paragraph index that contains this OfficeMath.
            Paragraph parentParagraph = om.ParentParagraph;
            int paragraphIndex = parentParagraph != null ? allParagraphs.IndexOf(parentParagraph) : -1;

            // Determine the section index that contains this OfficeMath.
            Section parentSection = parentParagraph?.ParentSection;
            int sectionIndex = parentSection != null ? doc.Sections.IndexOf(parentSection) : -1;

            reportBuilder.AppendLine(
                $"Equation {i + 1}: MathObjectType = {mathObjectType}, Section = {sectionIndex}, Paragraph = {paragraphIndex}");
        }

        // Write the report to a text file.
        string reportPath = "OfficeMathReport.txt";
        File.WriteAllText(reportPath, reportBuilder.ToString());

        // Validate that the report file was created.
        if (!File.Exists(reportPath))
            throw new InvalidOperationException("Report file was not created.");

        // Optionally, output the report location (no interactive input required).
        Console.WriteLine($"Report generated: {Path.GetFullPath(reportPath)}");
    }

    // Helper that inserts an EQ field, converts it to OfficeMath, and removes the field.
    private static OfficeMath InsertOfficeMath(DocumentBuilder builder, string eqArguments)
    {
        // Insert an empty EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the EQ arguments into the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(eqArguments);

        // Return to the paragraph that contains the field.
        builder.MoveTo(field.Start.ParentNode);

        // Convert the field to a real OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();

        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start node.
            field.Start.ParentNode.InsertBefore(officeMath, field.Start);
            // Remove the original field.
            field.Remove();
        }

        // Insert a new paragraph after the equation for readability.
        builder.InsertParagraph();

        return officeMath;
    }
}

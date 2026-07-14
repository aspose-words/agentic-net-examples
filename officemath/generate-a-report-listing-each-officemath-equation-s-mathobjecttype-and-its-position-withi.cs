using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class OfficeMathReportExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert several equations using the deterministic EQ‑field bootstrap workflow.
        InsertEquation(builder, @"\f(1,2)");          // Fraction 1/2
        InsertEquation(builder, @"\r(3,x)");          // Cube root of x
        InsertEquation(builder, @"\i \su(n=1,5,n)"); // Integral with summation

        // Save the document (optional, just to have an output file).
        string docPath = "Sample.docx";
        doc.Save(docPath);

        // Enumerate all OfficeMath nodes in the document.
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        StringBuilder reportBuilder = new StringBuilder();
        reportBuilder.AppendLine("OfficeMath Equation Report");
        reportBuilder.AppendLine("---------------------------");

        for (int i = 0; i < mathNodes.Count; i++)
        {
            OfficeMath officeMath = (OfficeMath)mathNodes[i];
            // Position is the zero‑based index of the equation in the document order.
            int position = i;
            // MathObjectType indicates the type of the OfficeMath node.
            MathObjectType type = officeMath.MathObjectType;

            reportBuilder.AppendLine($"Equation {i + 1}:");
            reportBuilder.AppendLine($"  MathObjectType = {type}");
            reportBuilder.AppendLine($"  Position       = {position}");
            reportBuilder.AppendLine();
        }

        // Write the report to a text file.
        string reportPath = "OfficeMathReport.txt";
        File.WriteAllText(reportPath, reportBuilder.ToString());

        // Validate that the report file was created.
        if (!File.Exists(reportPath))
            throw new InvalidOperationException("Report file was not created.");

        // Output the report path to the console for verification.
        Console.WriteLine($"Report generated: {Path.GetFullPath(reportPath)}");
    }

    // Helper method that inserts an EQ field, converts it to OfficeMath, and removes the field.
    private static void InsertEquation(DocumentBuilder builder, string eqArguments)
    {
        // Insert an empty EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Move to the field separator and write the EQ arguments.
        builder.MoveTo(field.Separator);
        builder.Write(eqArguments);

        // Update the field so that the EQ code is processed.
        field.Update();

        // Return the builder to the field start position.
        builder.MoveTo(field.Start.ParentNode);

        // Convert the field to a real OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Insert the OfficeMath node before the field start.
        field.Start.ParentNode.InsertBefore(officeMath, field.Start);
        // Remove the original field.
        field.Remove();

        // Insert a paragraph break after the equation for readability.
        builder.InsertParagraph();
    }
}

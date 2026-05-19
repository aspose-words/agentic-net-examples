using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class OfficeMathReportGenerator
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert several equations using the deterministic EQ-field bootstrap workflow.
        InsertEquation(builder, @"\f(1,2)");                     // Simple fraction 1/2
        builder.Writeln();                                     // New paragraph
        InsertEquation(builder, @"\r(3,x)");                     // Cube root of x
        builder.Writeln();                                     // New paragraph
        InsertEquation(builder, @"\i \su(n=1,5,n)");            // Integral with summation
        builder.Writeln();                                     // New paragraph
        InsertEquation(builder, @"\s \up8(Superscript) \s \do8(Subscript)"); // Superscript and subscript

        // Save the document containing the equations.
        string docPath = Path.Combine(Environment.CurrentDirectory, "OfficeMathReport.docx");
        doc.Save(docPath);

        // Generate a report listing each OfficeMath equation's MathObjectType and its position.
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        NodeCollection paragraphNodes = doc.GetChildNodes(NodeType.Paragraph, true);

        StringBuilder reportBuilder = new StringBuilder();
        reportBuilder.AppendLine("OfficeMath Equations Report");
        reportBuilder.AppendLine("---------------------------");

        for (int i = 0; i < officeMathNodes.Count; i++)
        {
            OfficeMath officeMath = (OfficeMath)officeMathNodes[i];
            string mathObjectType = officeMath.MathObjectType.ToString();

            // Determine the index of the parent paragraph within the document.
            Paragraph parentParagraph = officeMath.ParentParagraph;
            int paragraphIndex = paragraphNodes.IndexOf(parentParagraph);

            reportBuilder.AppendLine($"Equation {i + 1}:");
            reportBuilder.AppendLine($"  MathObjectType : {mathObjectType}");
            reportBuilder.AppendLine($"  ParagraphIndex : {paragraphIndex}");
            reportBuilder.AppendLine();
        }

        // Write the report to a text file.
        string reportPath = Path.Combine(Environment.CurrentDirectory, "OfficeMathReport.txt");
        File.WriteAllText(reportPath, reportBuilder.ToString());

        // Validate that the report file was created.
        if (!File.Exists(reportPath))
            throw new InvalidOperationException("Report file was not created.");

        Console.WriteLine($"Document saved to: {docPath}");
        Console.WriteLine($"Report saved to: {reportPath}");
    }

    // Inserts an equation into the document using the EQ field bootstrap method.
    private static void InsertEquation(DocumentBuilder builder, string equationSwitches)
    {
        // Insert an EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the equation arguments into the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(equationSwitches);

        // Update the field to ensure the equation code is processed.
        field.Update();

        // Convert the EQ field to a real OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Insert the OfficeMath node before the field start node.
        field.Start.ParentNode.InsertBefore(officeMath, field.Start);
        // Remove the original field from the document.
        field.Remove();
    }
}

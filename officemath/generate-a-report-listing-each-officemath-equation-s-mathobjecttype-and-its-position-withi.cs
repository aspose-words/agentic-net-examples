using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;   // Needed for OfficeMath

public class Program
{
    public static void Main()
    {
        // Create a new document and builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a title paragraph.
        builder.Writeln("Sample document with OfficeMath equations:");
        builder.Writeln();

        // Insert three simple equations using the deterministic EQ-field bootstrap workflow.
        InsertEquation(builder, @"\f(1,2)"); // fraction
        builder.Writeln(); // separate paragraphs
        InsertEquation(builder, @"\r(3,x)"); // root
        builder.Writeln();
        InsertEquation(builder, @"\s(5)"); // sigma

        // Save the sample document.
        string docPath = "SampleWithEquations.docx";
        doc.Save(docPath);

        // Reload the document to ensure a clean state.
        Document loadedDoc = new Document(docPath);

        // Prepare the report.
        StringBuilder reportBuilder = new StringBuilder();
        reportBuilder.AppendLine("OfficeMath Equation Report");
        reportBuilder.AppendLine("---------------------------");

        // Get all OfficeMath nodes.
        NodeCollection officeMathNodes = loadedDoc.GetChildNodes(NodeType.OfficeMath, true);
        int equationIndex = 0;

        // Get all paragraphs once for position calculations.
        NodeCollection allParagraphs = loadedDoc.GetChildNodes(NodeType.Paragraph, true);

        foreach (OfficeMath om in officeMathNodes)
        {
            equationIndex++;

            // MathObjectType of the equation.
            var mathObjectType = om.MathObjectType;

            // Parent paragraph.
            Paragraph parentParagraph = om.ParentParagraph;

            // Paragraph index (1‑based).
            int paragraphIndex = allParagraphs.IndexOf(parentParagraph) + 1;

            // Section index (1‑based).
            Section parentSection = parentParagraph?.ParentSection;
            int sectionIndex = loadedDoc.Sections.IndexOf(parentSection) + 1;

            reportBuilder.AppendLine(
                $"Equation {equationIndex}: MathObjectType={mathObjectType}, Section={sectionIndex}, Paragraph={paragraphIndex}");
        }

        // Write the report to a text file.
        string reportPath = "OfficeMathReport.txt";
        File.WriteAllText(reportPath, reportBuilder.ToString());

        // Validate that the report file was created.
        if (!File.Exists(reportPath))
            throw new InvalidOperationException("Report file was not created.");

        // Optionally, write a brief console message (no user interaction required).
        Console.WriteLine($"Report generated: {Path.GetFullPath(reportPath)}");
    }

    // Helper method to insert an equation using the EQ-field bootstrap workflow.
    private static void InsertEquation(DocumentBuilder builder, string eqSwitch)
    {
        // Insert an EQ field.
        Field field = builder.InsertField(FieldType.FieldEquation, true);
        if (field is not FieldEQ fieldEq)
            return; // Safety check.

        // Move to the field separator and write the EQ argument.
        builder.MoveTo(fieldEq.Separator);
        builder.Write(eqSwitch);

        // Convert the field to an OfficeMath node.
        OfficeMath officeMath = fieldEq.AsOfficeMath();

        // If conversion succeeded, replace the field with the real OfficeMath node.
        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start.
            Node fieldStart = field.Start;
            fieldStart.ParentNode.InsertBefore(officeMath, fieldStart);

            // Remove the original field from the document.
            field.Remove();
        }
    }
}

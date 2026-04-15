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

        // Insert a few deterministic equations using the EQ field bootstrap workflow.
        InsertEquation(builder, @"\f(1,2)");   // Fraction 1/2
        InsertEquation(builder, @"\r(3,x)");  // Cube root of x
        InsertEquation(builder, @"\i");       // Integral symbol

        // Save the sample document.
        string docPath = "SampleOfficeMath.docx";
        doc.Save(docPath);
        if (!File.Exists(docPath))
            throw new Exception("Failed to save the sample document.");

        // Generate a report that lists each OfficeMath equation's MathObjectType and its position.
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        StringBuilder reportBuilder = new StringBuilder();

        for (int i = 0; i < mathNodes.Count; i++)
        {
            OfficeMath om = (OfficeMath)mathNodes[i];
            string mathObjectType = om.MathObjectType.ToString();

            Paragraph para = om.ParentParagraph;
            Section sec = para?.ParentSection;

            int sectionIndex = sec != null ? doc.Sections.IndexOf(sec) : -1;
            int paragraphIndexInSection = -1;
            if (sec != null && para != null)
                paragraphIndexInSection = sec.Body.Paragraphs.IndexOf(para);

            reportBuilder.AppendLine(
                $"Equation {i + 1}: MathObjectType = {mathObjectType}, Section = {sectionIndex}, Paragraph = {paragraphIndexInSection}");
        }

        // Save the report to a text file.
        string reportPath = "OfficeMathReport.txt";
        File.WriteAllText(reportPath, reportBuilder.ToString());

        // Validate that the report was created.
        if (!File.Exists(reportPath))
            throw new Exception("Failed to create the OfficeMath report.");

        // The example finishes without requiring any user interaction.
    }

    // Helper method that inserts an EQ field, converts it to a real OfficeMath node,
    // and then removes the original field.
    private static void InsertEquation(DocumentBuilder builder, string eqArguments)
    {
        // Insert an empty EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the EQ arguments into the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(eqArguments);

        // Return the builder to the paragraph that contains the field.
        builder.MoveTo(field.Start.ParentNode);

        // Convert the field to an OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();
        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start.
            field.Start.ParentNode.InsertBefore(officeMath, field.Start);

            // Remove the original field from the document.
            field.Remove();

            // Position the builder after the newly inserted OfficeMath node.
            builder.MoveTo(officeMath);
            // Add a new paragraph to separate equations.
            builder.Writeln();
        }
    }
}

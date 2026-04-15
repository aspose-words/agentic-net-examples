using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Aspose.Words.Saving;

public class ChangeOfficeMathDisplay
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert several paragraphs each containing an inline OfficeMath equation.
        for (int i = 1; i <= 5; i++)
        {
            // Add some introductory text.
            builder.Writeln($"Paragraph {i} with an inline equation:");

            // Insert an EQ field which will be converted to a real OfficeMath node.
            FieldEQ eqField = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
            // Move to the field separator to write the EQ arguments.
            builder.MoveTo(eqField.Separator);
            // Use a simple fraction as the equation content.
            builder.Write(@"\f(1,2)");
            // Return the builder to the paragraph that contains the field.
            builder.MoveTo(eqField.Start.ParentNode);

            // Convert the EQ field to an OfficeMath object.
            OfficeMath officeMath = eqField.AsOfficeMath();
            if (officeMath != null)
            {
                // Insert the OfficeMath node before the field start node.
                eqField.Start.ParentNode.InsertBefore(officeMath, eqField.Start);
                // Remove the original EQ field from the document.
                eqField.Remove();

                // Ensure the equation is initially displayed inline.
                officeMath.DisplayType = OfficeMathDisplayType.Inline;
            }

            // Add an empty line after each equation for readability.
            builder.Writeln();
        }

        // Save the sample document (optional, demonstrates input creation).
        string inputPath = Path.Combine(Environment.CurrentDirectory, "SampleInput.docx");
        doc.Save(inputPath, SaveFormat.Docx);

        // Load the document back (simulating processing of an existing large report).
        Document loadedDoc = new Document(inputPath);

        // Enumerate all OfficeMath nodes in the document.
        NodeCollection mathNodes = loadedDoc.GetChildNodes(NodeType.OfficeMath, true);
        foreach (OfficeMath math in mathNodes)
        {
            // Target only top‑level equations (OMathPara) that are currently inline.
            if (math.MathObjectType == MathObjectType.OMathPara &&
                math.DisplayType == OfficeMathDisplayType.Inline)
            {
                // Change the display type to separate line (Display) and left‑justify.
                math.DisplayType = OfficeMathDisplayType.Display;
                math.Justification = OfficeMathJustification.Left;
            }
        }

        // Save the modified document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "ModifiedReport.docx");
        loadedDoc.Save(outputPath, SaveFormat.Docx);

        // Simple validation to ensure the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}

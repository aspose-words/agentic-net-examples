using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class ApplyOfficeMathJustification
{
    public static void Main()
    {
        // Folder for output files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert three sections, each containing two simple equations.
        for (int sectionIndex = 0; sectionIndex < 3; sectionIndex++)
        {
            // Add a heading for the section.
            builder.Writeln($"Section {sectionIndex + 1}");
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;

            // Insert two equations in the current section.
            for (int eqIndex = 0; eqIndex < 2; eqIndex++)
            {
                // Insert an EQ field.
                Field field = builder.InsertField(FieldType.FieldEquation, true);
                FieldEQ fieldEQ = (FieldEQ)field;

                // Write a simple fraction as the equation argument.
                builder.MoveTo(fieldEQ.Separator);
                builder.Write(@"\f(1,2)"); // fraction 1/2
                builder.MoveTo(fieldEQ.Start.ParentNode);

                // Convert the field to a real OfficeMath object.
                OfficeMath officeMath = fieldEQ.AsOfficeMath();
                if (officeMath != null)
                {
                    // Insert the OfficeMath node before the field start and remove the field.
                    fieldEQ.Start.ParentNode.InsertBefore(officeMath, fieldEQ.Start);
                    fieldEQ.Remove();
                }

                // Add a paragraph break after the equation.
                builder.Writeln();
            }

            // Insert a section break after each section except the last.
            if (sectionIndex < 2)
                builder.InsertBreak(BreakType.SectionBreakNewPage);
        }

        // Save the document containing the equations.
        string initialPath = Path.Combine(outputDir, "Initial.docx");
        doc.Save(initialPath, SaveFormat.Docx);

        // Apply a uniform justification to all top‑level OfficeMath equations.
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        foreach (OfficeMath om in mathNodes)
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
            {
                // Set display type before justification (required by the API).
                om.DisplayType = OfficeMathDisplayType.Display;
                om.Justification = OfficeMathJustification.Center;
            }
        }

        // Save the modified document.
        string finalPath = Path.Combine(outputDir, "Justified.docx");
        doc.Save(finalPath, SaveFormat.Docx);

        // Simple validation that the output file was created.
        if (!File.Exists(finalPath))
            throw new InvalidOperationException("The justified document was not saved correctly.");
    }
}

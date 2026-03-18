using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Load the source DOCX file.
        Document doc = new Document("Input.docx");

        // Identifier that marks the equation to be replaced (e.g., a placeholder inside the OfficeMath).
        string identifier = "EQ1";

        // New equation text that will replace the identified equation.
        string newEquation = "a+b=c";

        // Configure find/replace options to include OfficeMath nodes.
        FindReplaceOptions replaceOptions = new FindReplaceOptions();
        replaceOptions.IgnoreOfficeMath = false; // Allow replacement inside OfficeMath objects.

        // Perform the replacement throughout the document.
        doc.Range.Replace(identifier, newEquation, replaceOptions);

        // Save the modified document as DOCX.
        doc.Save("Output.docx");
    }
}

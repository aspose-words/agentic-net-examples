using Aspose.Words;
using Aspose.Words.Math;
using Aspose.Words.Loading;

// Paths to the source HTML file and the destination document.
string htmlPath = "input.html";
string outputPath = "output.docx";

// Load the HTML document. HtmlLoadOptions can be customized if needed.
HtmlLoadOptions loadOptions = new HtmlLoadOptions();
Document doc = new Document(htmlPath, loadOptions);

// Iterate through all OfficeMath objects in the document.
foreach (OfficeMath officeMath in doc.GetChildNodes(NodeType.OfficeMath, true))
{
    // OfficeMath must be in Display mode before setting justification.
    officeMath.DisplayType = OfficeMathDisplayType.Display;

    // Apply the desired justification (e.g., center the equation group).
    officeMath.Justification = OfficeMathJustification.CenterGroup;
}

// Save the modified document.
doc.Save(outputPath);

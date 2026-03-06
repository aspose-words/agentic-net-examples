using System;
using Aspose.Words;
using Aspose.Words.Settings;
using Aspose.Words;

// Load an existing RTF document.
string inputPath = @"C:\Docs\Input.rtf";
Document doc = new Document(inputPath);

// Enable automatic hyphenation for the whole document.
doc.HyphenationOptions.AutoHyphenation = true;
doc.HyphenationOptions.HyphenateCaps = true;          // Hyphenate words in all caps.
doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;    // Max consecutive hyphenated lines.
doc.HyphenationOptions.HyphenationZone = 720;        // 0.5 inch from the right margin.

// (Optional) Insert explicit optional hyphen characters into a sample paragraph.
// ControlChar.OptionalHyphenChar represents an optional hyphen point.
DocumentBuilder builder = new DocumentBuilder(doc);
builder.MoveToDocumentStart();
builder.Write("Hy" + ControlChar.OptionalHyphenChar + "phen" + ControlChar.OptionalHyphenChar + "ation");

// Save the modified document back to RTF.
string outputPath = @"C:\Docs\Output.rtf";
doc.Save(outputPath);

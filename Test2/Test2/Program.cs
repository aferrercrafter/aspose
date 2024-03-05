// See https://aka.ms/new-console-template for more information
using Aspose.Words;
using Aspose.Words.Tables;

Console.WriteLine("Starting process");

var document = new Document(@".\Docs\Doc2.docx");

var newDoc = (Document)document.Clone(false);
var builder = new DocumentBuilder(newDoc);

var table = document.FirstSection.GetChild(NodeType.Table, 0, true) as Table;

if (table is null)
{
    Console.WriteLine("No Table found");
    return;
}

NodeImporter importer = new NodeImporter(document, newDoc, ImportFormatMode.KeepSourceFormatting);

var sectionIndex = 0;
var lastSectionIndex = table.Rows.Count - 1;
foreach (Row row in table.Rows)
{
    var cellIndex = 0;
    foreach (Cell cell in row.Cells)
    {
        if (cellIndex == 0)
        {
            var destNode = importer.ImportNode(cell.FirstParagraph, true);
            newDoc.Sections[sectionIndex].Body.AppendChild(destNode);
        }
        else
        {
            var nodes = cell.GetChildNodes(NodeType.Paragraph, true);
            foreach (var node in nodes)
            {
                var destNode = importer.ImportNode(node, true);
                newDoc.Sections[sectionIndex].Body.AppendChild(destNode);
            }
        }
        cellIndex++;
    }

    if (sectionIndex < lastSectionIndex)
    {
        builder.MoveToSection(sectionIndex);
        builder.MoveTo(newDoc.Sections[sectionIndex].Body.LastChild);
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        sectionIndex++;
    }
}

newDoc.Save(@".\Docs\out.docx");

Console.WriteLine("Process ends...");
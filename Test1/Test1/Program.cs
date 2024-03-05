// See https://aka.ms/new-console-template for more information
using Aspose.Words;
using Aspose.Words.Fields;


Console.WriteLine("Starting process");

var document = new Document(@".\Docs\Doc1.docx");
var builder = new DocumentBuilder(document);

var sectionIndex = 0;
foreach (Section section in document.Sections)
{
    section.PageSetup.RestartPageNumbering = false;
    builder.MoveToSection(1);

    var footer = section.HeadersFooters[HeaderFooterType.FooterPrimary];
    var footer2 = section.HeadersFooters[HeaderFooterType.FooterFirst];
    var footer3 = section.HeadersFooters[HeaderFooterType.FooterEven];
    var hasPageNumber = ContainsPageNumber(footer);
    if (!hasPageNumber)
    {
        builder.MoveToSection(sectionIndex);
        builder.MoveTo(footer.LastChild);
        builder.InsertField(FieldType.FieldPage, true);
    }

    sectionIndex++;
}
    

document.Save(@".\Docs\out.docx");

Console.WriteLine("Process ends...");

static bool ContainsPageNumber(HeaderFooter footer)
{
    var fields = footer.GetChildNodes(NodeType.FieldStart, true);

    return fields.Any(x => ((FieldStart)x).FieldType == FieldType.FieldPage);
}

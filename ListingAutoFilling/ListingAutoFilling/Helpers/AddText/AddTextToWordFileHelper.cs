using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using ListingAutoFilling.Helpers.GetText;
using ListingAutoFilling.Helpers.GetText.Data;

namespace ListingAutoFilling.Helpers.AddText;

public class AddTextToWordFileHelper : IAddTextToWordFileHelper
{
    private readonly IGetCodeTextHelper _getCodeTextHelper = new GetCodeTextHelper();
    public void AddTextToWordFile(string folderToSearch, Body body)
    {
        foreach (var file in Directory.GetFiles(folderToSearch, "*.cs")) // TODO: передавать, какие типы файлов
        {
            FileInfo fileInfo = new FileInfo(file);
            var fileData = _getCodeTextHelper.GetFileData(folderToSearch, fileInfo.Name);
            AddTextFromFileData(body, fileData);
        }

        foreach (var directory in Directory.GetDirectories(folderToSearch))
        {
            AddTextToWordFile(directory, body);
        }
    }

    private void AddTextFromFileData(Body body, FileData fileData)
    {
        var titleParagraph = GetTitleParagraph(fileData.TitleText);
        body.Append(titleParagraph);

        var codeParagraph = GetCodeParagraph(fileData.CodeText);
        body.Append(codeParagraph);
    }

    private Paragraph GetCodeParagraph(string codeText)
    {
        Run run = new Run();
        RunProperties runProps = new RunProperties();
        FontSize fontSize = new FontSize
        {
            Val = new StringValue("20")
        };
        RunFonts runFont = new RunFonts { Ascii = "Times New Roman" };
            
        runProps.Append(fontSize);
        runProps.Append(runFont);
            
        run.Append(runProps);
            
        var text = new Text("\n" + codeText + "\n");
        run.Append(text);

        Paragraph para = new Paragraph();
        para.Append(run);
        return para;
    }

    private Paragraph GetTitleParagraph(string titleText)
    {
        Run run = new Run();
        RunProperties runProps = new RunProperties();
        Bold bold = new Bold();
        FontSize fontSize = new FontSize
        {
            Val = new StringValue("28")
        };
        RunFonts runFont = new RunFonts { Ascii = "Times New Roman" };
            
        runProps.Append(bold);
        runProps.Append(fontSize);
        runProps.Append(runFont);
            
        run.Append(runProps);
            
        var text = new Text("\n" + titleText);
        run.Append(text);

        Paragraph para = new Paragraph();
        para.Append(run);
        return para;
    }
}
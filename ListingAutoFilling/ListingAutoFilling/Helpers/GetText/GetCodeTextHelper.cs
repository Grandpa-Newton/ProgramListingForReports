using System.Text;
using ListingAutoFilling.Helpers.GetText.Data;

namespace ListingAutoFilling.Helpers.GetText;

public class GetCodeTextHelper : IGetCodeTextHelper
{
    private const string ListingTitleText = "\tЛистинг класса {0}";
    public FileData GetFileData(string path, string fileName)
    {
        return new FileData
        {
            CodeText = GetTextFromFile(path, fileName),
            TitleText = string.Format(ListingTitleText, fileName)
        };
    }

    private string GetTextFromFile(string path, string fileName)
    {
        var fullPath = Path.Combine(path, fileName);
            
        var stringBuilder = new StringBuilder();

        stringBuilder.AppendLine(File.ReadAllText(fullPath));
        stringBuilder.AppendLine();
            
        return stringBuilder.ToString();
    }
}
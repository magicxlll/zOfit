using System.Text;

namespace ZOfit.Engine.Parsers;

/// <summary>
/// Converts Vietnamese TCVN3 (ABC) character encoding commonly used in AutoCAD drawings (.VnTime, .VnArial)
/// to standard Unicode.
/// </summary>
public static class Tcvn3Converter
{
    private static readonly (string Tcvn, string Placeholder, string Unicode)[] CadPhrases = new[]
    {
        ("MÆT C¾T", "__MATCAT__", "MẶT CẮT"),
        ("THÐP TR£N", "__THEP_TREN__", "THÉP TRÊN"),
        ("THÐP D¦íI", "__THEP_DUOI__", "THÉP DƯỚI"),
        ("THÐP §AI", "__THEP_DAI__", "THÉP ĐAI"),
        ("THÐP B£N", "__THEP_BEN__", "THÉP BÊN"),
        ("THÐP", "__THEP__", "THÉP"),
        ("VÞ TRÝ", "__VITRI__", "VỊ TRÍ"),
        ("T£N", "__TEN__", "TÊN"),
        ("TR¸I", "__TRAI__", "TRÁI"),
        ("GI÷A", "__GIUA__", "GIỮA"),
        ("PH¶I", "__PHAI__", "PHẢI"),
        ("TR£N", "__TREN__", "TRÊN"),
        ("D¦íI", "__DUOI__", "DƯỚI"),
        ("§AI", "__DAI__", "ĐAI"),
        ("B£N", "__BEN__", "BÊN")
    };

    private static readonly Dictionary<char, char> Tcvn3ToUnicodeMap = new()
    {
        // Lowercase vowels
        { '\u00B8', 'á' }, { '\u00B5', 'à' }, { '\u00B6', 'ả' }, { '\u00B7', 'ã' }, { '\u00A1', 'ạ' },
        { '\u00A8', 'ă' }, { '\u00BE', 'ắ' }, { '\u00BB', 'ằ' }, { '\u00BC', 'ẳ' }, { '\u00BD', 'ẵ' }, { '\u00A2', 'ặ' },
        { '\u00A9', 'â' }, { '\u00CA', 'ấ' }, { '\u00C7', 'ầ' }, { '\u00C8', 'ẩ' }, { '\u00C9', 'ẫ' }, { '\u00A3', 'ậ' },
        { '\u00D0', 'é' }, { '\u00CC', 'è' }, { '\u00CD', 'ẻ' }, { '\u00CE', 'ẽ' }, { '\u00A4', 'ẹ' },
        { '\u00AA', 'ê' }, { '\u00D5', 'ế' }, { '\u00D1', 'ề' }, { '\u00D2', 'ể' }, { '\u00D3', 'ễ' }, { '\u00A5', 'ệ' },
        { '\u00DD', 'í' }, { '\u00D7', 'ì' }, { '\u00D8', 'ỉ' }, { '\u00DC', 'ĩ' }, { '\u00DE', 'ị' },
        { '\u00E3', 'ó' }, { '\u00DF', 'ò' }, { '\u00E1', 'ỏ' }, { '\u00E2', 'õ' }, { '\u00E4', 'ọ' },
        { '\u00AB', 'ô' }, { '\u00E9', 'ố' }, { '\u00E5', 'ồ' }, { '\u00E6', 'ổ' }, { '\u00E8', 'ỗ' }, { '\u00EA', 'ộ' },
        { '\u00AC', 'ơ' }, { '\u00EF', 'ớ' }, { '\u00EB', 'ờ' }, { '\u00EC', 'ở' }, { '\u00EE', 'ỡ' }, { '\u00F1', 'ợ' },
        { '\u00F6', 'ú' }, { '\u00F2', 'ù' }, { '\u00F3', 'ủ' }, { '\u00F4', 'ũ' }, { '\u00F7', 'ụ' },
        { '\u00AD', 'ư' }, { '\u00FC', 'ứ' }, { '\u00F8', 'ừ' }, { '\u00F9', 'ử' }, { '\u00FB', 'ữ' }, { '\u00FD', 'ự' },
        { '\u00FE', 'ỳ' }, { '\u00FF', 'ỷ' },
        { '\u00A7', 'đ' },

        // CAD Uppercase variants
        { '\u00C6', 'Ặ' }, // MÆT -> MẶT
        { '\u00A6', 'Ư' }, // D¦íI -> DƯỚI
    };

    public static string ToUnicode(string? text)
    {
        if (string.IsNullOrEmpty(text)) return string.Empty;

        // 1. First replace common full uppercase CAD phrase tokens with ASCII placeholders
        string working = text;
        foreach (var p in CadPhrases)
        {
            working = working.Replace(p.Tcvn, p.Placeholder, StringComparison.OrdinalIgnoreCase);
        }

        // 2. Map individual TCVN3 characters to Unicode
        var sb = new StringBuilder(working.Length);
        foreach (char c in working)
        {
            if (Tcvn3ToUnicodeMap.TryGetValue(c, out char uni))
            {
                sb.Append(uni);
            }
            else
            {
                sb.Append(c);
            }
        }

        // 3. Restore the full Unicode phrases from placeholders
        string result = sb.ToString();
        foreach (var p in CadPhrases)
        {
            result = result.Replace(p.Placeholder, p.Unicode, StringComparison.OrdinalIgnoreCase);
        }

        return result;
    }
}

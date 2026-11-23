using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using System.Text;

namespace OfficeAutoVersion.Core
{
    /// <summary>
    /// 极简 JSON 读写（零第三方依赖），仅覆盖本插件用到的数据类型：
    /// 对象（公开属性）、字典、数组、字符串、数字、布尔、null。
    /// </summary>
    public static class MiniJson
    {
        /// <summary>把对象序列化为 JSON 文本。</summary>
        public static string Serialize(object value)
        {
            var builder = new StringBuilder();
            WriteValue(builder, value);
            return builder.ToString();
        }

        private static void WriteValue(StringBuilder builder, object value)
        {
            if (value == null)
            {
                builder.Append("null");
                return;
            }
            if (value is string)
            {
                WriteString(builder, (string)value);
                return;
            }
            if (value is bool)
            {
                builder.Append(((bool)value) ? "true" : "false");
                return;
            }
            if (value is int || value is long || value is short || value is byte || value is uint || value is ulong || value is ushort || value is sbyte)
            {
                builder.Append(Convert.ToInt64(value, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture));
                return;
            }
            if (value is double || value is float || value is decimal)
            {
                builder.Append(Convert.ToDouble(value, CultureInfo.InvariantCulture).ToString("R", CultureInfo.InvariantCulture));
                return;
            }
            if (value is Enum)
            {
                builder.Append(Convert.ToInt64(value, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture));
                return;
            }
            if (value is Guid)
            {
                WriteString(builder, ((Guid)value).ToString());
                return;
            }

            var dictionary = value as IDictionary;
            if (dictionary != null)
            {
                builder.Append('{');
                var first = true;
                foreach (DictionaryEntry entry in dictionary)
                {
                    if (!first)
                    {
                        builder.Append(',');
                    }
                    first = false;
                    WriteString(builder, Convert.ToString(entry.Key, CultureInfo.InvariantCulture));
                    builder.Append(':');
                    WriteValue(builder, entry.Value);
                }
                builder.Append('}');
                return;
            }

            var enumerable = value as IEnumerable;
            if (enumerable != null)
            {
                builder.Append('[');
                var first = true;
                foreach (var item in enumerable)
                {
                    if (!first)
                    {
                        builder.Append(',');
                    }
                    first = false;
                    WriteValue(builder, item);
                }
                builder.Append(']');
                return;
            }

            // 普通对象 → 按公开实例属性序列化
            var type = value.GetType();
            builder.Append('{');
            var firstProperty = true;
            foreach (var property in type.GetProperties(BindingFlags.Public | BindingFlags.Instance))
            {
                if (!property.CanRead || property.GetIndexParameters().Length != 0)
                {
                    continue;
                }
                object propertyValue;
                try
                {
                    propertyValue = property.GetValue(value, null);
                }
                catch
                {
                    continue;
                }
                if (!firstProperty)
                {
                    builder.Append(',');
                }
                firstProperty = false;
                WriteString(builder, property.Name);
                builder.Append(':');
                WriteValue(builder, propertyValue);
            }
            builder.Append('}');
        }

        private static void WriteString(StringBuilder builder, string value)
        {
            builder.Append('"');
            foreach (var ch in value)
            {
                switch (ch)
                {
                    case '"': builder.Append("\\\""); break;
                    case '\\': builder.Append("\\\\"); break;
                    case '\b': builder.Append("\\b"); break;
                    case '\f': builder.Append("\\f"); break;
                    case '\n': builder.Append("\\n"); break;
                    case '\r': builder.Append("\\r"); break;
                    case '\t': builder.Append("\\t"); break;
                    default:
                        if (ch < 32)
                        {
                            builder.Append("\\u").Append(((int)ch).ToString("x4", CultureInfo.InvariantCulture));
                        }
                        else
                        {
                            builder.Append(ch);
                        }
                        break;
                }
            }
            builder.Append('"');
        }

        /// <summary>解析 JSON 文本，返回 Dictionary&lt;string, object&gt; / List&lt;object&gt; / string / bool / long / double / null。</summary>
        public static object Parse(string json)
        {
            if (string.IsNullOrWhiteSpace(json))
            {
                return null;
            }
            var index = 0;
            return ParseValue(json, ref index);
        }

        private static object ParseValue(string text, ref int index)
        {
            SkipWhitespace(text, ref index);
            if (index >= text.Length)
            {
                return null;
            }
            var ch = text[index];
            if (ch == '{')
            {
                return ParseObject(text, ref index);
            }
            if (ch == '[')
            {
                return ParseArray(text, ref index);
            }
            if (ch == '"')
            {
                return ParseString(text, ref index);
            }
            if (ch == 't')
            {
                Consume(text, ref index, "true");
                return true;
            }
            if (ch == 'f')
            {
                Consume(text, ref index, "false");
                return false;
            }
            if (ch == 'n')
            {
                Consume(text, ref index, "null");
                return null;
            }
            return ParseNumber(text, ref index);
        }

        private static Dictionary<string, object> ParseObject(string text, ref int index)
        {
            var result = new Dictionary<string, object>(StringComparer.Ordinal);
            index++; // 跳过 {
            SkipWhitespace(text, ref index);
            if (index < text.Length && text[index] == '}')
            {
                index++;
                return result;
            }
            while (index < text.Length)
            {
                SkipWhitespace(text, ref index);
                var key = ParseString(text, ref index);
                SkipWhitespace(text, ref index);
                if (index < text.Length && text[index] == ':')
                {
                    index++;
                }
                var value = ParseValue(text, ref index);
                if (key != null)
                {
                    result[key] = value;
                }
                SkipWhitespace(text, ref index);
                if (index < text.Length && text[index] == ',')
                {
                    index++;
                    continue;
                }
                if (index < text.Length && text[index] == '}')
                {
                    index++;
                }
                break;
            }
            return result;
        }

        private static List<object> ParseArray(string text, ref int index)
        {
            var result = new List<object>();
            index++; // 跳过 [
            SkipWhitespace(text, ref index);
            if (index < text.Length && text[index] == ']')
            {
                index++;
                return result;
            }
            while (index < text.Length)
            {
                result.Add(ParseValue(text, ref index));
                SkipWhitespace(text, ref index);
                if (index < text.Length && text[index] == ',')
                {
                    index++;
                    continue;
                }
                if (index < text.Length && text[index] == ']')
                {
                    index++;
                }
                break;
            }
            return result;
        }

        private static string ParseString(string text, ref int index)
        {
            SkipWhitespace(text, ref index);
            if (index >= text.Length || text[index] != '"')
            {
                throw new FormatException("JSON 字符串格式错误，位置 " + index.ToString(CultureInfo.InvariantCulture));
            }
            index++;
            var builder = new StringBuilder();
            while (index < text.Length)
            {
                var ch = text[index++];
                if (ch == '"')
                {
                    return builder.ToString();
                }
                if (ch == '\\')
                {
                    if (index >= text.Length)
                    {
                        break;
                    }
                    var escaped = text[index++];
                    switch (escaped)
                    {
                        case '"': builder.Append('"'); break;
                        case '\\': builder.Append('\\'); break;
                        case '/': builder.Append('/'); break;
                        case 'b': builder.Append('\b'); break;
                        case 'f': builder.Append('\f'); break;
                        case 'n': builder.Append('\n'); break;
                        case 'r': builder.Append('\r'); break;
                        case 't': builder.Append('\t'); break;
                        case 'u':
                            if (index + 4 <= text.Length)
                            {
                                var hex = text.Substring(index, 4);
                                index += 4;
                                int code;
                                if (int.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out code))
                                {
                                    builder.Append((char)code);
                                }
                            }
                            break;
                        default:
                            builder.Append(escaped);
                            break;
                    }
                }
                else
                {
                    builder.Append(ch);
                }
            }
            throw new FormatException("JSON 字符串未闭合。");
        }

        private static object ParseNumber(string text, ref int index)
        {
            var start = index;
            while (index < text.Length && "+-0123456789.eE".IndexOf(text[index]) >= 0)
            {
                index++;
            }
            var numberText = text.Substring(start, index - start);
            long longValue;
            if (long.TryParse(numberText, NumberStyles.Integer, CultureInfo.InvariantCulture, out longValue))
            {
                return longValue;
            }
            double doubleValue;
            if (double.TryParse(numberText, NumberStyles.Float, CultureInfo.InvariantCulture, out doubleValue))
            {
                return doubleValue;
            }
            return 0L;
        }

        private static void SkipWhitespace(string text, ref int index)
        {
            while (index < text.Length && char.IsWhiteSpace(text[index]))
            {
                index++;
            }
        }

        private static void Consume(string text, ref int index, string literal)
        {
            if (index + literal.Length > text.Length ||
                string.CompareOrdinal(text.Substring(index, literal.Length), literal) != 0)
            {
                throw new FormatException("JSON 字面量格式错误，位置 " + index.ToString(CultureInfo.InvariantCulture));
            }
            index += literal.Length;
        }

        /// <summary>把 JSON 字典转换为 VersionOptions（键名不区分大小写，缺省值保留默认）。</summary>
        public static VersionOptions ToVersionOptions(Dictionary<string, object> map)
        {
            if (map == null)
            {
                return null;
            }
            var options = new VersionOptions();
            bool flag;
            int number;
            foreach (var pair in map)
            {
                var key = (pair.Key ?? string.Empty).Trim();
                switch (key.ToLowerInvariant())
                {
                    case "enabled":
                        {
                            if (TryBoolean(pair.Value, out flag))
                            {
                                options.Enabled = flag;
                            }
                            break;
                        }
                    case "storagemode":
                        {
                            if (TryInt32(pair.Value, out number))
                            {
                                options.StorageMode = (VersionStorageMode)number;
                            }
                            break;
                        }
                    case "versionfoldername":
                        {
                            var text = AsString(pair.Value);
                            if (text != null)
                            {
                                options.VersionFolderName = text;
                            }
                            break;
                        }
                    case "centralroot":
                        {
                            var text = AsString(pair.Value);
                            if (text != null)
                            {
                                options.CentralRoot = text;
                            }
                            break;
                        }
                    case "maxversionstokeep":
                        {
                            if (TryInt32(pair.Value, out number))
                            {
                                options.MaxVersionsToKeep = number;
                            }
                            break;
                        }
                    case "skipifunchanged":
                        {
                            if (TryBoolean(pair.Value, out flag))
                            {
                                options.SkipIfUnchanged = flag;
                            }
                            break;
                        }
                    case "filenamepattern":
                        {
                            var text = AsString(pair.Value);
                            if (text != null)
                            {
                                options.FileNamePattern = text;
                            }
                            break;
                        }
                    case "logenabled":
                        {
                            if (TryBoolean(pair.Value, out flag))
                            {
                                options.LogEnabled = flag;
                            }
                            break;
                        }
                }
            }
            return options;
        }

        private static string AsString(object value)
        {
            return value as string;
        }

        private static bool TryBoolean(object value, out bool result)
        {
            result = false;
            if (value is bool)
            {
                result = (bool)value;
                return true;
            }
            var text = value as string;
            if (text != null)
            {
                var trimmed = text.Trim();
                if (trimmed.Equals("true", StringComparison.OrdinalIgnoreCase))
                {
                    result = true;
                    return true;
                }
                if (trimmed.Equals("false", StringComparison.OrdinalIgnoreCase))
                {
                    result = false;
                    return true;
                }
                return false;
            }
            if (value is long || value is double)
            {
                result = Convert.ToDouble(value, CultureInfo.InvariantCulture) != 0;
                return true;
            }
            return false;
        }

        private static bool TryInt32(object value, out int result)
        {
            result = 0;
            if (value is long)
            {
                result = (int)(long)value;
                return true;
            }
            if (value is double)
            {
                result = (int)Math.Round((double)value);
                return true;
            }
            var text = value as string;
            if (text != null)
            {
                return int.TryParse(text.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out result);
            }
            return false;
        }
    }
}

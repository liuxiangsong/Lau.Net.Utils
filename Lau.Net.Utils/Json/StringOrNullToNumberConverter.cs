using Newtonsoft.Json;
using System;

/// <summary>
/// 实体属性为Int、Decimal，json为空字符串或null时，转化为0
/// </summary>
/// <typeparam name="T"></typeparam>
public class StringOrNullToNumberConverter<T> : JsonConverter<T> where T : struct, IConvertible
{
    public override void WriteJson(JsonWriter writer, T value, JsonSerializer serializer)
    {
        writer.WriteValue(value);
    }

    public override T ReadJson(JsonReader reader, Type objectType, T existingValue, bool hasExistingValue, JsonSerializer serializer)
    {
        if (reader.TokenType == JsonToken.String)
        {
            string stringValue = (string)reader.Value;
            if (string.IsNullOrEmpty(stringValue))
            {
                return default(T); // default for value types (int, decimal) is 0
            }
            return (T)Convert.ChangeType(stringValue, typeof(T));
        }
        else if (reader.TokenType == JsonToken.Null)
        {
            return default(T); // default for value types (int, decimal) is 0
        }
        else if (reader.TokenType == JsonToken.Float || reader.TokenType == JsonToken.Integer)
        {
            return (T)Convert.ChangeType(reader.Value, typeof(T));
        }

        throw new JsonSerializationException($"Invalid value for {typeof(T).Name}");
    }
}
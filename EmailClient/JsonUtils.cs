using Newtonsoft.Json.Linq;
using Newtonsoft.Json.Schema;
using System;
using System.Activities.Validation;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace EmailClient
{
    public static class JsonUtils
    {
        public static string AssemblyDirectory => Path.GetDirectoryName(Uri.UnescapeDataString(new UriBuilder(Assembly.GetExecutingAssembly().CodeBase).Path));

        public static string StringValue(this JToken jToken, string path, string defaultValue = "") => (string)jToken.SelectToken(path) ?? defaultValue;

        public static int IntValue(this JToken jToken, string path, int defaultValue = 0)
        {
            int num = defaultValue;
            string s = (string)jToken.SelectToken(path);
            if (s != null)
                num = int.Parse(s);
            return num;
        }

        public static bool BoolValue(this JToken jToken, string path, bool defaultValue = false)
        {
            bool flag = defaultValue;
            string str = (string)jToken.SelectToken(path);
            if (str != null)
                flag = str.ToLower().Equals("true");
            return flag;
        }

        public static string StringValue(this JObject jObject, string path, string defaultValue = "") => (string)jObject.SelectToken(path) ?? defaultValue;

        public static int IntValue(this JObject jObject, string path, int defaultValue = 0)
        {
            int num = defaultValue;
            string s = (string)jObject.SelectToken(path);
            if (s != null)
                num = int.Parse(s);
            return num;
        }

        public static bool BoolValue(this JObject jObject, string path, bool defaultValue = false)
        {
            bool flag = defaultValue;
            string str = (string)jObject.SelectToken(path);
            if (str != null)
                flag = str.ToLower().Equals("true");
            return flag;
        }

        //public static bool IsValid(
        //  string jsonSettings,
        //  string schemaFileName,
        //  out ICollection<ValidationError> validationErrors)
        //{
        //    JsonSchema result = JsonSchema.FromFileAsync(JsonUtils.AssemblyDirectory + Path.DirectorySeparatorChar.ToString() + schemaFileName).GetAwaiter().GetResult();
        //    validationErrors = result.Validate(jsonSettings);
        //    return validationErrors.Count == 0;
        //}

    }
}

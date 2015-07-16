namespace IG_TableExporter
{
    // 서버쪽 테이블자동화를 위한 생성 파일명, 내용들
    // 다른 방법으로 구현할 방법이 없는 지 알아보자
    internal class MetaTable
    {
        private static string prefix = Properties.Settings.Default.MetaTable_Prefix;

        #region MiniJson.cs
        public static string MetaMiniJsonName = "MiniJson.cs";
        public static string MetaMiniJsonContent =
@"/*
 * Copyright (c) 2013 Calvin Rien
 *
 * Based on the JSON parser by Patrick van Bergen
 * http://techblog.procurios.nl/k/618/news/view/14605/14863/How-do-I-write-my-own-parser-for-JSON.html
 *
 * Simplified it so that it doesn't throw exceptions
 * and can be used in Unity iPhone with maximum code stripping.
 *
 * Permission is hereby granted, free of charge, to any person obtaining
 * a copy of this software and associated documentation files (the
 * ""Software""), to deal in the Software without restriction, including
 * without limitation the rights to use, copy, modify, merge, publish,
 * distribute, sublicense, and/or sell copies of the Software, and to
 * permit persons to whom the Software is furnished to do so, subject to
 * the following conditions:
 *
 * The above copyright notice and this permission notice shall be
 * included in all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND,
 * EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
 * MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
 * IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY
 * CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
 * TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
 * SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 */
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace IGData
{
    // Example usage:
    //
    //  using UnityEngine;
    //  using System.Collections;
    //  using System.Collections.Generic;
    //  using MiniJSON;
    //
    //  public class MiniJSONTest : MonoBehaviour {
    //      void Start () {
    //          var jsonString = ""{ \""array\"": [1.44,2,3], "" +
    //                          ""\""object\"": {\""key1\"":\""value1\"", \""key2\"":256}, "" +
    //                          ""\""string\"": \""The quick brown fox \\\""jumps\\\"" over the lazy dog \"", "" +
    //                          ""\""unicode\"": \""\\u3041 Men\u00fa sesi\u00f3n\"", "" +
    //                          ""\""int\"": 65536, "" +
    //                          ""\""float\"": 3.1415926, "" +
    //                          ""\""bool\"": true, "" +
    //                          ""\""null\"": null }"";
    //
    //          var dict = Json.Deserialize(jsonString) as Dictionary<string,object>;
    //
    //          Debug.Log(""deserialized: "" + dict.GetType());
    //          Debug.Log(""dict['array'][0]: "" + ((List<object>) dict[""array""])[0]);
    //          Debug.Log(""dict['string']: "" + (string) dict[""string""]);
    //          Debug.Log(""dict['float']: "" + (double) dict[""float""]); // floats come out as doubles
    //          Debug.Log(""dict['int']: "" + (long) dict[""int""]); // ints come out as longs
    //          Debug.Log(""dict['unicode']: "" + (string) dict[""unicode""]);
    //
    //          var str = Json.Serialize(dict);
    //
    //          Debug.Log(""serialized: "" + str);
    //      }
    //  }

    /// <summary>
    /// This class encodes and decodes JSON strings.
    /// Spec. details, see http://www.json.org/
    ///
    /// JSON uses Arrays and Objects. These correspond here to the datatypes IList and IDictionary.
    /// All numbers are parsed to doubles.
    /// </summary>
    public static class Json
    {
        /// <summary>
        /// Parses the string json into a value
        /// </summary>
        /// <param name=""json"">A JSON string.</param>
        /// <returns>An List&lt;object&gt;, a Dictionary&lt;string, object&gt;, a double, an integer,a string, null, true, or false</returns>
        public static object Deserialize(string json)
        {
            // save the string for debug information
            if (json == null)
            {
                return null;
            }

            return Parser.Parse(json);
        }

        sealed class Parser : IDisposable
        {
            const string WORD_BREAK = ""{}[],:\"""";

            public static bool IsWordBreak(char c)
            {
                return Char.IsWhiteSpace(c) || WORD_BREAK.IndexOf(c) != -1;
            }

            enum TOKEN
            {
                NONE,
                CURLY_OPEN,
                CURLY_CLOSE,
                SQUARED_OPEN,
                SQUARED_CLOSE,
                COLON,
                COMMA,
                STRING,
                NUMBER,
                TRUE,
                FALSE,
                NULL
            };

            StringReader json;

            Parser(string jsonString)
            {
                json = new StringReader(jsonString);
            }

            public static object Parse(string jsonString)
            {
                using (var instance = new Parser(jsonString))
                {
                    return instance.ParseValue();
                }
            }

            public void Dispose()
            {
                json.Dispose();
                json = null;
            }

            Dictionary<string, object> ParseObject()
            {
                Dictionary<string, object> table = new Dictionary<string, object>();

                // ditch opening brace
                json.Read();

                // {
                while (true)
                {
                    switch (NextToken)
                    {
                        case TOKEN.NONE:
                            return null;
                        case TOKEN.COMMA:
                            continue;
                        case TOKEN.CURLY_CLOSE:
                            return table;
                        default:
                            // name
                            string name = ParseString();
                            if (name == null)
                            {
                                return null;
                            }

                            // :
                            if (NextToken != TOKEN.COLON)
                            {
                                return null;
                            }
                            // ditch the colon
                            json.Read();

                            // value
                            table[name] = ParseValue();
                            break;
                    }
                }
            }

            List<object> ParseArray()
            {
                List<object> array = new List<object>();

                // ditch opening bracket
                json.Read();

                // [
                var parsing = true;
                while (parsing)
                {
                    TOKEN nextToken = NextToken;

                    switch (nextToken)
                    {
                        case TOKEN.NONE:
                            return null;
                        case TOKEN.COMMA:
                            continue;
                        case TOKEN.SQUARED_CLOSE:
                            parsing = false;
                            break;
                        default:
                            object value = ParseByToken(nextToken);

                            array.Add(value);
                            break;
                    }
                }

                return array;
            }

            object ParseValue()
            {
                TOKEN nextToken = NextToken;
                return ParseByToken(nextToken);
            }

            object ParseByToken(TOKEN token)
            {
                switch (token)
                {
                    case TOKEN.STRING:
                        return ParseString();
                    case TOKEN.NUMBER:
                        return ParseNumber();
                    case TOKEN.CURLY_OPEN:
                        return ParseObject();
                    case TOKEN.SQUARED_OPEN:
                        return ParseArray();
                    case TOKEN.TRUE:
                        return true;
                    case TOKEN.FALSE:
                        return false;
                    case TOKEN.NULL:
                        return null;
                    default:
                        return null;
                }
            }

            string ParseString()
            {
                StringBuilder s = new StringBuilder();
                char c;

                // ditch opening quote
                json.Read();

                bool parsing = true;
                while (parsing)
                {

                    if (json.Peek() == -1)
                    {
                        parsing = false;
                        break;
                    }

                    c = NextChar;
                    switch (c)
                    {
                        case '""':
                            parsing = false;
                            break;
                        case '\\':
                            if (json.Peek() == -1)
                            {
                                parsing = false;
                                break;
                            }

                            c = NextChar;
                            switch (c)
                            {
                                case '""':
                                case '\\':
                                case '/':
                                    s.Append(c);
                                    break;
                                case 'b':
                                    s.Append('\b');
                                    break;
                                case 'f':
                                    s.Append('\f');
                                    break;
                                case 'n':
                                    s.Append('\n');
                                    break;
                                case 'r':
                                    s.Append('\r');
                                    break;
                                case 't':
                                    s.Append('\t');
                                    break;
                                case 'u':
                                    var hex = new char[4];

                                    for (int i = 0; i < 4; i++)
                                    {
                                        hex[i] = NextChar;
                                    }

                                    s.Append((char)Convert.ToInt32(new string(hex), 16));
                                    break;
                            }
                            break;
                        default:
                            s.Append(c);
                            break;
                    }
                }

                return s.ToString();
            }

            object ParseNumber()
            {
                string number = NextWord;

                if (number.IndexOf('.') == -1)
                {
                    long parsedInt;
                    Int64.TryParse(number, out parsedInt);
                    return parsedInt;
                }

                double parsedDouble;
                Double.TryParse(number, out parsedDouble);
                return parsedDouble;
            }

            void EatWhitespace()
            {
                while (Char.IsWhiteSpace(PeekChar))
                {
                    json.Read();

                    if (json.Peek() == -1)
                    {
                        break;
                    }
                }
            }

            char PeekChar
            {
                get
                {
                    return Convert.ToChar(json.Peek());
                }
            }

            char NextChar
            {
                get
                {
                    return Convert.ToChar(json.Read());
                }
            }

            string NextWord
            {
                get
                {
                    StringBuilder word = new StringBuilder();

                    while (!IsWordBreak(PeekChar))
                    {
                        word.Append(NextChar);

                        if (json.Peek() == -1)
                        {
                            break;
                        }
                    }

                    return word.ToString();
                }
            }

            TOKEN NextToken
            {
                get
                {
                    EatWhitespace();

                    if (json.Peek() == -1)
                    {
                        return TOKEN.NONE;
                    }

                    switch (PeekChar)
                    {
                        case '{':
                            return TOKEN.CURLY_OPEN;
                        case '}':
                            json.Read();
                            return TOKEN.CURLY_CLOSE;
                        case '[':
                            return TOKEN.SQUARED_OPEN;
                        case ']':
                            json.Read();
                            return TOKEN.SQUARED_CLOSE;
                        case ',':
                            json.Read();
                            return TOKEN.COMMA;
                        case '""':
                            return TOKEN.STRING;
                        case ':':
                            return TOKEN.COLON;
                        case '0':
                        case '1':
                        case '2':
                        case '3':
                        case '4':
                        case '5':
                        case '6':
                        case '7':
                        case '8':
                        case '9':
                        case '-':
                            return TOKEN.NUMBER;
                    }

                    switch (NextWord)
                    {
                        case ""false"":
                            return TOKEN.FALSE;
                        case ""true"":
                            return TOKEN.TRUE;
                        case ""null"":
                            return TOKEN.NULL;
                    }

                    return TOKEN.NONE;
                }
            }
        }

        /// <summary>
        /// Converts a IDictionary / IList object or a simple type (string, int, etc.) into a JSON string
        /// </summary>
        /// <param name=""json"">A Dictionary&lt;string, object&gt; / List&lt;object&gt;</param>
        /// <returns>A JSON encoded string, or null if object 'json' is not serializable</returns>
        public static string Serialize(object obj)
        {
            return Serializer.Serialize(obj);
        }

        sealed class Serializer
        {
            StringBuilder builder;

            Serializer()
            {
                builder = new StringBuilder();
            }

            public static string Serialize(object obj)
            {
                var instance = new Serializer();

                instance.SerializeValue(obj);

                return instance.builder.ToString();
            }

            void SerializeValue(object value)
            {
                IList asList;
                IDictionary asDict;
                string asStr;

                if (value == null)
                {
                    builder.Append(""null"");
                }
                else if ((asStr = value as string) != null)
                {
                    SerializeString(asStr);
                }
                else if (value is bool)
                {
                    builder.Append((bool)value ? ""true"" : ""false"");
                }
                else if ((asList = value as IList) != null)
                {
                    SerializeArray(asList);
                }
                else if ((asDict = value as IDictionary) != null)
                {
                    SerializeObject(asDict);
                }
                else if (value is char)
                {
                    SerializeString(new string((char)value, 1));
                }
                else
                {
                    SerializeOther(value);
                }
            }

            void SerializeObject(IDictionary obj)
            {
                bool first = true;

                builder.Append('{');

                foreach (object e in obj.Keys)
                {
                    if (!first)
                    {
                        builder.Append(',');
                    }

                    SerializeString(e.ToString());
                    builder.Append(':');

                    SerializeValue(obj[e]);

                    first = false;
                }

                builder.Append('}');
            }

            void SerializeArray(IList anArray)
            {
                builder.Append('[');

                bool first = true;

                foreach (object obj in anArray)
                {
                    if (!first)
                    {
                        builder.Append(',');
                    }

                    SerializeValue(obj);

                    first = false;
                }

                builder.Append(']');
            }

            void SerializeString(string str)
            {
                builder.Append('\""');

                char[] charArray = str.ToCharArray();
                foreach (var c in charArray)
                {
                    switch (c)
                    {
                        case '""':
                            builder.Append(""\\\"""");
                            break;
                        case '\\':
                            builder.Append(""\\\\"");
                            break;
                        case '\b':
                            builder.Append(""\\b"");
                            break;
                        case '\f':
                            builder.Append(""\\f"");
                            break;
                        case '\n':
                            builder.Append(""\\n"");
                            break;
                        case '\r':
                            builder.Append(""\\r"");
                            break;
                        case '\t':
                            builder.Append(""\\t"");
                            break;
                        default:
                            int codepoint = Convert.ToInt32(c);
                            if ((codepoint >= 32) && (codepoint <= 126))
                            {
                                builder.Append(c);
                            }
                            else
                            {
                                builder.Append(""\\u"");
                                builder.Append(codepoint.ToString(""x4""));
                            }
                            break;
                    }
                }

                builder.Append('\""');
            }

            void SerializeOther(object value)
            {
                // NOTE: decimals lose precision during serialization.
                // They always have, I'm just letting you know.
                // Previously floats and doubles lost precision too.
                if (value is float)
                {
                    builder.Append(((float)value).ToString(""R""));
                }
                else if (value is int
                  || value is uint
                  || value is long
                  || value is sbyte
                  || value is byte
                  || value is short
                  || value is ushort
                  || value is ulong)
                {
                    builder.Append(value);
                }
                else if (value is double
                  || value is decimal)
                {
                    builder.Append(Convert.ToDouble(value).ToString(""R""));
                }
                else
                {
                    SerializeString(value.ToString());
                }
            }
        }
    }
}
";
        #endregion

        #region AbKVType.cs
        public static string MetaAbKVTypeName = "AbKVType.cs";
        public static string MetaAbKVTypeContent =
@"using System;
using System.Collections.Generic;

namespace " + prefix + @"Data
{
    public abstract class AbKVType<K, V, TClass> 
        where TClass : AbKVType<K, V, TClass>
    {
        protected static string typeName = null;

        protected K key;
        protected V val;

        protected string fieldName;
        protected static SortedList<K, TClass> members = new SortedList<K, TClass>();

        protected AbKVType() { }
        protected AbKVType(string fieldName, K k, V v)
        { this.fieldName = fieldName; key = k; val = v; members.Add(k, this as TClass); }

        public string TypeName { get { return typeName; } }
        public K Key { get { return key; } }
        public V Value { get { return val; } }
        public string FieldName { get { return fieldName; } }

        public static TClass FindByKName(string kName)
        {
            TClass result = null;
            if ( typeof(V) == typeof(string))
            {
                foreach (AbKVType<K, V, TClass> m in members.Values)
                {
                    if (m.Value as string == kName)
                        result = m as TClass;
                }
            }
            return result;
        }

        public static IList<TClass> Types { get { return members.Values; } }
        public static int GetCount() { return members.Count; }
    }
}";
        #endregion

        #region Object.cs
        public static string MetaObjectName = prefix + "Object.cs";
        public static string MetaObjectContent =
@"using System;
using System.Collections.Generic;

namespace " + prefix + @"Data
{

    public interface " + prefix + @"Object
    {
        int GetIndex();

        void Map(Dictionary<string, object> dic);
    }

    public class " + prefix + @"Container<T> where T : " + prefix + @"Object
    {
        private SortedList<int, T> map = new SortedList<int, T>();

        public IList<T> GetList()
        {
            return map.Values;
        }

        public IDictionary<int, T> GetMap()
        {
            return map;
        }

        public void Add(int index, T t)
        {
            map.Add(index, t);
        }

        public T Get(int index)
        {
            T ret = default(T);
            try
            {
                ret = map[index];
            }
            catch (Exception)
            {

            }
            return ret;
        }

        public void Clear()
        {
	        map.Clear();
        }
    }
}";
        #endregion

        #region Util.cs
        public static string MetaUtilName = prefix + "Util.cs";
        public static string MetaUtilContent =
@"using LitJson;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;

namespace " + prefix + @"Data
{
    class " + prefix + @"Util
    {

        // 텍스트 파일 => StringBuilder
        public static StringBuilder ReadTextFile(string fileName)
        {
            StringBuilder sb = new StringBuilder();
            {
                using (StreamReader sr = File.OpenText(fileName))
                {
                    string line = null;
                    while ((line = sr.ReadLine()) != null)
                    {
                        sb.AppendLine(line);
                    }
                }
            }
            return sb;
        }

        // 리플렉션을 이용해서 객체의 필드의 값을 표시함
        // 리플렉션을 사용하므로 디버그 모드에서만 사용할 것
        public static string ToStringByReflection(object obj)
        {
            StringBuilder sb = new StringBuilder();
            Type t = obj.GetType();
            foreach (FieldInfo f in t.GetFields())
            {
                sb.Append(String.Format(""\t{0}:{1}\n"", f.Name, f.GetValue(obj)));
            }
            //foreach (PropertyInfo f in t.GetProperties())
            //{
            //    sb.Append(String.Format(""\t{0}:{1}\n"", f.Name, f.GetValue(obj)));
            //}
            return sb.ToString();
        }

        public static string ToStringByJson(object obj)
        {
            return JsonMapper.ToJson(obj);
        }

#if USE_LITJSON
        public static List<T> LoadJson2List<T>(string dataArrayJson) where T : class, " + prefix + @"Object
        {
            List<T> list = new List<T>();
            JsonData dataArrayJd = JsonMapper.ToObject(dataArrayJson);
            for (int i = 0; i < dataArrayJd.Count; ++i)
            {
                T t = JsonMapper.ToObject<T>(dataArrayJd[i].ToJson());
                list.Add(t);
                //Console.WriteLine(t.ToString());
            }
            return list;
        }

		public static " + prefix + @"Container<T> LoadJson2" + prefix + @"Container<T>(string dataArrayJson, string className) where T : class, " + prefix + @"Object
        {
            " + prefix + @"Container<T> container = new " + prefix + @"Container<T>();

            JsonData dataArrayJd = JsonMapper.ToObject(dataArrayJson);
            
            for (int i = 0; i < dataArrayJd.Count; ++i)
            {
                T t = JsonMapper.ToObject<T>(dataArrayJd[i].ToJson());
                container.Add(t.GetIndex(), t);
                //Console.WriteLine(t.ToString());
            }

            return container;
        }

        public static Dictionary<int, T> LoadJson2Dic<T>(string dataArrayJson) where T : class, " + prefix + @"Object
        {
            Dictionary<int, T> dic = new Dictionary<int, T>();
            JsonData dataArrayJd = JsonMapper.ToObject(dataArrayJson);
            for (int i = 0; i < dataArrayJd.Count; ++i)
            {
                T t = JsonMapper.ToObject<T>(dataArrayJd[i].ToJson());
                dic.Add(t.GetIndex(), t);
                //Console.WriteLine(t.ToString());
            }
            return dic;
        }
#endif//USE_LITJSON
        public static IGContainer<T> LoadMiniJson2GDContainer<T>(Dictionary<string, object> jsonDic, string className) where T : class, IGObject, new()
        {
            IGContainer<T> container = new IGContainer<T>();
            var dicList = jsonDic[className] as List<object>;
            foreach (var d in dicList)
            {
                var dt = d as Dictionary<string, object>;
                T t = new T();
                t.Map(dt);

                container.Add(t.GetIndex(), t);
            }
            return container;
        }
    }

    class " + prefix + @"ManagerConfig
    {
        public string SchemeJsonFileName = ""schemes.json"";
        public string TypeListFileName = ""type_list.json"";
        public string DataFilePath = ""Data"";

        override public string ToString()
        {
            return " + prefix + @"Util.ToStringByReflection(this);
        }
    }


}
";
        #endregion

        #region Table.cs
        public static string[] AdditionalMethods = { "CharacterLevel", "PetLevel" };
        public static string MetaTableName = ".cs";
        public static string[] MetaTableContent = {
@"using System;
using System.Collections.Generic;

namespace " + prefix + @"Data 
{
",
@"
		public int GetIndex()
		{
			return Index;
		}

		public override string ToString()
		{
			return IGUtil.ToStringByReflection(this);
		}",
@"
		public int GetStatByType(IGTypeSTAT statType)
		{
			if(statType == IGTypeSTAT.HP) return HP;
			if(statType == IGTypeSTAT.MP) return MP;
			if(statType == IGTypeSTAT.PATK) return PAtk;
			if(statType == IGTypeSTAT.MATK) return MAtk;
			if(statType == IGTypeSTAT.PDEF) return PDef;
			if(statType == IGTypeSTAT.CRIT) return Critical;
			if(statType == IGTypeSTAT.DODGE) return Dodge;
			
			return 0;
		}
",
@"	}
}"
        };
        #endregion

        #region Type.cs
        public static string MetaTypeName = ".cs";
        public static string[] MetaTypeContent = {
@"using System;
using System.Collections.Generic;

namespace " + prefix + @"Data 
{",
@"			try
			{
				result = members[key];
			}
			catch (Exception ) { }
			return result;
		}
	}
}"};
        #endregion

        #region Manager.cs
        public static string MetaManagerName = prefix + @"Manager.cs";
        public static string[] MetaManagerContent = {
@"using LitJson;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using ICSharpCode.SharpZipLib.Zip;

namespace " + prefix + @"Data
{
    /// <summary>
    /// 게임의 데이터를 관리하기 위한 매니저
    /// </summary>
    public class " + prefix + @"Manager
    {
        /// <summary>
        /// 싱글톤 패턴
        /// </summary>
        private static " + prefix + @"Manager inst = new " + prefix + @"Manager();

        /// <summary>
        /// 싱글톤의 인스턴스를 참조
        /// </summary>
        public static " + prefix + @"Manager Inst { get { return inst; } }
				

#region Fields

        private List<Thread> m_listThread = new List<Thread>(); // 쓰레드 리스트

        private volatile int m_nCurrLoadComplete = 0;       // 현재 완료된 테이블 수
        private int m_nTotalLoadCount = 0;                  // 총 매핑해야할 테이블 수
        private bool m_bLoadComplete = false;               // 로드 완료했는지

#endregion

#region Properties

        public int TotalTableCount
        {
            get { return m_nTotalLoadCount; }
            set { m_nTotalLoadCount = value; }
        }

        public bool IsLoadComplete 
        { 
            get { return m_bLoadComplete; }
            set { m_bLoadComplete = value; }
        }

#endregion

        public delegate void InvokeLog(string szLog);
        public InvokeLog Log;
",
@"        /// <summary>
        /// json 파일 형식의 데이터를 List 형식으로 메모리에 적재
        /// </summary>
        /// <param name=""dataFilePath"">데이터 json 파일이 있는 폴더의 절대 경로</param>
        public void LoadByPath(string dataFilePath)
        {
            string[] jsonFileNames = Directory.GetFiles(dataFilePath, ""*.json"");
            string[] dataStrings = new string[jsonFileNames.Length];
            for (int i = 0; i < jsonFileNames.Length; ++i )
            {
                StringBuilder sb = " + prefix + @"Util.ReadTextFile(jsonFileNames[i]);
                dataStrings[i] = sb.ToString();
            }
            Load(dataStrings, false);
        }

		/// <summary>
        /// 압축 해제
        /// </summary>
        /// <param name=""szZipFileName""></param>
        /// <param name=""szExtractDir""></param>
        public bool Extract(string szZipFileName, string szExtractDir, string szPassword)
        {
            //===================================================================
            // FastZip 방식 - 파일로 압축해제
            //===================================================================
            FastZipEvents pkFastZipEvent = new FastZipEvents();
            FastZip pkFastZip = new FastZip(pkFastZipEvent);
            pkFastZip.Password = szPassword;
            
            pkFastZip.ExtractZip(szZipFileName, szExtractDir, null);

            szZipFileName = szZipFileName.Replace("".fz"", "".txt"");
            
            //===================================================================
            // 암호체크
            //===================================================================
            using (FileStream stream = File.Open(szZipFileName, FileMode.Open))
            {
               // Log(""파일 : "" + szZipFileName + "", 사이즈 : "" + stream.Length);
                if (stream.Length > 0)
                    return true;
            }
            
            return false;
        }

        public void Load(string[] dataStrings, bool bUseThread)
        {
            IGTypeInit();

            m_bLoadComplete = false;
            m_listThread.Clear();
            m_nTotalLoadCount = dataStrings.Length;

            foreach (string dataString in dataStrings)
            {                
                //===================================================================
                // Use Thread
                //===================================================================
                if (bUseThread)
                {
                    Thread thread = new Thread(new ParameterizedThreadStart(Mapping));
                    thread.Start((object)dataString);
                    m_listThread.Add(thread);
                }
                else
                {
                    Mapping((object)dataString);
                }
            }
        }",

@"        public void Mapping(object data)
        {
            string szData = (string)data;

#if USE_LITJSON
            JsonData jd = JsonMapper.ToObject(szData);

            string className = """";
            foreach (string key in jd.Keys)
            {
                className = key; break;
            }
            //Log(""Class : "" + className + ""  Start"");

            if (string.IsNullOrEmpty(className))
                return;
            string js = jd[className].ToJson();
#else
            Dictionary<string, object> jsonDic = Json.Deserialize(szData) as Dictionary<string, object>;
            string[] classNames = {""""};
            jsonDic.Keys.CopyTo(classNames, 0);
            string className = classNames[0];
#endif

            // 속성에 데이터 로딩 : gdHeroCards = GDUtil.LoadJson2GDContainer<GDHeroCard>(js);",

@"            //Log(""Class : "" + className + ""  Load Complete"");

            m_nCurrLoadComplete++;
            LoadCheck();

        }

        private void LoadCheck()
        {
            if (m_nTotalLoadCount <= m_nCurrLoadComplete)
            {
                m_bLoadComplete = true;
            }
        }

        public void ThreadAbort()
        {
            foreach (Thread thread in m_listThread)
            {
                if (thread != null)
                {
                    thread.Abort();
                }
            }
        }

        /// <summary>
        /// 각 리스트의 필드와 프로퍼티의 값을 문자열로 반환함
        /// </summary>
        /// <typeparam name=""T""></typeparam>
        /// <param name=""listObj""></param>
        /// <returns>각 리스트의 필드와 프로퍼티의 값 문자열</returns>
        public string ToListString<T>(IList<T> listObj)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine(""===  "" + typeof(T).Name);
            foreach (T t in listObj)
            {
                sb.AppendLine(t.ToString());
            }
            sb.AppendLine(""==========================="");
            return sb.ToString();
        }

        public static void LoadExample()
        {
            //string dataFilePath = Directory.GetCurrentDirectory() + @""/Data"";

            //" + prefix + @"Manager.Inst.Load(dataFilePath);
            //string values = " + prefix + @"Manager.Inst.ToListString<" + prefix + @"PCCard>(" + prefix + @"Manager.Inst.PCCard.GetList());
        }
    }
}"
        };
        #endregion
    }
}

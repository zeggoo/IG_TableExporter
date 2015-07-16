using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Newtonsoft.Json;

namespace IG_TableExporter
{
    // 테이블 구조 정의
    // 나중에 json.net의 serialize 적용해보든지 하자
    public class IG_Table : System.Collections.SortedList
    {
        private string name;
        private bool uniqueKey;

        private StringBuilder sb;
        private StringWriter sw;
        private JsonTextWriter json;

        public enum DataType
        {
            INTEGER,
            FLOAT,
            BOOL,
            TEXT
        }

        public string Name
        {
            get
            {
                return this.name.Replace(Properties.Settings.Default.TablePostfix, "");
            }
            set
            {
                this.name = value;
            }
        }

        public bool UniqueKey
        {
            get
            {
                return this.uniqueKey;
            }
            set
            {
                this.uniqueKey = value;
            }
        }

        public IG_Table(string name)
        {
            Name = name;
            uniqueKey = true;

            sb = new StringBuilder();
            sw = new StringWriter(sb);
            json = new JsonTextWriter(sw);            
            json.Formatting = Formatting.Indented;
            
            json.WriteStartObject();
            json.WritePropertyName(Properties.Settings.Default.MetaTable_Prefix + name.Replace(Properties.Settings.Default.TablePostfix, String.Empty));
            json.WriteStartArray();
        }

        public void StartAdd()
        {
            json.WriteStartObject();
        }

        public void StartAdd(int key)
        {            
            if (base.ContainsKey(key))
                throw new Exception(String.Format("{0} 인덱스가 중복됩니다.", key));
            base.Add(key, null);

            StartAdd();
        }

        public void EndAdd()
        {
            json.WriteEndObject();
        }

        public void AddElement(string key, string value, string dataType, Dictionary<string, Dictionary<string, string>> subgroups)
        {
            string name = Name.Trim();

            json.WritePropertyName(key);
            //json.WriteValue(value);
            //json.WriteRawValue(value);

            // 데이터 타입에 따라 출력방식을 다르게 함
            switch (dataType.ToUpper())
            {
                case "INTEGER":                
                case "BOOL":
                    json.WriteRawValue(value);
                    break;
                case "TEXT":
                    json.WriteValue(value);
                    break;
                case "FLOAT":
                    json.WriteRawValue(value);
                    //json.WriteComment(String.Format("{0:P1}", Convert.ToDouble(value) / Properties.Settings.Default.PermilFactor));
                    break;
                default:
                    // SUBGROUP이 존재함
                    json.WriteRawValue(subgroups[dataType.ToUpper()][value].Split(Properties.Settings.Default.SubgroupSeperator)[0]);
                    //json.WriteComment(String.Format("{0}: {1}", value, subgroups[dataType.ToUpper()][value].Split(Properties.Settings.Default.SubgroupSeperator)[1]));
                    break;
            }
        }

        public override string ToString()
        {
            json.WriteEndArray();
            json.WriteEndObject();
            return sb.ToString();
        }
    }
}

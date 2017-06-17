using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Newtonsoft.Json;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace IG_TableExporter
{
    // 테이블 구조 정의
    // 나중에 json.net의 serialize 적용해보든지 하자
    public class IG_Table : System.Collections.SortedList
    {
        private string name;
        private bool hasUniqueKey;

        private int wsRow, wsCol;

        private Excel.Application xl;
        private Excel.Workbook wb;
        private Excel.Worksheet ws;

        private StringBuilder sb;
        private StringWriter sw;
        private JsonTextWriter json;

        public enum DataType
        {
            UNIQUEKEY,
            KEY,
            BYTE,
            USHORT,
            UINT,
            SHORT,
            INT,
            LONG,
            FLOAT,
            FLOAT_1K,
            FLOAT_10K,
            FLOAT_1M,
            BOOL,
            TEXT,
            ARRAY
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
                return this.hasUniqueKey;
            }
            set
            {
                this.hasUniqueKey = value;
            }
        }

        public IG_Table(string name) : this(name, true) { }

        public IG_Table(string name, bool hasUniqueKey = true)
        {
            Name = name;
            UniqueKey = hasUniqueKey;

            sb = new StringBuilder();
            sw = new StringWriter(sb);
            json = new JsonTextWriter(sw);
            json.Formatting = Formatting.Indented;

            // 여기서 엑셀시트 하나 열어주자            

            //json.WriteStartObject();
            //json.WritePropertyName(Properties.Settings.Default.MetaTable_Prefix + name.Replace(Properties.Settings.Default.TablePostfix, String.Empty));
            json.WriteStartArray();
        }

        public void StartAdd()
        {
            json.WriteStartObject();
            wsRow++;
            wsCol = 1;
        }

        public void StartAdd(int key)
        {            
            if (hasUniqueKey && base.ContainsKey(key))
                throw new Exception(String.Format("{0} 인덱스가 중복됩니다.", key));
            if (hasUniqueKey)
                base.Add(key, null);
            else
                base.Add(this.Count + 1, null);
            StartAdd();
        }

        public void EndAdd()
        {
            json.WriteEndObject();
        }

        // 엑셀시트 열기
        public void CreateMetaTable()
        {
            xl = new Excel.Application();            
            wb = xl.Workbooks.Add();
            ws = wb.ActiveSheet;
            ws.Name = Properties.Settings.Default.XLSX_SHEET;
            wsRow = wsCol = 0;
        }

        public void AddMetaTableInfos(Dictionary<string, string> define, Dictionary<string, string>dataType, Dictionary<string, string>desc, Dictionary<string, string>descCHN)
        {
            try
            {
                // 1행 2행 데이터
                int cnt = 0;
                foreach (var k in define.Keys)
                {
                    cnt++;
                    ws.Cells[1, cnt] = k;
                    ws.Cells[2, cnt] = GetMetaTableDataType(dataType[k]);                    
                }
                // 3행 데이터
                cnt = 0;
                foreach (var v in desc.Values)
                {
                    cnt++;
                    ws.Cells[4, cnt] = v;
                }
                // 4행 데이터
                cnt = 0;
                foreach (var v in descCHN.Values)
                {
                    cnt++;
                    ws.Cells[3, cnt] = v;
                }
                wsRow = 4;
            }
            catch(Exception e)
            {
                throw new Exception("서버 xlsx 파일 추출 중 에러발생");
            }
        }

        public void AddMetaTableEntry(string value)
        {
            ws.Cells[wsRow, wsCol++] = value;
        }

        public void CloseMetaTable()
        {
            var ServerExcelPath = Globals.IG_PlanAddIn.Application.ActiveWorkbook.Path + Path.DirectorySeparatorChar + Properties.Settings.Default.XLSX_PATH;
            
            // 저장 경고 무시
            xl.DisplayAlerts = false;

            if (!Directory.Exists(ServerExcelPath))
                Directory.CreateDirectory(ServerExcelPath);

            wb.SaveAs(Globals.IG_PlanAddIn.Application.ActiveWorkbook.Path + Path.DirectorySeparatorChar + Properties.Settings.Default.XLSX_PATH + this.name,
                AccessMode: Excel.XlSaveAsAccessMode.xlExclusive);

            xl.DisplayAlerts = true;

            wb.Close();
            xl.Quit();

            ReleaseMetaTable();
        }

        public bool ExistsMetaTable()
        {
            return ws != null;
        }

        public void ReleaseMetaTable()
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xl);
        }

        // DataType을 참조하여 실제 데이터타입명을 출력
        private string GetMetaTableDataType(string dataType)
        {
            switch (dataType.ToUpper())
            {
                case "UNIQUEKEY":
                case "KEY":
                case "FLOAT_1K":
                case "FLOAT_10K":
                case "FLOAT_1M":
                case "UINT":
                    return "unsigned int";
                case "BOOL":
                case "BYTE":
                    return "byte";
                case "USHORT":
                    return "unsigned short";
                case "SHORT":
                    return "short";
                case "INTEGER":
                case "INT":
                    return "int";
                case "LONG":
                    return "long";
                case "TEXT":
                    return "string";
                case "FLOAT":
                    return "float";
                case "ARRAY":
                    return "jsonobject";
                default:
                    // subgroup일 경우, byte
                    return "byte";
            }
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
                // UniqueKEY,KEY,BYTE,USHORT,UINT,SHORT,INT,FLOAT_1K,FLOAT_10K,FLOAT_1M,BOOL,TEXT,ARRAY
                case "UNIQUEKEY":
                case "KEY":
                case "BYTE":
                case "USHORT":
                case "UINT":
                case "SHORT":
                case "INT":       
                case "LONG":
                case "INTEGER":                
                case "BOOL":
                    json.WriteRawValue(value);
                    AddMetaTableEntry(value);
                    break;
                case "TEXT":
                    json.WriteValue(value);
                    AddMetaTableEntry(value);
                    break;
                case "FLOAT_1K":
                case "FLOAT_10K":
                case "FLOAT_1M":
                case "FLOAT":
                    json.WriteRawValue(value);
                    AddMetaTableEntry(value);
                    //json.WriteComment(String.Format("{0:P1}", Convert.ToDouble(value) / Properties.Settings.Default.PermilFactor));
                    break;
                case "ARRAY":
                    // ARRAY 검증
                    json.WriteRawValue(value);
                    AddMetaTableEntry(value);
                    break;
                default:
                    // SUBGROUP이 존재함
                    json.WriteRawValue(subgroups[dataType.ToUpper()][value].Split(Properties.Settings.Default.SubgroupSeperator)[0]);
                    AddMetaTableEntry(subgroups[dataType.ToUpper()][value].Split(Properties.Settings.Default.SubgroupSeperator)[0]);
                    //json.WriteComment(String.Format("{0}: {1}", value, subgroups[dataType.ToUpper()][value].Split(Properties.Settings.Default.SubgroupSeperator)[1]));
                    break;
            }
        }

        public override string ToString()
        {
            json.WriteEndArray();
            //json.WriteEndObject();
            return sb.ToString();
        }
    }
}


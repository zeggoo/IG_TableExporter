using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Drawing;
using System.Diagnostics;
using System.Collections;
using System.Reflection;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Newtonsoft.Json;

namespace IG_TableExporter
{
    public struct MonsterInfo
    {
        // 몬스터ID	타입	스프라이트	경험치	포인트	골드 HP 공격력
        public int index;
        public string stage;
        public string type;
        public string sprite;
        public int exp;
        public int point;
        public int goldMin;
        public int goldMax;
        public int HP;
        public int atk;
        public double speed;
        public double scale;
        public Color color;

        public int GetGold()
        {
            return (int)((goldMin + goldMax) / 2);
        }
    }

    // 테이블 개편으로 몬스터 타입 enum값이 필요해짐 - 개편 시, 수작업 필요
    public enum MonsterType
    {
        Monster = 1,
        Veteran,
        Boss,
        Object,
        Wall,
        Chest,        
        Angel,
        Demon,
        Gold,
        EndBox,
        Warning,
        Sanctuary,
        BigChest,
        BigGold,
        TrapChest,
        BombTop,
        BombMid,
        BombBottom,
        Banner
    }

    public struct MonsterInfoPathVerification
    {
        public bool monsterTable;
        public bool resourcePathTable;
        public bool monsterSprite;
        public bool stageTable;
    }
    
    public partial class IG_PlanAddIn
    {
        private string[] branchList;
        private Dictionary<string, Dictionary<string, string>> branchDefines;
        private Dictionary<string, Dictionary<string, string>> branchAliases;
        private Dictionary<string, Dictionary<string, string>> branchDataTypes;
        private Dictionary<string, Dictionary<string, string>> branchDataDescriptions;
        private Dictionary<string, Dictionary<string, string>> branchDataDescriptionsCHN;
        
        // min/max 처리용
        private Dictionary<string, Dictionary<string, string>> branchMins;
        private Dictionary<string, Dictionary<string, string>> branchMaxes;

        private string[] tableInfo;
        private Dictionary<string, string> monsterSpritePaths;

        public MonsterInfoPathVerification mInfoPath;

        public string[] BranchList
        {
            get
            {
                if (this.branchList == null) this.branchList = GetBranchList();
                return branchList;
                //return GetBranchList();
            }
        }

        public Dictionary<string, Dictionary<string, string>> BranchDefines
        {
            get
            {
                if (this.branchDefines == null) this.branchDefines = GetBranchDefines();
                return branchDefines;
                //return GetBranchDefines();
            }
        }

        public Dictionary<string, Dictionary<string, string>> BranchAliases
        {
            get
            {
                if (this.branchAliases == null) this.branchAliases = GetBranchAliases();
                return branchAliases;
                //return GetBranchAliases();
            }
        }

        public Dictionary<string, Dictionary<string, string>> BranchDataTypes
        {
            get
            {
                if (this.branchDataTypes == null) this.branchDataTypes = GetBranchDataTypes();
                return branchDataTypes;
                //return GetBranchDataTypes();
            }
        }

        public Dictionary<string, Dictionary<string, string>> BranchDataDescriptions
        {
            get
            {
                if (this.branchDataDescriptions == null) this.branchDataDescriptions = GetBranchDataDescriptions();                
                return branchDataDescriptions;
            }
        }
        public Dictionary<string, Dictionary<string, string>> BranchDataDescriptionsCHN
        {
            get
            {
                //if (this.branchDataDescriptionsCHN == null) this.branchDataDescriptionsCHN = GetBranchDataDescriptionsCHN();
                if (this.branchDataDescriptionsCHN == null) this.branchDataDescriptionsCHN = GetBranchDataDescriptions("(CHN)");
                return branchDataDescriptionsCHN;
            }
        }

        public Dictionary<string, string> MonsterSpritePaths
        {
            get
            {
                if (this.monsterSpritePaths == null) this.monsterSpritePaths = GetMonsterSpritePaths();
                return monsterSpritePaths;
            }
        }
        
        public string[] TableInfo
        {
            get
            {
                if (this.tableInfo == null) this.tableInfo = GetTableInfo();
                return tableInfo;
                //return GetTableInfo();
            }
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {   
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region 테이블처리 코드
       
        public IG_Table ExportTable(string branch)
        {
            // 각종 테이블정보 초기화
            InitiateInfo();

            int cnt = 0;

            Dictionary<string, string> define;
            Dictionary<string, string> desc;        // 중국서버 xlsx 파일 생성용
            Dictionary<string, string> descCHN;        // 중국서버 xlsx 파일 생성용
            Dictionary<string, string> dataType;
            Dictionary<string, Dictionary<string, string>> subgroups;

            // 필드 최소/최대값 검증을 위한 필드
            Dictionary<string, Dictionary<string, string>> mins;
            Dictionary<string, Dictionary<string, string>> maxes;
            
            /*
            IG_Table table = new IG_Table(GetTableName());
            if (table == null) throw new Exception("테이블명이 제대로 설정되지 않았습니다.");
             */

            if (!BranchDefines.ContainsKey(branch)) throw new Exception("[" + branch + "] 브랜치 설정이 존재하지 않습니다.");

            define = BranchDefines[branch];
            ChangeBranch(branch);

            IG_Table table = new IG_Table(GetTableName(), define.First().Key == Properties.Settings.Default.UniqueKeyName);
            if (table == null) throw new Exception("테이블명이 제대로 설정되지 않았습니다.");

            dataType = BranchDataTypes[branch];
            desc = BranchDataDescriptions[branch];
            descCHN = BranchDataDescriptionsCHN[branch];
            subgroups = GetSubgroups();

            mins = GetBranchMins();
            maxes = GetBranchMaxes();

            // row 순서대로 브랜치를 체크하여 데이터로 출력
            foreach (Excel.Worksheet ws in Globals.IG_PlanAddIn.Application.Worksheets)
            {
                foreach (Excel.ListObject lo in ws.ListObjects)
                {
                    if (lo.Name.Length >= Properties.Settings.Default.TablePrefix.Length && lo.Name.Substring(0, Properties.Settings.Default.TablePrefix.Length).Equals(Properties.Settings.Default.TablePrefix))
                    {
                        // contentsID가 없는 테이블은 무사통과
                        bool hasContentsID = false;

                        // 컨텐츠ID 체크
                        for (int i = 1; i <= lo.ListColumns.Count; i++)
                            if (lo.ListColumns.get_Item(i).Name.ToUpper().Equals("contentsID".ToUpper()))
                            {
                                hasContentsID = true;
                                break;
                            }

                        // 필드가 define에 존재하는지 체크
                        // Alias 필드인지 확인
                        Dictionary<string, int> indexMatch = new Dictionary<string, int>();

                        foreach (string k in BranchDefines[branch].Keys)
                        {
                            try
                            {
                                if (BranchAliases[branch][k] == null) indexMatch.Add(k, lo.ListColumns[k].Index);
                                else indexMatch.Add(k, lo.ListColumns[BranchAliases[branch][k]].Index);
                            }
                            catch
                            {
                                indexMatch.Add(k, 0);
                            }
                        }

                        // 유일키 인덱스 구하기
                        int keyIndex = lo.ListColumns[BranchDefines[branch].First().Key].Index;

                        // 주석 인덱스 구하기
                        int commentIndex = 0;
                        for (int c = 1; c <= lo.ListColumns.Count; c++)
                            if (lo.HeaderRowRange[1, c].value2 == Properties.Settings.Default.CommentString1 || lo.HeaderRowRange[1, c].value2 == Properties.Settings.Default.CommentString2)
                            {
                                commentIndex = c;
                                break;
                            }

                            // IG_Table에 define / desc 정보 보내기: 중국서버 xlsx 파일 생성용
                            if (!table.ExistsMetaTable() && ExportXLSX())
                            {
                                table.CreateMetaTable();
                                table.AddMetaTableInfos(define, dataType, desc, descCHN);
                            }

                        for (int r = 1; r <= lo.ListRows.Count; r++)
                        {
                            // tableinfo에서 설정된 contents만 출력                            
                            if (hasContentsID == false || IsValidContentsId(lo.DataBodyRange[r, lo.ListColumns["contentsID"].Index].value2))
                            {
                                cnt++;
                                int id = Convert.ToInt32(lo.DataBodyRange[r, keyIndex].value2);
                                if (id > 0)
                                {
                                    if (table.ContainsKey(id) && table.UniqueKey)
                                        throw new Exception(String.Format("인덱스 중복오류: {0}", id));       
                             
                                    // 주석 삽입
                                    if (commentIndex > 0)
                                        table.StartAdd(id, Convert.ToString(lo.DataBodyRange[r, commentIndex].value2));
                                    else
                                        table.StartAdd(id, "");

                                    foreach (string k in BranchDefines[branch].Keys)
                                    {
                                        object tmp = lo.DataBodyRange[r, indexMatch[k]].value2;

                                        // 데이터 타입 검증 후, 출력
                                        try
                                        {
                                            // 필드정의에 존재하나 실데이터가 없는 경우, 기본값 출력   
                                            if (indexMatch[k] > 0 && tmp != null && dataType[k] != null)
                                                table.AddElement(k, ToXmlString(GetValidateData(Convert.ToString(tmp), dataType[k], subgroups, mins[branch][k], maxes[branch][k])), dataType[k], subgroups);
                                            else
                                                table.AddElement(k, ToXmlString(GetValidateData(Convert.ToString(BranchDefines[branch][k]), dataType[k], subgroups, mins[branch][k], maxes[branch][k])), dataType[k], subgroups);
                                        }
                                        catch (Exception e)
                                        {
                                            System.Windows.Forms.Clipboard.Clear();
                                            System.Windows.Forms.Clipboard.SetText(Convert.ToString(id));
                                            if (dataType[k] != null)
                                            {
                                                //if (table.ExistsMetaTable())
                                                 //   table.ReleaseMetaTable();
                                                throw new Exception(String.Format("[데이터타입 오류]\n{0}\n인덱스: {1}\n필드명: {2}", e.Message, id, k));
                                            }
                                            else
                                            {
                                                //if (table.ExistsMetaTable())
                                                 //   table.ReleaseMetaTable();
                                                throw new Exception(String.Format("[데이터타입 미설정 오류]\n인덱스: {0}", k));
                                            }
                                        }
                                    }
                                    table.EndAdd();

                                    if (r % 5 == 0)
                                        Globals.IG_PlanAddIn.Application.StatusBar = "계산 중: " + r + "/" + lo.ListRows.Count + "(" + lo.Name + ")";
                                }
                            }
                        }
                        //table.CloseMetaTable();
                    }
                }
            }

            if (table.ExistsMetaTable())
            {
                table.SaveMetaTable();
                table.CloseMetaTable();
            }
            if (cnt <= 0) throw new Exception("테이블명이 정확하지 않습니다.");
                return table;
        }
        #endregion

        #region DEFINE 처리 코드

        public String ExportDefine(string branch)
        {
            // 각종 테이블정보 초기화
            InitiateInfo();
            var sb = new StringBuilder();
            var wr = new JsonTextWriter(new StringWriter(sb));

            Dictionary<string, string> define;
            Dictionary<string, string> dataType;
            Dictionary<string, Dictionary<string, string>> subgroups;

            // 필드 최소/최대값 검증을 위한 필드
            Dictionary<string, string> min;
            Dictionary<string, string> max;

            // 필드 참조 테이블 검증을 위한 필드
            Dictionary<string, string> refTable;
            Dictionary<string, string> refField;

            if (!BranchDefines.ContainsKey(branch)) throw new Exception("[" + branch + "] 브랜치 설정이 존재하지 않습니다.");

            define = BranchDefines[branch];
            ChangeBranch(branch);

            //IG_Table table = new IG_Table(GetTableName(), define.First().Key == Properties.Settings.Default.UniqueKeyName);
            //if (table == null) throw new Exception("테이블명이 제대로 설정되지 않았습니다.");

            dataType = BranchDataTypes[branch];
            subgroups = GetSubgroups();
            char[] subgroupSeparator = {';'};
            min = GetBranchMins()[branch];
            max = GetBranchMaxes()[branch];
            refTable = GetBranchRefTable()[branch];
            refField = GetBranchRefField()[branch];

            // 여기서부터 def-######.json 파일 추출
            wr.Formatting = Formatting.Indented;
            wr.WriteStartObject();
            wr.WritePropertyName("tablename");
            wr.WriteValue(GetTableName());
            wr.WritePropertyName("tablefields");
            wr.WriteStartArray();
            foreach(var fieldName in define.Keys)
            {
                wr.WriteStartObject();
                wr.WritePropertyName("name");
                wr.WriteValue(fieldName);
                wr.WritePropertyName("datatype");
                wr.WriteValue(IG_Table.GetMetaTableDataType(dataType[fieldName]));
                wr.WritePropertyName("isunique");                
                wr.WriteValue(dataType[fieldName] == "UniqueKEY");
                if (subgroups.ContainsKey(dataType[fieldName].ToUpper()))
                {
                    wr.WritePropertyName("limitedvalue");
                    wr.WriteStartArray();
                    // 서브그룹 설정
                    foreach (var v in subgroups[dataType[fieldName].ToUpper()].Values)
                        wr.WriteRawValue(v.Split(subgroupSeparator)[0]);
                    wr.WriteEndArray();
                }
                if (!String.IsNullOrEmpty(min[fieldName]))
                {
                    wr.WritePropertyName("minvalue");
                    wr.WriteRawValue(min[fieldName]);
                }
                if (!String.IsNullOrEmpty(max[fieldName]))
                {
                    wr.WritePropertyName("maxvalue");
                    wr.WriteRawValue(max[fieldName]);
                }

                if (!String.IsNullOrEmpty(refTable[fieldName]) && !String.IsNullOrEmpty(refField[fieldName]))
                {
                    var tmpTables = refTable[fieldName].Split(",".ToCharArray());
                    var tmpFields = refField[fieldName].Split(",".ToCharArray());

                    if (tmpTables.Length != tmpFields.Length) throw new Exception("[" + fieldName + "] 필드의 ref_table/ref_field 설정에 오류가 있습니다.");

                    wr.WritePropertyName("ref_tables");
                    wr.WriteStartArray();
                    for (int i = 0; i < tmpTables.Length; i++)
                    {
                        wr.WriteStartObject();
                        wr.WritePropertyName("tableName");
                        wr.WriteValue(tmpTables[i]);
                        wr.WritePropertyName("tableField");
                        wr.WriteValue(tmpFields[i]);
                        wr.WriteEndObject();
                    }
                    wr.WriteEndArray();
                }
                wr.WriteEndObject();

            }
            wr.WriteEndArray();
            wr.WriteEndObject();

            return sb.ToString();
        }
        #endregion

        #region 몬스터테이블 처리코드
        public IG_Table[] ExportMonsterTable(string branch, string[] stages)
        {
            // 각종 테이블정보 초기화
            InitiateInfo();

            int cnt = 0;
            string tableName = GetTableName();

            Dictionary<string, string> define;
            Dictionary<string, string> dataType;
            Dictionary<string, Dictionary<string, string>> subgroups;
            
            if (!BranchDefines.ContainsKey(branch)) throw new Exception("[" + branch + "] 브랜치 설정이 존재하지 않습니다.");
            
            define = BranchDefines[branch];
            ChangeBranch(branch);

            dataType = BranchDataTypes[branch];
            subgroups = GetSubgroups();

            Dictionary<string, IG_Table> tables = new Dictionary<string, IG_Table>();

            // row 순서대로 브랜치를 체크하여 데이터로 출력
            foreach (Excel.Worksheet ws in Globals.IG_PlanAddIn.Application.Worksheets)
            {
                foreach (Excel.ListObject lo in ws.ListObjects)
                {
                    if (lo.Name.Length >= Properties.Settings.Default.TablePrefix.Length && lo.Name.Substring(0, Properties.Settings.Default.TablePrefix.Length).Equals(Properties.Settings.Default.TablePrefix))
                    {
                        // contentsID가 없는 테이블은 무사통과
                        bool hasContentsID = false;

                        // 컨텐츠ID 체크
                        for (int i = 1; i <= lo.ListColumns.Count; i++)
                            if (lo.ListColumns.get_Item(i).Name.Equals("contentsID"))
                            {
                                hasContentsID = true;
                                break;
                            }

                        // 필드가 define에 존재하는지 체크
                        // Alias 필드인지 확인
                        Dictionary<string, int> indexMatch = new Dictionary<string, int>();

                        foreach (string k in BranchDefines[branch].Keys)
                        {
                            try
                            {
                                if (BranchAliases[branch][k] == null) indexMatch.Add(k, lo.ListColumns[k].Index);
                                else indexMatch.Add(k, lo.ListColumns[BranchAliases[branch][k]].Index);
                            }
                            catch
                            {
                                indexMatch.Add(k, 0);
                            }
                        }

                        // 유일키 인덱스 구하기
                        int keyIndex = lo.ListColumns[BranchDefines[branch].First().Key].Index;

                        for (int r = 1; r <= lo.ListRows.Count; r++)
                        {
                            // tableinfo에서 설정된 contents만 출력                            
                            if (hasContentsID == false || IsValidContentsId(lo.DataBodyRange[r, lo.ListColumns["contentsID"].Index].value2))
                            {
                                string stage = Convert.ToString(lo.DataBodyRange[r, lo.ListColumns["Stage"].Index].value2);

                                if (String.IsNullOrWhiteSpace(stage)) stage = Properties.Settings.Default.MonsterTableDefaultStageName;

                                if (!tables.ContainsKey(stage))
                                    tables.Add(stage, new IG_Table(tableName));

                                cnt++;
                                int id = Convert.ToInt32(lo.DataBodyRange[r, keyIndex].value2);
                                if (id > 0)
                                {
                                    if (tables[stage].ContainsKey(id))
                                        throw new Exception(String.Format("인덱스 중복오류: {0}", id));

                                    tables[stage].StartAdd(id);

                                    foreach (string k in BranchDefines[branch].Keys)
                                    {
                                        object tmp = lo.DataBodyRange[r, indexMatch[k]].value2;

                                        // 데이터 타입 검증 후, 출력
                                        try
                                        {
                                            // 필드정의에 존재하나 실데이터가 없는 경우, 기본값 출력   
                                            if (indexMatch[k] > 0 && tmp != null && dataType[k] != null)
                                                tables[stage].AddElement(k, ToXmlString(GetValidateData(Convert.ToString(tmp), dataType[k], subgroups)), dataType[k], subgroups);
                                            else
                                                tables[stage].AddElement(k, ToXmlString(GetValidateData(Convert.ToString(BranchDefines[branch][k]), dataType[k], subgroups)), dataType[k], subgroups);
                                        }
                                        catch (Exception e)
                                        {
                                            System.Windows.Forms.Clipboard.Clear();
                                            System.Windows.Forms.Clipboard.SetText(Convert.ToString(id));
                                            if (dataType[k] != null)
                                                throw new Exception(String.Format("[{0} 데이터타입 오류]\n인덱스: {1}\n필드명: {2}", e.Message, id, k));
                                            else
                                                throw new Exception(String.Format("[데이터타입 미설정 오류]\n필드명: {0}", k));
                                        }
                                    }
                                    tables[stage].EndAdd();
                                }
                            }
                        }
                    }
                }
            }
            if (cnt <= 0) throw new Exception("테이블명이 정확하지 않습니다.");

            return tables.Values.ToArray();          
        }
        #endregion

        #region 스테이지노트 처리 코드
        public string ExportNote()
        {
            int cnt = 0;
            int currentRound = 0;
            int nextRound = 0;

            // 스테이지 최대 길이 및 보상정보 구하기
            int maxRound = GetNoteLength();
            IG_StageNote note = new IG_StageNote(maxRound);

            // row 순서대로 브랜치를 체크하여 데이터로 출력
            foreach (Excel.Worksheet ws in Globals.IG_PlanAddIn.Application.Worksheets)
            {
                foreach (Excel.ListObject lo in ws.ListObjects)
                {
                    if (lo.Name.Length >= Properties.Settings.Default.NotePrefix.Length && lo.Name.Substring(0, Properties.Settings.Default.NotePrefix.Length).Equals(Properties.Settings.Default.NotePrefix))
                    {
                        for (int r = 1; r <= note.Length; r++)
                        {
                            if (r > currentRound) cnt++;
                            currentRound = Convert.ToInt32(lo.DataBodyRange[cnt, lo.ListColumns["Round"].Index].value2);
                            nextRound = Convert.ToInt32(lo.DataBodyRange[cnt+1, lo.ListColumns["Round"].Index].value2);

                            List<Tuple<int, int, float, int>> noteElement = new List<Tuple<int, int, float, int>>();
                            if (r == currentRound)
                            {
                                if (note.Count <= 0) note.StartNote(currentRound);
                                note.StartAdd(cnt);
                                for (int i = 1; i <= Properties.Settings.Default.NoteMaxSpawn; i++)
                                {
                                    int spawn = 0;
                                    try
                                    {
                                        spawn = Convert.ToInt32(lo.DataBodyRange[cnt, lo.ListColumns["Spawn" + i].Index].value2);
                                    }
                                    catch(FormatException)
                                    {
                                        spawn = 0;
                                    }

                                    // 1~5까지 등장위치지정(기본값: 0)
                                    int pos = 0;
                                    try
                                    {
                                        pos = Convert.ToInt32(lo.DataBodyRange[cnt, lo.ListColumns["Pos" + i].Index].value2);
                                        
                                    }
                                    catch (Exception)
                                    {
                                        pos = 0;
                                    }
                                    //int prob = Convert.ToInt32(lo.DataBodyRange[cnt, lo.ListColumns["Prob" + i].Index].value2);

                                    if (spawn > 0)
                                        if (r < note.Length)
                                            noteElement.Add(new Tuple<int, int, float, int>(spawn, ValidateNoteProb(r, lo.DataBodyRange[cnt, lo.ListColumns["Prob" + i].Index].value2), (float)(nextRound - currentRound) / 10, pos));
                                        else
                                            noteElement.Add(new Tuple<int, int, float, int>(spawn, ValidateNoteProb(r, lo.DataBodyRange[cnt, lo.ListColumns["Prob" + i].Index].value2), 0f, pos));
                                }

                                if (IsValidNote(r, noteElement)) note.AddElement(noteElement);

                                note.EndAdd();
                            }
                        }
                    }
                }
            }
            if (cnt <= 0) throw new Exception("테이블명이 정확하지 않습니다.");

            return note.ToString();
        }
        #endregion

        #region 몬스터정보 처리
        public void RefreshMonsterColor(List<MonsterInfo> monList)
        {
            // 몬스터정보 리스트를 인덱스순 dictionary로 변경함
            Dictionary<long, Color> monDic = new Dictionary<long, Color>();
            foreach(var info in monList)           
                if (!monDic.ContainsKey(info.index))
                    monDic.Add(info.index, info.color);

            int cnt = 0;
            int maxRound = GetNoteLength();

            foreach (Excel.Worksheet ws in Globals.IG_PlanAddIn.Application.Worksheets)
            {
                foreach (Excel.ListObject lo in ws.ListObjects)
                {
                    if (lo.Name.Length >= Properties.Settings.Default.NotePrefix.Length && lo.Name.Substring(0, Properties.Settings.Default.NotePrefix.Length).Equals(Properties.Settings.Default.NotePrefix))
                    {
                        for (int r = 1; r <= maxRound; r++)
                        {
                            for (int i = 1; i <= Properties.Settings.Default.NoteMaxSpawn; i++)
                            {
                                int spawn = 0;
                                try
                                {
                                    spawn = Convert.ToInt32(lo.DataBodyRange[r, lo.ListColumns["Spawn" + i].Index].value2);                                    
                                }
                                catch (FormatException)
                                {
                                    spawn = 0;
                                }

                                if (r < maxRound)
                                {
                                    if (spawn > 0)
                                        lo.DataBodyRange[r, lo.ListColumns["Spawn" + i].Index].Interior.Color = System.Drawing.ColorTranslator.ToOle(monDic[spawn]);
                                    else
                                        lo.DataBodyRange[r, lo.ListColumns["Spawn" + i].Index].Interior.ColorIndex = 0;
                                }                                        
                            }

                            cnt++;
                        }
                    }
                }
            }

            if (cnt <= 0) throw new Exception("테이블명이 정확하지 않습니다.");
        }

        public string GetStageName()
        {
            // 스테이지노트명을 직접 참조하도록 코드 변경됨
            string stageName = Path.GetFileNameWithoutExtension(Application.ActiveWorkbook.Name);
            return stageName;

            /*
            string name = "";
            
            try
            {
                string stageTable = File.ReadAllText(Properties.Settings.Default.StageTablePath);
                JsonTextReader reader = new JsonTextReader(new StringReader(stageTable));

                while (reader.Read())
                {
                    if (reader.TokenType == JsonToken.PropertyName)
                    {
                        
                        if ((string)reader.Value == "StageName") name = reader.ReadAsString();
                        if ((string)reader.Value == "StageNoteFile")
                        {
                            string[] tmp = reader.ReadAsString().Split(Path.AltDirectorySeparatorChar);                            
                            //System.Windows.Forms.MessageBox.Show(Path.GetFileNameWithoutExtension(reader.ReadAsString()));
                            if (Path.GetFileNameWithoutExtension(tmp[tmp.Length - 1]) == stageName) return name;
                            
                        }
                        if (Path.GetFileNameWithoutExtension(reader.ReadAsString()).Equals(stageName)) return name;                                                 
                    }
                }
            }
            catch (IOException ioe)
            {
                mInfoPath.stageTable = false;
                throw ioe;
            }
            catch (ArgumentException)
            {
                mInfoPath.stageTable = false;
                throw new IOException();
            }
            catch (Exception)
            {
                throw new Exception("스테이지 테이블을 읽는 과정에서 오류가 발생하였습니다.");
            }

            throw new Exception("스테이지노트파일명과 일치하는 스테이지를 찾을 수 없습니다.");
            */
        }

        public Dictionary<int, string> GetMonsterSprite()
        { 
            // 스프라이트 데이터 읽어오기
            Dictionary<int, string> spriteNames = new Dictionary<int, string>();
            try
            {
                string resourcePathTable = File.ReadAllText(Properties.Settings.Default.ResourcePathTablePath);
                JsonTextReader reader = new JsonTextReader(new StringReader(resourcePathTable));

                int key = 0;
                while (reader.Read())
                {
                    if (reader.TokenType == JsonToken.PropertyName)
                    {
                        if ((string)reader.Value == "Index")
                        {
                            key = Convert.ToInt32(reader.ReadAsInt32());
                            spriteNames.Add(key, null);
                        }
                        else if ((string)reader.Value == "Path")
                            spriteNames[key] = reader.ReadAsString();
                    }
                }
            }
            catch (IOException ioe)
            {
                mInfoPath.resourcePathTable = false;
                throw ioe;
            }
            catch (ArgumentException)
            {
                mInfoPath.resourcePathTable = false;
                throw new IOException();
            }
            catch (Exception)
            {
            }

            return spriteNames;
        }

        public List<MonsterInfo> GetMonsterInfos(string stageName, Dictionary<int, string> spriteNames)
        {
            List<MonsterInfo> monsterInfos = new List<MonsterInfo>();

            try
            {
                // 스테이지몬스터 정보불러오기
                GetMonsterInfo(stageName, spriteNames, monsterInfos);

                // 공통몬스터 정보불러오기
                GetMonsterInfo("Common", spriteNames, monsterInfos);
            }
            catch (IOException ioe)
            {
                mInfoPath.monsterTable = false;
                throw ioe;
            }
            catch (ArgumentException)
            {
                mInfoPath.monsterTable = false;
                throw new IOException();
            }
            catch (Exception e)
            {
                throw e;
            }

            return monsterInfos;
        }

        private void GetMonsterInfo(string stage, Dictionary<int, string> spriteNames, List<MonsterInfo> monsterInfos)
        {

            JsonTextReader reader;
            MonsterInfo tmpInfo;

            var monsterTable = "";

            // 몬스터 데이터 읽어오기
            if (stage == "Common")
                monsterTable = File.ReadAllText(Path.Combine(Properties.Settings.Default.MonsterTablePath, Properties.Settings.Default.MonsterTableExportName + "_" + stage + ".json"));
            else
                monsterTable = File.ReadAllText(Path.Combine(Properties.Settings.Default.MonsterTablePath, Properties.Settings.Default.StageMonsterTablePath, Properties.Settings.Default.MonsterTableExportName + "_" + stage + ".json"));

            /*JsonTextReader*/
            reader = new JsonTextReader(new StringReader(monsterTable));

            int cnt = 0;
            /*MonsterInfo*/
            tmpInfo = new MonsterInfo();
            while (reader.Read())
            {
                if (reader.TokenType == JsonToken.PropertyName)
                    switch ((string)reader.Value)
                    {
                        case "Index":
                            if (cnt > 0)
                                monsterInfos.Add(tmpInfo);
                            cnt++;
                            tmpInfo.index = Convert.ToInt32(reader.ReadAsInt32());
                            break;
                        case "Stage":
                            tmpInfo.stage = reader.ReadAsString();
                            break;
                        case "Type":
                            try
                            {
                                tmpInfo.type = //reader.ReadAsString();
                                Enum.GetName(typeof(MonsterType), Convert.ToInt32(reader.ReadAsInt32()));
                            }
                            catch (ArgumentException)
                            {
                                tmpInfo.type = "NONE";
                            }
                            break;
                        case "Monster_Sprite":
                            int tmpIndex = Convert.ToInt32(reader.ReadAsInt32());
                            if (spriteNames.ContainsKey(tmpIndex))
                                tmpInfo.sprite = spriteNames[tmpIndex];
                            else
                                tmpInfo.sprite = "";
                            break;
                        case "MonsterExp":
                            tmpInfo.exp = Convert.ToInt32(reader.ReadAsInt32());
                            break;
                        case "MonsterPoint":
                            tmpInfo.point = Convert.ToInt32(reader.ReadAsInt32());
                            break;
                        case "MonsterMinGold":
                            tmpInfo.goldMin = Convert.ToInt32(reader.ReadAsInt32());
                            break;
                        case "MonsterMaxGold":
                            tmpInfo.goldMax = Convert.ToInt32(reader.ReadAsInt32());
                            break;
                        case "MonsterHP":
                            tmpInfo.HP = Convert.ToInt32(reader.ReadAsInt32());
                            break;
                        case "MonsterPAtk":
                            tmpInfo.atk = Convert.ToInt32(reader.ReadAsInt32());
                            break;
                        case "Speed":
                            tmpInfo.speed = Convert.ToDouble(reader.ReadAsString()) / Properties.Settings.Default.PermilFactor;
                            break;
                        case "MonsterScale":
                            tmpInfo.scale = Convert.ToDouble(reader.ReadAsString()) / Properties.Settings.Default.PermilFactor;

                            // 타입/스케일 정보로 색깔값 지정
                            tmpInfo.color = Globals.IG_PlanAddIn.GetMonsterTypeColor(tmpInfo.type, tmpInfo.scale);
                            break;
                        default:
                            break;
                    }
            }
            // 마지막 인덱스 몬스터 추가
            if (tmpInfo.index > 0)
                monsterInfos.Add(tmpInfo);
            
        }

        // 몬스터타입에 따른 색깔구분
        public Color GetMonsterTypeColor(String monType, double monScale = 1.0d)
        {
            var color = Color.White;

            switch (monType.ToUpper())
            {
                case "MONSTER":
                    color = Color.LightPink;
                    break;
                case "VETERAN":
                    color = Color.Pink;
                    break;
                case "BOSS":
                    color = Color.LightCoral;
                    break;
                case "OBJECT":
                    color = Color.LightCyan;
                    break;
                case "WALL":
                    color = Color.LightBlue;
                    break;
                case "TRAPCHEST":
                    color = Color.Lime;
                    break;
                case "BIGCHEST":
                    color = Color.Yellow;
                    break;
                case "CHEST":
                    color = Color.LightYellow;
                    break;
                case "ANGEL":
                    color = Color.LightGreen;
                    break;
                case "DEMON":
                    color = Color.LightSalmon;
                    break;
                case "GOLD":
                    color = Color.Gold;
                    break;
                case "BIGGOLD":
                    color = Color.Goldenrod;
                    break;
                case "ENDBOX":
                    color = Color.LightSlateGray;
                    break;
                case "WARNING":
                    color = Color.Purple;
                    break;
                case "SANCTUARY":
                    color = Color.LightSkyBlue;
                    break;
                case "BOMBTOP":
                    color = Color.LightPink;
                    break;
                case "BOMBMID":
                    color = Color.Pink;
                    break;
                case "BOMBBOTTOM":
                    color = Color.DeepPink;
                    break;
                case "BANNER":
                    color = Color.DarkGray;;
                    break;
                default:
                    break;
            }

            return Color.FromArgb((int)(color.A * Math.Min(1.0d, monScale)), color);
        }

        // 밸런스문서에서 처리함
        //public int RefreshMonsterInfoTable(List<MonsterInfo> monsterInfos, string stage)
        //{
        //    int cntTable = 0;
        //    int cnt = 0;
        //    foreach (Excel.Worksheet ws in Globals.IG_PlanAddIn.Application.Worksheets)
        //    {
        //        foreach (Excel.ListObject lo in ws.ListObjects)
        //        {
        //            if (lo.Name.Equals(Properties.Settings.Default.MonsterInfoTableName))
        //            {
        //                cntTable++;

        //                foreach (Excel.Shape shape in ws.Shapes)
        //                    shape.Delete();

        //                lo.DataBodyRange.ClearContents();

        //                foreach (MonsterInfo info in monsterInfos)
        //                {
        //                    if (String.IsNullOrEmpty(info.stage) || info.stage.Trim() == stage.Trim())
        //                    {
        //                        cnt++;
        //                        lo.DataBodyRange[cnt, lo.ListColumns["인덱스"].Index].value2 = info.index;
        //                        lo.DataBodyRange[cnt, lo.ListColumns["타입"].Index].value2 = info.type;
        //                        lo.DataBodyRange[cnt, lo.ListColumns["경험치"].Index].value2 = info.exp;
        //                        lo.DataBodyRange[cnt, lo.ListColumns["포인트"].Index].value2 = info.point;
        //                        lo.DataBodyRange[cnt, lo.ListColumns["골드"].Index].value2 = info.GetGold();

        //                        //lo.DataBodyRange[cnt, lo.ListColumns["스프라이트"].Index].value2 = info.sprite;         
        //                        InsertMonsterSpriteImage(info.sprite, ws, lo, cnt, lo.ListColumns["스프라이트"].Index);
        //                    }
        //                }

        //                lo.Resize(ws.Range[ws.Cells[lo.HeaderRowRange.Row, lo.HeaderRowRange.Column], ws.Cells[lo.HeaderRowRange.Row + cnt, lo.HeaderRowRange.Column + lo.ListColumns.Count - 1]]);                        
        //            }
        //        }
        //    }

        //    if (cntTable <= 0)
        //        throw new Exception("[" + Properties.Settings.Default.MonsterInfoTableName + "]표가 존재하지 않습니다.");

        //    return cnt;
        //}

        // 밸런스문서에서 처리함
        //private void InsertMonsterSpriteImage(string spriteName, Excel.Worksheet ws, Excel.ListObject lo, int row, int col)
        //{
        //    Excel.Range rng = ws.Range[ws.Cells[lo.HeaderRowRange.Row + row, lo.HeaderRowRange.Column + col-1], ws.Cells[lo.HeaderRowRange.Row + row, lo.HeaderRowRange.Column + col-1]];            
           
        //    float size = Properties.Settings.Default.SpriteImageSize;

        //    ws.Rows[lo.HeaderRowRange.Row + row].RowHeight = size;
        //    //ws.Columns[lo.HeaderRowRange.Column + col - 1].ColumnWidth = size;

        //    if (!System.IO.Directory.Exists(Properties.Settings.Default.MonsterSpritePath))
        //    {
        //        mInfoPath.monsterSprite = false;
        //        throw new FileNotFoundException("스프라이트 폴더 패스 오류", Properties.Settings.Default.MonsterSpritePath);
        //    }
                
        //    try
        //    {                
        //        //ws.Shapes.AddPicture(Properties.Settings.Default.MonsterSpritePath + Path.DirectorySeparatorChar + spriteName + Properties.Settings.Default.MonsterSpriteExtension,
        //        if (!String.IsNullOrEmpty(spriteName))
        //        {
        //            ws.Shapes.AddPicture(MonsterSpritePaths[spriteName],
        //                       Office.MsoTriState.msoTrue, Office.MsoTriState.msoTrue,
        //                       rng.Left /*+ ( rng.Left - size ) / 2*/, rng.Top, Convert.ToInt32(size / 2), Convert.ToInt32(size / 2));

        //            rng.AutoFit();
        //        }
        //    }
        //    catch (Exception)
        //    {
        //    }
        //}

        private Dictionary<string, string> GetMonsterSpritePaths()
        {
            //string[] _Result    = Directory.GetFiles("/assets", "*.png", SearchOption.AllDirectories);
            //Dictionary<string, string> tmpPaths = new Dictionary<string, string>();
            //System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(Properties.Settings.Default.MonsterSpritePath);

            //GetMonsterSpritePath(di, tmpPaths);

            //return tmpPaths;

            string[] _results = Directory.GetFiles(Properties.Settings.Default.MonsterSpritePath, "*.png", SearchOption.AllDirectories);

            Dictionary<string, string> paths = new Dictionary<string, string>();

            foreach(var result in _results)
            {
                if (!paths.ContainsKey(Path.GetFileNameWithoutExtension(result)))
                    paths.Add(Path.GetFileNameWithoutExtension(result), result);
            }

            return paths;

        }

        private void GetMonsterSpritePath(System.IO.DirectoryInfo di, Dictionary<string, string> paths)
        {
             foreach (System.IO.FileInfo f in di.GetFiles())
                if (!paths.ContainsKey(Path.GetFileNameWithoutExtension(f.FullName)))
                    paths.Add(Path.GetFileNameWithoutExtension(f.FullName), f.FullName);

            foreach (System.IO.DirectoryInfo sdi in di.GetDirectories())
                GetMonsterSpritePath(sdi, paths);
        }

        internal void InitiateMonsterInfo()
        {
            mInfoPath.monsterTable = true;
            mInfoPath.resourcePathTable = true;
            mInfoPath.monsterSprite = true;
            mInfoPath.stageTable = true;       
        }

        internal void InsertMonsterIndex(int index)
        {
            int lastRow = 0;
            int lastCol = 1;
            foreach (Excel.Worksheet ws in Globals.IG_PlanAddIn.Application.Worksheets)
            {
                foreach (Excel.ListObject lo in ws.ListObjects)
                {
                    if (lo.Name.Length >= Properties.Settings.Default.NotePrefix.Length && lo.Name.Substring(0, Properties.Settings.Default.NotePrefix.Length).Equals(Properties.Settings.Default.NotePrefix))
                    {
                        for (int r = 1; r <= lo.ListRows.Count; r++)
                        {
                            int round = Convert.ToInt32(lo.DataBodyRange[r, lo.ListColumns["Round"].Index].value2);
                            int mCnt = 0;
                            for (int c = 1; c <= lo.ListColumns.Count; c++)
                                if (lo.HeaderRowRange[1, c].value2.Length >= "Spawn".Length && (string)lo.HeaderRowRange[1, c].value2.Substring(0, "Spawn".Length) == "Spawn")
                                    if (Convert.ToInt32(lo.DataBodyRange[r, c].value2) > 0) mCnt++;
                                
                            // 라운드 설정이 있지만 몬스터정보가 없거나, 몬스터정보가 없는 셀 우선 입력     
                            if (!(round > 0 || round <= 0 && mCnt > 0))
                            {
                                lastRow = r;
                                break;
                            }
                        }


                        // 인덱스 기입
                        if (lastRow <= 0)                                
                        {
                            lo.ListRows.AddEx();
                            lastRow = lo.ListRows.Count;
                        }
                        lo.DataBodyRange[lastRow, lo.ListColumns["Spawn" + lastCol].Index].value2 = (double)index;

                        ws.Activate();
                        ws.Range[ws.Cells[lo.HeaderRowRange.Row + lastRow, lo.HeaderRowRange.Column + lo.ListColumns["Spawn" + lastCol].Index - 1], ws.Cells[lo.HeaderRowRange.Row + lastRow, lo.HeaderRowRange.Column + lo.ListColumns["Spawn" + lastCol].Index - 1]].Select();
                    }
                }
            }
        }
        #endregion

        #region 리소스패스 검증 처리코드
        public int VerifyResourcePathTable()
        {
            // 에셋폴더 순환체크
            Dictionary<string, string> resourcePaths = new Dictionary<string, string>();

            if (!System.IO.Directory.Exists(Properties.Settings.Default.ResourceAssetPath))
                throw new FileNotFoundException("에셋폴더 패스 오류", Properties.Settings.Default.ResourceAssetPath);

            System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(Properties.Settings.Default.ResourceAssetPath);

            GetResourcePaths(di, resourcePaths);

            // 리소스패스 테이블 체크
            int cnt = 0;
            int tableCnt = 0;

            foreach (Excel.Worksheet ws in Globals.IG_PlanAddIn.Application.Worksheets)
            {
                foreach (Excel.ListObject lo in ws.ListObjects)
                {
                    if (lo.Name.Length >= Properties.Settings.Default.ResourcePathTableName.Length && lo.Name.Substring(0, Properties.Settings.Default.ResourcePathTableName.Length).Equals(Properties.Settings.Default.ResourcePathTableName))
                    {
                        tableCnt++;

                        string tmpPath;
                        int verificationIndex = 0;
                        try
                        {
                            verificationIndex = lo.ListColumns[Properties.Settings.Default.ResourceVerificationFieldName].Index;
                        }
                        catch
                        {
                            lo.ListColumns.Add();
                            verificationIndex = lo.ListColumns.Count;
                            lo.HeaderRowRange[1, verificationIndex].value2 = Properties.Settings.Default.ResourceVerificationFieldName;
                        }
                        finally
                        {
                            lo.ListColumns[verificationIndex].DataBodyRange.ClearContents();
                        }

                        for (int r = 1; r <= lo.ListRows.Count; r++)
                        {
                            tmpPath = (string)lo.DataBodyRange[r, lo.ListColumns["Path"].Index].value2;
                            if (!String.IsNullOrEmpty(tmpPath))
                            {
                                tmpPath = Path.GetFileNameWithoutExtension(tmpPath).ToUpper();
                                if (resourcePaths.ContainsKey(tmpPath.ToUpper()))
                                //if (resourcePaths.ContainsKey(Path.GetFileNameWithoutExtension(tmpPath).ToUpper()))
                                {
                                    lo.DataBodyRange[r, verificationIndex].value2 = resourcePaths[tmpPath.ToUpper()];
                                }
                                else
                                {
                                    cnt++;
                                    ws.Activate();
                                    ws.Range[ws.Cells[lo.HeaderRowRange.Row + r, lo.HeaderRowRange.Column + lo.ListColumns[Properties.Settings.Default.ResourceVerificationFieldName].Index - 1], ws.Cells[lo.HeaderRowRange.Row + r, lo.HeaderRowRange.Column + lo.ListColumns[Properties.Settings.Default.ResourceVerificationFieldName].Index - 1]].Select();
                                    //ws.Activate();
                                    //ws.Range[ws.Cells[lo.HeaderRowRange.Row + r, lo.HeaderRowRange.Column + verificationIndex - 1], ws.Cells[lo.HeaderRowRange.Row + r, lo.HeaderRowRange.Column + verificationIndex - 1]].Select();
                                }
                            }
                        }
                        if (verificationIndex > 0)
                        {
                            Excel.Range rng = ws.Range[ws.Cells[lo.HeaderRowRange.Row, lo.HeaderRowRange.Column + verificationIndex - 1], ws.Cells[lo.HeaderRowRange.Row + lo.ListRows.Count, lo.HeaderRowRange.Column + verificationIndex - 1]];
                            //lo.ListColumns[Properties.Settings.Default.ResourceVerificationFieldName].AutoFit();
                            //ws.Range[ws.Cells[lo.HeaderRowRange.Row, lo.HeaderRowRange.Column + verificationIndex], ws.Cells[lo.HeaderRowRange.Row + lo.ListRows.Count - 1, lo.HeaderRowRange.Column + verificationIndex]].AutoFit();
                            //rng.AutoFit();
                            rng.Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;    
                            //ws.Range[ws.Cells[lo.HeaderRowRange.Row, lo.HeaderRowRange.Column + verificationIndex], ws.Cells[lo.HeaderRowRange.Row + lo.ListRows.Count - 1, lo.HeaderRowRange.Column + verificationIndex]].Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;                            
                        }
                    }
                }
            }

            if (tableCnt <= 0) throw new Exception("[" + Properties.Settings.Default.ResourcePathTableName + "]가 존재하지 않습니다.");

            return cnt;
        }

        private void GetResourcePaths(System.IO.DirectoryInfo di, Dictionary<string, string> paths)
        {
            foreach (System.IO.FileInfo f in di.GetFiles())
                if (!paths.ContainsKey(Path.GetFileNameWithoutExtension(f.FullName).ToUpper()))
                    paths.Add(Path.GetFileNameWithoutExtension(f.FullName).ToUpper(), f.FullName);

            foreach (System.IO.DirectoryInfo sdi in di.GetDirectories())
                GetResourcePaths(sdi, paths);
        }

        #endregion

        #region svn diff 처리코드
        internal bool SVNDiff(string fileName)
        {
            var p = new Process();
            bool result = true;

            try
            {
                p.StartInfo.UseShellExecute = false;
                p.StartInfo.RedirectStandardOutput = true;
                p.StartInfo.FileName = "tortoiseproc";
                p.StartInfo.Arguments = string.Format("/command:diff /path:{0}", fileName);
                p.Start();
                p.WaitForExit();
            }
            catch
            {
                result = false;
            }

            return result;
        }
        #endregion

        #region 기타 스테이지노트 처리코드
        private int GetNoteLength()
        {
            int maxLength = -1;
            string maxRound = "";
            bool hasEndbox = false;

            foreach (Excel.Worksheet ws in Globals.IG_PlanAddIn.Application.Worksheets)
            {
                foreach (Excel.ListObject lo in ws.ListObjects)
                {
                    if (lo.Name.Length >= Properties.Settings.Default.NotePrefix.Length && lo.Name.Substring(0, Properties.Settings.Default.NotePrefix.Length).Equals(Properties.Settings.Default.NotePrefix))
                    {
                        for (int r = 1; r <= lo.ListRows.Count; r++)
                        {
                            string lastRound = "";
                            for (int i = 1; i <= Properties.Settings.Default.NoteMaxSpawn; i++)
                            {
                                int spawn = Convert.ToInt32(lo.DataBodyRange[r, lo.ListColumns["Spawn" + i].Index].value2);
                                if (spawn == Properties.Settings.Default.EndboxIndex)
                                    hasEndbox = true;
                                lastRound += spawn;                                
                            }

                            if (lo.DataBodyRange[r, lo.ListColumns["Round"].Index].value2 < maxLength)
                            {
                                // 라운드값이 순차적이지 않을 때
                                throw new Exception(r + "번째 라운드(" + lo.DataBodyRange[r, lo.ListColumns["Round"].Index].value2 + "ms)의 시간값이 잘못 기입되어있습니다.");
                            }
                            else if (lo.DataBodyRange[r, lo.ListColumns["Round"].Index].value2 > maxLength && maxRound != lastRound)
                            {
                                maxLength = Convert.ToInt32(lo.DataBodyRange[r, lo.ListColumns["Round"].Index].value2);
                                maxRound = lastRound;
                            }
                        }
                    }
                }
            }

            if (!hasEndbox)
                throw new Exception("ENDBOX가 존재하지 않습니다.");
            if (maxLength < 0)
                throw new Exception("라운드 정보가 잘못 기입되었습니다.");

            return maxLength;
        }

        // 밸런스문서에서 처리하도록 변경
        //private int GetStagePoint(int maxRound)
        //{         
        //    foreach (Excel.Worksheet ws in Globals.IG_PlanAddIn.Application.Worksheets)
        //    {
        //        foreach (Excel.ListObject lo in ws.ListObjects)
        //        {
        //            if (lo.Name.Length >= Properties.Settings.Default.NotePrefix.Length && lo.Name.Substring(0, Properties.Settings.Default.NotePrefix.Length).Equals(Properties.Settings.Default.NotePrefix))
        //            {
        //                for (int r = 1; r <= lo.ListRows.Count; r++)
        //                    if (lo.DataBodyRange[r, lo.ListColumns["Round"].Index].value == maxRound)                            
        //                        return Convert.ToInt32(lo.DataBodyRange[r, lo.ListColumns["누적포인트"].Index].value);
        //            }
        //        }
        //    }

        //    throw new Exception("스테이지 누적포인트 정보가 잘못 기입되었습니다.");
        //}

        // 밸런스문서에서 처리하도록 변경
        //private int GetStageGold(int maxRound)
        //{
        //    foreach (Excel.Worksheet ws in Globals.IG_PlanAddIn.Application.Worksheets)
        //    {
        //        foreach (Excel.ListObject lo in ws.ListObjects)
        //        {
        //            if (lo.Name.Length >= Properties.Settings.Default.NotePrefix.Length && lo.Name.Substring(0, Properties.Settings.Default.NotePrefix.Length).Equals(Properties.Settings.Default.NotePrefix))
        //            {
        //                for (int r = 1; r <= lo.ListRows.Count; r++)
        //                    if (lo.DataBodyRange[r, lo.ListColumns["Round"].Index].value == maxRound)
        //                        return Convert.ToInt32(lo.DataBodyRange[r, lo.ListColumns["누적골드"].Index].value);
        //            }
        //        }
        //    }

        //    throw new Exception("스테이지 누적골드 정보가 잘못 기입되었습니다.");
        //}

        private bool IsValidNote(int r, List<Tuple<int, int, float, int>> noteElement)
        {
            bool isValid = true;

            //int totalProb = 0;
            int spawnCount = 0;

            // 라운드 번호는 양수여야함
            if (r <= 0) isValid = false;
            
            
            for (int i = 0; i < noteElement.Count; i++)
            {
                if (noteElement.ElementAt(i).Item1 > 0) spawnCount++;
                if (noteElement.ElementAt(i).Item2 > Properties.Settings.Default.PermilFactor) isValid = false;
                if (noteElement.ElementAt(i).Item3 < 0) isValid = false;
                if (noteElement.ElementAt(i).Item4 < 0 || noteElement.ElementAt(i).Item4 > Properties.Settings.Default.NoteMaxPosition) isValid = false;
            }

            // 최소한 하나의 몬스터아이디가 기입되어있어야하며 아이디는 양수여야 함
            if (spawnCount <= 0) isValid = false;
  
            // 등장확률의 합계는 100%를 초과할 수 없음
            //if (totalProb > Properties.Settings.Default.PermilFactor) isValid = false;
            
            //return true;
            if (isValid)
                return true;
            else
                throw new Exception(r + "라운드의 정보가 잘못 기입되었습니다.");           
        }
        private int ValidateNoteProb(int round, object prob)
        {
            if (prob != null && (int)(Convert.ToDouble(prob) * Properties.Settings.Default.PermilFactor) < 0)
                throw new Exception(round + "라운드의 확률이 잘못 기입되었습니다.");
            else if (prob == null)
                return Properties.Settings.Default.PermilFactor;
            else
                return (int)(Convert.ToDouble(prob) * Properties.Settings.Default.PermilFactor);                
        }

        internal void WriteStageNote(string xmlName, string xmlString)
        {
            if (xmlName != "")
            {
                File.WriteAllText(xmlName, xmlString, Encoding.UTF8);
            }
            else
            {
                throw new Exception("파일명이 정확하지 않습니다.");
            }
        }
        #endregion

        #region 기타 테이블처리코드
        internal void InitiateInfo()
        {
            branchList = null;
            branchDefines = null;
            branchAliases = null;
            tableInfo = null;
            branchDataTypes = null;
            branchDataDescriptions = null;
            branchDataDescriptionsCHN = null;

            branchMins = null;
            branchMaxes = null;
        }

        internal void WriteTable(string fileName, IG_Table table)
        {
            if (fileName != "")
            {
                File.WriteAllText(fileName, table.ToString(), Encoding.UTF8);
            }
            else
            {
                throw new Exception("파일명 - " + fileName + " - 이 정확하지 않습니다.");
            }
        }

        internal void WriteTable(string[] fileNames, IG_Table[] tables)
        {
            for (int i = 0; i < fileNames.Length; i++)
                WriteTable(fileNames[i], tables[i]);
        }

        internal void WriteDefine(string fileName, String def)
        {
            if (fileName != "")
            {
                File.WriteAllText(fileName, def, Encoding.UTF8);
            }
            else
            {
                throw new Exception("파일명 - " + fileName + " - 이 정확하지 않습니다.");
            }
        }

        private bool IsValidContentsId(object contentsid)
        {
            //string[] contentsList = TableInfo;

            if (contentsid == null)
                return true;
            else
            {
                foreach (string content in TableInfo)
                {
                    if (content.Equals(Convert.ToString(contentsid)))
                    {
                        return true;
                    }
                }
                return false;
            }
        }

        private void ChangeBranch(string branch)
        {
            foreach (Excel.Name name in Globals.IG_PlanAddIn.Application.ActiveWorkbook.Names)
            {
                if (name.Name.Equals(Properties.Settings.Default.BranchName))
                {
                    name.RefersToRange.Value2 = branch;
                    
                    // 레퍼런스 재계산
                    Globals.IG_PlanAddIn.Application.CalculateFull();
                    break;
                }
            }
        }

        // 노드명이 있을 경우엔 노드명으로 그 외엔 테이블명으로
        internal string GetTableName()
        {
            foreach (Excel.Name name in Globals.IG_PlanAddIn.Application.ActiveWorkbook.Names)
            {
                if (name.Name.Equals(Properties.Settings.Default.TableName))
                {
                    string tmp = name.RefersToRange.Value2;

                    if (tmp != null) return tmp;

                }
            }
            return null;
        }

        // 서버용 XLSX 추출여부
        internal bool ExportXLSX()
        {
            foreach (Excel.Name name in Globals.IG_PlanAddIn.Application.ActiveWorkbook.Names)
            {
                if (name.Name.Equals(Properties.Settings.Default.XLSX_EXPORT))
                {
                    var tmp = name.RefersToRange.Value2;

                    if (tmp != null) return tmp;

                }
            }
            return true;
        }

        // 몬스터테이블의 스테이지목록 가져오기
        internal string[] GetStagesFromMonsterTable()
        {
            List<string> stages = new List<string>();
            
            foreach (Excel.Worksheet ws in Globals.IG_PlanAddIn.Application.Worksheets)
            {
                foreach (Excel.ListObject lo in ws.ListObjects)
                {
                    if (lo.Name.Length >= Properties.Settings.Default.MonsterTableName.Length && lo.Name.Substring(0, Properties.Settings.Default.MonsterTableName.Length).Equals(Properties.Settings.Default.MonsterTableName))
                    {
                        for (int r = 1; r <= lo.ListRows.Count; r++)
                        {
                            string stage = Convert.ToString(lo.DataBodyRange[r, lo.ListColumns["Stage"].Index].value2);
                            if (String.IsNullOrWhiteSpace(stage)) stage = Properties.Settings.Default.MonsterTableDefaultStageName;
                            if (!stages.Contains(stage)) stages.Add(stage);
                        }
                    }
                }
            }

            return stages.ToArray();
        }

        // 유일키 설정불러오기
        private bool IsUniqueKey()
        {
            foreach (Excel.Name name in Globals.IG_PlanAddIn.Application.ActiveWorkbook.Names)
            {
                if (name.Name.Equals(Properties.Settings.Default.UniqueKeyName))                
                    return !(bool)name.RefersToRange.Value2;
            }

            return true;
        }

        // 브랜치 목록 가져오기
        private string[] GetBranchList()
        {   
            foreach (Excel.Worksheet ws in Globals.IG_PlanAddIn.Application.Worksheets)
            {
                string[] tmp;

                foreach (Excel.ListObject lo in ws.ListObjects)
                {
                    if (lo.Name.Equals(Properties.Settings.Default.BranchListName))
                    {
                        tmp = new string[lo.ListRows.Count];
                        for (int r = 0; r < lo.ListRows.Count; r++)
                        {
                            tmp[r] = lo.DataBodyRange[r + 1, lo.ListColumns["branchID"].Index].Value2.Trim();
                        }
                        if (tmp.Length > 0) return tmp;
                    }
                }
            }

            throw new Exception("브랜치 정보가 설정되지 않았습니다.");
        }

        // 브랜치정보 가져오기
        private Dictionary<string, Dictionary<string, string>> GetBranchDefines()
        {
            Dictionary<string, string> tmpDefine;
            Dictionary<string, Dictionary<string, string>> tmpDefines = new Dictionary<string, Dictionary<string, string>>();

            foreach (string branch in BranchList)
            {
                tmpDefine = new Dictionary<string, string>();

                foreach (Excel.Worksheet ws in Globals.IG_PlanAddIn.Application.Worksheets)
                {
                    foreach (Excel.ListObject lo in ws.ListObjects)
                    {
                        if (lo.Name.Equals(Properties.Settings.Default.BranchDefineName))
                        {
                            string k, v;
                            for (int r = 0; r < lo.ListRows.Count; r++)
                            {
                                if (branch.Equals(lo.DataBodyRange[r + 1, lo.ListColumns["branchID"].Index].value2))
                                {
                                    k = (string)lo.DataBodyRange[r + 1, lo.ListColumns["field"].Index].value2;
                                    try
                                    {
                                        v = "" + lo.DataBodyRange[r + 1, lo.ListColumns["default"].Index].value2;
                                        tmpDefine.Add(k, v);
                                    }
                                    catch
                                    {
                                        tmpDefine.Add(k, "");
                                    }
                                }
                            }
                        }
                    }
                }
                tmpDefines.Add(branch, tmpDefine);
            }
            return tmpDefines;
        }

        // DEFINE 최소/최대값 검증
        private Dictionary<string, Dictionary<string, string>> GetBranchMins()
        {
            Dictionary<string, string> tmpMin;
            Dictionary<string, Dictionary<string, string>> tmpMins = new Dictionary<string, Dictionary<string, string>>();

            foreach (string branch in BranchList)
            {
                tmpMin = new Dictionary<string, string>();

                foreach (Excel.Worksheet ws in Globals.IG_PlanAddIn.Application.Worksheets)
                {
                    foreach (Excel.ListObject lo in ws.ListObjects)
                    {
                        if (lo.Name.Equals(Properties.Settings.Default.BranchDefineName))
                        {
                            string k, v;
                            for (int r = 0; r < lo.ListRows.Count; r++)
                            {
                                if (branch.Equals(lo.DataBodyRange[r + 1, lo.ListColumns["branchID"].Index].value2))
                                {
                                    k = (string)lo.DataBodyRange[r + 1, lo.ListColumns["field"].Index].value2;
                                    try
                                    {
                                        v = "" + lo.DataBodyRange[r + 1, lo.ListColumns["min"].Index].value2;
                                        tmpMin.Add(k, v);
                                    }
                                    catch
                                    {
                                        tmpMin.Add(k, "");
                                    }
                                }
                            }
                        }
                    }
                }
                tmpMins.Add(branch, tmpMin);
            }
            return tmpMins;                
        }

        private Dictionary<string, Dictionary<string, string>> GetBranchMaxes()
        {
            Dictionary<string, string> tmpMax;
            Dictionary<string, Dictionary<string, string>> tmpMaxes = new Dictionary<string, Dictionary<string, string>>();

            foreach (string branch in BranchList)
            {
                tmpMax = new Dictionary<string, string>();

                foreach (Excel.Worksheet ws in Globals.IG_PlanAddIn.Application.Worksheets)
                {
                    foreach (Excel.ListObject lo in ws.ListObjects)
                    {
                        if (lo.Name.Equals(Properties.Settings.Default.BranchDefineName))
                        {
                            string k, v;
                            for (int r = 0; r < lo.ListRows.Count; r++)
                            {
                                if (branch.Equals(lo.DataBodyRange[r + 1, lo.ListColumns["branchID"].Index].value2))
                                {
                                    k = (string)lo.DataBodyRange[r + 1, lo.ListColumns["field"].Index].value2;
                                    try
                                    {
                                        v = "" + lo.DataBodyRange[r + 1, lo.ListColumns["max"].Index].value2;
                                        tmpMax.Add(k, v);
                                    }
                                    catch
                                    {
                                        tmpMax.Add(k, "");
                                    }
                                }
                            }
                        }
                    }
                }
                tmpMaxes.Add(branch, tmpMax);
            }
            return tmpMaxes;
        }

        private Dictionary<string, Dictionary<string, string>> GetBranchRefTable()
        {
            Dictionary<string, string> tmpTable;
            Dictionary<string, Dictionary<string, string>> tmpTables = new Dictionary<string, Dictionary<string, string>>();

            foreach (string branch in BranchList)
            {
                tmpTable = new Dictionary<string, string>();

                foreach (Excel.Worksheet ws in Globals.IG_PlanAddIn.Application.Worksheets)
                {
                    foreach (Excel.ListObject lo in ws.ListObjects)
                    {
                        if (lo.Name.Equals(Properties.Settings.Default.BranchDefineName))
                        {
                            string k, v;
                            for (int r = 0; r < lo.ListRows.Count; r++)
                            {
                                if (branch.Equals(lo.DataBodyRange[r + 1, lo.ListColumns["branchID"].Index].value2))
                                {
                                    k = (string)lo.DataBodyRange[r + 1, lo.ListColumns["field"].Index].value2;
                                    try
                                    {
                                        v = "" + lo.DataBodyRange[r + 1, lo.ListColumns["ref_table"].Index].value2;
                                        tmpTable.Add(k, v);
                                    }
                                    catch
                                    {
                                        tmpTable.Add(k, "");
                                    }
                                }
                            }
                        }
                    }
                }
                tmpTables.Add(branch, tmpTable);
            }
            return tmpTables;
        }

        private Dictionary<string, Dictionary<string, string>> GetBranchRefField()
        {
            Dictionary<string, string> tmpField;
            Dictionary<string, Dictionary<string, string>> tmpFields = new Dictionary<string, Dictionary<string, string>>();

            foreach (string branch in BranchList)
            {
                tmpField = new Dictionary<string, string>();

                foreach (Excel.Worksheet ws in Globals.IG_PlanAddIn.Application.Worksheets)
                {
                    foreach (Excel.ListObject lo in ws.ListObjects)
                    {
                        if (lo.Name.Equals(Properties.Settings.Default.BranchDefineName))
                        {
                            string k, v;
                            for (int r = 0; r < lo.ListRows.Count; r++)
                            {
                                if (branch.Equals(lo.DataBodyRange[r + 1, lo.ListColumns["branchID"].Index].value2))
                                {
                                    k = (string)lo.DataBodyRange[r + 1, lo.ListColumns["field"].Index].value2;
                                    try
                                    {
                                        v = "" + lo.DataBodyRange[r + 1, lo.ListColumns["ref_field"].Index].value2;
                                        tmpField.Add(k, v);
                                    }
                                    catch
                                    {
                                        tmpField.Add(k, "");
                                    }
                                }
                            }
                        }
                    }
                }
                tmpFields.Add(branch, tmpField);
            }
            return tmpFields;
        }

        // Alias로 설정해야할 필드인지 체크
        private Dictionary<string, Dictionary<string, string>> GetBranchAliases()
        {
            Dictionary<string, string> tmpAlias;
            Dictionary<string, Dictionary<string, string>> tmpAliases = new Dictionary<string, Dictionary<string, string>>();

            foreach (string branch in BranchList)
            {
                tmpAlias = new Dictionary<string, string>();

                foreach (Excel.Worksheet ws in Globals.IG_PlanAddIn.Application.Worksheets)
                {
                    foreach (Excel.ListObject lo in ws.ListObjects)
                    {
                        if (lo.Name.Equals(Properties.Settings.Default.BranchDefineName))
                        {
                            string k, v;
                            for (int r = 0; r < lo.ListRows.Count; r++)
                            {
                                if (branch.Equals(lo.DataBodyRange[r + 1, lo.ListColumns["branchID"].Index].value2))
                                {
                                    k = (string)lo.DataBodyRange[r + 1, lo.ListColumns["field"].Index].value2;
                                    try
                                    {
                                        v = (string)lo.DataBodyRange[r + 1, lo.ListColumns["alias"].Index].value2;
                                        tmpAlias.Add(k, v);
                                    }
                                    catch
                                    {
                                        tmpAlias.Add(k, null);
                                    }
                                }
                            }
                        }
                    }
                }
                tmpAliases.Add(branch, tmpAlias);
            }
            return tmpAliases;
        }

        // 필드 데이터타입 참조
        private Dictionary<string, Dictionary<string, string>> GetBranchDataTypes()
        {
            Dictionary<string, string> tmpDataType;
            Dictionary<string, Dictionary<string, string>> tmpDataTypes = new Dictionary<string, Dictionary<string, string>>();

            foreach (string branch in BranchList)
            {
                tmpDataType = new Dictionary<string, string>();

                foreach (Excel.Worksheet ws in Globals.IG_PlanAddIn.Application.Worksheets)
                {
                    foreach (Excel.ListObject lo in ws.ListObjects)
                    {
                        if (lo.Name.Equals(Properties.Settings.Default.BranchDefineName))
                        {
                            string k;
                            string v;
                            for (int r = 0; r < lo.ListRows.Count; r++)
                            {
                                if (branch.Equals(lo.DataBodyRange[r + 1, lo.ListColumns["branchID"].Index].value2))
                                {
                                    k = (string)lo.DataBodyRange[r + 1, lo.ListColumns["field"].Index].value2;
                                    try
                                    {
                                        v = (string)lo.DataBodyRange[r + 1, lo.ListColumns["type"].Index].value2;
                                        tmpDataType.Add(k, v);
                                    }
                                    catch
                                    {
                                        tmpDataType.Add(k, null);
                                    }
                                }
                            }
                        }
                    }
                }
                tmpDataTypes.Add(branch, tmpDataType);
            }
            return tmpDataTypes;
        }

        // 필드 설명 참조
        private Dictionary<string, Dictionary<string, string>> GetBranchDataDescriptions(string descType = "")
        {
            Dictionary<string, string> tmpDataDescription;
            Dictionary<string, Dictionary<string, string>> tmpDatatmpDataDescriptions = new Dictionary<string, Dictionary<string, string>>();

            foreach (string branch in BranchList)
            {
                tmpDataDescription = new Dictionary<string, string>();

                foreach (Excel.Worksheet ws in Globals.IG_PlanAddIn.Application.Worksheets)
                {
                    foreach (Excel.ListObject lo in ws.ListObjects)
                    {
                        if (lo.Name.Equals(Properties.Settings.Default.BranchDefineName))
                        {
                            string k;
                            string v;
                            for (int r = 0; r < lo.ListRows.Count; r++)
                            {
                                if (branch.Equals(lo.DataBodyRange[r + 1, lo.ListColumns["branchID"].Index].value2))
                                {
                                    k = (string)lo.DataBodyRange[r + 1, lo.ListColumns["field"].Index].value2;
                                    try
                                    {
                                        //v = (string)lo.DataBodyRange[r + 1, lo.ListColumns["description"+descType].Index].value2;
                                        v = Encoding.UTF8.GetString(Encoding.UTF8.GetBytes(lo.DataBodyRange[r + 1, lo.ListColumns["description" + descType].Index].value2));
                                        tmpDataDescription.Add(k, v);
                                    }
                                    catch
                                    {
                                        tmpDataDescription.Add(k, null);
                                    }
                                }
                            }
                        }
                    }
                }
                tmpDatatmpDataDescriptions.Add(branch, tmpDataDescription);
            }
            return tmpDatatmpDataDescriptions;
        }
        
        private string GetValidateData(string data, string dataType, Dictionary<string, Dictionary<string, string>> subgroups, string min = null, string max = null)
        {
            bool validate = true;
            string tmp = data;

            try
            {
                // 이 곳에서 유효성검증을 하자
                switch (dataType.ToUpper())
                {
                    case "UNIQUEKEY":
                    case "KEY":
                    case "UINT":
                        tmp = Convert.ToUInt32(data).ToString();
                        validate = !(!String.IsNullOrEmpty(min) && Convert.ToUInt32(min) > Convert.ToUInt32(tmp)) && !(!String.IsNullOrEmpty(max) && Convert.ToUInt32(max) < Convert.ToUInt32(tmp));
                        break;
                    case "ULONG":
                        tmp = Convert.ToUInt64(data).ToString();
                        validate = !(!String.IsNullOrEmpty(min) && Convert.ToUInt64(min) > Convert.ToUInt64(tmp)) && !(!String.IsNullOrEmpty(max) && Convert.ToUInt64(max) < Convert.ToUInt64(tmp));
                        break;
                    case "BYTE":
                        tmp = Convert.ToByte(data).ToString();
                        validate = !(!String.IsNullOrEmpty(min) && Convert.ToByte(min) > Convert.ToByte(tmp)) && !(!String.IsNullOrEmpty(max) && Convert.ToByte(max) < Convert.ToByte(tmp));
                        break;
                    case "USHORT":
                        tmp = Convert.ToUInt16(data).ToString();
                        validate = !(!String.IsNullOrEmpty(min) && Convert.ToUInt16(min) > Convert.ToUInt16(tmp)) && !(!String.IsNullOrEmpty(max) && Convert.ToUInt16(max) < Convert.ToUInt16(tmp));                        
                        break;
                    case "SHORT":
                        tmp = Convert.ToInt16(data).ToString();
                        validate = !(!String.IsNullOrEmpty(min) && Convert.ToInt16(min) > Convert.ToInt16(tmp)) && !(!String.IsNullOrEmpty(max) && Convert.ToInt16(max) < Convert.ToInt16(tmp));                        
                        break;                        
                    case "INT":
                        tmp = Convert.ToInt32(data).ToString();
                        validate = !(!String.IsNullOrEmpty(min) && Convert.ToInt32(min) > Convert.ToInt32(tmp)) && !(!String.IsNullOrEmpty(max) && Convert.ToInt32(max) < Convert.ToInt32(tmp));
                        break;
                    case "LONG":
                        tmp = Convert.ToInt64(data).ToString();
                        validate = !(!String.IsNullOrEmpty(min) && Convert.ToInt64(min) > Convert.ToInt64(tmp)) && !(!String.IsNullOrEmpty(max) && Convert.ToInt64(max) < Convert.ToInt64(tmp));
                        break;
                    case "FLOAT":
                        tmp = Convert.ToDouble(data).ToString();
                        validate = !(!String.IsNullOrEmpty(min) && Convert.ToDouble(min) > Convert.ToDouble(tmp)) && !(!String.IsNullOrEmpty(max) && Convert.ToDouble(max) < Convert.ToDouble(tmp));
                        break;
                    // 실수형은 천분율, 만분율, 백만분율로 타입에 따라 변환
                    case "FLOAT_1K":                        
                        tmp = Convert.ToInt32(Convert.ToDouble(tmp) * Properties.Settings.Default.FLOATFACTOR_1K).ToString();
                        validate = !(!String.IsNullOrEmpty(min) && Convert.ToInt32(Convert.ToDouble(min) * Properties.Settings.Default.FLOATFACTOR_1K) > Convert.ToInt32(tmp))
                                && !(!String.IsNullOrEmpty(max) && Convert.ToInt32(Convert.ToDouble(max) * Properties.Settings.Default.FLOATFACTOR_1K) < Convert.ToInt32(tmp));
                        break;
                    case "FLOAT_10K":
                        tmp = Convert.ToInt32(Convert.ToDouble(tmp) * Properties.Settings.Default.FLOATFACTOR_10K).ToString();
                        validate = !(!String.IsNullOrEmpty(min) && Convert.ToInt32(Convert.ToDouble(min) * Properties.Settings.Default.FLOATFACTOR_10K) > Convert.ToInt32(tmp))
                                && !(!String.IsNullOrEmpty(max) && Convert.ToInt32(Convert.ToDouble(max) * Properties.Settings.Default.FLOATFACTOR_10K) < Convert.ToInt32(tmp));
                        break;
                    case "FLOAT_1M":
                        tmp = Convert.ToInt32(Convert.ToDouble(tmp) * Properties.Settings.Default.FLOATFACTOR_1M).ToString();
                        validate = !(!String.IsNullOrEmpty(min) && Convert.ToInt32(Convert.ToDouble(min) * Properties.Settings.Default.FLOATFACTOR_1M) > Convert.ToInt32(tmp))
                                && !(!String.IsNullOrEmpty(max) && Convert.ToInt32(Convert.ToDouble(max) * Properties.Settings.Default.FLOATFACTOR_1M) < Convert.ToInt32(tmp));
                        break;
                    case "BOOL":
                        tmp = Convert.ToBoolean(data) ? "1" : "0";
                        break;
                    case "TEXT":
                        Convert.ToString(data);
                        break;
                    case "ARRAY":
                        tmp = Convert.ToString(data);                        
                        if (tmp.First().ToString() == Properties.Settings.Default.ARRAY_PREFIX && tmp.Last().ToString() == Properties.Settings.Default.ARRAY_POSTFIX)
                        {                            
                            // Array의 빈 자리가 있는지 체크
                            var arr = tmp.Substring(1, tmp.Length - 2);
                            var tmpArr = arr.Split(Properties.Settings.Default.ARRAY_SEPARATOR.ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
                            //validate = tmpArr.Length == arr.Split(Properties.Settings.Default.ARRAY_SEPARATOR.ToCharArray(), StringSplitOptions.None).Length;                            
                         

                            // min/max 체크(double type으로)
                            foreach(var v in tmpArr)
                            {
                                validate &= !(!String.IsNullOrEmpty(min) && Convert.ToDouble(min) > Convert.ToDouble(v)) && !(!String.IsNullOrEmpty(max) && Convert.ToDouble(max) < Convert.ToDouble(v));
                            }
                        }                        

                        break;
                    default:
                        // SUBGROUP이 존재할 때
                        if (subgroups.ContainsKey(dataType.ToUpper()))
                        {
                            validate = subgroups[dataType.ToUpper()].ContainsKey(data.ToUpper()) ? true : false;
                            tmp = data.ToUpper();
                        }
                        else
                            validate = false;
                        break;
                }
            }
            catch
            {
                validate = false;
            }

            if (!validate)
                //throw new Exception(dataType.ToUpper());
                throw new Exception(dataType.ToUpper() + " 유효성 검증실패");
            

            return tmp;
        }

        // 서브그룹명: ( 각 서브그룹 : 인덱스;설명 ) 으로 구성된 서브그룹 리스트구하기
        private Dictionary<string, Dictionary<string, string>> GetSubgroups()
        {
            Dictionary<string, Dictionary<string, string>> tmpSubgroups = new Dictionary<string, Dictionary<string, string>>();

            foreach (Excel.Worksheet ws in Globals.IG_PlanAddIn.Application.Worksheets)
            {
                foreach (Excel.ListObject lo in ws.ListObjects)
                {
                    if (lo.Name.Equals(Properties.Settings.Default.BranchSubgroupName))
                    {
                        string k, v, v2;
                        int v1;
                        int cnt = 0;
                        for (int r = 0; r < lo.ListRows.Count; r++)
                        {
                            if (lo.DataBodyRange[r + 1, lo.ListColumns["subgroups"].Index].value2 != null)
                            {
                                //k = ((string)lo.DataBodyRange[r + 1, lo.ListColumns["subgroups"].Index].value2).ToUpper().Replace("_", "");
                                k = ((string)lo.DataBodyRange[r + 1, lo.ListColumns["subgroups"].Index].value2).ToUpper();
                                v = (string)lo.DataBodyRange[r + 1, lo.ListColumns["subgroup_name"].Index].value2;

                                try
                                {
                                    if (!tmpSubgroups.ContainsKey(k))
                                        tmpSubgroups.Add(k, new Dictionary<string, string>());
                                    
                                    object tmp = lo.DataBodyRange[r + 1, lo.ListColumns["subgroup_value"].Index].value2;
                                    if (tmp == null || !(Convert.GetTypeCode(tmp) == TypeCode.Int32 || Convert.GetTypeCode(tmp) == TypeCode.Double))
                                        v1 = cnt;
                                    else
                                        v1 = Convert.ToInt32(tmp);

                                    cnt++;
                                    v2 = (string)lo.DataBodyRange[r + 1, lo.ListColumns["subgroup_desc"].Index].value2;
                                    //if (!tmpSubgroups[k].ContainsKey(v.ToUpper().Replace("_", "")))
                                    tmpSubgroups[k].Add(v.ToUpper()/*.Replace("_", "")*/, v1 + Properties.Settings.Default.SubgroupSeperator.ToString() + v2);
                                    
                                }
                                catch (ArgumentException)
                                {
                                    throw new ArgumentException(String.Format("[Subgroup에 중복값이 존재합니다.: {0} -> {1}", k.ToUpper(), v.ToUpper()));
                                }
                                catch
                                {
                                }
                            }
                        }
                    }
                }
            }

            return tmpSubgroups;
        }

        // xml 특수문자 처리
        private string ToXmlString(string str)
        {
            // XML용 특수문자 처리용
            // json은 별도 요청이 없는 한 처리하지 않음
            /*
            StringBuilder sw = new StringBuilder();
                        
            foreach (char c in str)
            {
                switch (c)
                {
                    case '<':
                        sw.Append("&lt;");
                        break;
                    case '>':
                        sw.Append("&gt;");
                        break;
                    case '&':                        
                        sw.Append("&amp;");
                        break;
                    case '\':
                        sw.Append(String.Empty);
                        break;
                    default:
                        sw.Append(c);
                        break;
                }
        
            }
             */

            return str;// sw.ToString();             
        }

        private string[] GetTableInfo()
        {
            string[] tmp = null;
            string[] result = null;
            int cnt = 0;

            foreach (Excel.Worksheet wb in Globals.IG_PlanAddIn.Application.Worksheets)
            {
                foreach (Excel.ListObject lo in wb.ListObjects)
                {
                    if (lo.Name.Length >= Properties.Settings.Default.TableInfoName.Length && lo.Name.Substring(0, Properties.Settings.Default.TableInfoName.Length).Equals(Properties.Settings.Default.TableInfoName))
                    {
                        tmp = new string[lo.ListRows.Count];

                        for (int r = 1; r <= lo.ListRows.Count; r++)
                        {
                            object contents = lo.DataBodyRange[r, lo.ListColumns["contentsid"].Index].value2;
                            object reference = lo.DataBodyRange[r, lo.ListColumns["reference"].Index].value2;

                            try
                            {
                                if ((bool)reference == true)
                                {
                                    tmp[cnt] = Convert.ToString(contents);
                                    cnt++;
                                }
                            }
                            catch
                            {
                                if (contents != null && ((string)contents).Length > 0)
                                {
                                    tmp[cnt] = Convert.ToString(reference);
                                    cnt++;
                                }
                            }                            
                        }
                        result = new string[cnt];
                        for (int i = 0; i < result.Length; i++)
                            result[i] = tmp[i];
                        break;
                    }
                }
            }

            return result;
        }        
        #endregion

        #region 메타테이블 처리 코드
        internal void WriteMetaTableMiniJson()
        {
            File.WriteAllText(Path.GetFullPath(Properties.Settings.Default.MetaTablePath + Path.DirectorySeparatorChar + MetaTable.MetaMiniJsonName),
                MetaTable.MetaMiniJsonContent, Encoding.UTF8);
        }

        internal void WriteMetaTableAbKVType()
        {
            File.WriteAllText(Path.GetFullPath(Properties.Settings.Default.MetaTablePath + Path.DirectorySeparatorChar + MetaTable.MetaAbKVTypeName),
                MetaTable.MetaAbKVTypeContent, Encoding.UTF8);
        }

        internal void WriteMetaTableObject()
        {
            File.WriteAllText(Path.GetFullPath(Properties.Settings.Default.MetaTablePath + Path.DirectorySeparatorChar + MetaTable.MetaObjectName),
                MetaTable.MetaObjectContent, Encoding.UTF8);
        }

        internal void WriteMetaTableUtil()
        {
            File.WriteAllText(Path.GetFullPath(Properties.Settings.Default.MetaTablePath + Path.DirectorySeparatorChar + MetaTable.MetaUtilName),
                MetaTable.MetaUtilContent, Encoding.UTF8);
        }

        internal void WriteMetaTableTable(string branch)
        {
            StringBuilder sb = new StringBuilder();
            string filePath = null;

            // 각종 테이블정보 초기화
            InitiateInfo();
                        
            IG_Table table = new IG_Table(GetTableName());
            if (table == null) throw new Exception("테이블명이 제대로 설정되지 않았습니다.");

            Dictionary<string, string> define;
            Dictionary<string, string> dataType;
            Dictionary<string, string> description;
            //Dictionary<string, List<string>> subgroups;

            if (!BranchDefines.ContainsKey(branch)) throw new Exception("[" + branch + "] 브랜치 설정이 존재하지 않습니다.");

            define = BranchDefines[branch];
            dataType = BranchDataTypes[branch];
            description = BranchDataDescriptions[branch];
            //subgroups = GetSubgroups();

            sb.Append(MetaTable.MetaTableContent[0]);

            sb.Append(
@"	public	class	" + Properties.Settings.Default.MetaTable_Prefix + table.Name + @" : " + Properties.Settings.Default.MetaTable_Prefix + @"Object {
		public	" + Properties.Settings.Default.MetaTable_Prefix + table.Name + @"() {}");
            sb.AppendLine();

            foreach (string k in dataType.Keys)
            {
                sb.Append(@"		public ");
                switch(dataType[k].ToUpper())
                {
                    case "INTEGER":
                    case "BOOL":    // 논리값은 1 or 0으로 출력
                    case "FLOAT":   // 실수값은 천분율 정수값으로 출력 
                        sb.Append("int " + k);
                        break;
                    case "TEXT":
                        sb.Append("string " + k);
                        break;
                    default:
                        // SUBGROUP이 존재할 때
                        if (dataType.ContainsKey(k))
                            //sb.Append(Properties.Settings.Default.MetaTableSubgroup_Prefix + dataType[k].ToUpper() + " " + k);
                            sb.Append("int " + k);
                        else
                            throw new Exception(String.Format("[{0} -> {1}]", k, dataType[k]));
                        break;
                }
                sb.Append(@"{ get; set; }");

                if (!string.IsNullOrEmpty(description[k]))
                    sb.Append("\t" + @"// " + description[k]);

                sb.AppendLine();
            }
            sb.Append(MetaTable.MetaTableContent[1]);
            sb.AppendLine();

            // GetStatByType 메소드가 필요한 테이블일 경우, 메소드 추가
            if (MetaTable.AdditionalMethods.Contains(table.Name))
            {
                sb.Append(MetaTable.MetaTableContent[2]);
                sb.AppendLine();
            }

            sb.Append(
@"		public void Map(Dictionary<string, object> dic)
		{
");
            foreach(string k in dataType.Keys)
            {
                sb.Append(@"			");
                switch (dataType[k].ToUpper())
                {
                    case "INTEGER":
                    case "BOOL":    // 논리값은 1 or 0으로 출력
                    case "FLOAT":   // 실수값은 천분율 정수값으로 출력 
                        sb.Append(k + @" = (int)(long)dic[""" + k + @"""];");                        
                        break;
                    case "TEXT":
                        sb.Append(k + @" = (string)dic[""" + k + @"""];");      
                        break;
                    default:
                        // SUBGROUP이 존재할 때
                        if (dataType.ContainsKey(k))
                            sb.Append(k + @" = (int)(long)dic[""" + k + @"""];");    
                        else
                            throw new Exception(String.Format("[{0} -> {1}]", k, dataType[k]));
                        break;
                }
                sb.AppendLine();
            }

            sb.Append(@"		}");
            sb.AppendLine();
            sb.Append(MetaTable.MetaTableContent[3]);

            filePath = Path.GetFullPath(Properties.Settings.Default.MetaTablePath + Path.DirectorySeparatorChar + Properties.Settings.Default.MetaTable_Prefix + table.Name + MetaTable.MetaTableName);
            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);

            SVNDiff(filePath);
        }
        
        internal void WriteMetaTableType(string branch)
        {
            StringBuilder sb = new StringBuilder();
            string filePath = null;

            // 각종 테이블정보 초기화
            InitiateInfo();

            IG_Table table = new IG_Table(GetTableName());
            if (table == null) throw new Exception("테이블명이 제대로 설정되지 않았습니다.");

            Dictionary<string, string> define;
            Dictionary<string, string> dataType;
            Dictionary<string, Dictionary<string, string>> subgroups;

            if (!BranchDefines.ContainsKey(branch)) throw new Exception("[" + branch + "] 브랜치 설정이 존재하지 않습니다.");

            define = BranchDefines[branch];
            dataType = BranchDataTypes[branch];
            subgroups = GetSubgroups();

            // Subgroup마다 파일 생성
            foreach (string field in subgroups.Keys)
            {
                string type = Properties.Settings.Default.MetaTableSubgroup_Prefix + field;

                // 생성자에서 초기값 설정
                var defaultValue = subgroups[field].Keys.FirstOrDefault();

                sb.Append(MetaTable.MetaTypeContent[0]);
                sb.AppendLine();

                sb.Append(
@"	public	class	" + type + @" : " + Path.GetFileNameWithoutExtension(MetaTable.MetaAbKVTypeName) + @"<int, string, " + type + @">
	{
		public " + type + @"() { " + @"var obj = " + type + @"." + defaultValue + @"; }" + @"
		protected " + type + @"(string fieldName, int k, string v)
			: base(fieldName, k, v) {	typeName = """";}"); // 일단 typeName은 빈칸: 입력받는 곳이 없음
                sb.AppendLine();

                foreach (string k in subgroups[field].Keys)
                {
                    string[] typeValue = subgroups[field][k].Split(Properties.Settings.Default.SubgroupSeperator);
                    sb.Append(
@"		public static " + type + @" " + k + @" = new " + type + @"(""" + k + @""", " + typeValue[0] + @", """ + typeValue[1] + @""");");
                    sb.AppendLine();
                }

                sb.Append(
@"		public static " + type + @" Get" + Properties.Settings.Default.MetaTableSubgroup_Prefix + @"Type(int key)
		{
			" + type + @" result = null;");
                sb.AppendLine();

                sb.Append(MetaTable.MetaTypeContent[1]);
                sb.AppendLine();

                filePath = Path.GetFullPath(Properties.Settings.Default.MetaTablePath + Path.DirectorySeparatorChar + type + MetaTable.MetaTableName);
                File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
                SVNDiff(filePath);

                sb.Clear();
            }
        }

        internal void WriteMetaTableManager(string branch)
        {
            StringBuilder sb = new StringBuilder();
            List<string> tables = new List<string>();
            List<string> subgroups = new List<string>();

            string filePath = null;

            DirectoryInfo di = new DirectoryInfo(Properties.Settings.Default.MetaTablePath);

            // 테이블 및 서브그룹 파일 검색
            foreach (FileInfo fi in di.GetFiles())
            {
                string tmp = Path.GetFileNameWithoutExtension(fi.Name);

                // 기본 생성파일 및 Type파일 제외하고 나머지 테이블이름 가져오기
                if (fi.Extension.Equals(".cs") && tmp.Length >= Properties.Settings.Default.MetaTable_Prefix.Length &&
                    tmp.Substring(0, Properties.Settings.Default.MetaTable_Prefix.Length).Equals(Properties.Settings.Default.MetaTable_Prefix) &&
                    !tmp.Equals(Path.GetFileNameWithoutExtension(MetaTable.MetaAbKVTypeName)) &&
                    !tmp.Equals(Path.GetFileNameWithoutExtension(MetaTable.MetaObjectName)) &&
                    !tmp.Equals(Path.GetFileNameWithoutExtension(MetaTable.MetaManagerName)) &&
                    !tmp.Equals(Path.GetFileNameWithoutExtension(MetaTable.MetaUtilName)))
                    if (tmp.Length >= Properties.Settings.Default.MetaTableSubgroup_Prefix.Length && tmp.Substring(0, Properties.Settings.Default.MetaTableSubgroup_Prefix.Length).Equals(Properties.Settings.Default.MetaTableSubgroup_Prefix))
                        subgroups.Add(tmp);
                    else
                        tables.Add(tmp);

                /*
                if (fi.Extension.Equals(".cs") && tmp.Length >= Properties.Settings.Default.MetaTable_Prefix.Length &&
                    tmp.Substring(0, Properties.Settings.Default.MetaTable_Prefix.Length).Equals(Properties.Settings.Default.MetaTable_Prefix))
                    if (!tmp.Equals(Path.GetFileNameWithoutExtension(MetaTable.MetaAbKVTypeName)) &&
                        !tmp.Equals(Path.GetFileNameWithoutExtension(MetaTable.MetaObjectName)) &&
                        !tmp.Equals(Path.GetFileNameWithoutExtension(MetaTable.MetaManagerName)) &&
                        !tmp.Equals(Path.GetFileNameWithoutExtension(MetaTable.MetaUtilName)) &&
                        !(tmp.Length >= Properties.Settings.Default.MetaTableSubgroup_Prefix.Length && tmp.Substring(0, Properties.Settings.Default.MetaTableSubgroup_Prefix.Length).Equals(Properties.Settings.Default.MetaTableSubgroup_Prefix)))
                        tables.Add(tmp);
                 */

            }

            sb.Append(MetaTable.MetaManagerContent[0]);
            sb.AppendLine();

            // 타입 초기화 함수
            sb.Append(
@"        private void IGTypeInit()
        {");
            sb.AppendLine();
            
            foreach (string subgroup in subgroups)
            {
                sb.Append(
@"            { var obj = new " + subgroup + @"(); }");
                sb.AppendLine();
            }
            sb.Append(
@"        }");
            sb.AppendLine();

            // 테이블을 속성으로 추가함
            sb.Append(
@"		// 속성리스트 : ");
            sb.AppendLine();
            foreach(string table in tables)
            {
                sb.AppendLine(
@"        protected " + Properties.Settings.Default.MetaTable_Prefix + @"Container<" + table + @">	" + table.Replace(Properties.Settings.Default.MetaTable_Prefix, Properties.Settings.Default.MetaTable_Prefix.ToLower()) + @";
        public " + Properties.Settings.Default.MetaTable_Prefix + @"Container<" + table + @">	" + table + @"	{ get { return " + table.Replace(Properties.Settings.Default.MetaTable_Prefix, Properties.Settings.Default.MetaTable_Prefix.ToLower()) + @"; } }");
                sb.AppendLine();
            }

            sb.Append(MetaTable.MetaManagerContent[1]);
            sb.AppendLine();

            sb.Append(
@"
        public void ClearContainer()
        {");
            sb.AppendLine();
            foreach (string table in tables)
            {
                sb.Append(
@"            " + table.Replace(Properties.Settings.Default.MetaTable_Prefix, Properties.Settings.Default.MetaTable_Prefix.ToLower()) + @".Clear();");
                sb.AppendLine();
            }
            sb.Append(
@"        }");
            sb.AppendLine();

            sb.Append(MetaTable.MetaManagerContent[2]);
            sb.AppendLine();

            sb.Append(
@"            switch (className)
            {");
            sb.AppendLine();
            foreach(string table in tables)
            {
                sb.Append(@"                case """ + table + @""":");
                sb.AppendLine();
                sb.Append(@"#if USE_LITJSON");
                sb.AppendLine();
                sb.Append(@"                    " + table.Replace(Properties.Settings.Default.MetaTable_Prefix, Properties.Settings.Default.MetaTable_Prefix.ToLower()) + " = " + Properties.Settings.Default.MetaTable_Prefix + @"Util.LoadJson2" + Properties.Settings.Default.MetaTable_Prefix + @"Container<" + table + @">(js, className);break;");                
                sb.AppendLine();
                sb.Append(@"#else
                    " + table.Replace(Properties.Settings.Default.MetaTable_Prefix, Properties.Settings.Default.MetaTable_Prefix.ToLower()) + " = " + Properties.Settings.Default.MetaTable_Prefix + @"Util.LoadMiniJson2GDContainer<" + table + @">(jsonDic, className);break;
#endif");
                sb.AppendLine();
            }
            sb.Append(
@"            }");
            sb.AppendLine();

            sb.Append(MetaTable.MetaManagerContent[3]);
            sb.AppendLine();

            filePath = Path.GetFullPath(Properties.Settings.Default.MetaTablePath + Path.DirectorySeparatorChar + MetaTable.MetaManagerName);
            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);

            SVNDiff(filePath);
        }
        #endregion

        #region VSTO에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다.
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마십시오.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion        
    }
}
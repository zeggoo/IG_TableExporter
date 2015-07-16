﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Newtonsoft.Json;

namespace IG_TableExporter
{
    // 테이블 구조 정의
    // 나중에 json.net의 serialize 적용해보든지 하자
    public class IG_StageNote
    {
        private int length;
        private int count;

        private int point;
        private int gold;

        private StringBuilder sb;
        private StringWriter sw;
        private JsonTextWriter json;

        public int Length
        {
            get
            {
                return this.length;
            }
        }

        public int Count
        {
            get
            {
                return this.count;
            }
        }

        public int Point
        {
            get
            {
                return this.point;
            }
        }

        public int Gold
        {
            get
            {
                return this.gold;
            }
        }

        public IG_StageNote(int length, int point, int gold)
        {
            this.length = length;
            this.point = point;
            this.gold = gold;
            this.count = 0;

            sb = new StringBuilder();
            sw = new StringWriter(sb);
            json = new JsonTextWriter(sw);
            json.Formatting = Formatting.Indented;
            json.WriteStartObject();

            InitiateNote();
        }

        private void InitiateNote()
        {
            // 스테이지노트 메타정보 추가
            json.WritePropertyName("Length");
            json.WriteRawValue((this.length / 10f).ToString());

            json.WritePropertyName("Point");
            json.WriteRawValue(this.point.ToString());

            json.WritePropertyName("Gold");
            json.WriteRawValue(this.gold.ToString());
        }

        public void StartAdd(int key)
        {
            this.count++;
            json.WritePropertyName(Convert.ToString(key));

            json.WriteStartArray();
            //json.WriteStartObject();
        }

        public void EndAdd()
        {
            json.WriteEndArray();
            //json.WriteEnd();
            //json.WriteEndObject();
        }

        public void AddElement(List<Tuple<int, int, float>> element)
        {
            for (int i = 0; i < element.Count; i++)
            //foreach (int k in element.Keys)
            {
                json.WriteStartObject();

                // 몬스터인덱스 기입
                json.WritePropertyName(Properties.Settings.Default.SpawnPropertyName);
                json.WriteValue(Convert.ToString(element.ElementAt(i).Item1));

                // 등장확률 기입
                json.WritePropertyName(Properties.Settings.Default.ProbPropertyName);
                json.WriteValue(Convert.ToString(element.ElementAt(i).Item2));

                // 다음 등장까지 딜레이 시간 기입
                json.WritePropertyName(Properties.Settings.Default.NextTimeName);
                json.WriteValue(Convert.ToString(element.ElementAt(i).Item3));

                json.WriteEndObject();
            }
        }

        public override string ToString()
        {
            json.WriteEndObject();
            return sb.ToString();
        }


        internal void StartNote(int nextRound)
        {
            StartAdd(0);
            var firstList = new List<Tuple<int, int, float>>();
            firstList.Add(new Tuple<int, int, float>(0, 0, (float)nextRound / 10f));
            AddElement(firstList);
            EndAdd();
        }
    }
}

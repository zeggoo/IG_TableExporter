using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Excel;

namespace IG_TableExporter
{
    public partial class MonsterInfoTask : UserControl
    {
        private string stageName;
        private Dictionary<int, int> monsterIndexs;

        public string StageName
        {
            get
            {
                return this.stageName;
            }
        }

        public MonsterInfoTask()
        {
            InitializeComponent();
        }

        public MonsterInfoTask(string name)
        {
            this.stageName = name;            
            InitializeComponent();
        }

        private void MonsterInfoTask_Load(object sender, EventArgs e)
        {            
        }

        internal void MonsterInfoRefresh(List<MonsterInfo> monsterInfos, string stage)
        {
            int cnt = 0;

            monsterDataGridView.Rows.Clear();
            monsterIndexs = new Dictionary<int, int>();

            foreach (MonsterInfo info in monsterInfos)
                if (info.stage == null || info.stage == "" || info.stage.Trim() == stage.Trim()) 
                {
                    monsterIndexs.Add(cnt, info.index);
                    monsterDataGridView.Rows.Add(new object[9]
                    {
                        Convert.ToString(info.index),   
                        ( Globals.IG_PlanAddIn.MonsterSpritePaths.ContainsKey(info.sprite) ?
                            Image.FromFile(Globals.IG_PlanAddIn.MonsterSpritePaths[info.sprite]) : null
                        ),
                        String.Format("{0:P0}", info.speed),
                        String.Format("{0:P0}", info.scale),
                        String.Format("{0:N0}", info.HP),
                        Convert.ToString(info.atk),                                                  
                        Convert.ToString(info.type),
                        String.Format("{0:N0}", info.point),
                        String.Format("{0:N0}", info.GetGold())
                    });

                    // 스테이지 전용 몬스터는 볼드체
                    if (info.stage.Trim() == stage.Trim())
                    {
                        monsterDataGridView.Rows[cnt].Cells[0].Style.Font = new Font(monsterDataGridView.Font, FontStyle.Bold);
                        monsterDataGridView.Rows[cnt].Cells[2].Style.Font = new Font(monsterDataGridView.Font, FontStyle.Bold);
                        monsterDataGridView.Rows[cnt].Cells[3].Style.Font = new Font(monsterDataGridView.Font, FontStyle.Bold);
                        monsterDataGridView.Rows[cnt].Cells[4].Style.Font = new Font(monsterDataGridView.Font, FontStyle.Bold);
                        monsterDataGridView.Rows[cnt].Cells[5].Style.Font = new Font(monsterDataGridView.Font, FontStyle.Bold);
                        monsterDataGridView.Rows[cnt].Cells[6].Style.Font = new Font(monsterDataGridView.Font, FontStyle.Bold);
                        monsterDataGridView.Rows[cnt].Cells[7].Style.Font = new Font(monsterDataGridView.Font, FontStyle.Bold);
                        monsterDataGridView.Rows[cnt].Cells[8].Style.Font = new Font(monsterDataGridView.Font, FontStyle.Bold);
                    }

                    // 몬스터 타입별로 색깔입히기
                    //monsterDataGridView.Rows[cnt].DefaultCellStyle.BackColor = Globals.IG_PlanAddIn.GetMonsterTypeColor(info.type, info.scale);
                    monsterDataGridView.Rows[cnt].DefaultCellStyle.BackColor = info.color;

                    cnt++;
                }
            //monsterDataGridView.Update();
        }

        private void monsterDataGridView_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                //Globals.IG_PlanAddIn.InsertMonsterIndex(monsterIndexs[e.RowIndex]);
                Globals.IG_PlanAddIn.InsertMonsterIndex(Convert.ToInt32(monsterDataGridView.Rows[e.RowIndex].Cells[0].Value));
            }
            catch
            {
                MessageBox.Show("몬스터 인덱스 기입과정에서 오류가 발생하였습니다.");
            }
        }

        private void monsterDataGridView_KeyUp(object sender, System.Windows.Forms.KeyEventArgs e)
        {
            if (monsterDataGridView.SelectedRows != null)
                if (e.Control && e.KeyCode == Keys.C)
                {
                    Clipboard.Clear();
                    Clipboard.SetText(Convert.ToString(monsterDataGridView.Rows[monsterDataGridView.SelectedRows[0].Index].Cells[0].Value));
                }

        }
    }
}


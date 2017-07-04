using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Deployment.Application;
using System.Diagnostics;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

namespace IG_TableExporter
{
    public partial class IG_Ribbon
    {
        private bool branchLoaded;
        private string branch;

        // TASK PANE 테스트
        private MonsterInfoTask monsterInfoTask;
        private Microsoft.Office.Tools.CustomTaskPane monsterInfoTaskPane;

        private void IG_Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            branchLoaded = false;
            btnExportTable.Enabled = false;
            btnExportMonsterTable.Enabled = false;
            btnExportMetaTable.Enabled = false;
            btnExportVerification.Enabled = false;
            //IG_TableGroup.Label = FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location).ProductVersion;
            branch = null;
        }

        private void btnExport_Click(object sender, RibbonControlEventArgs e)
        {
            saveTableFileDialog.Filter = "JSON Data|*.json";
            saveTableFileDialog.Title = "Export Table";

            var tableName = Globals.IG_PlanAddIn.GetTableName();
            if (String.IsNullOrEmpty(tableName))
                saveTableFileDialog.FileName = System.IO.Path.GetFileNameWithoutExtension(Globals.IG_PlanAddIn.Application.ActiveWorkbook.Name);
            else
                saveTableFileDialog.FileName = tableName;

            saveTableFileDialog.DefaultExt = "json";
            saveTableFileDialog.ShowDialog();
        }

        private void saveTableFileDialog_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {
            try
            {
                Stopwatch sw = new Stopwatch();                
                sw.Start();

                Globals.IG_PlanAddIn.WriteTable(saveTableFileDialog.FileName, Globals.IG_PlanAddIn.ExportTable(branch));                
                sw.Stop();

                System.Windows.Forms.MessageBox.Show("[" + Path.GetFileName(saveTableFileDialog.FileName) + "] 저장완료 (수행시간: " + sw.ElapsedMilliseconds.ToString() + "msec)");

                // 강제 diff
                Globals.IG_PlanAddIn.SVNDiff(saveTableFileDialog.FileName);
            }
            catch (Exception except)
            {                
                MessageBox.Show(except.Message, "테이블 추출 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnExportVerification_Click(object sender, RibbonControlEventArgs e)
        {
            var tableName = Globals.IG_PlanAddIn.GetTableName();
            var fileName = Globals.IG_PlanAddIn.Application.ActiveWorkbook.Path + Path.DirectorySeparatorChar +
                                    Properties.Settings.Default.DefinePath + Properties.Settings.Default.DefinePrefix + tableName + "." + Properties.Settings.Default.DefineExt;
            if (String.IsNullOrEmpty(tableName))
                saveTableFileDialog.FileName = System.IO.Path.GetFileNameWithoutExtension(Globals.IG_PlanAddIn.Application.ActiveWorkbook.Name);
            else
                saveTableFileDialog.FileName = tableName;

            // 테이블유효성 검증을 위한 json 파일 추출
            try
            {
                Globals.IG_PlanAddIn.WriteDefine(fileName, Globals.IG_PlanAddIn.ExportDefine(branch));
                MessageBox.Show(Properties.Settings.Default.DefinePrefix + tableName + "을 추출하였습니다.", "define 추출 성공", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception except)
            {
                MessageBox.Show(except.Message, "define 추출 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            // 강제 diff
            Globals.IG_PlanAddIn.SVNDiff(fileName);
        }

        private void branchComboBox_TextChanged(object sender, RibbonControlEventArgs e)
        {
            btnExport_Refresh();
            btnExportMonsterTable_Refresh();
            btnExportMetaTable_Refresh();
            btnExportVerification_Refresh();
            branch = branchComboBox.Text;
        }

        #region 브랜치 정보 불러오기
        private void branchComboBox_Refresh()
        {
            branchComboBox.Items.Clear();            

            Globals.IG_PlanAddIn.InitiateInfo();

            try
            {
                foreach (string item in Globals.IG_PlanAddIn.BranchList)
                {
                    RibbonDropDownItem i = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                    i.Label = item;                    
                    branchComboBox.Items.Add(i);
                    branchComboBox.Text = branchComboBox.Items[0].Label;
                    branch = branchComboBox.Text;           
                }
                this.branchLoaded = true;

            }
            catch
            {
                this.branchLoaded = false;
                branchComboBox.Text = null;                
                branch = branchComboBox.Text;                
                throw new Exception("브랜치 정보가 존재하지 않습니다.");
            }
            finally
            {
                btnExport_Refresh();
                btnExportVerification_Refresh();
                btnExportMonsterTable_Refresh();
                btnExportMetaTable_Refresh();
            }
        }

        private void btnExport_Refresh()
        {
            if (branchLoaded && Globals.IG_PlanAddIn.BranchList.Contains(branchComboBox.Text))
                btnExportTable.Enabled = true;
            else
                btnExportTable.Enabled = false;
        }

        private void btnExportVerification_Refresh()
        {
            if (branchLoaded && Globals.IG_PlanAddIn.BranchList.Contains(branchComboBox.Text))
                btnExportVerification.Enabled = true;
            else
                btnExportVerification.Enabled = false;
        }

        private void btnExportMonsterTable_Refresh()
        {
            if (branchLoaded && Globals.IG_PlanAddIn.BranchList.Contains(branchComboBox.Text))
                btnExportMonsterTable.Enabled = true;
            else
                btnExportMonsterTable.Enabled = false;
        }

        private void btnExportMetaTable_Refresh()
        {
            if (branchLoaded && Globals.IG_PlanAddIn.BranchList.Contains(branchComboBox.Text))
                btnExportMetaTable.Enabled = true;
            else
                btnExportMetaTable.Enabled = false;
        }
        #endregion

        private void btnCheckBranch_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                this.branchComboBox_Refresh();
            }
            catch (Exception except)
            {
                MessageBox.Show(except.Message, "테이블 추출 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnExportStageNote_Click(object sender, RibbonControlEventArgs e)
        {
            saveNoteFileDialog.Filter = "JSON Data|*.json";
            saveNoteFileDialog.Title = "Export StageNote";

            saveNoteFileDialog.FileName = System.IO.Path.GetFileNameWithoutExtension(Globals.IG_PlanAddIn.Application.ActiveWorkbook.Name);
            saveNoteFileDialog.DefaultExt = "json";
            saveNoteFileDialog.ShowDialog();
        }

        private void saveNoteFileDialog_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {
            try
            {
                Stopwatch sw = new Stopwatch();
                sw.Start();

                Globals.IG_PlanAddIn.WriteStageNote(saveNoteFileDialog.FileName, Globals.IG_PlanAddIn.ExportNote());
                sw.Stop();

                System.Windows.Forms.MessageBox.Show("[" + Path.GetFileName(saveNoteFileDialog.FileName) + "] 저장완료 (수행시간: " + sw.ElapsedMilliseconds.ToString() + "msec)");

                // 강제 diff
                Globals.IG_PlanAddIn.SVNDiff(saveNoteFileDialog.FileName);
            }
            catch (Exception except)
            {
                MessageBox.Show(except.Message, "노트 추출 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnVerifyMonsters_Click(object sender, RibbonControlEventArgs e)
        {
            List<MonsterInfo> monsterInfo;
            Dictionary<int, string> spriteNames;
            string stageName;

            Globals.IG_PlanAddIn.InitiateMonsterInfo();

            try
            {
                Globals.IG_PlanAddIn.Application.ScreenUpdating = false;
                Globals.IG_PlanAddIn.Application.Calculation = Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationManual;

                //Stopwatch sw = new Stopwatch();
                //sw.Start();

                // 스테이지 인덱스 불러오기
                stageName = Globals.IG_PlanAddIn.GetStageName();

                // 몬스터 스프라이트 불러오기
                spriteNames = Globals.IG_PlanAddIn.GetMonsterSprite();
                
                // 몬스터 데이터 불러오기                
                monsterInfo = Globals.IG_PlanAddIn.GetMonsterInfos(stageName, spriteNames);
                
                // 밸런스문서에서 처리함
                // 몬스터데이터 표에 정보 기입하기
                //cnt = Globals.IG_PlanAddIn.RefreshMonsterInfoTable(monsterInfo, stageName);

                // 스테이지노트의 Spawn 필드를 검색하여 몬스터타입별로 색깔을 새로 지정함
                Globals.IG_PlanAddIn.RefreshMonsterColor(monsterInfo);

                // TASK PANE 테스트
                if (monsterInfoTask == null || monsterInfoTask.StageName != stageName)
                {
                    monsterInfoTask = new MonsterInfoTask(stageName);
                    monsterInfoTaskPane = Globals.IG_PlanAddIn.CustomTaskPanes.Add(monsterInfoTask, "몬스터정보");
                }
                monsterInfoTask.MonsterInfoRefresh(monsterInfo, stageName);

                //sw.Stop();

                // TASK PANE 테스트
                // System.Windows.Forms.MessageBox.Show("몬스터 정보 불러오기 완료: " + cnt + "마리(수행시간: " + sw.ElapsedMilliseconds.ToString() + "msec)");

                // TASK PANE 테스트
                monsterInfoTaskPane.Width = monsterInfoTask.monsterDataGridView.Width;
                //monsterInfoTaskPane.Height = monsterInfoTask.monsterDataGridView.Height;
                monsterInfoTaskPane.Visible = true;

            }

            // 파일 패스가 없거나 잘못 되었을 때, 새로운 정보불러오기
            catch (IOException)
            {
                if (!Globals.IG_PlanAddIn.mInfoPath.monsterTable)
                    MessageBox.Show("몬스터 테이블을 찾을 수 없습니다.");
                else if (!Globals.IG_PlanAddIn.mInfoPath.resourcePathTable)
                    MessageBox.Show("리소스패스 테이블을 찾을 수 없습니다.");
                else if (!Globals.IG_PlanAddIn.mInfoPath.stageTable)
                    MessageBox.Show("스테이지 테이블을 찾을 수 없습니다.");
                else if (!Globals.IG_PlanAddIn.mInfoPath.monsterSprite)
                    MessageBox.Show("몬스터 스프라이트가 저장된 폴더를 찾을 수 없습니다.");
            }
            catch (Exception except)
            {
                MessageBox.Show(except.Message, "몬스터 정보 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Properties.Settings.Default.Save();
                Globals.IG_PlanAddIn.Application.Calculation = Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationAutomatic;
                Globals.IG_PlanAddIn.Application.ScreenUpdating = true;
            }
        }

        private void btnVerifyResourcePaths_Click(object sender, RibbonControlEventArgs e)
        {
            int notMatchCnt = 0;
            try
            {
                Globals.IG_PlanAddIn.Application.ScreenUpdating = false;
                Globals.IG_PlanAddIn.Application.Calculation = Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationManual;

                notMatchCnt = Globals.IG_PlanAddIn.VerifyResourcePathTable();

                if (notMatchCnt > 0)
                    MessageBox.Show(notMatchCnt + "개의 패스가 일치하지 않습니다.");
                else
                    MessageBox.Show("모든 리소스패스가 검증되었습니다.");
            }
            catch (FileNotFoundException)
            {
                MessageBox.Show("에셋폴더를 읽는 과정에서 오류가 발생하였습니다.");
            }
            catch (IOException)
            {
                MessageBox.Show("에셋폴더를 읽는 과정에서 오류가 발생하였습니다.");
            }
            catch (Exception except)
            {
                MessageBox.Show(except.Message, "리소스패스 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Properties.Settings.Default.Save();
                Globals.IG_PlanAddIn.Application.Calculation = Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationAutomatic;
                Globals.IG_PlanAddIn.Application.ScreenUpdating = true;
            }
        }

        private void monsterTablePathFileDialog_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {
            Properties.Settings.Default.MonsterTablePath = monsterTablePathFileDialog.FileName;
            Globals.IG_PlanAddIn.mInfoPath.monsterTable = true;
        }

        private void resourcePathTablePathFileDialog_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {
            Properties.Settings.Default.ResourcePathTablePath = resourcePathTablePathFileDialog.FileName;
            Globals.IG_PlanAddIn.mInfoPath.resourcePathTable = true;
        }
        
        private void stageTableFileDialog_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {
            Properties.Settings.Default.StageTablePath = stageTableFileDialog.FileName;
            Globals.IG_PlanAddIn.mInfoPath.stageTable = true;
        }

        private void btnSetResourcePathProperties_Click(object sender, RibbonControlEventArgs e)
        {
            setResourcePathProperties();
        }
        private void btnSetPathProperties_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.IG_PlanAddIn.mInfoPath.monsterTable = false;
            Globals.IG_PlanAddIn.mInfoPath.resourcePathTable = false;
            Globals.IG_PlanAddIn.mInfoPath.monsterSprite = false;
            Globals.IG_PlanAddIn.mInfoPath.stageTable = false;

            setMonsterPathProperties();
        }

        private void setMonsterPathProperties()
        {
            //if (!Globals.IG_PlanAddIn.mInfoPath.monsterTable)
            //{
            //    monsterTablePathFileDialog.Filter = "JSON Data|*.json";
            //    monsterTablePathFileDialog.Title = "Monster Table";
            //    monsterTablePathFileDialog.FileName = "MonsterTable";
            //    monsterTablePathFileDialog.DefaultExt = "json";
            //    monsterTablePathFileDialog.ShowDialog();
            //}

            // 쪼개진 몬스터테이블로 인해 테이블 대신 폴더를 지정함
            if (!Globals.IG_PlanAddIn.mInfoPath.monsterTable)
            {
                monsterTablePathBrowserDialog.ShowDialog();
                Properties.Settings.Default.MonsterTablePath = monsterTablePathBrowserDialog.SelectedPath;
                Globals.IG_PlanAddIn.mInfoPath.monsterTable = true;
            }

            if (!Globals.IG_PlanAddIn.mInfoPath.resourcePathTable)
            {
                resourcePathTablePathFileDialog.Filter = "JSON Data|*.json";
                resourcePathTablePathFileDialog.Title = "ResourcePath Table";
                resourcePathTablePathFileDialog.FileName = "ResourcePathTable";
                resourcePathTablePathFileDialog.DefaultExt = "json";
                resourcePathTablePathFileDialog.ShowDialog();
            }

            if (!Globals.IG_PlanAddIn.mInfoPath.stageTable)
            {
                stageTableFileDialog.Filter = "JSON Data|*.json";
                stageTableFileDialog.Title = "Stage Table";
                stageTableFileDialog.FileName = "StageTable";
                stageTableFileDialog.DefaultExt = "json";
                stageTableFileDialog.ShowDialog();
            }

            if (!Globals.IG_PlanAddIn.mInfoPath.monsterSprite)
            {
                monsterSpritePathBrowserDialog.ShowDialog();
                Properties.Settings.Default.MonsterSpritePath = monsterSpritePathBrowserDialog.SelectedPath;
                Globals.IG_PlanAddIn.mInfoPath.monsterSprite = true;
            }
        }

        private void setResourcePathProperties()
        {
            assetPathBrowserDialog.ShowDialog();
            Properties.Settings.Default.ResourceAssetPath = assetPathBrowserDialog.SelectedPath;
        }

        private void setMonsterTablePathProperties()
        {
            setMonsterPathProperties();
        }

        private void setMetaTablePathProperties()
        {
            metaTablePathBrowserDialog.ShowDialog();
            Properties.Settings.Default.MetaTablePath = metaTablePathBrowserDialog.SelectedPath;
        }

        private void btnSetMetaTablePathProperties_Click(object sender, RibbonControlEventArgs e)
        {
            setMetaTablePathProperties();
        }

        private void btnExportMetaTable_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Globals.IG_PlanAddIn.Application.ScreenUpdating = false;
                Globals.IG_PlanAddIn.Application.Calculation = Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationManual;

                // 0. MiniJson 추출
                Globals.IG_PlanAddIn.WriteMetaTableMiniJson();

                // 1. AbKVType 추출
                Globals.IG_PlanAddIn.WriteMetaTableAbKVType();                

                // 2. Object 추출
                Globals.IG_PlanAddIn.WriteMetaTableObject();

                // 3. Util 추출
                Globals.IG_PlanAddIn.WriteMetaTableUtil();

                // 4. 이 테이블의 테이블 추출
                Globals.IG_PlanAddIn.WriteMetaTableTable(branch);

                // 5. 이 테이블의 서브그룹 추출
                Globals.IG_PlanAddIn.WriteMetaTableType(branch);

                // 6. 테이블 및 서브그룹의 매니저 추출
                Globals.IG_PlanAddIn.WriteMetaTableManager(branch);

                // 설정 패스로 탐색기 열기
                Process.Start("explorer.exe", @"/e,/root, " + Properties.Settings.Default.MetaTablePath);
            }
            catch (IOException)
            {
                MessageBox.Show("메타테이블 추출 폴더를 읽는 과정에서 오류가 발생하였습니다.");
            }
            catch (Exception except)
            {
                MessageBox.Show(except.Message, "메타테이블 추출 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Properties.Settings.Default.Save();
                Globals.IG_PlanAddIn.Application.Calculation = Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationAutomatic;
                Globals.IG_PlanAddIn.Application.ScreenUpdating = true;
            }
        }

        // 몬스터테이블 추출 전용
        private void setTablePathProperties()
        {
            tablePathBrowserDialog.ShowDialog();
            Properties.Settings.Default.TablePath = tablePathBrowserDialog.SelectedPath;
            Properties.Settings.Default.Save();
        }

        private void btnExportMonsterTable_Click(object sender, RibbonControlEventArgs e)
        {
            string[] stages = Globals.IG_PlanAddIn.GetStagesFromMonsterTable();
            if (!(stages.Length > 0))
            {
                MessageBox.Show("몬스터테이블만 추출가능합니다.");
                return;
            }

            List<string> fileNames = new List<string>();

            Array.ForEach(stages, delegate(string stage) { 
                if (stage == "Common")
                    fileNames.Add(Path.Combine(Properties.Settings.Default.TablePath, Properties.Settings.Default.MonsterTableExportName + "_" + stage + ".json"));
                else
                    fileNames.Add(Path.Combine(Properties.Settings.Default.TablePath, Properties.Settings.Default.StageMonsterTablePath, Properties.Settings.Default.MonsterTableExportName + "_" + stage + ".json"));
            });

            // 일단 패스설정부터
            setTablePathProperties();

            try
            {
                if (String.IsNullOrWhiteSpace(Properties.Settings.Default.TablePath)) throw new IOException();

                Stopwatch sw = new Stopwatch();
                sw.Start();
                Globals.IG_PlanAddIn.WriteTable(fileNames.ToArray(), Globals.IG_PlanAddIn.ExportMonsterTable(branch, stages));                
                sw.Stop();

                System.Windows.Forms.MessageBox.Show("[몬스터테이블] " + stages.Length + "개 저장완료 (수행시간: " + sw.ElapsedMilliseconds.ToString() + "msec)");
                
                // 설정 패스로 탐색기 열기
                Process.Start("explorer.exe", @"/e,/root, " + Properties.Settings.Default.TablePath);
            }
            catch (IOException)
            {
                MessageBox.Show("몬스터테이블을 추출할 폴더가 정확하지 않습니다.");
                setTablePathProperties();
            }
            catch (Exception except)
            {
                MessageBox.Show(except.Message, "몬스터테이블 추출 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}


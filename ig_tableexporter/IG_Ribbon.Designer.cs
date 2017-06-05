namespace IG_TableExporter
{
    partial class IG_Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 필수 디자이너 변수입니다.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public IG_Ribbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 사용 중인 모든 리소스를 정리합니다.
        /// </summary>
        /// <param name="disposing">관리되는 리소스를 삭제해야 하면 true이고, 그렇지 않으면 false입니다.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 구성 요소 디자이너에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다.
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마십시오.
        /// </summary>
        private void InitializeComponent()
        {
            this.IG_Tab = this.Factory.CreateRibbonTab();
            this.IG_TableGroup = this.Factory.CreateRibbonGroup();
            this.branchComboBox = this.Factory.CreateRibbonComboBox();
            this.btnCheckBranch = this.Factory.CreateRibbonButton();
            this.btnExportMonsterTable = this.Factory.CreateRibbonButton();
            this.btnExportTable = this.Factory.CreateRibbonButton();
            this.IG_MetaTableGroup = this.Factory.CreateRibbonGroup();
            this.btnSetMetaTablePathProperties = this.Factory.CreateRibbonButton();
            this.btnExportMetaTable = this.Factory.CreateRibbonButton();
            this.IG_ResourcePathGroup = this.Factory.CreateRibbonGroup();
            this.btnSetResourcePathProperties = this.Factory.CreateRibbonButton();
            this.btnVerifyResourcePaths = this.Factory.CreateRibbonButton();
            this.IG_StageNoteGroup = this.Factory.CreateRibbonGroup();
            this.btnSetNotePathProperties = this.Factory.CreateRibbonButton();
            this.btnVerifyMonsters = this.Factory.CreateRibbonButton();
            this.btnExportStageNote = this.Factory.CreateRibbonButton();
            this.saveTableFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.saveNoteFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.monsterTablePathFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.resourcePathTablePathFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.monsterSpritePathBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.stageTableFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.assetPathBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.metaTablePathBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.tablePathBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.monsterTablePathBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.IG_Tab.SuspendLayout();
            this.IG_TableGroup.SuspendLayout();
            this.IG_MetaTableGroup.SuspendLayout();
            this.IG_ResourcePathGroup.SuspendLayout();
            this.IG_StageNoteGroup.SuspendLayout();
            // 
            // IG_Tab
            // 
            this.IG_Tab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.IG_Tab.Groups.Add(this.IG_TableGroup);
            this.IG_Tab.Groups.Add(this.IG_MetaTableGroup);
            this.IG_Tab.Groups.Add(this.IG_ResourcePathGroup);
            this.IG_Tab.Groups.Add(this.IG_StageNoteGroup);
            this.IG_Tab.Label = "MagnetGames";
            this.IG_Tab.Name = "IG_Tab";
            // 
            // IG_TableGroup
            // 
            this.IG_TableGroup.Items.Add(this.branchComboBox);
            this.IG_TableGroup.Items.Add(this.btnCheckBranch);
            this.IG_TableGroup.Items.Add(this.btnExportMonsterTable);
            this.IG_TableGroup.Items.Add(this.btnExportTable);
            this.IG_TableGroup.Label = "테이블 추출";
            this.IG_TableGroup.Name = "IG_TableGroup";
            // 
            // branchComboBox
            // 
            this.branchComboBox.Label = "브랜치";
            this.branchComboBox.Name = "branchComboBox";
            this.branchComboBox.Text = null;
            this.branchComboBox.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.branchComboBox_TextChanged);
            // 
            // btnCheckBranch
            // 
            this.btnCheckBranch.Image = global::IG_TableExporter.Properties.Resources._1424878236_view_refresh_512;
            this.btnCheckBranch.Label = "브랜치 확인";
            this.btnCheckBranch.Name = "btnCheckBranch";
            this.btnCheckBranch.ShowImage = true;
            this.btnCheckBranch.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCheckBranch_Click);
            // 
            // btnExportMonsterTable
            // 
            this.btnExportMonsterTable.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnExportMonsterTable.Image = global::IG_TableExporter.Properties.Resources.CORING;
            this.btnExportMonsterTable.Label = "몬스터테이블 추출";
            this.btnExportMonsterTable.Name = "btnExportMonsterTable";
            this.btnExportMonsterTable.ShowImage = true;
            this.btnExportMonsterTable.Visible = false;
            this.btnExportMonsterTable.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExportMonsterTable_Click);
            // 
            // btnExportTable
            // 
            this.btnExportTable.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnExportTable.Image = global::IG_TableExporter.Properties.Resources._1424876079_x_office_spreadsheet_512;
            this.btnExportTable.Label = "테이블 추출";
            this.btnExportTable.Name = "btnExportTable";
            this.btnExportTable.ShowImage = true;
            this.btnExportTable.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExport_Click);
            // 
            // IG_MetaTableGroup
            // 
            this.IG_MetaTableGroup.Items.Add(this.btnSetMetaTablePathProperties);
            this.IG_MetaTableGroup.Items.Add(this.btnExportMetaTable);
            this.IG_MetaTableGroup.Label = "메타테이블";
            this.IG_MetaTableGroup.Name = "IG_MetaTableGroup";
            this.IG_MetaTableGroup.Visible = false;
            // 
            // btnSetMetaTablePathProperties
            // 
            this.btnSetMetaTablePathProperties.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSetMetaTablePathProperties.Image = global::IG_TableExporter.Properties.Resources._1426432298_Settings_5_512;
            this.btnSetMetaTablePathProperties.Label = "경로 설정";
            this.btnSetMetaTablePathProperties.Name = "btnSetMetaTablePathProperties";
            this.btnSetMetaTablePathProperties.ShowImage = true;
            this.btnSetMetaTablePathProperties.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSetMetaTablePathProperties_Click);
            // 
            // btnExportMetaTable
            // 
            this.btnExportMetaTable.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnExportMetaTable.Image = global::IG_TableExporter.Properties.Resources._1426247862_19_512;
            this.btnExportMetaTable.Label = "테이블 메타정보";
            this.btnExportMetaTable.Name = "btnExportMetaTable";
            this.btnExportMetaTable.ShowImage = true;
            this.btnExportMetaTable.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExportMetaTable_Click);
            // 
            // IG_ResourcePathGroup
            // 
            this.IG_ResourcePathGroup.Items.Add(this.btnSetResourcePathProperties);
            this.IG_ResourcePathGroup.Items.Add(this.btnVerifyResourcePaths);
            this.IG_ResourcePathGroup.Label = "리소스";
            this.IG_ResourcePathGroup.Name = "IG_ResourcePathGroup";
            this.IG_ResourcePathGroup.Visible = false;
            // 
            // btnSetResourcePathProperties
            // 
            this.btnSetResourcePathProperties.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSetResourcePathProperties.Image = global::IG_TableExporter.Properties.Resources._1425331429_wrench_5121;
            this.btnSetResourcePathProperties.Label = "경로 설정";
            this.btnSetResourcePathProperties.Name = "btnSetResourcePathProperties";
            this.btnSetResourcePathProperties.ShowImage = true;
            this.btnSetResourcePathProperties.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSetResourcePathProperties_Click);
            // 
            // btnVerifyResourcePaths
            // 
            this.btnVerifyResourcePaths.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnVerifyResourcePaths.Image = global::IG_TableExporter.Properties.Resources._1425315321_11_512;
            this.btnVerifyResourcePaths.Label = "리소스 검증";
            this.btnVerifyResourcePaths.Name = "btnVerifyResourcePaths";
            this.btnVerifyResourcePaths.ShowImage = true;
            this.btnVerifyResourcePaths.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnVerifyResourcePaths_Click);
            // 
            // IG_StageNoteGroup
            // 
            this.IG_StageNoteGroup.Items.Add(this.btnSetNotePathProperties);
            this.IG_StageNoteGroup.Items.Add(this.btnVerifyMonsters);
            this.IG_StageNoteGroup.Items.Add(this.btnExportStageNote);
            this.IG_StageNoteGroup.Label = "스테이지 노트";
            this.IG_StageNoteGroup.Name = "IG_StageNoteGroup";
            this.IG_StageNoteGroup.Visible = false;
            // 
            // btnSetNotePathProperties
            // 
            this.btnSetNotePathProperties.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSetNotePathProperties.Image = global::IG_TableExporter.Properties.Resources._1426247587_config;
            this.btnSetNotePathProperties.Label = "경로 설정";
            this.btnSetNotePathProperties.Name = "btnSetNotePathProperties";
            this.btnSetNotePathProperties.ShowImage = true;
            this.btnSetNotePathProperties.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSetPathProperties_Click);
            // 
            // btnVerifyMonsters
            // 
            this.btnVerifyMonsters.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnVerifyMonsters.Image = global::IG_TableExporter.Properties.Resources._1424895247_frankenstein_monster_icon;
            this.btnVerifyMonsters.Label = "몬스터 데이터";
            this.btnVerifyMonsters.Name = "btnVerifyMonsters";
            this.btnVerifyMonsters.ShowImage = true;
            this.btnVerifyMonsters.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnVerifyMonsters_Click);
            // 
            // btnExportStageNote
            // 
            this.btnExportStageNote.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnExportStageNote.Image = global::IG_TableExporter.Properties.Resources._1424876076_accessories_text_editor_512;
            this.btnExportStageNote.Label = "노트 추출";
            this.btnExportStageNote.Name = "btnExportStageNote";
            this.btnExportStageNote.ShowImage = true;
            this.btnExportStageNote.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExportStageNote_Click);
            // 
            // saveTableFileDialog
            // 
            this.saveTableFileDialog.OverwritePrompt = false;
            this.saveTableFileDialog.RestoreDirectory = true;
            this.saveTableFileDialog.FileOk += new System.ComponentModel.CancelEventHandler(this.saveTableFileDialog_FileOk);
            // 
            // saveNoteFileDialog
            // 
            this.saveNoteFileDialog.OverwritePrompt = false;
            this.saveNoteFileDialog.RestoreDirectory = true;
            this.saveNoteFileDialog.FileOk += new System.ComponentModel.CancelEventHandler(this.saveNoteFileDialog_FileOk);
            // 
            // monsterTablePathFileDialog
            // 
            this.monsterTablePathFileDialog.DefaultExt = "json";
            this.monsterTablePathFileDialog.FileName = "monsterTablePathFileDialog";
            this.monsterTablePathFileDialog.FileOk += new System.ComponentModel.CancelEventHandler(this.monsterTablePathFileDialog_FileOk);
            // 
            // resourcePathTablePathFileDialog
            // 
            this.resourcePathTablePathFileDialog.DefaultExt = "json";
            this.resourcePathTablePathFileDialog.FileName = "resourcePathTablePathFileDialog";
            this.resourcePathTablePathFileDialog.FileOk += new System.ComponentModel.CancelEventHandler(this.resourcePathTablePathFileDialog_FileOk);
            // 
            // monsterSpritePathBrowserDialog
            // 
            this.monsterSpritePathBrowserDialog.Description = "에셋 폴더를 지정해주세요";
            this.monsterSpritePathBrowserDialog.RootFolder = System.Environment.SpecialFolder.MyComputer;
            this.monsterSpritePathBrowserDialog.ShowNewFolderButton = false;
            // 
            // stageTableFileDialog
            // 
            this.stageTableFileDialog.FileName = "stageTableFileDialog";
            this.stageTableFileDialog.FileOk += new System.ComponentModel.CancelEventHandler(this.stageTableFileDialog_FileOk);
            // 
            // assetPathBrowserDialog
            // 
            this.assetPathBrowserDialog.Description = "에셋 폴더를 지정해주세요";
            // 
            // metaTablePathBrowserDialog
            // 
            this.metaTablePathBrowserDialog.Description = "메타테이블 폴더를 지정해주세요";
            this.metaTablePathBrowserDialog.ShowNewFolderButton = false;
            // 
            // tablePathBrowserDialog
            // 
            this.tablePathBrowserDialog.Description = "몬스터테이블이 저장될 폴더를 지정하세요";
            // 
            // monsterTablePathBrowserDialog
            // 
            this.monsterTablePathBrowserDialog.Description = "몬스터테이블이 저장된 폴더를 지정하세요";
            this.monsterTablePathBrowserDialog.ShowNewFolderButton = false;
            // 
            // IG_Ribbon
            // 
            this.Name = "IG_Ribbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.IG_Tab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.IG_Ribbon_Load);
            this.IG_Tab.ResumeLayout(false);
            this.IG_Tab.PerformLayout();
            this.IG_TableGroup.ResumeLayout(false);
            this.IG_TableGroup.PerformLayout();
            this.IG_MetaTableGroup.ResumeLayout(false);
            this.IG_MetaTableGroup.PerformLayout();
            this.IG_ResourcePathGroup.ResumeLayout(false);
            this.IG_ResourcePathGroup.PerformLayout();
            this.IG_StageNoteGroup.ResumeLayout(false);
            this.IG_StageNoteGroup.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab IG_Tab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup IG_TableGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExportTable;
        private System.Windows.Forms.SaveFileDialog saveTableFileDialog;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox branchComboBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCheckBranch;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExportStageNote;
        private System.Windows.Forms.SaveFileDialog saveNoteFileDialog;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnVerifyMonsters;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup IG_StageNoteGroup;
        private System.Windows.Forms.OpenFileDialog monsterTablePathFileDialog;
        private System.Windows.Forms.OpenFileDialog resourcePathTablePathFileDialog;
        private System.Windows.Forms.FolderBrowserDialog monsterSpritePathBrowserDialog;
        private System.Windows.Forms.OpenFileDialog stageTableFileDialog;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup IG_ResourcePathGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnVerifyResourcePaths;
        private System.Windows.Forms.FolderBrowserDialog assetPathBrowserDialog;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSetNotePathProperties;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSetResourcePathProperties;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExportMetaTable;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup IG_MetaTableGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSetMetaTablePathProperties;
        private System.Windows.Forms.FolderBrowserDialog metaTablePathBrowserDialog;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExportMonsterTable;
        private System.Windows.Forms.FolderBrowserDialog tablePathBrowserDialog;
        private System.Windows.Forms.FolderBrowserDialog monsterTablePathBrowserDialog;
    }

    partial class ThisRibbonCollection
    {
        internal IG_Ribbon IG_Ribbon
        {
            get { return this.GetRibbon<IG_Ribbon>(); }
        }
    }
}

namespace IG_TableExporter
{
    partial class MonsterInfoTask
    {
        /// <summary> 
        /// 필수 디자이너 변수입니다.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MonsterInfoTask));
            this.monsterDataGridView = new System.Windows.Forms.DataGridView();
            this.mIndex = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.mSprite = new System.Windows.Forms.DataGridViewImageColumn();
            this.mSpeed = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.mScale = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.mHP = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.mAtk = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.mType = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.mPoint = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.mGold = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.monsterDataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // monsterDataGridView
            // 
            this.monsterDataGridView.AllowDrop = true;
            this.monsterDataGridView.AllowUserToAddRows = false;
            this.monsterDataGridView.AllowUserToDeleteRows = false;
            this.monsterDataGridView.AllowUserToOrderColumns = true;
            this.monsterDataGridView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.monsterDataGridView.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.monsterDataGridView.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.monsterDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.monsterDataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.mIndex,
            this.mSprite,
            this.mSpeed,
            this.mScale,
            this.mHP,
            this.mAtk,
            this.mType,
            this.mPoint,
            this.mGold});
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.monsterDataGridView.DefaultCellStyle = dataGridViewCellStyle3;
            this.monsterDataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.monsterDataGridView.Location = new System.Drawing.Point(0, 0);
            this.monsterDataGridView.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.monsterDataGridView.MultiSelect = false;
            this.monsterDataGridView.Name = "monsterDataGridView";
            this.monsterDataGridView.ReadOnly = true;
            this.monsterDataGridView.RowHeadersVisible = false;
            this.monsterDataGridView.RowTemplate.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this.monsterDataGridView.RowTemplate.Height = 64;
            this.monsterDataGridView.RowTemplate.ReadOnly = true;
            this.monsterDataGridView.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.monsterDataGridView.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.monsterDataGridView.ShowCellErrors = false;
            this.monsterDataGridView.ShowRowErrors = false;
            this.monsterDataGridView.Size = new System.Drawing.Size(1029, 200);
            this.monsterDataGridView.TabIndex = 0;
            this.monsterDataGridView.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.monsterDataGridView_CellDoubleClick);
            this.monsterDataGridView.KeyUp += new System.Windows.Forms.KeyEventHandler(this.monsterDataGridView_KeyUp);
            // 
            // mIndex
            // 
            this.mIndex.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.mIndex.Frozen = true;
            this.mIndex.HeaderText = "인덱스";
            this.mIndex.Name = "mIndex";
            this.mIndex.ReadOnly = true;
            this.mIndex.Width = 128;
            // 
            // mSprite
            // 
            this.mSprite.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle2.NullValue = ((object)(resources.GetObject("dataGridViewCellStyle2.NullValue")));
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.mSprite.DefaultCellStyle = dataGridViewCellStyle2;
            this.mSprite.Frozen = true;
            this.mSprite.HeaderText = "이미지";
            this.mSprite.ImageLayout = System.Windows.Forms.DataGridViewImageCellLayout.Zoom;
            this.mSprite.Name = "mSprite";
            this.mSprite.ReadOnly = true;
            this.mSprite.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.mSprite.Width = 48;
            // 
            // mSpeed
            // 
            this.mSpeed.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.mSpeed.Frozen = true;
            this.mSpeed.HeaderText = "스피드";
            this.mSpeed.Name = "mSpeed";
            this.mSpeed.ReadOnly = true;
            this.mSpeed.Width = 128;
            // 
            // mScale
            // 
            this.mScale.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.mScale.Frozen = true;
            this.mScale.HeaderText = "스케일";
            this.mScale.Name = "mScale";
            this.mScale.ReadOnly = true;
            this.mScale.Width = 128;
            // 
            // mHP
            // 
            this.mHP.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.mHP.Frozen = true;
            this.mHP.HeaderText = "HP";
            this.mHP.Name = "mHP";
            this.mHP.ReadOnly = true;
            this.mHP.Width = 79;
            // 
            // mAtk
            // 
            this.mAtk.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.mAtk.Frozen = true;
            this.mAtk.HeaderText = "공격력";
            this.mAtk.Name = "mAtk";
            this.mAtk.ReadOnly = true;
            this.mAtk.Width = 128;
            // 
            // mType
            // 
            this.mType.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.mType.Frozen = true;
            this.mType.HeaderText = "타입";
            this.mType.Name = "mType";
            this.mType.ReadOnly = true;
            this.mType.Width = 98;
            // 
            // mPoint
            // 
            this.mPoint.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.mPoint.Frozen = true;
            this.mPoint.HeaderText = "포인트";
            this.mPoint.Name = "mPoint";
            this.mPoint.ReadOnly = true;
            this.mPoint.Width = 128;
            // 
            // mGold
            // 
            this.mGold.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.mGold.Frozen = true;
            this.mGold.HeaderText = "골드";
            this.mGold.Name = "mGold";
            this.mGold.ReadOnly = true;
            this.mGold.Width = 98;
            // 
            // MonsterInfoTask
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(16F, 31F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.Controls.Add(this.monsterDataGridView);
            this.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.MinimumSize = new System.Drawing.Size(1029, 200);
            this.Name = "MonsterInfoTask";
            this.Size = new System.Drawing.Size(1029, 200);
            this.Load += new System.EventHandler(this.MonsterInfoTask_Load);
            ((System.ComponentModel.ISupportInitialize)(this.monsterDataGridView)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        internal System.Windows.Forms.DataGridView monsterDataGridView;
        private System.Windows.Forms.DataGridViewTextBoxColumn mIndex;
        private System.Windows.Forms.DataGridViewImageColumn mSprite;
        private System.Windows.Forms.DataGridViewTextBoxColumn mSpeed;
        private System.Windows.Forms.DataGridViewTextBoxColumn mScale;
        private System.Windows.Forms.DataGridViewTextBoxColumn mHP;
        private System.Windows.Forms.DataGridViewTextBoxColumn mAtk;
        private System.Windows.Forms.DataGridViewTextBoxColumn mType;
        private System.Windows.Forms.DataGridViewTextBoxColumn mPoint;
        private System.Windows.Forms.DataGridViewTextBoxColumn mGold;

    }
}

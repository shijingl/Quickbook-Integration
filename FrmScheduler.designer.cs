namespace QbIntegration
{
    partial class FrmScheduler
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.kryptonPanel1 = new ComponentFactory.Krypton.Toolkit.KryptonPanel();
            this.lblFrmName = new System.Windows.Forms.Label();
            this.btnSave = new ComponentFactory.Krypton.Toolkit.KryptonButton();
            this.btnClose = new ComponentFactory.Krypton.Toolkit.KryptonButton();
            this.kryptonPanel3 = new ComponentFactory.Krypton.Toolkit.KryptonPanel();
            this.kryptonPanel6 = new ComponentFactory.Krypton.Toolkit.KryptonPanel();
            this.kryptonPanel2 = new ComponentFactory.Krypton.Toolkit.KryptonPanel();
            this.kryptonPanel5 = new ComponentFactory.Krypton.Toolkit.KryptonPanel();
            this.kryptonPanel4 = new ComponentFactory.Krypton.Toolkit.KryptonPanel();
            this.dgScheduler = new ComponentFactory.Krypton.Toolkit.KryptonDataGridView();
            this.Day = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.StartTime = new ComponentFactory.Krypton.Toolkit.KryptonDataGridViewMaskedTextBoxColumn();
            this.EndTime = new ComponentFactory.Krypton.Toolkit.KryptonDataGridViewMaskedTextBoxColumn();
            this.ExecutionTime = new ComponentFactory.Krypton.Toolkit.KryptonDataGridViewMaskedTextBoxColumn();
            this.Recursive = new ComponentFactory.Krypton.Toolkit.KryptonDataGridViewNumericUpDownColumn();
            this.lblErrorMsg = new ComponentFactory.Krypton.Toolkit.KryptonWrapLabel();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonPanel1)).BeginInit();
            this.kryptonPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonPanel3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonPanel6)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonPanel2)).BeginInit();
            this.kryptonPanel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonPanel5)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonPanel4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgScheduler)).BeginInit();
            this.SuspendLayout();
            // 
            // kryptonPanel1
            // 
            this.kryptonPanel1.Controls.Add(this.lblFrmName);
            this.kryptonPanel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.kryptonPanel1.Location = new System.Drawing.Point(0, 0);
            this.kryptonPanel1.Name = "kryptonPanel1";
            this.kryptonPanel1.Size = new System.Drawing.Size(756, 31);
            this.kryptonPanel1.StateCommon.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(43)))), ((int)(((byte)(84)))));
            this.kryptonPanel1.StateCommon.Color2 = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(43)))), ((int)(((byte)(84)))));
            this.kryptonPanel1.TabIndex = 112;
            // 
            // lblFrmName
            // 
            this.lblFrmName.AutoSize = true;
            this.lblFrmName.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(43)))), ((int)(((byte)(84)))));
            this.lblFrmName.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFrmName.ForeColor = System.Drawing.SystemColors.ButtonFace;
            this.lblFrmName.Location = new System.Drawing.Point(12, 9);
            this.lblFrmName.Name = "lblFrmName";
            this.lblFrmName.Size = new System.Drawing.Size(64, 13);
            this.lblFrmName.TabIndex = 86;
            this.lblFrmName.Text = "Scheduler";
            // 
            // btnSave
            // 
            this.btnSave.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.btnSave.Location = new System.Drawing.Point(615, 37);
            this.btnSave.Name = "btnSave";
            this.btnSave.OverrideDefault.Back.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(211)))), ((int)(((byte)(222)))), ((int)(((byte)(229)))));
            this.btnSave.OverrideDefault.Back.Color2 = System.Drawing.Color.FromArgb(((int)(((byte)(211)))), ((int)(((byte)(222)))), ((int)(((byte)(229)))));
            this.btnSave.OverrideDefault.Border.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(47)))), ((int)(((byte)(79)))), ((int)(((byte)(117)))));
            this.btnSave.OverrideDefault.Border.Color2 = System.Drawing.Color.FromArgb(((int)(((byte)(47)))), ((int)(((byte)(79)))), ((int)(((byte)(117)))));
            this.btnSave.OverrideDefault.Border.DrawBorders = ((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right)));
            this.btnSave.OverrideDefault.Border.Rounding = 5;
            this.btnSave.OverrideDefault.Border.Width = 2;
            this.btnSave.OverrideDefault.Content.ShortText.Color1 = System.Drawing.Color.Black;
            this.btnSave.OverrideDefault.Content.ShortText.Color2 = System.Drawing.Color.Black;
            this.btnSave.OverrideDefault.Content.ShortText.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold);
            this.btnSave.OverrideFocus.Back.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(47)))), ((int)(((byte)(79)))), ((int)(((byte)(117)))));
            this.btnSave.OverrideFocus.Back.Color2 = System.Drawing.Color.FromArgb(((int)(((byte)(47)))), ((int)(((byte)(79)))), ((int)(((byte)(117)))));
            this.btnSave.OverrideFocus.Border.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(47)))), ((int)(((byte)(79)))), ((int)(((byte)(117)))));
            this.btnSave.OverrideFocus.Border.Color2 = System.Drawing.Color.FromArgb(((int)(((byte)(47)))), ((int)(((byte)(79)))), ((int)(((byte)(117)))));
            this.btnSave.OverrideFocus.Border.DrawBorders = ((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right)));
            this.btnSave.OverrideFocus.Border.Rounding = 5;
            this.btnSave.OverrideFocus.Border.Width = 2;
            this.btnSave.OverrideFocus.Content.ShortText.Color1 = System.Drawing.Color.White;
            this.btnSave.OverrideFocus.Content.ShortText.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSave.Size = new System.Drawing.Size(57, 29);
            this.btnSave.StateCommon.Back.Color1 = System.Drawing.Color.White;
            this.btnSave.StateCommon.Back.Color2 = System.Drawing.Color.White;
            this.btnSave.StateCommon.Border.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(47)))), ((int)(((byte)(79)))), ((int)(((byte)(117)))));
            this.btnSave.StateCommon.Border.Color2 = System.Drawing.Color.FromArgb(((int)(((byte)(47)))), ((int)(((byte)(79)))), ((int)(((byte)(117)))));
            this.btnSave.StateCommon.Border.DrawBorders = ((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right)));
            this.btnSave.StateCommon.Border.Rounding = 5;
            this.btnSave.StateCommon.Border.Width = 2;
            this.btnSave.StateCommon.Content.ShortText.Color1 = System.Drawing.Color.Black;
            this.btnSave.StateCommon.Content.ShortText.Color2 = System.Drawing.Color.Black;
            this.btnSave.StateCommon.Content.ShortText.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold);
            this.btnSave.StateDisabled.Back.Color1 = System.Drawing.Color.White;
            this.btnSave.StateDisabled.Back.Color2 = System.Drawing.Color.White;
            this.btnSave.StateDisabled.Content.ShortText.Color1 = System.Drawing.Color.Black;
            this.btnSave.StateDisabled.Content.ShortText.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold);
            this.btnSave.StateNormal.Back.Color1 = System.Drawing.Color.White;
            this.btnSave.StateNormal.Back.Color2 = System.Drawing.Color.White;
            this.btnSave.StateNormal.Content.ShortText.Color1 = System.Drawing.Color.Black;
            this.btnSave.StateNormal.Content.ShortText.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold);
            this.btnSave.StatePressed.Back.Color1 = System.Drawing.Color.White;
            this.btnSave.StatePressed.Back.Color2 = System.Drawing.Color.White;
            this.btnSave.StatePressed.Content.ShortText.Color1 = System.Drawing.Color.White;
            this.btnSave.StatePressed.Content.ShortText.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold);
            this.btnSave.StateTracking.Back.Color1 = System.Drawing.Color.White;
            this.btnSave.StateTracking.Back.Color2 = System.Drawing.Color.White;
            this.btnSave.StateTracking.Content.ShortText.Color1 = System.Drawing.Color.Black;
            this.btnSave.StateTracking.Content.ShortText.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold);
            this.btnSave.TabIndex = 127;
            this.btnSave.Values.Text = "Save";
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnClose
            // 
            this.btnClose.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.btnClose.Location = new System.Drawing.Point(678, 37);
            this.btnClose.Name = "btnClose";
            this.btnClose.OverrideDefault.Back.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(211)))), ((int)(((byte)(222)))), ((int)(((byte)(229)))));
            this.btnClose.OverrideDefault.Back.Color2 = System.Drawing.Color.FromArgb(((int)(((byte)(211)))), ((int)(((byte)(222)))), ((int)(((byte)(229)))));
            this.btnClose.OverrideDefault.Border.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(47)))), ((int)(((byte)(79)))), ((int)(((byte)(117)))));
            this.btnClose.OverrideDefault.Border.Color2 = System.Drawing.Color.FromArgb(((int)(((byte)(47)))), ((int)(((byte)(79)))), ((int)(((byte)(117)))));
            this.btnClose.OverrideDefault.Border.DrawBorders = ((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right)));
            this.btnClose.OverrideDefault.Border.Rounding = 5;
            this.btnClose.OverrideDefault.Border.Width = 2;
            this.btnClose.OverrideDefault.Content.ShortText.Color1 = System.Drawing.Color.Black;
            this.btnClose.OverrideDefault.Content.ShortText.Color2 = System.Drawing.Color.Black;
            this.btnClose.OverrideDefault.Content.ShortText.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold);
            this.btnClose.OverrideFocus.Back.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(47)))), ((int)(((byte)(79)))), ((int)(((byte)(117)))));
            this.btnClose.OverrideFocus.Back.Color2 = System.Drawing.Color.FromArgb(((int)(((byte)(47)))), ((int)(((byte)(79)))), ((int)(((byte)(117)))));
            this.btnClose.OverrideFocus.Border.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(47)))), ((int)(((byte)(79)))), ((int)(((byte)(117)))));
            this.btnClose.OverrideFocus.Border.Color2 = System.Drawing.Color.FromArgb(((int)(((byte)(47)))), ((int)(((byte)(79)))), ((int)(((byte)(117)))));
            this.btnClose.OverrideFocus.Border.DrawBorders = ((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right)));
            this.btnClose.OverrideFocus.Border.Rounding = 5;
            this.btnClose.OverrideFocus.Border.Width = 2;
            this.btnClose.OverrideFocus.Content.ShortText.Color1 = System.Drawing.Color.White;
            this.btnClose.OverrideFocus.Content.ShortText.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnClose.Size = new System.Drawing.Size(57, 29);
            this.btnClose.StateCommon.Back.Color1 = System.Drawing.Color.White;
            this.btnClose.StateCommon.Back.Color2 = System.Drawing.Color.White;
            this.btnClose.StateCommon.Border.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(47)))), ((int)(((byte)(79)))), ((int)(((byte)(117)))));
            this.btnClose.StateCommon.Border.Color2 = System.Drawing.Color.FromArgb(((int)(((byte)(47)))), ((int)(((byte)(79)))), ((int)(((byte)(117)))));
            this.btnClose.StateCommon.Border.DrawBorders = ((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right)));
            this.btnClose.StateCommon.Border.Rounding = 5;
            this.btnClose.StateCommon.Border.Width = 2;
            this.btnClose.StateCommon.Content.ShortText.Color1 = System.Drawing.Color.Black;
            this.btnClose.StateCommon.Content.ShortText.Color2 = System.Drawing.Color.Black;
            this.btnClose.StateCommon.Content.ShortText.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold);
            this.btnClose.StateDisabled.Back.Color1 = System.Drawing.Color.White;
            this.btnClose.StateDisabled.Back.Color2 = System.Drawing.Color.White;
            this.btnClose.StateDisabled.Content.ShortText.Color1 = System.Drawing.Color.Black;
            this.btnClose.StateDisabled.Content.ShortText.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold);
            this.btnClose.StateNormal.Back.Color1 = System.Drawing.Color.White;
            this.btnClose.StateNormal.Back.Color2 = System.Drawing.Color.White;
            this.btnClose.StateNormal.Content.ShortText.Color1 = System.Drawing.Color.Black;
            this.btnClose.StateNormal.Content.ShortText.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold);
            this.btnClose.StatePressed.Back.Color1 = System.Drawing.Color.White;
            this.btnClose.StatePressed.Back.Color2 = System.Drawing.Color.White;
            this.btnClose.StatePressed.Content.ShortText.Color1 = System.Drawing.Color.White;
            this.btnClose.StatePressed.Content.ShortText.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold);
            this.btnClose.StateTracking.Back.Color1 = System.Drawing.Color.White;
            this.btnClose.StateTracking.Back.Color2 = System.Drawing.Color.White;
            this.btnClose.StateTracking.Content.ShortText.Color1 = System.Drawing.Color.Black;
            this.btnClose.StateTracking.Content.ShortText.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold);
            this.btnClose.TabIndex = 126;
            this.btnClose.Values.Text = "Close";
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // kryptonPanel3
            // 
            this.kryptonPanel3.Dock = System.Windows.Forms.DockStyle.Left;
            this.kryptonPanel3.Location = new System.Drawing.Point(0, 31);
            this.kryptonPanel3.Name = "kryptonPanel3";
            this.kryptonPanel3.Size = new System.Drawing.Size(5, 313);
            this.kryptonPanel3.StateCommon.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(43)))), ((int)(((byte)(84)))));
            this.kryptonPanel3.StateCommon.Color2 = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(43)))), ((int)(((byte)(84)))));
            this.kryptonPanel3.TabIndex = 128;
            // 
            // kryptonPanel6
            // 
            this.kryptonPanel6.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.kryptonPanel6.Location = new System.Drawing.Point(5, 339);
            this.kryptonPanel6.Name = "kryptonPanel6";
            this.kryptonPanel6.Size = new System.Drawing.Size(751, 5);
            this.kryptonPanel6.StateCommon.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(43)))), ((int)(((byte)(84)))));
            this.kryptonPanel6.StateCommon.Color2 = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(43)))), ((int)(((byte)(84)))));
            this.kryptonPanel6.TabIndex = 129;
            // 
            // kryptonPanel2
            // 
            this.kryptonPanel2.Controls.Add(this.kryptonPanel5);
            this.kryptonPanel2.Controls.Add(this.kryptonPanel4);
            this.kryptonPanel2.Dock = System.Windows.Forms.DockStyle.Right;
            this.kryptonPanel2.Location = new System.Drawing.Point(751, 31);
            this.kryptonPanel2.Name = "kryptonPanel2";
            this.kryptonPanel2.Size = new System.Drawing.Size(5, 308);
            this.kryptonPanel2.StateCommon.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(43)))), ((int)(((byte)(84)))));
            this.kryptonPanel2.StateCommon.Color2 = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(43)))), ((int)(((byte)(84)))));
            this.kryptonPanel2.TabIndex = 130;
            // 
            // kryptonPanel5
            // 
            this.kryptonPanel5.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.kryptonPanel5.Location = new System.Drawing.Point(0, 298);
            this.kryptonPanel5.Name = "kryptonPanel5";
            this.kryptonPanel5.Size = new System.Drawing.Size(5, 5);
            this.kryptonPanel5.StateCommon.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(43)))), ((int)(((byte)(84)))));
            this.kryptonPanel5.StateCommon.Color2 = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(43)))), ((int)(((byte)(84)))));
            this.kryptonPanel5.TabIndex = 117;
            // 
            // kryptonPanel4
            // 
            this.kryptonPanel4.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.kryptonPanel4.Location = new System.Drawing.Point(0, 303);
            this.kryptonPanel4.Name = "kryptonPanel4";
            this.kryptonPanel4.Size = new System.Drawing.Size(5, 5);
            this.kryptonPanel4.StateCommon.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(43)))), ((int)(((byte)(84)))));
            this.kryptonPanel4.StateCommon.Color2 = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(43)))), ((int)(((byte)(84)))));
            this.kryptonPanel4.TabIndex = 112;
            // 
            // dgScheduler
            // 
            this.dgScheduler.AllowUserToAddRows = false;
            this.dgScheduler.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgScheduler.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Day,
            this.StartTime,
            this.EndTime,
            this.ExecutionTime,
            this.Recursive});
            this.dgScheduler.Location = new System.Drawing.Point(15, 72);
            this.dgScheduler.Name = "dgScheduler";
            this.dgScheduler.RowHeadersVisible = false;
            this.dgScheduler.Size = new System.Drawing.Size(720, 221);
            this.dgScheduler.StateCommon.Background.Color1 = System.Drawing.Color.White;
            this.dgScheduler.StateCommon.Background.Color2 = System.Drawing.Color.White;
            this.dgScheduler.StateCommon.BackStyle = ComponentFactory.Krypton.Toolkit.PaletteBackStyle.GridBackgroundList;
            this.dgScheduler.StateCommon.HeaderColumn.Back.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(43)))), ((int)(((byte)(84)))));
            this.dgScheduler.StateCommon.HeaderColumn.Back.Color2 = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(43)))), ((int)(((byte)(84)))));
            this.dgScheduler.StateCommon.HeaderColumn.Content.Color1 = System.Drawing.Color.White;
            this.dgScheduler.StateCommon.HeaderColumn.Content.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dgScheduler.StateCommon.HeaderRow.Content.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(43)))), ((int)(((byte)(84)))));
            this.dgScheduler.TabIndex = 131;
            // 
            // Day
            // 
            this.Day.HeaderText = "Day";
            this.Day.Name = "Day";
            this.Day.ReadOnly = true;
            // 
            // StartTime
            // 
            this.StartTime.HeaderText = "StartTime";
            this.StartTime.Mask = "00:00:00";
            this.StartTime.Name = "StartTime";
            this.StartTime.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.StartTime.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.StartTime.ToolTipText = "00:00:00";
            this.StartTime.Width = 144;
            // 
            // EndTime
            // 
            this.EndTime.HeaderText = "EndTime";
            this.EndTime.Mask = "00:00:00";
            this.EndTime.Name = "EndTime";
            this.EndTime.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.EndTime.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.EndTime.Width = 143;
            // 
            // ExecutionTime
            // 
            this.ExecutionTime.HeaderText = "ExecutionTime";
            this.ExecutionTime.Mask = "00:00:00";
            this.ExecutionTime.Name = "ExecutionTime";
            this.ExecutionTime.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.ExecutionTime.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.ExecutionTime.Width = 144;
            // 
            // Recursive
            // 
            this.Recursive.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.Recursive.HeaderText = "Recursive (InMintute)";
            this.Recursive.Name = "Recursive";
            this.Recursive.Width = 144;
            // 
            // lblErrorMsg
            // 
            this.lblErrorMsg.AutoSize = false;
            this.lblErrorMsg.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblErrorMsg.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(30)))), ((int)(((byte)(57)))), ((int)(((byte)(91)))));
            this.lblErrorMsg.Location = new System.Drawing.Point(7, 305);
            this.lblErrorMsg.Name = "lblErrorMsg";
            this.lblErrorMsg.Size = new System.Drawing.Size(737, 32);
            this.lblErrorMsg.StateCommon.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            // 
            // FrmScheduler
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(756, 344);
            this.Controls.Add(this.lblErrorMsg);
            this.Controls.Add(this.dgScheduler);
            this.Controls.Add(this.kryptonPanel2);
            this.Controls.Add(this.kryptonPanel6);
            this.Controls.Add(this.kryptonPanel3);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.kryptonPanel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "FrmScheduler";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "FrmScheduler";
            this.Load += new System.EventHandler(this.FrmScheduler_Load);
            ((System.ComponentModel.ISupportInitialize)(this.kryptonPanel1)).EndInit();
            this.kryptonPanel1.ResumeLayout(false);
            this.kryptonPanel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonPanel3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonPanel6)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonPanel2)).EndInit();
            this.kryptonPanel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.kryptonPanel5)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonPanel4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgScheduler)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private ComponentFactory.Krypton.Toolkit.KryptonPanel kryptonPanel1;
        private System.Windows.Forms.Label lblFrmName;
        private ComponentFactory.Krypton.Toolkit.KryptonButton btnSave;
        private ComponentFactory.Krypton.Toolkit.KryptonButton btnClose;
        private ComponentFactory.Krypton.Toolkit.KryptonPanel kryptonPanel3;
        private ComponentFactory.Krypton.Toolkit.KryptonPanel kryptonPanel6;
        private ComponentFactory.Krypton.Toolkit.KryptonPanel kryptonPanel2;
        private ComponentFactory.Krypton.Toolkit.KryptonPanel kryptonPanel5;
        private ComponentFactory.Krypton.Toolkit.KryptonPanel kryptonPanel4;
        private ComponentFactory.Krypton.Toolkit.KryptonDataGridView dgScheduler;
        private System.Windows.Forms.DataGridViewTextBoxColumn Day;
        private ComponentFactory.Krypton.Toolkit.KryptonDataGridViewMaskedTextBoxColumn StartTime;
        private ComponentFactory.Krypton.Toolkit.KryptonDataGridViewMaskedTextBoxColumn EndTime;
        private ComponentFactory.Krypton.Toolkit.KryptonDataGridViewMaskedTextBoxColumn ExecutionTime;
        private ComponentFactory.Krypton.Toolkit.KryptonDataGridViewNumericUpDownColumn Recursive;
        private ComponentFactory.Krypton.Toolkit.KryptonWrapLabel lblErrorMsg;
    }
}
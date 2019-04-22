namespace QbIntegration
{
    partial class FrmHistory
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
            this.components = new System.ComponentModel.Container();
            this.kryptonManager = new ComponentFactory.Krypton.Toolkit.KryptonManager(this.components);
            this.kryptonManager1 = new ComponentFactory.Krypton.Toolkit.KryptonManager(this.components);
            this.lblErrorMessage = new ComponentFactory.Krypton.Toolkit.KryptonWrapLabel();
            this.dtViewDate = new ComponentFactory.Krypton.Toolkit.KryptonDateTimePicker();
            this.kryptonLabel1 = new ComponentFactory.Krypton.Toolkit.KryptonLabel();
            this.btnClose = new ComponentFactory.Krypton.Toolkit.KryptonButton();
            this.dgHistory = new ComponentFactory.Krypton.Toolkit.KryptonDataGridView();
            this.Srno = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Time = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Type = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Total = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.New = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Update = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Skip = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Fail = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dgHistory)).BeginInit();
            this.SuspendLayout();
            // 
            // lblErrorMessage
            // 
            this.lblErrorMessage.AutoSize = false;
            this.lblErrorMessage.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.lblErrorMessage.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblErrorMessage.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(30)))), ((int)(((byte)(57)))), ((int)(((byte)(91)))));
            this.lblErrorMessage.Location = new System.Drawing.Point(0, 401);
            this.lblErrorMessage.Name = "lblErrorMessage";
            this.lblErrorMessage.Size = new System.Drawing.Size(833, 33);
            this.lblErrorMessage.StateCommon.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            // 
            // dtViewDate
            // 
            this.dtViewDate.CustomFormat = "yyyy-MM-dd";
            this.dtViewDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtViewDate.Location = new System.Drawing.Point(211, 13);
            this.dtViewDate.Name = "dtViewDate";
            this.dtViewDate.Size = new System.Drawing.Size(233, 19);
            this.dtViewDate.StateActive.Content.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(47)))), ((int)(((byte)(79)))), ((int)(((byte)(117)))));
            this.dtViewDate.StateActive.Content.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dtViewDate.StateCommon.Content.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(47)))), ((int)(((byte)(79)))), ((int)(((byte)(117)))));
            this.dtViewDate.StateCommon.Content.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dtViewDate.TabIndex = 158;
            this.dtViewDate.ValueChanged += new System.EventHandler(this.dtViewDate_ValueChanged);
            // 
            // kryptonLabel1
            // 
            this.kryptonLabel1.Location = new System.Drawing.Point(24, 14);
            this.kryptonLabel1.Name = "kryptonLabel1";
            this.kryptonLabel1.Size = new System.Drawing.Size(154, 17);
            this.kryptonLabel1.StateCommon.ShortText.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(47)))), ((int)(((byte)(79)))), ((int)(((byte)(117)))));
            this.kryptonLabel1.StateCommon.ShortText.Color2 = System.Drawing.Color.FromArgb(((int)(((byte)(47)))), ((int)(((byte)(79)))), ((int)(((byte)(117)))));
            this.kryptonLabel1.StateCommon.ShortText.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.kryptonLabel1.StateNormal.ShortText.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(47)))), ((int)(((byte)(79)))), ((int)(((byte)(117)))));
            this.kryptonLabel1.StateNormal.ShortText.Color2 = System.Drawing.Color.FromArgb(((int)(((byte)(47)))), ((int)(((byte)(79)))), ((int)(((byte)(117)))));
            this.kryptonLabel1.StateNormal.ShortText.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.kryptonLabel1.TabIndex = 159;
            this.kryptonLabel1.Values.Text = "Select Date to History:";
            // 
            // btnClose
            // 
            this.btnClose.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.btnClose.Location = new System.Drawing.Point(762, 12);
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
            this.btnClose.TabIndex = 161;
            this.btnClose.Values.Text = "Close";
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // dgHistory
            // 
            this.dgHistory.AllowUserToAddRows = false;
            this.dgHistory.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgHistory.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Srno,
            this.Time,
            this.Type,
            this.Total,
            this.New,
            this.Update,
            this.Skip,
            this.Fail});
            this.dgHistory.Location = new System.Drawing.Point(13, 50);
            this.dgHistory.Name = "dgHistory";
            this.dgHistory.RowHeadersVisible = false;
            this.dgHistory.Size = new System.Drawing.Size(806, 334);
            this.dgHistory.StateCommon.Background.Color1 = System.Drawing.Color.White;
            this.dgHistory.StateCommon.Background.Color2 = System.Drawing.Color.White;
            this.dgHistory.StateCommon.BackStyle = ComponentFactory.Krypton.Toolkit.PaletteBackStyle.GridBackgroundList;
            this.dgHistory.StateCommon.HeaderColumn.Back.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(64)))));
            this.dgHistory.StateCommon.HeaderColumn.Back.Color2 = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(64)))));
            this.dgHistory.StateCommon.HeaderColumn.Content.Color1 = System.Drawing.Color.White;
            this.dgHistory.StateCommon.HeaderColumn.Content.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dgHistory.StateCommon.HeaderRow.Back.Color1 = System.Drawing.Color.White;
            this.dgHistory.StateCommon.HeaderRow.Back.Color2 = System.Drawing.Color.White;
            this.dgHistory.StateCommon.HeaderRow.Content.Color1 = System.Drawing.Color.Black;
            this.dgHistory.TabIndex = 162;
            // 
            // Srno
            // 
            this.Srno.FillWeight = 73.22935F;
            this.Srno.HeaderText = "Sr. No.";
            this.Srno.Name = "Srno";
            this.Srno.ReadOnly = true;
            // 
            // Time
            // 
            this.Time.FillWeight = 73.22935F;
            this.Time.HeaderText = "Time";
            this.Time.Name = "Time";
            this.Time.ReadOnly = true;
            // 
            // Type
            // 
            this.Type.FillWeight = 254.8485F;
            this.Type.HeaderText = "Type";
            this.Type.Name = "Type";
            this.Type.ReadOnly = true;
            // 
            // Total
            // 
            this.Total.FillWeight = 81.21828F;
            this.Total.HeaderText = "Total";
            this.Total.Name = "Total";
            // 
            // New
            // 
            this.New.FillWeight = 80.40722F;
            this.New.HeaderText = "New";
            this.New.Name = "New";
            // 
            // Update
            // 
            this.Update.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader;
            this.Update.HeaderText = "Update";
            this.Update.Name = "Update";
            this.Update.ReadOnly = true;
            this.Update.Width = 82;
            // 
            // Skip
            // 
            this.Skip.FillWeight = 79.02377F;
            this.Skip.HeaderText = "Skip";
            this.Skip.Name = "Skip";
            // 
            // Fail
            // 
            this.Fail.FillWeight = 78.36504F;
            this.Fail.HeaderText = "Fail";
            this.Fail.Name = "Fail";
            // 
            // FrmHistory
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(833, 434);
            this.ControlBox = false;
            this.Controls.Add(this.dgHistory);
            this.Controls.Add(this.lblErrorMessage);
            this.Controls.Add(this.dtViewDate);
            this.Controls.Add(this.kryptonLabel1);
            this.Controls.Add(this.btnClose);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Name = "FrmHistory";
            this.StateCommon.Back.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(64)))));
            this.StateCommon.Back.Color2 = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(64)))));
            this.StateCommon.Border.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(64)))));
            this.StateCommon.Border.Color2 = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(64)))));
            this.StateCommon.Border.DrawBorders = ((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right)));
            this.StateCommon.Header.Back.Color1 = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(64)))));
            this.StateCommon.Header.Back.Color2 = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(64)))));
            this.StateCommon.Header.Content.ShortText.Color1 = System.Drawing.Color.White;
            this.StateCommon.Header.Content.ShortText.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Text = "History";
            this.Load += new System.EventHandler(this.FrmHistory_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgHistory)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private ComponentFactory.Krypton.Toolkit.KryptonManager kryptonManager;
        private ComponentFactory.Krypton.Toolkit.KryptonManager kryptonManager1;
        private ComponentFactory.Krypton.Toolkit.KryptonWrapLabel lblErrorMessage;
        private ComponentFactory.Krypton.Toolkit.KryptonDateTimePicker dtViewDate;
        private ComponentFactory.Krypton.Toolkit.KryptonLabel kryptonLabel1;
        private ComponentFactory.Krypton.Toolkit.KryptonButton btnClose;
        private ComponentFactory.Krypton.Toolkit.KryptonDataGridView dgHistory;
        private System.Windows.Forms.DataGridViewTextBoxColumn Srno;
        private System.Windows.Forms.DataGridViewTextBoxColumn Time;
        private System.Windows.Forms.DataGridViewTextBoxColumn Type;
        private System.Windows.Forms.DataGridViewTextBoxColumn Total;
        private System.Windows.Forms.DataGridViewTextBoxColumn New;
        private System.Windows.Forms.DataGridViewTextBoxColumn Update;
        private System.Windows.Forms.DataGridViewTextBoxColumn Skip;
        private System.Windows.Forms.DataGridViewTextBoxColumn Fail;
    }
}


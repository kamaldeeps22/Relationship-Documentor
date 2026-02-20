using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using McTools.Xrm.Connection;
using XrmToolBox.Extensibility;
using XrmToolBox.Extensibility.Interfaces;
using Excel = Microsoft.Office.Interop.Excel;
using Label = System.Windows.Forms.Label;

namespace Relationship_Documentor
{
    public partial class MyPluginControl : PluginControlBase, IAboutPlugin
    {
        // fields
        private Settings mySettings;
        private List<EntityItem> allEntities;
        private string _currentSolutionUniqueName = null;
        private HashSet<Guid> _relationshipFilter = null;
        private Dictionary<string, TableTabData> _tableTabs = new Dictionary<string, TableTabData>();
        private HashSet<string> _checkedTableLogicalNames = new HashSet<string>();
        private bool _suppressItemCheck = false;

        private void panelSearchSeparator_Paint(object sender, PaintEventArgs e)
        {
            var panel = (Panel)sender;
            int y = panel.Height / 2;
            using (var pen = new Pen(Color.LightGray))
                e.Graphics.DrawLine(pen, 4, y, panel.Width - 4, y);
        }

        public void ShowAboutDialog()
        {
            var aboutForm = new Form
            {
                Text = "About Relationship Documentor",
                Size = new Size(400, 300),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false
            };
            var panel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(20) };
            var lblTitle = new Label { Text = "Relationship Documentor", Font = new Font("Segoe UI", 14F, FontStyle.Bold), AutoSize = true, Location = new Point(20, 20) };
            var lblVersion = new Label { Text = "Version 1.2.0.0", AutoSize = true, Location = new Point(20, 55) };
            var lblDescription = new Label { Text = "View and document entity relationships (N:1, 1:N, N:N)\nwith complete relationship behaviors.", AutoSize = true, MaximumSize = new Size(340, 0), Location = new Point(20, 85) };
            var lblAuthor = new Label { Text = "Developer: Kamaldeep Singh", AutoSize = true, Location = new Point(20, 130) };
            var linkWebsite = new LinkLabel { Text = "Website: kamaldeepsingh.com", AutoSize = true, Location = new Point(20, 155) };
            linkWebsite.LinkClicked += (s, e) => { try { System.Diagnostics.Process.Start("https://kamaldeepsingh.com"); } catch { } };
            var linkEmail = new LinkLabel { Text = "Email: hello@kamaldeepsingh.com", AutoSize = true, Location = new Point(20, 180) };
            linkEmail.LinkClicked += (s, e) => { try { System.Diagnostics.Process.Start("mailto:hello@kamaldeepsingh.com"); } catch { } };
            var btnOk = new Button { Text = "OK", DialogResult = DialogResult.OK, Size = new Size(75, 30), Location = new Point(290, 220) };
            panel.Controls.Add(lblTitle); panel.Controls.Add(lblVersion); panel.Controls.Add(lblDescription);
            panel.Controls.Add(lblAuthor); panel.Controls.Add(linkWebsite); panel.Controls.Add(linkEmail); panel.Controls.Add(btnOk);
            aboutForm.Controls.Add(panel);
            aboutForm.AcceptButton = btnOk;
            aboutForm.ShowDialog(this);
        }

        private void LogInfo(string message, params object[] args) => System.Diagnostics.Debug.WriteLine($"[INFO] {(args.Length > 0 ? string.Format(message, args) : message)}");
        private void LogWarning(string message) => System.Diagnostics.Debug.WriteLine($"[WARNING] {message}");
        private void LogError(string message, Exception ex = null) => System.Diagnostics.Debug.WriteLine($"[ERROR] {message}{(ex != null ? " - " + ex.Message : "")}");

        public MyPluginControl() { InitializeComponent(); }

        private void MyPluginControl_Load(object sender, EventArgs e)
        {
            if (!SettingsManager.Instance.TryLoad(GetType(), out mySettings))
            { mySettings = new Settings(); LogWarning("Settings not found => a new settings file has been created!"); }
            else LogInfo("Settings found and loaded");
        }

        private void tsbLoadTables_ButtonClick(object sender, EventArgs e) { _currentSolutionUniqueName = null; _relationshipFilter = null; ExecuteMethod(() => ExecuteLoadTables(null)); }
        private void tsmiLoadDefault_Click(object sender, EventArgs e) { _currentSolutionUniqueName = null; _relationshipFilter = null; ExecuteMethod(() => ExecuteLoadTables(null)); }

        private void tsmiLoadFromSolution_Click(object sender, EventArgs e)
        {
            WorkAsync(new WorkAsyncInfo
            {
                Message = "Loading solutions...",
                Work = (worker, args) =>
                {
                    var query = new QueryExpression("solution") { ColumnSet = new ColumnSet("uniquename", "friendlyname", "version", "publisherid"), Criteria = new FilterExpression() };
                    query.Criteria.AddCondition("isvisible", ConditionOperator.Equal, true);
                    var pubLink = query.AddLink("publisher", "publisherid", "publisherid");
                    pubLink.Columns = new ColumnSet("friendlyname"); pubLink.EntityAlias = "pub";
                    query.AddOrder("friendlyname", OrderType.Ascending);
                    args.Result = Service.RetrieveMultiple(query).Entities;
                },
                PostWorkCallBack = (args) =>
                {
                    if (args.Error != null) { MessageBox.Show($"Error loading solutions:\n{args.Error.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
                    var solutions = (DataCollection<Entity>)args.Result;
                    if (solutions == null || solutions.Count == 0) { MessageBox.Show("No solutions found.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
                    using (var picker = new SolutionPickerDialog(solutions))
                    {
                        if (picker.ShowDialog(this) == DialogResult.OK && picker.SelectedSolution != null)
                        { _currentSolutionUniqueName = picker.SelectedSolution.UniqueName; ExecuteMethod(() => ExecuteLoadTables(_currentSolutionUniqueName)); }
                    }
                }
            });
        }

        private void ExecuteLoadTables(string solutionUniqueName)
        {
            checkedListBoxTables.Items.Clear(); _tableTabs.Clear(); tabControlTables.TabPages.Clear();
            _checkedTableLogicalNames.Clear(); listBoxSelectedTables.Items.Clear();
            tsbShowRelationships.Enabled = false; _relationshipFilter = null;

            string progressMsg = solutionUniqueName == null ? "Loading tables from organization..." : $"Loading tables from solution '{solutionUniqueName}'...";

            WorkAsync(new WorkAsyncInfo
            {
                Message = progressMsg,
                IsCancelable = false,
                MessageWidth = 340,
                MessageHeight = 150,
                Work = (worker, args) =>
                {
                    List<EntityMetadata> entities;
                    if (solutionUniqueName == null)
                    {
                        var request = new RetrieveAllEntitiesRequest { EntityFilters = EntityFilters.Entity, RetrieveAsIfPublished = false };
                        var response = (RetrieveAllEntitiesResponse)Service.Execute(request);
                        entities = response.EntityMetadata.Where(e => e.LogicalName != null).OrderBy(e => e.DisplayName?.UserLocalizedLabel?.Label ?? e.LogicalName).ToList();
                        args.Result = new { Entities = entities, RelationshipFilter = (HashSet<Guid>)null };
                    }
                    else
                    {
                        var compQuery = new QueryExpression("solutioncomponent") { ColumnSet = new ColumnSet("objectid"), Criteria = new FilterExpression() };
                        compQuery.Criteria.AddCondition("componenttype", ConditionOperator.Equal, 1);
                        var solLink = compQuery.AddLink("solution", "solutionid", "solutionid");
                        solLink.LinkCriteria.AddCondition("uniquename", ConditionOperator.Equal, solutionUniqueName);
                        var entityIds = new HashSet<Guid>(Service.RetrieveMultiple(compQuery).Entities.Select(c => (Guid)c["objectid"]));

                        var allResp = (RetrieveAllEntitiesResponse)Service.Execute(new RetrieveAllEntitiesRequest { EntityFilters = EntityFilters.Entity, RetrieveAsIfPublished = false });
                        entities = allResp.EntityMetadata.Where(e => entityIds.Contains(e.MetadataId.GetValueOrDefault())).OrderBy(e => e.DisplayName?.UserLocalizedLabel?.Label ?? e.LogicalName).ToList();

                        var relCompQuery = new QueryExpression("solutioncomponent") { ColumnSet = new ColumnSet("objectid"), Criteria = new FilterExpression() };
                        relCompQuery.Criteria.AddCondition("componenttype", ConditionOperator.Equal, 10);
                        var relSolLink = relCompQuery.AddLink("solution", "solutionid", "solutionid");
                        relSolLink.LinkCriteria.AddCondition("uniquename", ConditionOperator.Equal, solutionUniqueName);
                        var relationshipIds = new HashSet<Guid>(Service.RetrieveMultiple(relCompQuery).Entities.Select(c => (Guid)c["objectid"]));
                        args.Result = new { Entities = entities, RelationshipFilter = relationshipIds };
                    }
                },
                PostWorkCallBack = (args) =>
                {
                    if (args.Error != null) { MessageBox.Show($"Error loading tables:\n{args.Error.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
                    var result = args.Result as dynamic;
                    var entities = result.Entities as List<EntityMetadata>;
                    _relationshipFilter = result.RelationshipFilter as HashSet<Guid>;
                    allEntities = entities.Select(entity => new EntityItem { DisplayName = entity.DisplayName?.UserLocalizedLabel?.Label ?? entity.LogicalName, LogicalName = entity.LogicalName, Metadata = entity }).ToList();
                    checkedListBoxTables.Items.Clear();
                    foreach (var item in allEntities) checkedListBoxTables.Items.Add(item, false);
                    string source = solutionUniqueName == null ? "organization (all tables)" : $"solution '{solutionUniqueName}'";
                    labelInfo.Text = $"Loaded {allEntities.Count} tables from {source}. Check tables and click 'Show Relationships'.";
                }
            });
        }

        private void checkedListBoxTables_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            if (_suppressItemCheck) return;
            BeginInvoke(new Action(() =>
            {
                var item = checkedListBoxTables.Items[e.Index] as EntityItem;
                if (item != null)
                {
                    if (e.NewValue == CheckState.Checked) _checkedTableLogicalNames.Add(item.LogicalName);
                    else _checkedTableLogicalNames.Remove(item.LogicalName);
                }
                UpdateSelectedTablesDisplay();
                tsbShowRelationships.Enabled = _checkedTableLogicalNames.Count > 0;
            }));
        }

        private void UpdateSelectedTablesDisplay()
        {
            listBoxSelectedTables.Items.Clear();
            if (_checkedTableLogicalNames.Count == 0) { labelSelectedTables.Text = "Selected Tables (0)"; listBoxSelectedTables.Items.Add("(No tables selected)"); return; }
            labelSelectedTables.Text = $"Selected Tables ({_checkedTableLogicalNames.Count})";
            foreach (var item in allEntities.Where(i => _checkedTableLogicalNames.Contains(i.LogicalName)).OrderBy(i => i.DisplayName))
                listBoxSelectedTables.Items.Add(item);
        }

        private void btnSelectAll_Click(object sender, EventArgs e)
        {
            if (checkedListBoxTables.Items.Count == 0) return;
            btnSelectAll.Enabled = false; btnDeselectAll.Enabled = false; tsbLoadTables.Enabled = false; tsbShowRelationships.Enabled = false;
            int totalItems = checkedListBoxTables.Items.Count; const int chunkSize = 10;
            _suppressItemCheck = true;
            var worker = new System.ComponentModel.BackgroundWorker(); worker.WorkerReportsProgress = true;
            worker.DoWork += (s, args) =>
            {
                for (int i = 0; i < totalItems; i += chunkSize)
                {
                    int end = Math.Min(i + chunkSize, totalItems);
                    checkedListBoxTables.Invoke(new Action(() => { checkedListBoxTables.BeginUpdate(); for (int j = i; j < end; j++) checkedListBoxTables.SetItemChecked(j, true); checkedListBoxTables.EndUpdate(); }));
                    worker.ReportProgress((int)((end / (double)totalItems) * 100), $"Selecting tables... {end} of {totalItems}");
                    System.Threading.Thread.Sleep(10);
                }
            };
            worker.ProgressChanged += (s, args) => { labelInfo.Text = args.UserState?.ToString(); };
            worker.RunWorkerCompleted += (s, args) =>
            {
                _suppressItemCheck = false;
                _checkedTableLogicalNames.Clear();
                foreach (var item in checkedListBoxTables.CheckedItems.Cast<EntityItem>()) _checkedTableLogicalNames.Add(item.LogicalName);
                UpdateSelectedTablesDisplay();
                tsbShowRelationships.Enabled = _checkedTableLogicalNames.Count > 0;
                labelInfo.Text = $"All {totalItems} table(s) selected. Click 'Show Relationships' to load.";
                btnSelectAll.Enabled = true; btnDeselectAll.Enabled = true; tsbLoadTables.Enabled = true;
            };
            worker.RunWorkerAsync();
        }

        private void btnDeselectAll_Click(object sender, EventArgs e)
        {
            if (checkedListBoxTables.CheckedItems.Count == 0) return;
            btnSelectAll.Enabled = false; btnDeselectAll.Enabled = false; tsbLoadTables.Enabled = false; tsbShowRelationships.Enabled = false;
            int totalItems = checkedListBoxTables.Items.Count; const int chunkSize = 10;
            _suppressItemCheck = true;
            var worker = new System.ComponentModel.BackgroundWorker(); worker.WorkerReportsProgress = true;
            worker.DoWork += (s, args) =>
            {
                for (int i = 0; i < totalItems; i += chunkSize)
                {
                    int end = Math.Min(i + chunkSize, totalItems);
                    checkedListBoxTables.Invoke(new Action(() => { checkedListBoxTables.BeginUpdate(); for (int j = i; j < end; j++) checkedListBoxTables.SetItemChecked(j, false); checkedListBoxTables.EndUpdate(); }));
                    worker.ReportProgress((int)((end / (double)totalItems) * 100), $"Deselecting tables... {end} of {totalItems}");
                    System.Threading.Thread.Sleep(10);
                }
            };
            worker.ProgressChanged += (s, args) => { labelInfo.Text = args.UserState?.ToString(); };
            worker.RunWorkerCompleted += (s, args) =>
            {
                _suppressItemCheck = false; _checkedTableLogicalNames.Clear(); UpdateSelectedTablesDisplay();
                tsbShowRelationships.Enabled = false; labelInfo.Text = "All tables deselected.";
                btnSelectAll.Enabled = true; btnDeselectAll.Enabled = true; tsbLoadTables.Enabled = true;
            };
            worker.RunWorkerAsync();
        }

        private void txtFilterTables_TextChanged(object sender, EventArgs e)
        {
            if (allEntities == null) return;
            string filter = txtFilterTables.Text.ToLowerInvariant();
            checkedListBoxTables.Items.Clear();
            var filtered = string.IsNullOrWhiteSpace(filter) ? allEntities : allEntities.Where(item => item.DisplayName.ToLowerInvariant().Contains(filter) || item.LogicalName.ToLowerInvariant().Contains(filter));
            foreach (var item in filtered) checkedListBoxTables.Items.Add(item, _checkedTableLogicalNames.Contains(item.LogicalName));
        }

        private void tsbShowRelationships_Click(object sender, EventArgs e)
        {
            if (_checkedTableLogicalNames.Count == 0) { MessageBox.Show("Please check at least one table.", "No Tables Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            var checkedTables = allEntities.Where(item => _checkedTableLogicalNames.Contains(item.LogicalName)).ToList();
            _tableTabs.Clear(); tabControlTables.TabPages.Clear();
            tsbShowRelationships.Enabled = false; tsbLoadTables.Enabled = false; tsbExportSelected.Enabled = false; tsbExportAllTables.Enabled = false;
            labelInfo.Text = $"Loading relationships for {checkedTables.Count} table(s)...";

            WorkAsync(new WorkAsyncInfo
            {
                Message = $"Loading relationships for {checkedTables.Count} table(s)...",
                IsCancelable = false,
                MessageWidth = 340,
                MessageHeight = 150,
                Work = (worker, args) =>
                {
                    var results = new List<EntityMetadata>();
                    int total = checkedTables.Count;
                    for (int i = 0; i < total; i++)
                    {
                        var entityItem = checkedTables[i];
                        try
                        {
                            worker.ReportProgress((int)((i / (double)total) * 100), $"Loading {i + 1} of {total}: {entityItem.DisplayName}...");
                            var response = (RetrieveEntityResponse)Service.Execute(new RetrieveEntityRequest { LogicalName = entityItem.LogicalName, EntityFilters = EntityFilters.Relationships | EntityFilters.Entity, RetrieveAsIfPublished = false });
                            results.Add(response.EntityMetadata);
                        }
                        catch (Exception ex) { LogError($"Error loading {entityItem.LogicalName}: {ex.Message}", ex); }
                    }
                    args.Result = results;
                },
                ProgressChanged = (args) => { string msg = args.UserState?.ToString(); labelInfo.Text = msg; SetWorkingMessage(msg); },
                PostWorkCallBack = (args) =>
                {
                    tsbShowRelationships.Enabled = true; tsbLoadTables.Enabled = true; tsbExportSelected.Enabled = true; tsbExportAllTables.Enabled = true;
                    if (args.Error != null) { MessageBox.Show($"Error:\n{args.Error.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); labelInfo.Text = "Error loading relationships."; return; }
                    var results = args.Result as List<EntityMetadata>;
                    if (results == null || results.Count == 0) { labelInfo.Text = "No relationship data returned."; return; }
                    tabControlTables.SuspendLayout();
                    int created = 0;
                    foreach (var entity in results)
                    {
                        try { CreateTableTab(entity); created++; if (created % 5 == 0) Application.DoEvents(); }
                        catch (Exception ex) { LogError($"Error creating tab for {entity.LogicalName}: {ex.Message}", ex); }
                    }
                    tabControlTables.ResumeLayout();
                    labelInfo.Text = $"Loaded {created} of {checkedTables.Count} table(s). Click tabs to view relationships.";
                }
            });
        }

        private void CreateTableTab(EntityMetadata entity)
        {
            var displayName = entity.DisplayName?.UserLocalizedLabel?.Label ?? entity.LogicalName;
            var tableTab = new TabPage($"{displayName} ({entity.LogicalName})") { Name = $"tab_{entity.LogicalName}", Tag = entity };
            var nestedTabControl = new TabControl { Dock = DockStyle.Fill, Font = new Font("Segoe UI", 9F) };
            var tab1N = new TabPage("1:N Relationships"); var tabN1 = new TabPage("N:1 Relationships"); var tabNN = new TabPage("N:N Relationships");
            var dgv1N = CreateDataGridView(); var dgvN1 = CreateDataGridView(); var dgvNN = CreateDataGridView();
            tab1N.Controls.Add(dgv1N); tabN1.Controls.Add(dgvN1); tabNN.Controls.Add(dgvNN);
            nestedTabControl.TabPages.Add(tab1N); nestedTabControl.TabPages.Add(tabN1); nestedTabControl.TabPages.Add(tabNN);
            tableTab.Controls.Add(nestedTabControl); tabControlTables.TabPages.Add(tableTab);
            _tableTabs[entity.LogicalName] = new TableTabData { Entity = entity, TableTab = tableTab, NestedTabControl = nestedTabControl, DataGridView1N = dgv1N, DataGridViewN1 = dgvN1, DataGridViewNN = dgvNN };
            Display1NRelationships(entity, dgv1N, tab1N); DisplayN1Relationships(entity, dgvN1, tabN1); DisplayNNRelationships(entity, dgvNN, tabNN);
        }

        private DataGridView CreateDataGridView() => new DataGridView
        {
            Dock = DockStyle.Fill,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells,
            AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.LightGray },
            EnableHeadersVisualStyles = false,
            ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle { Font = new Font("Segoe UI", 9F, FontStyle.Bold), BackColor = Color.Navy, ForeColor = Color.White },
            ReadOnly = true,
            Font = new Font("Segoe UI", 9F)
        };

        private void Display1NRelationships(EntityMetadata entity, DataGridView dgv, TabPage tab)
        {
            var dt = new DataTable();
            foreach (var c in new[] { "Relationship Name", "Referencing Table", "Referencing Attribute", "Relationship Type", "Is Customizable", "Delete", "Assign", "Share", "Unshare", "Reparent", "Merge", "Rollup View" }) dt.Columns.Add(c);
            if (entity.OneToManyRelationships != null)
            {
                IEnumerable<OneToManyRelationshipMetadata> rels = entity.OneToManyRelationships.OrderBy(r => r.SchemaName);
                if (_relationshipFilter != null) rels = rels.Where(r => _relationshipFilter.Contains(r.MetadataId.GetValueOrDefault()));
                foreach (var rel in rels)
                    dt.Rows.Add(rel.SchemaName, rel.ReferencingEntity, rel.ReferencingAttribute, GetRelationshipType(rel.CascadeConfiguration), GetBooleanDisplay(rel.IsCustomizable?.Value),
                        GetCascadeBehavior(rel.CascadeConfiguration?.Delete), GetCascadeBehavior(rel.CascadeConfiguration?.Assign), GetCascadeBehavior(rel.CascadeConfiguration?.Share),
                        GetCascadeBehavior(rel.CascadeConfiguration?.Unshare), GetCascadeBehavior(rel.CascadeConfiguration?.Reparent), GetCascadeBehavior(rel.CascadeConfiguration?.Merge), GetCascadeBehavior(rel.CascadeConfiguration?.RollupView));
            }
            dgv.DataSource = dt; tab.Text = $"1:N Relationships ({dt.Rows.Count})";
        }

        private void DisplayN1Relationships(EntityMetadata entity, DataGridView dgv, TabPage tab)
        {
            var dt = new DataTable();
            foreach (var c in new[] { "Relationship Name", "Referenced Table", "Referencing Attribute", "Relationship Type", "Is Customizable", "Delete", "Assign", "Share", "Unshare", "Reparent", "Merge", "Rollup View" }) dt.Columns.Add(c);
            if (entity.ManyToOneRelationships != null)
            {
                IEnumerable<OneToManyRelationshipMetadata> rels = entity.ManyToOneRelationships.OrderBy(r => r.SchemaName);
                if (_relationshipFilter != null) rels = rels.Where(r => _relationshipFilter.Contains(r.MetadataId.GetValueOrDefault()));
                foreach (var rel in rels)
                    dt.Rows.Add(rel.SchemaName, rel.ReferencedEntity, rel.ReferencingAttribute, GetRelationshipType(rel.CascadeConfiguration), GetBooleanDisplay(rel.IsCustomizable?.Value),
                        GetCascadeBehavior(rel.CascadeConfiguration?.Delete), GetCascadeBehavior(rel.CascadeConfiguration?.Assign), GetCascadeBehavior(rel.CascadeConfiguration?.Share),
                        GetCascadeBehavior(rel.CascadeConfiguration?.Unshare), GetCascadeBehavior(rel.CascadeConfiguration?.Reparent), GetCascadeBehavior(rel.CascadeConfiguration?.Merge), GetCascadeBehavior(rel.CascadeConfiguration?.RollupView));
            }
            dgv.DataSource = dt; tab.Text = $"N:1 Relationships ({dt.Rows.Count})";
        }

        private void DisplayNNRelationships(EntityMetadata entity, DataGridView dgv, TabPage tab)
        {
            var dt = new DataTable();
            foreach (var c in new[] { "Relationship Name", "Intersect Table", "Entity 1", "Entity 1 Attribute", "Entity 2", "Entity 2 Attribute", "Is Customizable", "Is Managed" }) dt.Columns.Add(c);
            if (entity.ManyToManyRelationships != null)
            {
                IEnumerable<ManyToManyRelationshipMetadata> rels = entity.ManyToManyRelationships.OrderBy(r => r.SchemaName);
                if (_relationshipFilter != null) rels = rels.Where(r => _relationshipFilter.Contains(r.MetadataId.GetValueOrDefault()));
                foreach (var rel in rels)
                    dt.Rows.Add(rel.SchemaName, rel.IntersectEntityName, rel.Entity1LogicalName, rel.Entity1IntersectAttribute, rel.Entity2LogicalName, rel.Entity2IntersectAttribute, GetBooleanDisplay(rel.IsCustomizable?.Value), GetBooleanDisplay(rel.IsManaged));
            }
            dgv.DataSource = dt; tab.Text = $"N:N Relationships ({dt.Rows.Count})";
        }

        private void tsbExportSelected_Click(object sender, EventArgs e)
        {
            if (_checkedTableLogicalNames.Count == 0) { MessageBox.Show("Please check at least one table to export.", "No Tables Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            var checkedTables = allEntities.Where(item => _checkedTableLogicalNames.Contains(item.LogicalName)).ToList();
            var dlg = new SaveFileDialog { Filter = "Excel Files|*.xlsx", Title = "Export Selected Tables to Excel", FileName = $"SelectedTables_Relationships_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx", DefaultExt = "xlsx", InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) };
            if (dlg.ShowDialog() == DialogResult.OK) ExecuteMethod(() => ExportAllTablesToExcel(dlg.FileName, checkedTables.Select(i => i.Metadata).ToList()));
        }

        private void tsbExportAllTables_Click(object sender, EventArgs e)
        {
            if (allEntities == null || allEntities.Count == 0) { MessageBox.Show("Please load tables first.", "No Data", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            var dlg = new SaveFileDialog { Filter = "Excel Files|*.xlsx", Title = "Export All Tables to Excel", FileName = $"AllTables_Relationships_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx", DefaultExt = "xlsx", InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) };
            if (dlg.ShowDialog() == DialogResult.OK) ExecuteMethod(() => ExportAllTablesToExcel(dlg.FileName, allEntities.Select(i => i.Metadata).ToList()));
        }

        private void ExportAllTablesToExcel(string filePath, List<EntityMetadata> entitiesToExport)
        {
            Excel.Application excelApp = null; Excel.Workbook workbook = null;
            tsbExportSelected.Enabled = false; tsbExportAllTables.Enabled = false; tsbShowRelationships.Enabled = false;

            WorkAsync(new WorkAsyncInfo
            {
                Message = "Preparing export...",
                IsCancelable = false,
                MessageWidth = 340,
                MessageHeight = 150,
                Work = (worker, args) =>
                {
                    try
                    {
                        excelApp = new Excel.Application { Visible = false, DisplayAlerts = false };
                        workbook = excelApp.Workbooks.Add();
                        while (workbook.Worksheets.Count > 1) ((Excel.Worksheet)workbook.Worksheets[1]).Delete();

                        int tableCount = 0, processedCount = 0, totalTables = entitiesToExport.Count;
                        bool isFirstSheet = true;
                        var usedSheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                        foreach (var entity in entitiesToExport)
                        {
                            try
                            {
                                processedCount++;
                                string progressMsg = $"Exporting {processedCount} of {totalTables}: {entity.LogicalName}...";
                                worker.ReportProgress((int)((processedCount / (double)totalTables) * 100), progressMsg);

                                var response = (RetrieveEntityResponse)Service.Execute(new RetrieveEntityRequest { LogicalName = entity.LogicalName, EntityFilters = EntityFilters.Relationships | EntityFilters.Entity, RetrieveAsIfPublished = false });
                                var fullEntity = response.EntityMetadata;
                                var dt = BuildCombinedRelationshipsTable(fullEntity, _relationshipFilter);

                                if (dt.Rows.Count > 0)
                                {
                                    Excel.Worksheet sheet;
                                    if (isFirstSheet) { sheet = (Excel.Worksheet)workbook.Worksheets[1]; isFirstSheet = false; }
                                    else sheet = (Excel.Worksheet)workbook.Worksheets.Add(After: workbook.Worksheets[workbook.Worksheets.Count]);

                                    var displayName = SanitizeSheetName(fullEntity.DisplayName?.UserLocalizedLabel?.Label ?? fullEntity.LogicalName ?? "Table");
                                    var logicalName = SanitizeSheetName(fullEntity.LogicalName ?? "table");
                                    string baseSheetName = $"{displayName} ({logicalName})";
                                    if (baseSheetName.Length > 31)
                                    {
                                        string schemaName = $"({logicalName})"; int avail = 31 - schemaName.Length - 3;
                                        baseSheetName = avail > 3 ? displayName.Substring(0, Math.Min(displayName.Length, avail)) + "..." + schemaName : logicalName.Length <= 31 ? logicalName : logicalName.Substring(0, 31);
                                    }
                                    baseSheetName = SanitizeSheetName(baseSheetName);
                                    string sheetName = baseSheetName; int counter = 1;
                                    while (usedSheetNames.Contains(sheetName)) { string suffix = $"_{counter}"; int maxBase = 31 - suffix.Length; sheetName = maxBase > 0 ? baseSheetName.Substring(0, Math.Min(baseSheetName.Length, maxBase)) + suffix : $"Table{counter}"; counter++; }
                                    usedSheetNames.Add(sheetName); sheet.Name = sheetName;
                                    ExportDataTableToSheet(sheet, dt); tableCount++;
                                }
                            }
                            catch (Exception entityEx) { LogError($"Error processing {entity.LogicalName}: {entityEx.Message}", entityEx); }
                        }

                        if (tableCount == 0) { var sheet = (Excel.Worksheet)workbook.Worksheets[1]; sheet.Name = "No Data"; sheet.Cells[1, 1] = "No relationships found for any tables."; }
                        worker.ReportProgress(100, "Saving file...");
                        workbook.SaveAs(filePath);
                        args.Result = new { FilePath = filePath, TableCount = tableCount };
                    }
                    finally
                    {
                        workbook?.Close(false);
                        if (workbook != null) ReleaseObject(workbook);
                        if (excelApp != null) { excelApp.Quit(); ReleaseObject(excelApp); }
                        GC.Collect(); GC.WaitForPendingFinalizers();
                    }
                },
                ProgressChanged = (args) => { string msg = args.UserState?.ToString(); labelInfo.Text = msg; SetWorkingMessage(msg); },
                PostWorkCallBack = (args) =>
                {
                    tsbExportSelected.Enabled = true; tsbExportAllTables.Enabled = true; tsbShowRelationships.Enabled = _checkedTableLogicalNames.Count > 0;
                    if (args.Error != null) { labelInfo.Text = "Export failed."; MessageBox.Show($"Export failed:\n{args.Error.Message}", "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
                    var result = args.Result as dynamic; string path = result.FilePath; int count = result.TableCount;
                    labelInfo.Text = $"Export complete. {count} table(s) exported to: {System.IO.Path.GetFileName(path)}";
                    if (MessageBox.Show($"Export completed!\n\nExported {count} table(s).\n\nSaved to:\n{path}\n\nOpen the file?", "Export Complete", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                        try { System.Diagnostics.Process.Start(path); } catch (Exception ex) { MessageBox.Show($"Could not open file:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
                }
            });
        }

        private DataTable BuildCombinedRelationshipsTable(EntityMetadata entity, HashSet<Guid> relationshipFilter)
        {
            var dt = new DataTable();
            foreach (var c in new[] { "Relationship Type", "Relationship Name", "Related Table", "Attribute", "Cascade Delete", "Cascade Assign", "Cascade Share", "Cascade Unshare", "Cascade Reparent", "Cascade Merge", "Is Customizable", "Is Managed" }) dt.Columns.Add(c);

            if (entity.OneToManyRelationships != null)
            {
                IEnumerable<OneToManyRelationshipMetadata> rels = entity.OneToManyRelationships.OrderBy(r => r.SchemaName);
                if (relationshipFilter != null) rels = rels.Where(r => relationshipFilter.Contains(r.MetadataId.GetValueOrDefault()));
                foreach (var rel in rels) dt.Rows.Add("1:N", rel.SchemaName, rel.ReferencingEntity, rel.ReferencingAttribute, GetCascadeBehavior(rel.CascadeConfiguration?.Delete), GetCascadeBehavior(rel.CascadeConfiguration?.Assign), GetCascadeBehavior(rel.CascadeConfiguration?.Share), GetCascadeBehavior(rel.CascadeConfiguration?.Unshare), GetCascadeBehavior(rel.CascadeConfiguration?.Reparent), GetCascadeBehavior(rel.CascadeConfiguration?.Merge), GetBooleanDisplay(rel.IsCustomizable?.Value), GetBooleanDisplay(rel.IsManaged));
            }
            if (entity.ManyToOneRelationships != null)
            {
                IEnumerable<OneToManyRelationshipMetadata> rels = entity.ManyToOneRelationships.OrderBy(r => r.SchemaName);
                if (relationshipFilter != null) rels = rels.Where(r => relationshipFilter.Contains(r.MetadataId.GetValueOrDefault()));
                foreach (var rel in rels) dt.Rows.Add("N:1", rel.SchemaName, rel.ReferencedEntity, rel.ReferencingAttribute, GetCascadeBehavior(rel.CascadeConfiguration?.Delete), GetCascadeBehavior(rel.CascadeConfiguration?.Assign), GetCascadeBehavior(rel.CascadeConfiguration?.Share), GetCascadeBehavior(rel.CascadeConfiguration?.Unshare), GetCascadeBehavior(rel.CascadeConfiguration?.Reparent), GetCascadeBehavior(rel.CascadeConfiguration?.Merge), GetBooleanDisplay(rel.IsCustomizable?.Value), GetBooleanDisplay(rel.IsManaged));
            }
            if (entity.ManyToManyRelationships != null)
            {
                IEnumerable<ManyToManyRelationshipMetadata> rels = entity.ManyToManyRelationships.OrderBy(r => r.SchemaName);
                if (relationshipFilter != null) rels = rels.Where(r => relationshipFilter.Contains(r.MetadataId.GetValueOrDefault()));
                foreach (var rel in rels) { string related = rel.Entity1LogicalName == entity.LogicalName ? rel.Entity2LogicalName : rel.Entity1LogicalName; dt.Rows.Add("N:N", rel.SchemaName, related, rel.IntersectEntityName, "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", GetBooleanDisplay(rel.IsCustomizable?.Value), GetBooleanDisplay(rel.IsManaged)); }
            }
            return dt;
        }

        private void ExportDataTableToSheet(Excel.Worksheet sheet, DataTable dt)
        {
            for (int col = 0; col < dt.Columns.Count; col++) sheet.Cells[1, col + 1] = dt.Columns[col].ColumnName;
            Excel.Range headerRange = sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, dt.Columns.Count]];
            headerRange.Font.Bold = true; headerRange.Interior.Color = Excel.XlRgbColor.rgbNavy; headerRange.Font.Color = Excel.XlRgbColor.rgbWhite; headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            for (int row = 0; row < dt.Rows.Count; row++) for (int col = 0; col < dt.Columns.Count; col++) sheet.Cells[row + 2, col + 1] = dt.Rows[row][col]?.ToString() ?? "";
            sheet.Columns.AutoFit();
            sheet.Range[sheet.Cells[1, 1], sheet.Cells[dt.Rows.Count + 1, dt.Columns.Count]].AutoFilter(1);
            sheet.Application.ActiveWindow.SplitRow = 1; sheet.Application.ActiveWindow.FreezePanes = true;
        }

        private void ReleaseObject(object obj) { try { System.Runtime.InteropServices.Marshal.ReleaseComObject(obj); } catch { } }
        private void tsbEdit_Click(object sender, EventArgs e) { }
        private void tsbSave_Click(object sender, EventArgs e) { }
        private void tsbHelp_Click(object sender, EventArgs e) { ShowAboutDialog(); }

        private string GetRelationshipType(CascadeConfiguration config)
        {
            if (config == null) return "Unknown";
            if (config.Delete == CascadeType.Cascade && config.Assign == CascadeType.Cascade && config.Share == CascadeType.Cascade && config.Unshare == CascadeType.Cascade && config.Reparent == CascadeType.Cascade) return "Parental";
            if (config.Assign == CascadeType.NoCascade && config.Share == CascadeType.NoCascade && config.Unshare == CascadeType.NoCascade && config.Reparent == CascadeType.NoCascade) return "Referential";
            return "Custom";
        }

        private string GetCascadeBehavior(CascadeType? cascadeType)
        {
            if (cascadeType == null) return "N/A";
            switch (cascadeType.Value)
            {
                case CascadeType.NoCascade: return "No Cascade";
                case CascadeType.Cascade: return "Cascade All";
                case CascadeType.Active: return "Cascade Active";
                case CascadeType.UserOwned: return "Cascade User Owned";
                case CascadeType.RemoveLink: return "Remove Link";
                case CascadeType.Restrict: return "Restrict";
                default: return cascadeType.Value.ToString();
            }
        }

        private string GetBooleanDisplay(bool? value) => !value.HasValue ? "N/A" : value.Value ? "Yes" : "No";

        private string SanitizeSheetName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "Sheet";
            char[] invalidChars = { '\\', '/', '?', '*', '[', ']', ':' };
            string s = name; foreach (char c in invalidChars) s = s.Replace(c, '_');
            s = s.Trim(); if (string.IsNullOrWhiteSpace(s)) return "Sheet";
            if (s.Equals("History", StringComparison.OrdinalIgnoreCase)) return "History_";
            return s.Length > 31 ? s.Substring(0, 31) : s;
        }

        private void MyPluginControl_OnCloseTool(object sender, EventArgs e) { SettingsManager.Instance.Save(GetType(), mySettings); }

        public override void UpdateConnection(IOrganizationService newService, ConnectionDetail detail, string actionName, object parameter)
        {
            base.UpdateConnection(newService, detail, actionName, parameter);
            if (mySettings != null && detail != null) { mySettings.LastUsedOrganizationWebappUrl = detail.WebApplicationUrl; LogInfo("Connection changed to: {0}", detail.WebApplicationUrl); }
        }

        private class EntityItem { public string DisplayName { get; set; } public string LogicalName { get; set; } public EntityMetadata Metadata { get; set; } public override string ToString() => $"{DisplayName} ({LogicalName})"; }
        private class TableTabData { public EntityMetadata Entity { get; set; } public TabPage TableTab { get; set; } public TabControl NestedTabControl { get; set; } public DataGridView DataGridView1N { get; set; } public DataGridView DataGridViewN1 { get; set; } public DataGridView DataGridViewNN { get; set; } }
    }

    public class SolutionPickerDialog : Form
    {
        public SolutionItem SelectedSolution { get; private set; }
        private ListView lvSolutions; private Button btnOk; private Button btnCancel;

        public SolutionPickerDialog(DataCollection<Entity> solutions) { BuildLayout(); PopulateList(solutions); }

        private void BuildLayout()
        {
            Text = "Solutions Picker"; Size = new Size(700, 520); MinimumSize = new Size(520, 400);
            StartPosition = FormStartPosition.CenterParent; FormBorderStyle = FormBorderStyle.FixedDialog; MaximizeBox = false; MinimizeBox = false; Font = new Font("Segoe UI", 9F);
            var lblTitle = new System.Windows.Forms.Label { Text = "Select a solution to load tables from:", Dock = DockStyle.Top, Height = 38, Padding = new Padding(10, 10, 0, 0), Font = new Font("Segoe UI", 10F, FontStyle.Bold) };
            lvSolutions = new ListView { Dock = DockStyle.Fill, View = View.Details, FullRowSelect = true, GridLines = true, MultiSelect = false, HideSelection = false, Font = new Font("Segoe UI", 9F) };
            lvSolutions.Columns.Add("Display Name", 300); lvSolutions.Columns.Add("Version", 90); lvSolutions.Columns.Add("Publisher", 220);
            lvSolutions.MouseDoubleClick += (s, e) => AcceptSelection();
            lvSolutions.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) AcceptSelection(); };
            var separator = new Panel { Dock = DockStyle.Bottom, Height = 1, BackColor = Color.LightGray };
            var pnlButtons = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 50, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(8), WrapContents = false };
            btnCancel = new Button { Text = "Cancel", Width = 90, Height = 30, DialogResult = DialogResult.Cancel };
            btnOk = new Button { Text = "OK", Width = 90, Height = 30 }; btnOk.Click += (s, e) => AcceptSelection();
            pnlButtons.Controls.Add(btnCancel); pnlButtons.Controls.Add(btnOk);
            AcceptButton = btnOk; CancelButton = btnCancel;
            Controls.Add(lvSolutions); Controls.Add(separator); Controls.Add(pnlButtons); Controls.Add(lblTitle);
        }

        private void PopulateList(DataCollection<Entity> solutions)
        {
            lvSolutions.Items.Clear();
            foreach (var sol in solutions)
            {
                var friendlyName = sol.GetAttributeValue<string>("friendlyname") ?? "(unnamed)";
                var uniqueName = sol.GetAttributeValue<string>("uniquename") ?? "";
                var version = sol.GetAttributeValue<string>("version") ?? "";
                string publisher = string.Empty;
                if (sol.Contains("pub.friendlyname") && sol["pub.friendlyname"] is AliasedValue av && av.Value is string pubName) publisher = pubName;
                var item = new ListViewItem(friendlyName) { Tag = new SolutionItem { FriendlyName = friendlyName, UniqueName = uniqueName, Version = version, Publisher = publisher } };
                item.SubItems.Add(version); item.SubItems.Add(publisher); lvSolutions.Items.Add(item);
            }
            if (lvSolutions.Items.Count > 0) lvSolutions.Items[0].Selected = true;
        }

        private void AcceptSelection()
        {
            if (lvSolutions.SelectedItems.Count == 0) { MessageBox.Show("Please select a solution.", "Selection Required", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            SelectedSolution = (SolutionItem)lvSolutions.SelectedItems[0].Tag; DialogResult = DialogResult.OK; Close();
        }
    }

    public class SolutionItem { public string FriendlyName { get; set; } public string UniqueName { get; set; } public string Version { get; set; } public string Publisher { get; set; } public override string ToString() => FriendlyName; }
}
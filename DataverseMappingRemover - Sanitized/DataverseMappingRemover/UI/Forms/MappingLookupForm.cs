using DataverseMappingRemover.Services;
using DataverseMappingRemover.UI.Components;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Metadata.Query;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;

namespace DataverseMappingRemover.UI.Forms
{
    public partial class MappingLookupForm : Form
    {
        public string SourceEntityLogicalName { get; private set; } = string.Empty;
        public string SourceAttributeLogicalName { get; private set; } = string.Empty;

        private ComboBox sourceEntityComboBox = default!;
        private ComboBox sourceAttributeComboBox = default!;
        private RoundedButton submitButton = default!;
        private RoundedButton exitButton = default!;
        private RoundedButton deleteButton = default!;
        private DataGridView gridMappings = default!;
        private StatusStrip statusStrip = default!;
        private ToolStripStatusLabel statusLabel = default!;
        private readonly Microsoft.Xrm.Sdk.IOrganizationService _organizationService;
        private CancellationTokenSource? _ctsLookup;
        private readonly AutoCompleteStringCollection _entityAutoComplete = new();
        private readonly AutoCompleteStringCollection _attributeAutoComplete = new();
        private readonly Dictionary<string, string> _entitiesCache = new(StringComparer.OrdinalIgnoreCase);

        private readonly Dictionary<string, Dictionary<string, string>> _attributesDisplayCache = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, List<AttributeItem>> _attributesItemsCache = new(StringComparer.OrdinalIgnoreCase);
        private string _metadataversionstamp = string.Empty;
        private const int EntityBatchSize = 100;
        private BindingList<MappingResult>? _gridData;
        private static IEnumerable<List<T>> Chunk<T>(IEnumerable<T> source, int chunkSize)
        {
            var chunk = new List<T>(chunkSize);
            foreach (var item in source)
            {
                chunk.Add(item);
                if (chunk.Count == chunkSize)
                {
                    yield return chunk;
                    chunk = new List<T>(chunkSize);
                }
            }
            if (chunk.Count > 0)
            {
                yield return chunk;
            }
        }
        private string GetEntityDisplayName(string entityLogicalName) => 
            _entitiesCache.TryGetValue(entityLogicalName, out var displayName) 
            ? displayName : entityLogicalName;


        private string GetAttributeDisplayName(string entityLogicalName, string attributeLogicalName)
        {
            if (_attributesDisplayCache.TryGetValue(entityLogicalName, out var attributes) && 
                attributes.TryGetValue(attributeLogicalName, out var label))
            { 
                return label; 
            }
            return attributeLogicalName;
        }

        private CancellationTokenSource? _ctsMetadata;

        private class EntityItem
        {
            public string LogicalName { get; set; } = string.Empty;
            public string DisplayName { get; set; } = string.Empty;
            public override string ToString() => DisplayName;
        }

        private class AttributeItem
        {
            public string LogicalName { get; set; } = string.Empty;
            public string DisplayName { get; set; } = string.Empty;
            public override string ToString() => $"{DisplayName} ({LogicalName})";
        }

        internal static class WindowHelpers
        {
            [DllImport("user32.dll")]
            private static extern bool SetForegroundWindow(IntPtr hWnd);

            [DllImport("user32.dll")]
            private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
            private const int SW_SHOWNORMAL = 1;
            public static void ForceToForeground(Form form)
            {
                if (form == null) return;
                ShowWindow(form.Handle, SW_SHOWNORMAL);
                SetForegroundWindow(form.Handle);
                form.TopMost = true;
                form.BringToFront();
                form.Activate();
            }
        }

        private async Task OnSubmitAsync()
        {
            submitButton.Enabled = false;
            exitButton.Enabled = false;

            try
            {
                statusLabel.Text = "Sanitizing and validating input ...";
                SourceEntityLogicalName = (sourceEntityComboBox.SelectedValue as string ??
                    sourceEntityComboBox.Text ?? "").Trim().ToLowerInvariant();
                SourceAttributeLogicalName = (sourceAttributeComboBox.SelectedValue as string ??
                    sourceAttributeComboBox.Text ?? "").Trim().ToLowerInvariant();
                if (SourceEntityLogicalName.Contains(" "))
                {
                    MessageBox.Show("Entity logical name cannot contain spaces.", "Input Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    sourceEntityComboBox.Focus();
                    return;
                }
                if (SourceAttributeLogicalName.Contains(" "))
                {
                    MessageBox.Show("Field logical name cannot contain spaces.", "Input Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    sourceAttributeComboBox.Focus();
                    return;
                }
                if (string.IsNullOrWhiteSpace(SourceEntityLogicalName))
                {
                    MessageBox.Show("Entity logical name must be provided.", "Input Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    sourceEntityComboBox.Focus();
                    return;
                }
                else if (string.IsNullOrWhiteSpace(SourceAttributeLogicalName))
                {
                    MessageBox.Show("Field logical name must be provided.", "Input Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    sourceAttributeComboBox.Focus();
                    return;
                }

                _ctsLookup?.Dispose();
                _ctsLookup = new CancellationTokenSource();
                var token = _ctsLookup.Token;

                gridMappings.DataSource = null;

                statusLabel.Text = "Searching for mappings ...";
                var inspector = new Services.MappingInspectorService(_organizationService);
                var result = await inspector.FindAttributeMappingsAsync(SourceEntityLogicalName,
                    SourceAttributeLogicalName, token);

                _gridData = new BindingList<MappingResult>(result);

                if (!IsDisposed && gridMappings.IsHandleCreated)
                {
                    gridMappings.DataSource = _gridData;

                    if (result.Count == 1)
                    {
                        statusLabel.Text = "Found 1 mapping";
                    }
                    else if (result.Count == 0)
                    {
                        statusLabel.Text = "No mappings found";
                    }
                    else
                    {
                        statusLabel.Text = $"Found {result.Count} mappings";
                    }
                }
            }
            catch (OperationCanceledException)
            {
                statusLabel.Text = "Search canceled.";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                submitButton.Enabled = true;
                exitButton.Enabled = true;
            }
        }

        public MappingLookupForm(Microsoft.Xrm.Sdk.IOrganizationService organizationService)
        {
            _organizationService = organizationService ?? throw new ArgumentNullException(nameof(organizationService));

            Text = "Field Mappings Lookup";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = true;
            MinimizeBox = true;
            MinimumSize = new Size(700, 400);
   
            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 4,
                Padding = new Padding(12),
            };

            layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            var buttonsLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 1,
                AutoSize = false,
                Margin = new Padding(0, 4, 0, 0),
                Padding = new Padding(0),
                GrowStyle = TableLayoutPanelGrowStyle.FixedSize
            };

            buttonsLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33f));
            buttonsLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33f));
            buttonsLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33f));
            buttonsLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var entityLabel = new System.Windows.Forms.Label
            {
                Text = "Enter the Logical Name of the Entity:",
                AutoSize = true,
                Dock = DockStyle.Fill,
                TabStop = false
            };

            var attributeLabel = new System.Windows.Forms.Label
            {
                Text = "Enter the Logical Name of the Field:",
                AutoSize = true,
                Dock = DockStyle.Fill,
                TabStop = false
            };

            sourceEntityComboBox = new ComboBox
            {
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDown,
                AutoCompleteMode = AutoCompleteMode.SuggestAppend,
                AutoCompleteSource = AutoCompleteSource.CustomSource,
                TabIndex = 0,
                TabStop = true
            };

            sourceAttributeComboBox = new ComboBox
            {
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDown,
                AutoCompleteMode = AutoCompleteMode.SuggestAppend,
                AutoCompleteSource = AutoCompleteSource.CustomSource,
                TabIndex = 1,
                TabStop = true
            };

            submitButton = new RoundedButton
            {
                Text = "Search",
                AutoSize = true,
                Anchor = AnchorStyles.Left,
                Margin = new Padding(8),
                FlatStyle = FlatStyle.Flat,
                ForeColor = Color.White,
                BackColor = Color.FromArgb(28, 151, 234),
                TabIndex = 2,
                TabStop = true,
                Enabled = false
            };

            submitButton.Click += async (s, e) => await OnSubmitAsync();

            exitButton = new RoundedButton
            {
                Text = "Exit",
                AutoSize = true,
                Anchor = AnchorStyles.None,
                Margin = new Padding(8),
                FlatStyle = FlatStyle.Flat,
                ForeColor = Color.White,
                BackColor = Color.FromArgb(200, 50, 50),
                TabIndex = 3,
                TabStop = true
            };

            exitButton.Click += (_, __) =>
            {
                _ctsLookup?.Cancel();
                DialogResult = DialogResult.Cancel;
                Close();
            };

            deleteButton = new RoundedButton
            {
                Text = "Delete Mapping",
                AutoSize = true,
                Anchor = AnchorStyles.Right,
                Margin = new Padding(8),
                FlatStyle = FlatStyle.Flat,
                ForeColor = Color.White,
                BackColor = Color.FromArgb(187, 0, 18),
                TabIndex = 5,
                TabStop = true,
                Enabled = false
            };

            deleteButton.Click += async (s, e) => await OnDeleteButtonClick();

            gridMappings = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                TabIndex = 4,
                TabStop = true,
                AutoGenerateColumns = false
            };

            gridMappings.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "SourceEntity",
                HeaderText = "Source Entity",
                Name = "colSourceEntity",
                ReadOnly = true,
                MinimumWidth = 120
            });

            gridMappings.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "SourceAttribute",
                HeaderText = "Source Field",
                Name = "colSourceAttribute",
                ReadOnly = true,
                MinimumWidth = 140
            });

            gridMappings.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "TargetEntity",
                HeaderText = "Target Entity",
                Name = "colTargetEntity",
                ReadOnly = true,
                MinimumWidth = 120
            });

            gridMappings.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "TargetAttribute",
                HeaderText = "Target Field",
                Name = "colTargetAttribute",
                ReadOnly = true,
                MinimumWidth = 140
            });

            gridMappings.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "EntityMapId",
                HeaderText = "Entity Mapping ID",
                Name = "colEntityMapId",
                ReadOnly = true,
                MinimumWidth = 200
            }); 
            
            gridMappings.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "AttributeMapId",
                HeaderText = "Attribute Mapping ID",
                Name = "colAttributeMapId",
                ReadOnly = true,
                MinimumWidth = 200
            });

            gridMappings.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "Direction",
                HeaderText = "Mapping Direction",
                Name = "colDirection",
                ReadOnly = true,
                MinimumWidth = 120
            });

            statusStrip = new StatusStrip();
            statusLabel = new ToolStripStatusLabel
            {
                Text = "Ready"
            };
            statusStrip.Items.Add(statusLabel);

            var progressBar = new ToolStripProgressBar
            {
                Name = "statusProgressBar",
                Size = new System.Drawing.Size(100, 16),
                Style = ProgressBarStyle.Continuous,
                Visible = false
            };
            statusStrip.Items.Add(progressBar);

            var split = new SplitContainer
            {
                Orientation = Orientation.Horizontal,
                Dock = DockStyle.Fill,
                SplitterWidth = 6,
                IsSplitterFixed = false,
                Panel1MinSize = 0,
                Panel2MinSize = 100
            };

            buttonsLayout.Controls.Add(submitButton, 0, 0);
            buttonsLayout.Controls.Add(exitButton, 1, 0);
            buttonsLayout.Controls.Add(deleteButton, 2, 0);

            split.Panel1Collapsed = true;
            split.Panel2.Controls.Add(gridMappings);

            layout.Controls.Add(entityLabel, 0, 0);
            layout.Controls.Add(sourceEntityComboBox, 1, 0);
            layout.SetColumnSpan(sourceEntityComboBox, 2);
            layout.Controls.Add(attributeLabel, 0, 1);
            layout.Controls.Add(sourceAttributeComboBox, 1, 1);
            layout.SetColumnSpan(sourceAttributeComboBox, 2);
            layout.Controls.Add(buttonsLayout, 0, 2);
            layout.SetColumnSpan(buttonsLayout, 3);
            layout.Controls.Add(split, 0, 3);
            layout.SetColumnSpan(split, 3);

            Controls.Add(layout);
            Controls.Add(statusStrip);

            AcceptButton = submitButton;
            CancelButton = exitButton;

            gridMappings.SelectionChanged += (_, __) =>
            {
                deleteButton.Enabled = gridMappings.SelectedRows.Count == 1;
            };
        }

        protected override bool ShowFocusCues => true;

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);

            WindowHelpers.ForceToForeground(this);

            BeginInvoke(async () =>
            {
                await PopulateEntitiesAsync();
                sourceEntityComboBox.Focus();
            });
        }

        private async Task PopulateEntitiesAsync()
        {
            statusLabel.Text = "Loading entities ...";

            var progressBar = statusStrip.Items["statusProgressBar"] as ToolStripProgressBar;
            if (progressBar != null) 
            {
                progressBar.Visible = true;
                progressBar.Style = ProgressBarStyle.Marquee;
            }

            _ctsMetadata?.Dispose();
            _ctsMetadata = new CancellationTokenSource();
            var token = _ctsMetadata.Token;

            var entities = await Task.Run(() =>
            {
                token.ThrowIfCancellationRequested();

                var req = new Microsoft.Xrm.Sdk.Query.QueryExpression("entitymap")
                {
                    ColumnSet = new Microsoft.Xrm.Sdk.Query.ColumnSet("sourceentityname", "targetentityname"),
                };

                var resp = _organizationService.RetrieveMultiple(req);

                token.ThrowIfCancellationRequested();

                var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var entity in resp.Entities)
                {
                    var source = entity.GetAttributeValue<string>("sourceentityname");
                    var target = entity.GetAttributeValue<string>("targetentityname");
                    if (!string.IsNullOrEmpty(source))
                    {
                        set.Add(source);
                    }
                    if (!string.IsNullOrEmpty(target))
                    {
                        set.Add(target);
                    }
                }

                return set.ToList();
            }, token);

            statusLabel.Text = "Caching entity display names ...";

            await EnsureEntityDisplayNameCachedAsync(entities, token);
            
            var items = new List<EntityItem>(capacity: entities.Count);
            foreach (var logicalName in entities.OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
            {
                var displayName = GetEntityDisplayName(logicalName);

                items.Add(new EntityItem
                {
                    LogicalName = logicalName,
                    DisplayName = displayName
                });
            }

            sourceEntityComboBox.DataSource = items;
            sourceEntityComboBox.DisplayMember = nameof(EntityItem.DisplayName);
            sourceEntityComboBox.ValueMember = nameof(EntityItem.LogicalName);

            _entityAutoComplete.Clear();
            foreach (var entity in items)
            {
                _entityAutoComplete.Add(entity.DisplayName);
                _entityAutoComplete.Add(entity.LogicalName);
            }
            sourceEntityComboBox.AutoCompleteCustomSource = _entityAutoComplete;

            sourceEntityComboBox.SelectedIndexChanged += async (_, __) =>
            {
                if (!IsDisposed && gridMappings.IsHandleCreated)
                {
                    gridMappings.DataSource = null;
                }
                sourceAttributeComboBox.DataSource = null;
                _attributeAutoComplete.Clear();

                SourceEntityLogicalName = string.Empty;
                SourceAttributeLogicalName = string.Empty;
                submitButton.Enabled = false;

                var logical = sourceEntityComboBox.SelectedValue as string ?? sourceEntityComboBox.Text.Trim();
                if (!string.IsNullOrEmpty(logical))
                {
                    await PopulateAttributesAsync(logical);
                }
                ;
            };

            if (progressBar != null)
            {
                progressBar.Visible = false;
                progressBar.Style = ProgressBarStyle.Continuous;
            }

            statusLabel.Text = "Entities loaded.";
        }

        private async Task PopulateAttributesAsync(string entityLogicalName)
        {
            statusLabel.Text = $"Loading attributes for {entityLogicalName}...";

            if (_attributesItemsCache.TryGetValue(entityLogicalName, out var cachedAttributes))
            {
                BindAttributes(cachedAttributes);
                statusLabel.Text = $"Attributes for {entityLogicalName} loaded from cache.";
                return;
            }

            var token = _ctsMetadata?.Token ?? CancellationToken.None;

            var attributes = await Task.Run(() =>
            {
                token.ThrowIfCancellationRequested();

                var results = new List<string>();

                var sourceQuery = new Microsoft.Xrm.Sdk.Query.QueryExpression("attributemap")
                {
                    ColumnSet = new Microsoft.Xrm.Sdk.Query.ColumnSet("sourceattributename"),
                    Distinct = true
                };

                var linkSource = new Microsoft.Xrm.Sdk.Query.LinkEntity(
                    "attributemap",
                    "entitymap",
                    "entitymapid",
                    "entitymapid",
                    Microsoft.Xrm.Sdk.Query.JoinOperator.Inner)
                {
                    Columns = new Microsoft.Xrm.Sdk.Query.ColumnSet("sourceentityname"),
                    EntityAlias = "em_source"
                };

                linkSource.LinkCriteria.AddCondition(
                    "sourceentityname",
                    Microsoft.Xrm.Sdk.Query.ConditionOperator.Equal,
                    entityLogicalName);

                sourceQuery.LinkEntities.Add(linkSource);

                var request = new RetrieveEntityRequest
                {
                    EntityFilters = Microsoft.Xrm.Sdk.Metadata.EntityFilters.Attributes,
                    LogicalName = entityLogicalName,
                    RetrieveAsIfPublished = true
                };

                var sourceResponse = _organizationService.RetrieveMultiple(sourceQuery);

                foreach (var entity in sourceResponse.Entities)
                {
                    var attrName = entity.GetAttributeValue<string>("sourceattributename");
                    if (!string.IsNullOrEmpty(attrName))
                    {
                        results.Add(attrName);
                    }
                }

                token.ThrowIfCancellationRequested();

                var targetQuery = new Microsoft.Xrm.Sdk.Query.QueryExpression("attributemap")
                {
                    ColumnSet = new Microsoft.Xrm.Sdk.Query.ColumnSet("targetattributename"),
                    Distinct = true
                };

                var linkTarget = new Microsoft.Xrm.Sdk.Query.LinkEntity(
                    "attributemap",
                    "entitymap",
                    "entitymapid",
                    "entitymapid",
                    Microsoft.Xrm.Sdk.Query.JoinOperator.Inner)
                {
                    Columns = new Microsoft.Xrm.Sdk.Query.ColumnSet("targetentityname"),
                    EntityAlias = "em_target"
                };

                linkTarget.LinkCriteria.AddCondition(
                    "targetentityname",
                    Microsoft.Xrm.Sdk.Query.ConditionOperator.Equal,
                    entityLogicalName);

                targetQuery.LinkEntities.Add(linkTarget);

                var targetResponse = _organizationService.RetrieveMultiple(targetQuery);

                foreach (var entity in targetResponse.Entities)
                {
                    var attrName = entity.GetAttributeValue<string>("targetattributename");
                    if (!string.IsNullOrEmpty(attrName))
                    {
                        results.Add(attrName);
                    }
                }

                token.ThrowIfCancellationRequested();

                var deduped = results
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(ai => ai, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                return deduped;
            }, token);

            await EnsureAttributeDisplayCachedAsync(entityLogicalName, token);

            var items = attributes.Select(attrLogicalName => new AttributeItem
            {
                LogicalName = attrLogicalName,
                DisplayName = GetAttributeDisplayName(entityLogicalName, attrLogicalName)
            })
            .OrderBy(x => x.DisplayName)                
            .ToList();

            _attributesItemsCache[entityLogicalName] = items;
            BindAttributes(items);
            statusLabel.Text = $"Attributes for {entityLogicalName} loaded.";
        }

        private void BindAttributes(List<AttributeItem> attributes)
        {
            sourceAttributeComboBox.DataSource = attributes;
            sourceAttributeComboBox.DisplayMember = nameof(AttributeItem.DisplayName);
            sourceAttributeComboBox.ValueMember = nameof(AttributeItem.LogicalName);

            _attributeAutoComplete.Clear();
            foreach (var attr in attributes)
            {
                _attributeAutoComplete.Add(attr.DisplayName);
                _attributeAutoComplete.Add(attr.LogicalName);
            }
            sourceAttributeComboBox.AutoCompleteCustomSource = _attributeAutoComplete;
            submitButton.Enabled = true;
        }

        private async Task EnsureEntityDisplayNameCachedAsync(IEnumerable<string> entityLogicalName, 
            CancellationToken token)
        {

            statusLabel.Text = "Converting entity logical names to display names ...";

            var needed = entityLogicalName?
                .Where(n => !string.IsNullOrWhiteSpace(n))
                .Select(n => n.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Where(n => !_entitiesCache.ContainsKey(n))
                .ToList() ?? new List<string>();

            if (needed.Count == 0)
            {
                return;
            }

            await Task.Run(() =>
            {
                token.ThrowIfCancellationRequested();

                foreach (var batch in Chunk(needed, EntityBatchSize))
                {
                    token.ThrowIfCancellationRequested();

                    var request = new OrganizationRequestCollection();

                    foreach (var logical in batch)
                    {
                        request.Add(new RetrieveEntityRequest
                        {
                            LogicalName = logical,
                            EntityFilters = EntityFilters.Entity,
                            RetrieveAsIfPublished = true
                        });
                    }

                    var exec = new ExecuteMultipleRequest
                    {
                        Requests = request,
                        Settings = new Microsoft.Xrm.Sdk.ExecuteMultipleSettings
                        {
                            ContinueOnError = true,
                            ReturnResponses = true
                        }
                    };

                    var execResponse = (ExecuteMultipleResponse)_organizationService.Execute(exec);

                    foreach (var item in execResponse.Responses)
                    {
                        if (item.Fault != null)
                        {
                            var req = (RetrieveEntityRequest)exec.Requests[item.RequestIndex];
                            var logical = req.LogicalName;
                            _entitiesCache[logical] = logical;
                            continue;
                        }

                        var em = ((RetrieveEntityResponse)item.Response).EntityMetadata;
                        var label = em.DisplayName?.UserLocalizedLabel?.Label;
                        _entitiesCache[em.LogicalName] = string.IsNullOrEmpty(label) 
                        ? em.LogicalName 
                        : label;
                    }
                }
            }, token);
        }

        private async Task EnsureAttributeDisplayCachedAsync(string entityLogicalName, CancellationToken token)
        {
            if (string.IsNullOrWhiteSpace(entityLogicalName))
            {
                return;
            }
            if (_attributesDisplayCache.ContainsKey(entityLogicalName))
            {
                return;
            }

            await Task.Run(() =>
            {
                token.ThrowIfCancellationRequested();

                var request = new RetrieveEntityRequest
                {
                    LogicalName = entityLogicalName,
                    EntityFilters = Microsoft.Xrm.Sdk.Metadata.EntityFilters.Attributes,
                    RetrieveAsIfPublished = true
                };
                var response = (RetrieveEntityResponse)_organizationService.Execute(request);

                token.ThrowIfCancellationRequested();

                var attributeMetadata = response.EntityMetadata.Attributes
                .ToDictionary(
                    attr => attr.LogicalName,
                    attr => attr.DisplayName?.UserLocalizedLabel?.Label ?? attr.LogicalName,
                    StringComparer.OrdinalIgnoreCase);

                _attributesDisplayCache[entityLogicalName] = attributeMetadata;

            }, token);
        }

        private async Task OnDeleteButtonClick()
        {
            if (gridMappings.SelectedRows.Count != 1)
            {
                MessageBox.Show(this, "Please select exactly one mapping row to delete.", "Delete Mapping Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var row = gridMappings.SelectedRows[0];
            var mapping = row.DataBoundItem as MappingResult;
            if (mapping == null)
            {
                MessageBox.Show(this, "Selected mapping row is not bound to data.", "Unbound Mapping Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var confirmation = MessageBox.Show(this,
                $"Delete mapping: \n\n{mapping.SourceEntity}.{mapping.SourceAttribute} -> " +
                $"{mapping.TargetEntity}.{mapping.TargetAttribute}\n\n"
                + "This will remove the field-level mapping (attributemap). Continue?",
                "Confirm Deletion", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            
            if (confirmation != DialogResult.Yes)
            {
                return;
            }
            var progress = statusStrip.Items["statusProgressBar"] as ToolStripProgressBar;
            if (progress != null)
            {
                progress.Visible = true;
                progress.Style = ProgressBarStyle.Marquee;
            }
            statusLabel.Text = "Deleting mapping ...";

            try
            {
                await Task.Run(() =>
                {
                    _organizationService.Delete("attributemap", mapping.AttributeMapId);

                    var orphanQuery = new Microsoft.Xrm.Sdk.Query.QueryExpression("attributemap")
                    {
                        ColumnSet = new Microsoft.Xrm.Sdk.Query.ColumnSet("attributemapid"),
                        Criteria = new Microsoft.Xrm.Sdk.Query.FilterExpression
                        {
                            Conditions =
                            {
                                new Microsoft.Xrm.Sdk.Query.ConditionExpression("entitymapid",
                                Microsoft.Xrm.Sdk.Query.ConditionOperator.Equal, mapping.EntityMapId)

                            }
                        },
                        PageInfo = new Microsoft.Xrm.Sdk.Query.PagingInfo
                        {
                            Count = 1,
                            PageNumber = 1
                        }
                    };
                    var remaining = _organizationService.RetrieveMultiple(orphanQuery);
                    var parentHasChildern = remaining.Entities.Count > 0;

                    if (!parentHasChildern)
                    {
                        _organizationService.Delete("entitymap", mapping.EntityMapId);
                    }
                });

                _gridData.Remove(mapping);

                _attributeAutoComplete.Remove(mapping.SourceAttribute);
                _attributeAutoComplete.Remove(mapping.TargetAttribute);
                sourceAttributeComboBox.AutoCompleteCustomSource = _attributeAutoComplete;

                statusLabel.Text = "Mapping deleted.";
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Failed to delete mapping.\n\n" + ex.Message, "Delete Failed",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                statusLabel.Text = "Delete Failed.";
            }
            finally
            {
                if (progress != null)
                {
                    progress.Visible = false;
                    progress.Style = ProgressBarStyle.Continuous;
                };
                deleteButton.Enabled = gridMappings.SelectedRows.Count == 1;
            }
        }
    }
}

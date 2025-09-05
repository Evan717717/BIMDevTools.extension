# -*- coding: utf-8 -*-

__title__ = 'Create\nPanel Types'
__author__ = 'BIM Tools'

import clr
clr.AddReference('RevitAPI')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from Autodesk.Revit.DB import *
from System.Windows.Forms import *
from System.Drawing import *

doc = __revit__.ActiveUIDocument.Document

class PanelCreatorDialog(Form):
    def __init__(self):
        self.Text = 'Create Panel Types'
        self.Width = 650
        self.Height = 520
        self.StartPosition = FormStartPosition.CenterScreen
        
        # Title label
        title_label = Label()
        title_label.Text = 'Enter panel data (each row creates a new type)'
        title_label.Location = Point(10, 10)
        title_label.Width = 600
        title_label.Height = 20
        title_label.Font = Font(title_label.Font, FontStyle.Bold)
        
        # Instructions
        info_label = Label()
        info_label.Text = 'Panel Name will be used as Type Name and Panel Name (by Type)'
        info_label.Location = Point(10, 35)
        info_label.Width = 600
        info_label.Height = 20
        info_label.ForeColor = Color.Blue
        
        # DataGridView
        self.dgv = DataGridView()
        self.dgv.Location = Point(10, 60)
        self.dgv.Width = 610
        self.dgv.Height = 300
        self.dgv.AllowUserToAddRows = True
        self.dgv.AllowUserToDeleteRows = True
        self.dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        
        # Right-click context menu
        self.context_menu = ContextMenuStrip()
        
        delete_item = ToolStripMenuItem('Delete Row')
        delete_item.Click += self.delete_from_context
        
        mode_item = ToolStripMenuItem('Switch to Row Select Mode')
        mode_item.Click += self.toggle_selection_mode
        
        self.context_menu.Items.Add(delete_item)
        self.context_menu.Items.Add(mode_item)
        
        self.dgv.ContextMenuStrip = self.context_menu
        self.is_row_select_mode = False
        
        # Create columns
        col_name = DataGridViewTextBoxColumn()
        col_name.HeaderText = 'Panel Name'
        col_name.Name = 'PanelName'
        col_name.FillWeight = 25
        
        col_height = DataGridViewTextBoxColumn()
        col_height.HeaderText = 'Height (mm)'
        col_height.Name = 'Height'
        col_height.FillWeight = 20
        
        col_width = DataGridViewTextBoxColumn()
        col_width.HeaderText = 'Width (mm)'
        col_width.Name = 'Width'
        col_width.FillWeight = 20
        
        col_depth = DataGridViewTextBoxColumn()
        col_depth.HeaderText = 'Depth (mm)'
        col_depth.Name = 'Depth'
        col_depth.FillWeight = 20
        
        self.dgv.Columns.Add(col_name)
        self.dgv.Columns.Add(col_height)
        self.dgv.Columns.Add(col_width)
        self.dgv.Columns.Add(col_depth)
        
        # Add sample data
        self.dgv.Rows.Add(['PWP-1', '600', '400', '200'])
        self.dgv.Rows.Add(['PWP-2', '800', '600', '250'])
        self.dgv.Rows.Add(['LP-1', '1000', '800', '300'])
        
        # Mode indicator
        self.mode_label = Label()
        self.mode_label.Text = 'Mode: Cell Edit (Right-click to switch)'
        self.mode_label.Location = Point(10, 365)
        self.mode_label.Width = 300
        self.mode_label.Height = 20
        self.mode_label.ForeColor = Color.DarkGreen
        
        # Instructions
        delete_info = Label()
        delete_info.Text = 'Delete: Select row number + Delete key, or Right-click menu'
        delete_info.Location = Point(10, 385)
        delete_info.Width = 600
        delete_info.Height = 20
        delete_info.ForeColor = Color.DarkGray
        delete_info.Font = Font(delete_info.Font.FontFamily, 8)
        
        # Create button
        self.btn_create = Button()
        self.btn_create.Text = 'Create Types'
        self.btn_create.Location = Point(250, 410)
        self.btn_create.Width = 120
        self.btn_create.Height = 35
        self.btn_create.BackColor = Color.LightGreen
        self.btn_create.Click += self.create_types
        
        # Cancel button
        self.btn_cancel = Button()
        self.btn_cancel.Text = 'Cancel'
        self.btn_cancel.Location = Point(380, 410)
        self.btn_cancel.Width = 100
        self.btn_cancel.Height = 35
        self.btn_cancel.Click += self.cancel_dialog
        
        # Add controls
        self.Controls.Add(title_label)
        self.Controls.Add(info_label)
        self.Controls.Add(self.dgv)
        self.Controls.Add(self.mode_label)
        self.Controls.Add(delete_info)
        self.Controls.Add(self.btn_create)
        self.Controls.Add(self.btn_cancel)
    
    def toggle_selection_mode(self, sender, e):
        """Toggle between cell edit mode and row select mode"""
        if self.is_row_select_mode:
            # Switch to cell edit mode
            self.dgv.SelectionMode = DataGridViewSelectionMode.CellSelect
            self.dgv.EditMode = DataGridViewEditMode.EditOnEnter
            self.mode_label.Text = 'Mode: Cell Edit (Right-click to switch)'
            self.mode_label.ForeColor = Color.DarkGreen
            self.context_menu.Items[1].Text = 'Switch to Row Select Mode'
            self.is_row_select_mode = False
        else:
            # Switch to row select mode
            self.dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            self.dgv.EditMode = DataGridViewEditMode.EditOnF2
            self.mode_label.Text = 'Mode: Row Select (Press F2 to edit, Right-click to switch)'
            self.mode_label.ForeColor = Color.DarkBlue
            self.context_menu.Items[1].Text = 'Switch to Cell Edit Mode'
            self.is_row_select_mode = True
    
    def delete_from_context(self, sender, e):
        """Delete row from context menu"""
        if self.dgv.CurrentRow and not self.dgv.CurrentRow.IsNewRow:
            self.dgv.Rows.Remove(self.dgv.CurrentRow)
    
    def create_types(self, sender, e):
        """Create panel types based on user input"""
        
        # Check if in family document
        if not doc.IsFamilyDocument:
            MessageBox.Show('Please run this tool in Family Editor!', 'Error', 
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
            return
        
        # Get family manager
        fm = doc.FamilyManager
        
        # Collect data from grid
        panels_to_create = []
        for i in range(self.dgv.Rows.Count - 1):  # Exclude the new row placeholder
            row = self.dgv.Rows[i]
            if row.Cells[0].Value and str(row.Cells[0].Value).strip():
                try:
                    panel_data = {
                        'name': str(row.Cells[0].Value).strip(),
                        'height': float(row.Cells[1].Value) if row.Cells[1].Value else 600.0,
                        'width': float(row.Cells[2].Value) if row.Cells[2].Value else 400.0,
                        'depth': float(row.Cells[3].Value) if row.Cells[3].Value else 200.0
                    }
                    panels_to_create.append(panel_data)
                except:
                    MessageBox.Show('Invalid data in row {}. Please check numbers.'.format(i+1), 
                                  'Error', MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    return
        
        if not panels_to_create:
            MessageBox.Show('Please enter at least one panel!', 'Information',
                          MessageBoxButtons.OK, MessageBoxIcon.Information)
            return
        
        # Find required parameters
        param_height = None
        param_width = None
        param_depth = None
        param_panel_name = None
        
        for param in fm.GetParameters():
            if not param.IsInstance:  # Only type parameters
                param_name = param.Definition.Name
                
                if param_name == u'\u76e4\u9ad8\u5ea6':  # "盤高度"
                    param_height = param
                elif param_name == u'\u76e4\u5bec\u5ea6':  # "盤寬度"
                    param_width = param
                elif param_name == u'\u76e4\u539a\u5ea6':  # "盤厚度"
                    param_depth = param
                elif param_name == 'Panel Name (by Type)':
                    param_panel_name = param
        
        # Check if all parameters found
        missing = []
        if not param_height:
            missing.append('Height parameter not found')
        if not param_width:
            missing.append('Width parameter not found')
        if not param_depth:
            missing.append('Depth parameter not found')
        if not param_panel_name:
            missing.append('Panel Name (by Type) parameter not found')
        
        if missing:
            msg = 'Missing parameters:\n' + '\n'.join(missing)
            msg += '\n\nDo you want to continue creating types without setting these parameters?'
            result = MessageBox.Show(msg, 'Warning', MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
            if result != DialogResult.Yes:
                return
        
        # Start transaction
        t = Transaction(doc, 'Create Panel Types')
        t.Start()
        
        created = []
        updated = []
        errors = []
        
        try:
            for panel in panels_to_create:
                try:
                    type_name = panel['name']
                    
                    # Check if type exists
                    existing_type = None
                    for ft in fm.Types:
                        if ft.Name == type_name:
                            existing_type = ft
                            break
                    
                    if existing_type:
                        # Update existing type
                        fm.CurrentType = existing_type
                        updated.append(type_name)
                    else:
                        # Create new type
                        new_type = fm.NewType(type_name)
                        fm.CurrentType = new_type
                        created.append(type_name)
                    
                    # Convert mm to feet (Revit internal units)
                    height_ft = panel['height'] / 304.8
                    width_ft = panel['width'] / 304.8
                    depth_ft = panel['depth'] / 304.8
                    
                    # Set parameter values
                    if param_height:
                        fm.Set(param_height, height_ft)
                    
                    if param_width:
                        fm.Set(param_width, width_ft)
                    
                    if param_depth:
                        fm.Set(param_depth, depth_ft)
                    
                    if param_panel_name:
                        fm.Set(param_panel_name, type_name)
                    
                except Exception as ex:
                    errors.append('{}: {}'.format(panel['name'], str(ex)))
            
            t.Commit()
            
            # Show results
            result_msg = ''
            
            if created:
                result_msg += 'Created {} new types:\n{}\n\n'.format(
                    len(created), '\n'.join(['  - ' + n for n in created]))
            
            if updated:
                result_msg += 'Updated {} existing types:\n{}\n\n'.format(
                    len(updated), '\n'.join(['  - ' + n for n in updated]))
            
            if errors:
                result_msg += 'Errors:\n{}'.format('\n'.join(['  - ' + e for e in errors]))
            
            if created or updated:
                MessageBox.Show(result_msg, 'Success', 
                              MessageBoxButtons.OK, MessageBoxIcon.Information)
                self.Close()
            elif errors:
                MessageBox.Show(result_msg, 'Error', 
                              MessageBoxButtons.OK, MessageBoxIcon.Error)
            
        except Exception as ex:
            t.RollBack()
            MessageBox.Show('Transaction failed:\n' + str(ex), 'Error',
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
    
    def cancel_dialog(self, sender, e):
        """Close dialog without creating types"""
        self.Close()

# Main execution
if __name__ == '__main__':
    # Check if in family editor first
    if not doc.IsFamilyDocument:
        MessageBox.Show(
            'This tool must be run in Family Editor!\n\n' +
            'Please open a family file (.rfa) first.',
            'Wrong Document Type',
            MessageBoxButtons.OK,
            MessageBoxIcon.Warning
        )
    else:
        # Show dialog
        dialog = PanelCreatorDialog()
        dialog.ShowDialog()
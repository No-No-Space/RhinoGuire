#! python3
# # -*- coding: utf-8 -*-
# __title__ = "Baquiano"                           # Name of the button displayed in Revit UI
# __doc__ = """Version = 0.7
# Date    = 2026-02-20
# Author: Aquelon - aquelon@pm.me
# _____________________________________________________________________
# Description:
# Advanced search tool for Rhino object metadata (User Keys/Values)
# Supports include/exclude conditions and pre-selection filtering.
# The window is modeless — Rhino stays accessible while it is open.
# _____________________________________________________________________
# How-to:
# -> Run the script in Rhino 8 (RunPythonScript), a search window opens
# -> Optionally select objects in Rhino (form stays open, Rhino is accessible)
# -> Choose search scope: all objects in the model or currently selected objects
# -> Click the Key dropdown in any condition row to see available keys from the model
# -> Use "Refresh Keys" to re-scan the model if you added new objects/keys
# -> Add include conditions: pick or type a Key name, type a Value, and select a match type
# -> Optionally add exclude conditions to filter out unwanted results
# -> Use "+ Add" buttons to combine multiple conditions (AND for include, OR for exclude)
# -> Click "Search" to find and select matching objects in the viewport
# -> Repeat searches without closing the window
# -> Match types: Contains, Equals, Starts/Ends with, and their negations (Does not...)
# _____________________________________________________________________
# Last update:
# - [20.02.2026] - 0.7 Key field is now a ComboBox with suggestions from the model
# - [20.02.2026] - 0.6 Converted to modeless form; search runs without closing window
# - [14.02.2026] - 0.5 RELEASE
# _____________________________________________________________________
# To-Do:
# - UI needs improvement, it is functional at the moment, but default look.
# - Default folder locations need to be updated to neutral locations (desktop or documents) instead of script folder.
# _____________________________________________________________________


import rhinoscriptsyntax as rs
import Rhino
import Eto.Drawing as drawing
import Eto.Forms as forms


def get_all_user_text_keys():
    """Collect all unique user text keys from every object in the model, sorted."""
    all_objects = rs.AllObjects()
    if not all_objects:
        return []
    keys = set()
    for obj_guid in all_objects:
        obj_keys = rs.GetUserText(obj_guid)  # returns list of key strings, or None
        if obj_keys:
            for k in obj_keys:
                keys.add(k)
    return sorted(keys)


class SearchCondition:
    """Represents a single search condition (include or exclude)."""
    def __init__(self, key, value, match_type, is_exclude=False):
        self.key = key
        self.value = value
        self.match_type = match_type  # "contains", "equals", "starts_with", "ends_with"
        self.is_exclude = is_exclude

    def matches(self, obj_guid):
        """Check if the object matches this condition."""
        obj_value = rs.GetUserText(obj_guid, self.key)
        if obj_value is None:
            return False

        obj_value_lower = obj_value.lower()
        search_value_lower = self.value.lower()

        if self.match_type == "contains":
            return search_value_lower in obj_value_lower
        elif self.match_type == "equals":
            return obj_value_lower == search_value_lower
        elif self.match_type == "starts_with":
            return obj_value_lower.startswith(search_value_lower)
        elif self.match_type == "ends_with":
            return obj_value_lower.endswith(search_value_lower)
        elif self.match_type == "not_contains":
            return search_value_lower not in obj_value_lower
        elif self.match_type == "not_equals":
            return obj_value_lower != search_value_lower
        elif self.match_type == "not_starts_with":
            return not obj_value_lower.startswith(search_value_lower)
        elif self.match_type == "not_ends_with":
            return not obj_value_lower.endswith(search_value_lower)
        return False


class ConditionRow:
    """UI row for a single search condition."""
    def __init__(self, parent_form, is_exclude=False, available_keys=None):
        self.parent = parent_form
        self.is_exclude = is_exclude
        self.available_keys = available_keys or []
        self.create_controls()

    def create_controls(self):
        # ComboBox lets the user pick a key from the model OR type a custom one
        self.key_combo = forms.ComboBox()
        self.key_combo.DataStore = self.available_keys
        self.key_combo.PlaceholderText = "Key name"
        self.key_combo.Width = 150

        self.value_textbox = forms.TextBox()
        self.value_textbox.PlaceholderText = "Search value"
        self.value_textbox.Width = 150

        self.match_dropdown = forms.DropDown()
        self.match_dropdown.Items.Add("Contains")
        self.match_dropdown.Items.Add("Equals")
        self.match_dropdown.Items.Add("Starts with")
        self.match_dropdown.Items.Add("Ends with")
        self.match_dropdown.Items.Add("Does not contain")
        self.match_dropdown.Items.Add("Does not equal")
        self.match_dropdown.Items.Add("Does not start with")
        self.match_dropdown.Items.Add("Does not end with")
        self.match_dropdown.SelectedIndex = 0
        self.match_dropdown.Width = 140

        self.remove_button = forms.Button()
        self.remove_button.Text = "X"
        self.remove_button.Width = 30
        self.remove_button.Click += self.on_remove

        # Horizontal layout for this row
        self.row_layout = forms.StackLayout()
        self.row_layout.Orientation = forms.Orientation.Horizontal
        self.row_layout.Spacing = 5
        self.row_layout.Items.Add(forms.StackLayoutItem(self.key_combo))
        self.row_layout.Items.Add(forms.StackLayoutItem(self.value_textbox))
        self.row_layout.Items.Add(forms.StackLayoutItem(self.match_dropdown))
        self.row_layout.Items.Add(forms.StackLayoutItem(self.remove_button))

    def update_available_keys(self, keys):
        """Repopulate the key ComboBox after a model refresh, preserving current text."""
        current_text = self.key_combo.Text
        self.available_keys = keys
        self.key_combo.DataStore = keys
        self.key_combo.Text = current_text

    def on_remove(self, sender, e):
        self.parent.remove_condition(self)

    def get_condition(self):
        """Returns a SearchCondition object or None if empty."""
        key = self.key_combo.Text.strip()
        value = self.value_textbox.Text.strip()
        if not key or not value:
            return None

        match_types = [
            "contains", "equals", "starts_with", "ends_with",
            "not_contains", "not_equals", "not_starts_with", "not_ends_with"
        ]
        match_type = match_types[self.match_dropdown.SelectedIndex]

        return SearchCondition(key, value, match_type, self.is_exclude)


class BaquianoSearchForm(forms.Form):
    """Main search form — modeless, stays open between searches."""

    def __init__(self, preselection_count=0):
        super().__init__()
        self.preselection_count = preselection_count
        self.has_preselection = preselection_count > 0
        self.include_conditions = []
        self.exclude_conditions = []
        self.available_keys = get_all_user_text_keys()
        self.Title = "Baquiano Search Data"
        self.Padding = drawing.Padding(10)
        self.Resizable = True
        self.MinimumSize = drawing.Size(600, 450)
        self.Owner = Rhino.UI.RhinoEtoApp.MainWindow
        self.create_controls()

    def create_controls(self):
        layout = forms.DynamicLayout()
        layout.DefaultSpacing = drawing.Size(5, 5)

        # Header
        header = forms.Label()
        header.Text = "Search Rhino Objects by User Keys"
        header.Font = drawing.Font(drawing.SystemFont.Bold, 12)
        layout.AddRow(header)
        layout.AddRow(None)

        # Search scope
        scope_label = forms.Label()
        scope_label.Text = "Search Scope:"
        scope_label.Font = drawing.Font(drawing.SystemFont.Bold)
        layout.AddRow(scope_label)

        self.scope_all_radio = forms.RadioButton()
        self.scope_all_radio.Text = "Search all objects in model"
        self.scope_all_radio.Checked = not self.has_preselection

        self.scope_selected_radio = forms.RadioButton(self.scope_all_radio)
        self.scope_selected_radio.Text = "Search currently selected objects"
        self.scope_selected_radio.Checked = self.has_preselection

        hint = forms.Label()
        hint.Text = "(select objects in Rhino, then click Search)"
        hint.TextColor = drawing.Colors.Gray

        layout.AddRow(self.scope_all_radio)
        scope_row = forms.StackLayout()
        scope_row.Orientation = forms.Orientation.Horizontal
        scope_row.Spacing = 10
        scope_row.Items.Add(forms.StackLayoutItem(self.scope_selected_radio))
        scope_row.Items.Add(forms.StackLayoutItem(hint))
        layout.AddRow(scope_row)

        layout.AddRow(None)

        # Key suggestions info row
        keys_row = forms.StackLayout()
        keys_row.Orientation = forms.Orientation.Horizontal
        keys_row.Spacing = 10

        self.keys_info_label = forms.Label()
        self.keys_info_label.Text = self._keys_info_text()
        self.keys_info_label.TextColor = drawing.Colors.Gray

        refresh_btn = forms.Button()
        refresh_btn.Text = "Refresh Keys"
        refresh_btn.Click += self.on_refresh_keys

        keys_row.Items.Add(forms.StackLayoutItem(self.keys_info_label, True))
        keys_row.Items.Add(forms.StackLayoutItem(refresh_btn))
        layout.AddRow(keys_row)

        layout.AddRow(None)

        # Include conditions section
        include_label = forms.Label()
        include_label.Text = "Include Conditions (objects must match ALL):"
        include_label.Font = drawing.Font(drawing.SystemFont.Bold)
        layout.AddRow(include_label)

        # Column headers for include
        include_header = forms.StackLayout()
        include_header.Orientation = forms.Orientation.Horizontal
        include_header.Spacing = 5
        key_lbl = forms.Label()
        key_lbl.Text = "Key"
        key_lbl.Width = 150
        val_lbl = forms.Label()
        val_lbl.Text = "Value"
        val_lbl.Width = 150
        match_lbl = forms.Label()
        match_lbl.Text = "Match Type"
        match_lbl.Width = 140
        include_header.Items.Add(forms.StackLayoutItem(key_lbl))
        include_header.Items.Add(forms.StackLayoutItem(val_lbl))
        include_header.Items.Add(forms.StackLayoutItem(match_lbl))
        layout.AddRow(include_header)

        self.include_stack = forms.StackLayout()
        self.include_stack.Orientation = forms.Orientation.Vertical
        self.include_stack.Spacing = 3
        layout.AddRow(self.include_stack)

        add_include_btn = forms.Button()
        add_include_btn.Text = "+ Add Include Condition"
        add_include_btn.Click += self.on_add_include
        layout.AddRow(add_include_btn)

        layout.AddRow(None)

        # Exclude conditions section
        exclude_label = forms.Label()
        exclude_label.Text = "Exclude Conditions (objects matching ANY will be excluded):"
        exclude_label.Font = drawing.Font(drawing.SystemFont.Bold)
        layout.AddRow(exclude_label)

        # Column headers for exclude
        exclude_header = forms.StackLayout()
        exclude_header.Orientation = forms.Orientation.Horizontal
        exclude_header.Spacing = 5
        key_lbl2 = forms.Label()
        key_lbl2.Text = "Key"
        key_lbl2.Width = 150
        val_lbl2 = forms.Label()
        val_lbl2.Text = "Value"
        val_lbl2.Width = 150
        match_lbl2 = forms.Label()
        match_lbl2.Text = "Match Type"
        match_lbl2.Width = 140
        exclude_header.Items.Add(forms.StackLayoutItem(key_lbl2))
        exclude_header.Items.Add(forms.StackLayoutItem(val_lbl2))
        exclude_header.Items.Add(forms.StackLayoutItem(match_lbl2))
        layout.AddRow(exclude_header)

        self.exclude_stack = forms.StackLayout()
        self.exclude_stack.Orientation = forms.Orientation.Vertical
        self.exclude_stack.Spacing = 3
        layout.AddRow(self.exclude_stack)

        add_exclude_btn = forms.Button()
        add_exclude_btn.Text = "+ Add Exclude Condition"
        add_exclude_btn.Click += self.on_add_exclude
        layout.AddRow(add_exclude_btn)

        layout.AddRow(None)

        # Bottom button row
        button_layout = forms.StackLayout()
        button_layout.Orientation = forms.Orientation.Horizontal
        button_layout.Spacing = 10

        search_btn = forms.Button()
        search_btn.Text = "Search"
        search_btn.Click += self.on_search

        close_btn = forms.Button()
        close_btn.Text = "Close"
        close_btn.Click += self.on_close_btn

        button_layout.Items.Add(forms.StackLayoutItem(search_btn))
        button_layout.Items.Add(forms.StackLayoutItem(close_btn))

        layout.AddRow(button_layout)

        # Status label — updated after each search or refresh
        self.status_label = forms.Label()
        self.status_label.Text = "Ready."
        self.status_label.TextColor = drawing.Colors.Gray
        layout.AddRow(self.status_label)

        self.Content = layout

        # Add first include condition by default
        self.add_include_condition()

    # ------------------------------------------------------------------
    # Key helpers
    # ------------------------------------------------------------------

    def _keys_info_text(self):
        count = len(self.available_keys)
        if count == 0:
            return "No user text keys found in model."
        return f"Key suggestions: {count} unique key(s) loaded from model."

    def on_refresh_keys(self, _sender, _e):
        self.available_keys = get_all_user_text_keys()
        self.keys_info_label.Text = self._keys_info_text()
        for row in self.include_conditions + self.exclude_conditions:
            row.update_available_keys(self.available_keys)
        self.status_label.Text = f"Keys refreshed — {len(self.available_keys)} key(s) found."
        self.status_label.TextColor = drawing.Colors.Gray

    # ------------------------------------------------------------------
    # Condition management
    # ------------------------------------------------------------------

    def add_include_condition(self):
        cond = ConditionRow(self, is_exclude=False, available_keys=self.available_keys)
        self.include_conditions.append(cond)
        self.include_stack.Items.Add(forms.StackLayoutItem(cond.row_layout))

    def add_exclude_condition(self):
        cond = ConditionRow(self, is_exclude=True, available_keys=self.available_keys)
        self.exclude_conditions.append(cond)
        self.exclude_stack.Items.Add(forms.StackLayoutItem(cond.row_layout))

    def remove_condition(self, condition_row):
        if condition_row in self.include_conditions:
            self.include_conditions.remove(condition_row)
            for i in range(self.include_stack.Items.Count):
                if self.include_stack.Items[i].Control == condition_row.row_layout:
                    self.include_stack.Items.RemoveAt(i)
                    break
        elif condition_row in self.exclude_conditions:
            self.exclude_conditions.remove(condition_row)
            for i in range(self.exclude_stack.Items.Count):
                if self.exclude_stack.Items[i].Control == condition_row.row_layout:
                    self.exclude_stack.Items.RemoveAt(i)
                    break

    def on_add_include(self, sender, e):
        self.add_include_condition()

    def on_add_exclude(self, sender, e):
        self.add_exclude_condition()

    # ------------------------------------------------------------------
    # Search
    # ------------------------------------------------------------------

    def on_search(self, sender, e):
        include_conditions = self.get_include_conditions()
        exclude_conditions = self.get_exclude_conditions()

        if not include_conditions:
            self.status_label.Text = "Error: specify at least one include condition."
            self.status_label.TextColor = drawing.Colors.Red
            return

        # Determine search scope — reads live Rhino selection if "selected" is chosen
        if self.scope_all_radio.Checked:
            objects_to_search = rs.AllObjects() or []
        else:
            objects_to_search = rs.SelectedObjects() or []

        if not objects_to_search:
            self.status_label.Text = "No objects in the selected scope."
            self.status_label.TextColor = drawing.Colors.Red
            return

        results = perform_search(objects_to_search, include_conditions, exclude_conditions)

        rs.UnselectAllObjects()
        if results:
            rs.SelectObjects(results)
            inc_summary = ", ".join([f"{c.key}={c.value}" for c in include_conditions])
            scope_text = "all objects" if self.scope_all_radio.Checked else f"{len(objects_to_search)} selected"
            self.status_label.Text = (
                f"Found {len(results)} object(s) in {scope_text} | Include: {inc_summary}"
            )
            self.status_label.TextColor = drawing.Colors.Green
        else:
            self.status_label.Text = "No objects found matching the search criteria."
            self.status_label.TextColor = drawing.Colors.Orange

    def on_close_btn(self, _sender, _e):
        self.Close()

    def get_include_conditions(self):
        conditions = []
        for row in self.include_conditions:
            cond = row.get_condition()
            if cond:
                conditions.append(cond)
        return conditions

    def get_exclude_conditions(self):
        conditions = []
        for row in self.exclude_conditions:
            cond = row.get_condition()
            if cond:
                conditions.append(cond)
        return conditions


def perform_search(objects_to_search, include_conditions, exclude_conditions):
    """
    Perform the search with include/exclude conditions.

    - Objects must match ALL include conditions
    - Objects matching ANY exclude condition are removed
    """
    results = []

    for obj in objects_to_search:
        # Check all include conditions (AND logic)
        if include_conditions:
            matches_all_includes = all(cond.matches(obj) for cond in include_conditions)
            if not matches_all_includes:
                continue

        # Check exclude conditions (OR logic - any match excludes)
        if exclude_conditions:
            matches_any_exclude = any(cond.matches(obj) for cond in exclude_conditions)
            if matches_any_exclude:
                continue

        results.append(obj)

    return results


def main():
    # Capture any pre-selection to default scope radio accordingly
    preselected = rs.SelectedObjects()
    preselection_count = len(preselected) if preselected else 0

    form = BaquianoSearchForm(preselection_count=preselection_count)
    form.Show()


if __name__ == "__main__":
    main()

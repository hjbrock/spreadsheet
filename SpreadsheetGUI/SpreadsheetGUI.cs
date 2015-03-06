// Author: Hannah Brock
// CS 3500, Fall 2012

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using SpreadsheetUtilities;

namespace SS
{
    /// <summary>
    /// GUI front end for the Spreadsheet class.
    /// </summary>
    public partial class SpreadsheetForm : Form
    {
        //The spreadsheet that backs this form
        private Spreadsheet sheet;

        //The filename under which this spreadsheet is currently saved
        private string currentFile;

        /// <summary>
        /// Create a blank spreadsheet form
        /// </summary>
        public SpreadsheetForm()
        {
            InitializeComponent();

            //initialize an empty spreadsheet
            sheet = new Spreadsheet(IsValid, s => s.ToUpper(), "ps6");
            currentFile = null;

            //register display selection method
            spreadsheetPanel.SelectionChanged += displaySelection;
            spreadsheetPanel.SetSelection(0, 0);
            displaySelection(spreadsheetPanel);
        }

        /// <summary>
        /// Create a spreadsheet form from a file
        /// </summary>
        public SpreadsheetForm(string filename)
        {
            InitializeComponent();

            //initialize the spreadsheet, display message if it can't be read
            try
            {
                sheet = new Spreadsheet(filename, IsValid, s => s.ToUpper(), "ps6");
            }
            catch (SpreadsheetReadWriteException e)
            {
                MessageBox.Show("Could not open file: " + e.Message);
                Close();
            }
            currentFile = filename;
            Text = currentFile;

            //register display selection method
            spreadsheetPanel.SelectionChanged += displaySelection;
            spreadsheetPanel.SetSelection(0,0);

            //get cell values
            int col, row;
            foreach (string cell in sheet.GetNamesOfAllNonemptyCells())
            {
                GetColRow(cell, out col, out row);
                spreadsheetPanel.SetValue(col, row, sheet.GetCellValue(cell).ToString());
            }

            displaySelection(spreadsheetPanel);
        }

        /// <summary>
        /// Parses the column and row out of a cell name
        /// </summary>
        /// <param name="name"></param>
        /// <param name="col"></param>
        /// <param name="row"></param>
        private void GetColRow(string name, out int col, out int row)
        {
            col = name[0] - 65;
            row = Int32.Parse(name.Substring(1))-1;
        }

        /// <summary>
        /// Creates a cell name for the given column and row
        /// </summary>
        /// <param name="col">The column</param>
        /// <param name="row">The row</param>
        /// <returns>Returns the name of the cell in the form of [A-Z][0-9]{1,2}</returns>
        private string GetCellName(int col, int row)
        {
            char colLetter = (char)(65 + col);
            return colLetter + (row + 1).ToString();
        }

        /// <summary>
        /// Method to verify that a cell name is valid. A name is valid if 
        /// it is of the form [A-Za-z][0-9]{1,2}
        /// </summary>
        /// <param name="name">Cell name to check</param>
        /// <returns>Returns true if the name is valid. False otherwise.</returns>
        private bool IsValid(string name)
        {
            String pattern = @"^[a-zA-Z][0-9]{1,2}$";

            if (name == null || !Regex.IsMatch(name, pattern))
                return false;
            return true;
        }

        /// <summary>
        /// Updates the given cell's display
        /// </summary>
        /// <param name="name">The name of the cell to update</param>
        private void UpdateCellDisplay(string name)
        {
            int col, row;
            GetColRow(name, out col, out row);
            object value = sheet.GetCellValue(name);
            if (value is FormulaError)
                value = "#ERROR: " + ((FormulaError)value).Reason;
            spreadsheetPanel.SetValue(col, row, value.ToString());
        }

        /// <summary>
        /// Action for when a cell is selected. Display name, value, and contents in top bar.
        /// </summary>
        /// <param name="ss">The spreadsheet panel</param>
        private void displaySelection(SpreadsheetPanel ss)
        {
            int row, col;
            ss.GetSelection(out col, out row);

            //get name of cell
            string cellName = GetCellName(col, row);

            //add contents to edit box
            object contents = sheet.GetCellContents(cellName);
            if (contents is Formula)
                TxtCellContent.Text = "=" + contents.ToString();
            else
                TxtCellContent.Text = contents.ToString();

            //add value to value box, if it's a FormulaError, the value is the reason
            object cellVal = sheet.GetCellValue(cellName);
            if (cellVal is FormulaError)
                cellVal = "#ERROR: " + ((FormulaError)cellVal).Reason;
            TxtCellValue.Text = cellVal.ToString();

            //display cell name
            TxtCellName.Text = cellName;

            TxtCellContent.Focus();
        }

        /// <summary>
        /// Opens a new spreadsheet form.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void newToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Tell the application context to run the form on the same
            // thread as the other forms.
            SpreadsheetApplicationContext.getAppContext().RunForm(new SpreadsheetForm());
        }

        /// <summary>
        /// Prompts to save if the spreadsheet has changed. Then, closes the spreadsheet.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void closeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (sheet.Changed)
                if (MessageBox.Show(this, "Spreadsheet has changed. If you close now, you will lose any unsaved information.\nAre you sure you want to close?","Confirm Exit", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    Close();
        }


        /// <summary>
        /// Updates the displayed cell when the enter key is pressed
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TxtCellContent_KeyDown(object sender, KeyEventArgs e)
        {
            //if enter is pressed while in the content box, update cell
            if (e.KeyCode == Keys.Enter)
            {
                try
                {
                    //update cell and re-display any changed cells
                    foreach (string cell in sheet.SetContentsOfCell(TxtCellName.Text, TxtCellContent.Text))
                        UpdateCellDisplay(cell);
                    if (sheet.Changed && !Text.EndsWith("*"))
                        Text = Text + "*";
                    //proceed to the next cell down
                    int row, col;
                    GetColRow(TxtCellName.Text, out col, out row);
                    spreadsheetPanel.SetSelection(col, row+1);
                    displaySelection(spreadsheetPanel);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "Unable to update cell: " + ex.Message, "Error updating cell");
                }
            }
        }

        /// <summary>
        /// Save the spreadsheet if the "Save" menu item is clicked.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (currentFile != null)
                    sheet.Save(currentFile);
                else
                    SaveAs();
            }
            catch (SpreadsheetReadWriteException ex)
            {
                MessageBox.Show(this, "Unable to save spreadsheet: " + ex.Message, "Error saving spreadsheet");
                return;
            }

            //if a file was chosen and save, set the title text to that filename
            if (currentFile != null)
                Text = currentFile;
        }


        /// <summary>
        /// Save the spreadsheet as a new file.
        /// </summary>
        private void SaveAs()
        {
            string newFile = ShowSaveDialog();
            if (newFile != null && !newFile.Equals(""))
            {
                sheet.Save(newFile);
                currentFile = newFile;
            }
        }

        /// <summary>
        /// Show the Save dialog and get filename
        /// </summary>
        /// <returns></returns>
        private string ShowSaveDialog()
        {
            saveFileDialog.Filter = "SS Files|*.ss|All files|*.*";
            saveFileDialog.ShowDialog();
            return saveFileDialog.FileName;
        }

        /// <summary>
        /// Show the Open file dialog and get the filename
        /// </summary>
        /// <returns></returns>
        private string ShowOpenDialog()
        {
            openFileDialog.Filter = "SS Files|*.ss|All files|*.*";
            openFileDialog.ShowDialog();
            return openFileDialog.FileName;
        }

        /// <summary>
        /// Open an existing spreadsheet
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                string file = ShowOpenDialog();
                if (file != null && !file.Equals(""))
                    SpreadsheetApplicationContext.getAppContext().RunForm(new SpreadsheetForm(file));
            }
            catch (SpreadsheetReadWriteException ex)
            {
                MessageBox.Show(this,"Unable to open spreadsheet: " + ex.Message, "Error opening spreadsheet");
                return;
            }
        }

        /// <summary>
        /// Save the spreadsheet under a new filename (forces the opening of the save dialog)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void saveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                SaveAs();
                Text = currentFile;
            }
            catch (SpreadsheetReadWriteException ex)
            {
                MessageBox.Show(this, "Could not save spreadsheet: " + ex.Message, "Error saving spreadsheet");
            }
        }

        /// <summary>
        /// Displays the about information for this application.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show(this,"Author: Hannah Brock\nCreated for CS 3500, Fall 2012", "Application Information");
        }

        /// <summary>
        /// Displays the help dialog.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void helpContentsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show(this, "How to use this spreadsheet:"
                                    + "\n\nSELECTING A CELL:\nCells can be selected by clicking on the cell in the grid. Pressing the Enter key will proceed to the next cell down."
                                    + "\n\nEDITING A CELL:\nTo edit a cell, select the cell in the grid. Enter the content for the cell in the “Cell Contents” text box. Press the Enter key when you are satisfied with your entry."
                                    + "\n\nOPENING A SPREADSHEET:\nTo open an existing spreadsheet, choose File -> Open and select the spreadsheet you wish to open. It will open a new Spreadsheet form."
                                    + "\n\nSAVING A SPREADSHEET:\nIf you see a “*” after the name of your spreadsheet in the title bar of the form, you have unsaved changes. To save this, choose File -> Save. If you have not previously saved your spreadsheet, you will be prompted to choose a file path and name. Otherwise, your spreadsheet will be saved in its previous location. To choose a new path and name for your file, choose File -> Save As."
                                    + "\n\nCELL CONTENTS:\nThe contents of a cell can be a number, a text string, or a formula. To enter a formula, enter “=” as the first character in the contents box. If there is an issue evaluating a formula due to syntax, you will get an error message when you press the Enter key. If there is an issue looking up another cell’s value when evaluating the formula at any time, a formula error will appear in the cell. More information can be found in the “Cell Value” text box.", "Help");
        }
    }
}

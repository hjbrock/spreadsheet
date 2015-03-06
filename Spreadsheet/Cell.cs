// Author: Hannah Brock
// CS 3500, Fall 2012

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SpreadsheetUtilities;

namespace SS
{
    /// <summary>
    /// This class provides a representation of a cell and its contents in a spreadsheet.
    /// 
    /// A cell may contain a string, double, or formula. If the cell is empty, its contents
    /// are an empty string, "".
    /// </summary>
    class Cell
    {
        //the value of the cell, set to 0.0 if the cell contains a string. Otherwise, set to the double value from the last calculation of the cell's formula.
        private object value;

        //Whether or not the cell value has changed and needs to be recalculated
        public bool Changed
        {
            get;
            protected set;
        }

        /// <summary>
        /// Constructs an empty cell.
        /// </summary>
        /// <param name="name">The name of the cell</param>
        public Cell(string name)
        {
            this.Name = name;
            this.Contents = "";
            value = 0.0;
            Changed = false;
        }

        /// <summary>
        /// Constructs a cell with the given contents.
        /// </summary>
        /// <param name="name">The name of the cell</param>
        /// <param name="contents">The contents of the cell</param>
        public Cell(string name, object contents)
        {
            this.Name = name;
            this.Contents = contents;
            if (contents is double)
                value = (double)contents;
            else
                value = 0.0;
            Changed = true;
        }

        /// <summary>
        /// Sets/gets the contents of this cell.
        /// 
        ///The contents of the cell. This may represent a string, a formula or a double.
        ///If the cell is empty, Contents = ""
        /// </summary>
        public object Contents
        {
            get;
            set;
        }

        /// <summary>
        /// Returns the name of this cell.
        /// </summary>
        public string Name
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the current value of cell, as of the last recalculation.
        /// </summary>
        /// <returns>Returns the value of the cell</returns>
        public object GetValue()
        {
            return value;
        }

        /// <summary>
        /// Returns the value of calling Evaluate if the contents of this cell is a formula,
        /// using the lookup provided. If there is an issue with the formula, returns a FormulaError.
        /// Returns a double if the cell contains a double.
        /// Otherwise, throws an InvalidOperationException.
        /// </summary>
        public object Recalculate(Func<string, double> lookup)
        {
            //check if our cell contains a formula
            if (Contents is Formula)
            {
                object response = ((Formula)Contents).Evaluate(lookup);
                if (response is FormulaError)
                {
                    value = response;
                    return response;
                }
                else
                {
                    value = (double)response;
                    Changed = false;
                    return value;
                }
            }
            //otherwise, see if it's a double
            else if (Contents is double)
            {
                value = (double)Contents;
                Changed = false;
                return value;
            }
            //if not, throw an exception
            else
            {
                Changed = false;
                return 0.0;
            }
        }

        /// <summary>
        /// Generates a hash code for this cell based on its name.
        /// </summary>
        /// <returns>Returns the integer hash code</returns>
        public override int GetHashCode()
        {
            return Name.GetHashCode();
        }

        /// <summary>
        /// Two cells are considered equal (the same cell) if they
        /// have the same name.
        /// </summary>
        /// <param name="obj">The obj to check</param>
        /// <returns>Returns true if the object is not null and is a Cell and has the same name as this cell.
        /// Returns false otherwise.</returns>
        public override bool Equals(object obj)
        {
            if (obj != null && obj is Cell)
                return ((Cell)obj).Name == Name;
            return false;
        }

    }
}
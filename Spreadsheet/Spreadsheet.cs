// Author: Hannah Brock
// CS 3500, Fall 2012

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SpreadsheetUtilities;
using System.Text.RegularExpressions;
using System.Xml;

namespace SS
{
    /// <summary>
    /// Backend representation of a spreadsheet. 
    /// </summary>
    public class Spreadsheet : AbstractSpreadsheet
    {
        // Abstraction function: 
        //      Graph of all of the dependencies in the spreadsheet.
        // Representation invariant: At any given time, all and only the dependencies between cells are held in this graph.
        private DependencyGraph dependencies;

        // Abstraction function: 
        //      The set of all non-empty cells in the spreadsheet.
        // Representation invariant:
        //      A cell is included in this set if it has non-empty contents. As soon as
        //      a cell is emptied (its contents become an empty string), it is removed 
        //      from this set. Each cell in this set has a unique name.
        private Dictionary<String, Cell> cells;

        //Whether or not the spreadsheet has changed since the last save (or creation)
        //private bool HasChanged;

        /// <see cref="AbstractSpreadsheet.Changed"/>
        public override bool Changed
        {
            get;
            protected set;
        }

        /// <summary>
        /// Constructs an empty spreadsheet. All variables are considered valid if the 
        /// begin with one or more letters and end with one or more numbers. Cell names
        /// are normalized by being placed into all uppercase. "default" is the default version.
        /// </summary>
        public Spreadsheet()
            :base(s => true, s => s, "default")
        {
            dependencies = new DependencyGraph();
            cells = new Dictionary<String, Cell>();
            Changed = false;
        }

        /// <summary>
        /// Creates an empty spreadsheet using the provided validity and normalization methods and version.
        /// </summary>
        /// <param name="isValid">Method to determine validity of cell names</param>
        /// <param name="normalize">Method to normalize cell names</param>
        /// <param name="version">Spreadsheet version</param>
        public Spreadsheet(Func<string, bool> isValid, Func<string, string> normalize, string version)
            :base (isValid, normalize, version)
        {
            dependencies = new DependencyGraph();
            cells = new Dictionary<String, Cell>();
            Changed = true;
        }

        /// <summary>
        /// Constructs a spreadsheet from a file using the provided validity and normalizations methods and version.
        /// If the version in the file doesn't match the provided version, throws a SpreadsheetReadWriteException.
        /// </summary>
        /// <param name="filename">file to open</param>
        /// <param name="isValid">Method to determine validity of cell names</param>
        /// <param name="normalize">Method to normalize cell names</param>
        /// <param name="version">Spreadsheet version</param>
        public Spreadsheet(string filename, Func<string, bool> isValid, Func<string, string> normalize, string version)
            : base(isValid, normalize, version)
        {
            Changed = false;
            BuildFromFile(filename, version);
        }

        /// <summary>
        /// Blanks out the current spreadsheet and re-builds it from the given file.
        /// Throws a SpreadsheetReadWriteException if the provided version doesn't match
        /// the version in the file.
        /// </summary>
        /// <param name="filename">file to open</param>
        /// /// <param name="version">spreadsheet version</param>
        private void BuildFromFile(string filename, string version)
        {
            dependencies = new DependencyGraph();
            cells = new Dictionary<String, Cell>();

            bool haveName = false; //make sure we don't get two names in a cell
            bool haveContents = false; //make sure we don't get two contents in a cell
            string name = ""; //storage for name
            bool versionFound = false; //make sure version found
            try
            {
                using (XmlReader reader = XmlReader.Create(filename))
                {
                    while (reader.Read())
                    {
                        if (reader.IsStartElement())
                        {
                            switch (reader.Name)
                            {
                                //check that the spreadsheet is the correct version
                                case "spreadsheet":
                                    if (!reader["version"].Equals(version))
                                        throw new SpreadsheetReadWriteException("Mismatched versions");
                                    versionFound = true;
                                    break;

                                case "cell": //reset our name/contents checking
                                    haveName = false;
                                    haveContents = false;
                                    break;

                                case "name":
                                    if (haveName) //have a name already, our XML isn't formatted correctly
                                        throw new SpreadsheetReadWriteException("XML formatted incorrectly");
                                    haveName = true;
                                    reader.Read();
                                    name = reader.Value;
                                    CheckName(name); //check that the name is valid
                                    name = Normalize(name);
                                    break;

                                case "contents":
                                    if (haveContents || !haveName) //have contents already or don't have a name, our XML isn't formatted correctly
                                        throw new SpreadsheetReadWriteException("XML formatted incorrectly");
                                    haveContents = true;
                                    reader.Read();
                                    SetContentsOfCell(name, reader.Value); //add the cell to the spreadsheet
                                    break;

                                default:
                                    throw new SpreadsheetReadWriteException("Unknown element encountered");
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                throw new SpreadsheetReadWriteException(e.Message);
            }
            //if the version wasn't found, spreadsheet is bad
            if (!versionFound)
                throw new SpreadsheetReadWriteException("No version information found");
        }

        /// <see cref="AbstractSpreadsheet.GetNamesOfAllNonemptyCells"/>
        public override IEnumerable<String> GetNamesOfAllNonemptyCells()
        {
            return cells.Keys;
        }

        /// <see cref="AbstractSpreadsheet.GetCellContents"/>
        public override object GetCellContents(String name)
        {
            //check that the name is valid
            CheckName(name);
            name = Normalize(name);
            //if it is (since CheckName throws the exception if it isn't), get the cell from the set
            Cell c;
            if (cells.TryGetValue(name, out c))
                return c.Contents;
            //if we couldn't find a cell by that name, it means it's empty, so return an empty string
            return "";
        }

        /// <see cref="AbstractSpreadsheet.SetCellContents(string, double)"/>
        protected override ISet<String> SetCellContents(String name, double number)
        {
            //check that the name is a valid cell name
            CheckName(name);
            name = Normalize(name);
            //add the cell contents
            UpdateCell(name, number);
            //set its dependees to an empty set
            dependencies.ReplaceDependees(name, new HashSet<string>());

            return new HashSet<String>(GetCellsToRecalculate(name));
        }

        /// <see cref="AbstractSpreadsheet.SetCellContents(string, string)"/>
        protected override ISet<String> SetCellContents(String name, String text)
        {
            //check that the name is a valid cell name
            CheckName(name);
            name = Normalize(name);
            //if the text is the empty string, remove the cell from the set of cells
            if (text.Equals(""))
                cells.Remove(name);
            //otherwise, add/change the cell
            else
                UpdateCell(name, text);
            //set its dependees to an empty set
            dependencies.ReplaceDependees(name, new HashSet<string>());

            return new HashSet<String>(GetCellsToRecalculate(name));
        }

        /// <see cref="AbstractSpreadsheet.SetCellContents(String, Formula)"/>
        protected override ISet<String> SetCellContents(String name, Formula formula)
        {
            //check that the cell name is valid
            CheckName(name);
            name = Normalize(name);

            //set up a queue and check each variable in the formula and its dependees for an
            //already existing dependency relationship with "name". If one exists, we would 
            //end up with a circular dependency, so throw an exception.
            Queue<string> ToCheck = new Queue<string>(formula.GetVariables());
            //also, keep track of any cells that will depend directly on this formula if we add it
            HashSet<string> dependees = new HashSet<string>(ToCheck);

            //save the old cell contents, add the formula to the cell and update the dependency graph
            object oldContents = GetCellContents(name);
            UpdateCell(name, formula);

            //set up the dependees for this cell
            dependencies.ReplaceDependees(name, dependees);

            //add this cell as a dependent to all the variable cells it references
            foreach (string dependee in dependees)
                dependencies.AddDependency(dependee, name);

            ISet<String> deps;
            try
            {
                //check for circular dependency and get cells to update
                deps = new HashSet<String>(GetCellsToRecalculate(name));
            }
            catch (CircularException c)
            {
                //set cell back to its old value, then re-throw the exception
                UpdateCell(name, oldContents);
                throw c;
            }

            //return this cell's direct and indirect dependents
            return deps;
        }

        /// <see cref="AbstractSpreadsheet.GetDirectDependents"/>
        protected override IEnumerable<String> GetDirectDependents(String name)
        {
            if (name == null)
                throw new ArgumentNullException();

            //check that the name is valid for a cell name
            CheckName(name);
            name = Normalize(name);

            //if the cell name is valid, return its dependents
            return dependencies.GetDependents(name);
        }

        /// <summary>
        /// Checks if a given string is a valid cell name. A cell name is 
        /// valid if and only if it consists of one or more letters followed
        /// by one or more digits.
        /// </summary>
        /// <param name="name">The name to check</param>
        private void CheckName(String name)
        {
            //name pattern- one or more characters followed by one or more numbers
            String pattern = @"^[a-zA-Z]+[0-9]+$";

            if (name == null || !Regex.IsMatch(name, pattern) || !IsValid(name))
                throw new InvalidNameException();
        }

        /// <summary>
        /// Adds or updates the cell contents for the given name
        /// </summary>
        /// <param name="name">Cell name</param>
        /// <param name="contents">Cell contents</param>
        private void UpdateCell(string name, object contents)
        {
            Cell c;
            name = Normalize(name);
            //update the cell if we already have one by that name
            if (cells.TryGetValue(name, out c))
                c.Contents = contents;
            //otherwise, add a new cell
            else
                cells.Add(name, new Cell(name, contents));
        }

        /// <see cref="AbstractSpreadsheet.GetSavedVersion"/>
        public override string GetSavedVersion(String filename)
        {
            try
            {
                using (XmlReader reader = XmlReader.Create(filename))
                {
                    while (reader.Read())
                    {
                        if (reader.IsStartElement())
                        {
                            switch (reader.Name)
                            {
                                //check that the spreadsheet is the correct version
                                case "spreadsheet":
                                    return reader["version"];
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                throw new SpreadsheetReadWriteException(e.Message);
            }
            throw new SpreadsheetReadWriteException("No version information found!");
        }

        /// <see cref="AbstractSpreadsheet.Save"/>
        public override void Save(String filename)
        {
            Changed = false;

            try
            {
                using (XmlWriter writer = XmlWriter.Create(filename))
                {
                    writer.WriteStartDocument();
                    writer.WriteStartElement("spreadsheet");
                    writer.WriteAttributeString("version", Version);

                    //write an element for each non-empty cell
                    foreach (Cell c in cells.Values)
                    {
                        //start of cell
                        writer.WriteStartElement("cell");
                        writer.WriteElementString("name", c.Name);
                        writer.WriteElementString("contents", GetStringContents(c));
                        writer.WriteEndElement();
                        //end of cell
                    }

                    writer.WriteEndElement();
                    writer.WriteEndDocument();
                }
            }
            catch (Exception e)
            {
                throw new SpreadsheetReadWriteException(e.Message);
            }
        }

        /// <summary>
        /// Gets the contents of a cell as a string to be written into 
        /// an XML document.
        /// </summary>
        /// <param name="c">The cell</param>
        /// <returns>The cells contents as a string</returns>
        private string GetStringContents(Cell c)
        {
            if (c.Contents is string)
                return (string)c.Contents;
            else if (c.Contents is Formula)
                return "=" + c.Contents.ToString();
            return c.Contents.ToString();
        }

        /// <see cref="AbstractSpreadsheet.GetCellValue"/>
        public override object GetCellValue(String name)
        {
            //check that the name is valid
            CheckName(name);
            name = Normalize(name);
            //get the cell corresponding to that name
            Cell c;
            if (cells.TryGetValue(name, out c))
            {
                if (c.Contents is string || c.Contents is String)
                    return c.Contents;
                else
                    return c.GetValue();
            }
            //return an empty string if the cell is empty
            return "";
        }

        /// <see cref="AbstractSpreadsheet.SetContentsOfCell"/>
        public override ISet<String> SetContentsOfCell(String name, String content)
        {
            //if contents or name null, throw exception
            if (content == null)
                throw new ArgumentNullException();

            //check that the name is valid
            CheckName(name);
            name = Normalize(name);

            Changed = true;
            ISet<string> ChangedCells;

            double outnum;
            if (Double.TryParse(content, out outnum))
                ChangedCells = SetCellContents(name, outnum);
            else if (content.Length > 0 && content[0] == '=')
                ChangedCells = SetCellContents(name, new Formula(content.Substring(1), IsValid, Normalize));
            else
                ChangedCells = SetCellContents(name, content);

            //recalculate the cells
            foreach (string s in ChangedCells)
            {
                Cell c;
                if (cells.TryGetValue(s, out c))
                    c.Recalculate(GetDoubleValue);
            }

            return ChangedCells;
        }

        /// <summary>
        /// Lookup method to get double values from cells. Throws an argument exception
        /// if the value of the cell isn't a double.
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        private double GetDoubleValue(String name)
        {
            //check that the name is valid and normalizes it
            CheckName(name);
            name = Normalize(name);

            object val = GetCellValue(name);
            //if (val.Equals(""))
             //   return 0;
            if (val is double)
                return (double) val;
            else
                throw new ArgumentException("Cell value of " + name + " is not a double");
        }
    }
}

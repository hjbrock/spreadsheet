// Author: Hannah Brock
// CS 3500, Fall 2012

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace SpreadsheetUtilities
{
    /// <summary>
    /// Represents formulas written in standard infix notation using standard precedence
    /// rules.  The allowed symbols are non-negative numbers written using floating-point
    /// syntax, variables that consist of one or more letters followed by one or more 
    /// digits, parentheses, and the four operator symbols +, -, *, and /.
    /// </summary>
    public class Formula
    {
        //Representation invariant: no strings are upper- or lower-cased on input to this list
        //                          no tokens are removed once added to the list
        //Abstraction function: The tokens exactly as input in the Formula constructor
        private List<string> tokens;

        //The string used to construct this formula, with no spaces
        private string formula;

        //Representation invariant: No variable will appear twice in this set
        //                          After construction, all variables in the formula appear in the set and are never removed
        //Abstraction function: The set of all variables used by this formula
        private HashSet<string> variables;

        //Set of all possible operators, none are removed
        private HashSet<string> operators = new HashSet<string>(new string[] {"+", "*", "/", "-"});

        /// <summary>
        /// Method used to determine whether a string that consists of one or more letters
        /// followed by one or more digits is a valid variable name.
        /// </summary>
        public Func<string, bool> IsValid;

        /// <summary>
        /// Method used to convert a cell name to its standard form.  For example,
        /// Normalize might convert names to upper case.
        /// </summary>
        public Func<string, string> Normalize;

        /// <summary>
        /// Creates a Formula from a string that consists of an infix expression written as
        /// described in the class comment.  If the expression is syntacticaly invalid,
        /// throws a FormulaFormatException with an explanatory Message.
        /// </summary>
        public Formula(String formula, Func<string, bool> IsValid, Func<string, string> Normalize)
        {
            this.formula = "";
            this.IsValid = IsValid;
            this.Normalize = Normalize;
            tokens = new List<string>();
            variables = new HashSet<string>();
            int ltParensCount = 0; //number of left parens seen at any given time
            int rtParensCount = 0; //number of right parens seen at any given time
            string previous = null; //the previous token at any given time
            double num; //place to store result of TryParse when trying to parse a number

            foreach (string token in GetTokens(formula))
            {
                string trimmed = token.Trim();
                

                //increment parens counts if necessary
                if (trimmed.Equals("("))
                    ltParensCount++;
                else if (trimmed.Equals(")"))
                    rtParensCount++;
                //if token isn't a number, check if it's an operator, and if not, check if it's a variable
                else if (!Double.TryParse(trimmed, out num))
                {
                    if (!operators.Contains(trimmed))
                        trimmed = CheckVariable(trimmed);
                }
                else
                    trimmed = num.ToString();

                CheckParens(ltParensCount, rtParensCount);
                CheckPrevious(trimmed, previous);

                previous = token;
                this.formula += trimmed;
                tokens.Add(trimmed);
            }

            //check to make sure we had at least one token
            if (tokens.Count == 0)
                throw new FormulaFormatException("Formula must have at least one token");
            //check that parens count is equal
            if (ltParensCount != rtParensCount)
                throw new FormulaFormatException("Incorrect parenthesis matching");
            //check that last token is correct
            CheckPrevious(")", previous);
            //if (operators.Contains(previous) || previous.Equals("("))
                //throw new FormulaFormatException("The last token of an expression must be a number, a variable, or a closing parenthesis");

        }

        /// <summary>
        /// Checks if a token is a valid variable. Throws a FormulaFormatException if it is not.
        /// Adds it to the list of variables if it is.
        /// </summary>
        /// <param name="token">The current token</param>
        private string CheckVariable(string token)
        {
            String pattern = @"[a-zA-Z]+[0-9]+";

            if (!Regex.IsMatch(token, pattern) || !IsValid(token))
                throw new FormulaFormatException("Invalid variable format");

            token = Normalize(token);

            variables.Add(token);
            return token;
        }

        /// <summary>
        /// Checks if the left parens count is greater than or equal to the right and throws
        /// a FormulaFormatException if it is not.
        /// </summary>
        /// <param name="left">Current number of left parens</param>
        /// <param name="right">Current number of right parens</param>
        private void CheckParens(int left, int right)
        {
            if (!(left >= right))
                throw new FormulaFormatException("Incorrect parenthesis matching");
        }

        /// <summary>
        /// Checks that the current token is valid based on the previous token.
        /// Throws a FormulaFormatException if the token is not valid.
        /// </summary>
        /// <param name="current">Current token</param>
        /// <param name="previous">Previous token</param>
        private void CheckPrevious(string current, string previous)
        {
            //The first token of an expression must be a number, a variable, or an opening parenthesis.
            //Any token that immediately follows an opening parenthesis or an operator must be either a number, a variable, or an opening parenthesis.
            if (previous == null || previous.Equals("(") || operators.Contains(previous))
            {
                if (operators.Contains(current) || current.Equals(")"))
                    throw new FormulaFormatException("An expression may not begin with an operator or closing parenthesis");
            }
            //Any token that immediately follows a number, a variable, or a closing parenthesis must be either an operator or a closing parenthesis.
            else
            {
                if (!operators.Contains(current) && !current.Equals(")"))
                    throw new FormulaFormatException("Only operators or closing parenthesis may follow a number, variable, or closing parenthesis");
            }
        }

        /// <summary>
        /// Evaluates this Formula, using the lookup delegate to determine the values of
        /// variables.  
        /// 
        /// Given a variable symbol as its parameter, lookup returns the
        /// variable's value (if it has one) or throws an ArgumentException (otherwise).
        /// 
        /// If no undefined variables or divisions by zero are encountered when evaluating 
        /// this Formula, the value is returned.  Otherwise, a FormulaError is returned.  
        /// The Reason property of the FormulaError should have a meaningful explanation.
        /// </summary>
        public object Evaluate(Func<string, double> lookup)
        {
            Stack<string> operatorStack = new Stack<string>(); //store operators as we evaluate
            Stack<double> valueStack = new Stack<double>(); //store values as we evaluate
            double currentValue = 0;
            try
            {
                foreach (string t in tokens)
                {
                    //token is a double
                    if (Double.TryParse(t, out currentValue))
                    {
                        valueStack.Push(currentValue);
                        if (operatorStack.Count != 0 && (operatorStack.Peek().Equals("*") || operatorStack.Peek().Equals("/")))
                            valueStack.Push(EvaluateNext(valueStack, operatorStack));
                    }
                    //token is "("
                    else if (t.Equals("("))
                        operatorStack.Push(t);
                    //token is "+" or "-"
                    else if (t.Equals("+") || t.Equals("-"))
                    {
                        //see if we need to complete previous operations
                        if (operatorStack.Count != 0 && (operatorStack.Peek().Equals("+") || operatorStack.Peek().Equals("-")))
                            valueStack.Push(EvaluateNext(valueStack, operatorStack));

                        operatorStack.Push(t);
                    }
                    //token is "*" or "/"
                    else if (t.Equals("*") || t.Equals("/"))
                    {
                        operatorStack.Push(t);
                    }
                    //token is ")"
                    else if (t.Equals(")"))
                    {
                        if (!operatorStack.Peek().Equals("("))
                            valueStack.Push(EvaluateNext(valueStack, operatorStack));
                        //pop opening parens
                        operatorStack.Pop();
                        //check for pending multiplication
                        if (operatorStack.Count != 0 && (operatorStack.Peek().Equals("*") || operatorStack.Peek().Equals("/")))
                            valueStack.Push(EvaluateNext(valueStack, operatorStack));
                    }
                    else
                    {
                        //check if it's a variable, push to value stack if it is
                        try
                        {
                            valueStack.Push(lookup(t));
                        }
                        catch (Exception e)
                        {
                            return new FormulaError(e.Message);
                        }
                        //check for pending multiplication
                        if (operatorStack.Count != 0 && (operatorStack.Peek().Equals("*") || operatorStack.Peek().Equals("/")))
                            valueStack.Push(EvaluateNext(valueStack, operatorStack));
                    }
                }

                //finish up with remaining operators and values
                while (operatorStack.Count > 0)
                    valueStack.Push(EvaluateNext(valueStack, operatorStack));

                return valueStack.Pop();
            }
            catch (ArgumentException e)
            {
                return new FormulaError(e.Message);
            }
        }

        /// <summary>
        /// Performs a mathematical operation on two integers, defined by the given operator.
        /// </summary>
        /// <param name="v1">2nd argument (i.e. divisor for division)</param>
        /// <param name="v2">1st argument</param>
        /// <param name="op">Mathematical operator</param>
        /// <returns>Returns the integer value of the expression. Throws an ArgumentException if the operator isn't a valid mathematical operator.</returns>
        private static double EvaluateNext(double v2, double v1, string op)
        {
            if (op.Equals("+"))
                return v1 + v2;
            if (op.Equals("-"))
                return v1 - v2;
            if (op.Equals("*"))
                return v1 * v2;
            else
            {
                //don't allow division by 0
                if (v2 == 0)
                    throw new ArgumentException("Cannot divide by 0");
                return v1 / v2;
            }
        }

        /// <summary>
        /// Performs the next operation on a value stack and an operator stack and returns the result.
        /// </summary>
        /// <param name="valueStack">The stack of values</param>
        /// <param name="operatorStack">The stack of operators</param>
        /// <returns>Returns the value of the operation on the next two values from the value stack, defined by the next operator.</returns>
        private static double EvaluateNext(Stack<double> valueStack, Stack<string> operatorStack)
        {
            double v2 = valueStack.Pop();
            double v1 = valueStack.Pop();
            string op = operatorStack.Pop();
            return EvaluateNext(v2, v1, op);
        }

        /// <summary>
        /// Enumerates all of the variables that occur in this formula.  No variable
        /// may appear more than once in the enumeration, even if it appears more than
        /// once in this Formula.
        /// </summary>
        public IEnumerable<String> GetVariables()
        {
            return variables;
        }

        /// <summary>
        /// Returns a string containing no spaces which, if passed to the Formula
        /// constructor, will produce a Formula f such that this.Equals(f).
        /// </summary>
        public override string ToString()
        {
            return formula;
        }

        /// <summary>
        /// If obj is null or obj is not a Formula, returns false.  Otherwise, reports
        /// whether or not this Formula and obj are equal.
        /// 
        /// Two Formulae are considered equal if they consist of the same tokens in the
        /// same order.  All tokens are compared as strings except for numeric tokens,
        /// which are compared as doubles.
        /// 
        /// Here are some examples.  
        /// new Formula("x1+y2").Equals(new Formula("x1  +  y2")) is true
        /// new Formula("x1+y2").Equalas(new Formula("y2+x1")) is false
        /// new Formula("2.0 + x7").Equals(new Formula("2.000 + x7")) is true
        /// </summary>
        public override bool Equals(object obj)
        {
            if (obj == null || !(obj is Formula))
                return false;

            Formula compare = (Formula)obj;
            //if the number of tokens are different, the objects aren't equal
            if (tokens.Count != compare.tokens.Count)
                return false;

            double num; //place to put result of double's tryparse
            for (int i = 0; i < tokens.Count; i++)
            {
                if (!Double.TryParse(tokens[i], out num))
                {
                    if (!tokens[i].Equals(compare.tokens[i]))
                        return false;
                }
                else
                {
                    if (!CheckDouble(num, compare.tokens[i]))
                        return false;
                }
            }
            return true;
        }

        /// <summary>
        /// Checks if the string token is a double and if it matches the input double
        /// </summary>
        /// <param name="num">The double to match</param>
        /// <param name="token">The token to check</param>
        /// <returns>Returns true if the token is a double and matches the input number.</returns>
        private bool CheckDouble(double num, string token)
        {
            double num2;
            if (!Double.TryParse(token, out num2))
                return false;
            if (num != num2)
                return false;
            return true;
        }

        /// <summary>
        /// Reports whether f1 == f2, using the notion of equality from the Equals method.
        /// Note that if both f1 and f2 are null, this method should return true.  If one is
        /// null and one is not, this method should return false.
        /// </summary>
        public static bool operator ==(Formula f1, Formula f2)
        {
            if ((object)f1 == null && (object)f2 == null)
                return true;
            if ((object)f1 == null || (object)f2 == null)
                return false;
            return f1.Equals(f2);
        }

        /// <summary>
        /// Reports whether f1 != f2, using the notion of equality from the Equals method.
        /// Note that if both f1 and f2 are null, this method should return false.  If one is
        /// null and one is not, this method should return true.
        /// </summary>
        public static bool operator !=(Formula f1, Formula f2)
        {
            return !(f1 == f2);
        }

        /// <summary>
        /// Returns a hash code for this Formula.
        /// </summary>
        public override int GetHashCode()
        {
            return this.ToString().GetHashCode();
        }

        /// <summary>
        /// Given a formula, enumerates the tokens that compose it.  Tokens are left paren,
        /// right paren, one of the four operator symbols, a string consisting of one or more
        /// letters followed by one or more digits, a double literal, and anything that doesn't
        /// match one of those patterns.  There are no empty tokens, and no token contains white space.
        /// </summary>
        private static IEnumerable<string> GetTokens(String formula)
        {
            // Patterns for individual tokens
            String lpPattern = @"\(";
            String rpPattern = @"\)";
            String opPattern = @"[\+\-*/]";
            String varPattern = @"[a-zA-Z]+\d+";
            String doublePattern = @"(?: \d+\.\d* | \d*\.\d+ | \d+ ) (?: e[\+-]?\d+)?";
            String spacePattern = @"\s+";

            // Overall pattern
            String pattern = String.Format("({0}) | ({1}) | ({2}) | ({3}) | ({4}) | ({5})",
                                            lpPattern, rpPattern, opPattern, varPattern, doublePattern, spacePattern);

            // Enumerate matching tokens that don't consist solely of white space.
            foreach (String s in Regex.Split(formula, pattern, RegexOptions.IgnorePatternWhitespace))
            {
                if (!Regex.IsMatch(s, @"^\s*$", RegexOptions.Singleline))
                {
                    yield return s;
                }
            }

        }
    }

    /// <summary>
    /// Used to report syntactic errors in the argument to the Formula constructor.
    /// </summary>
    public class FormulaFormatException : Exception
    {
        /// <summary>
        /// Constructs a FormulaFormatException containing the explanatory message.
        /// </summary>
        public FormulaFormatException(String message)
            : base(message)
        {
        }
    }

    /// <summary>
    /// Used as a possible return value of the Formula.Evaluate method.
    /// </summary>
    public struct FormulaError
    {
        /// <summary>
        /// Constructs a FormulaError containing the explanatory reason.
        /// </summary>
        /// <param name="reason"></param>
        public FormulaError(String reason)
            : this()
        {
            Reason = reason;
        }

        /// <summary>
        ///  The reason why this FormulaError was created.
        /// </summary>
        public string Reason { get; private set; }
    }
}


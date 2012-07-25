using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;

namespace Phakisa
{
    class clsTableFormulas
    {
         #region PROPERTIES

            private string strDatabaseName = "";
            private string strTableName = "";
            private string strTableFormulaName = "";
            private string strTableFormulaCall = "";
            private string strCalcName = "";
            private string strCalcSeq = "";
            private string strA = "";
            private string strB = "";
            private string strC = "";
            private string strD = "";
            private string strE = "";
            private string strF = "";
            private string strG = "";
            private string strH = "";
            private string strI = "";
            private string strJ = "";
            private string strSQL = "";
            private string strSaveColumn = "";

            private string strTempCalcName = "";
            private string strTempTableFormulaCall = "";
            private string strTempTableFormulaName = "";

            public string DatabaseName
            {
                get { return strDatabaseName; }
                set { strDatabaseName = value; }
            }

            public string Tablename
            {
                get { return strTableName; }
                set { strTableName = value; }
            }

            public string TableFormulaName
            {
                get { return strTableFormulaName; }
                set { strTableFormulaName = value; }
            }

            public string propSQL
            {
                get { return strSQL; }
                set { strSQL = value; }
            }

            public string TableFormulaCall
            {
                get { return strTableFormulaCall; }
                set { strTableFormulaCall = value; }
            }

            public string CalcName
            {
                get { return strCalcName; }
                set { strCalcName = value; }
            }

            public string CalcSeq
            {
                get { return strCalcSeq; }
                set { strCalcSeq = value; }
            }

            public string A
            {
                get { return strA; }
                set { strA = value; }
            }

            public string B
            {
                get { return strB; }
                set { strB = value; }
            }


            public string C
            {
                get { return strC; }
                set { strC = value; }
            }

            public string D
            {
                get { return strD; }
                set { strD = value; }
            }


            public string E
            {
                get { return strE; }
                set { strE = value; }
            }


            public string F
            {
                get { return strF; }
                set { strF = value; }
            }


            public string G
            {
                get { return strG; }
                set { strG = value; }
            }

            public string H
            {
                get { return strH; }
                set { strH = value; }
            }

            public string I
            {
                get { return strI; }
                set { strI = value; }
            }

            public string J
            {
                get { return strJ; }
                set { strJ = value; }
            }

            public string SaveColumn
            {
                get { return strSaveColumn; }
                set { strSaveColumn = value; }
            }

            public string tempCalcName
            {
                get { return strTempCalcName; }
                set { strTempCalcName = value; }
            }

            public string tempTableFormulaCall
            {
                get { return strTempTableFormulaCall; }
                set { strTempTableFormulaCall = value; }
            }

            public string tempTableFormulaName
            {
                get { return strTempTableFormulaName; }
                set { strTempTableFormulaName = value; }
            }

            #endregion

#region Methods

            public string writeSQL(string inputstring, int position, string strCombiner, int intIncreaser)
            {
                string outputstring = "";
                if (position > 0)
                // return outputstring = inputstring.Substring(0, inputstring.IndexOf(strCombiner) + intIncreaser); 
                {
                    //return outputstring = inputstring.Substring(0, inputstring.IndexOf(strCombiner) + intIncreaser); 
                    return outputstring = inputstring.Substring(0, position + intIncreaser);
                }
                else
                {

                    outputstring = inputstring.Substring(0);
                    return outputstring;
                }

            }

            public void divideSQLmerge(int intIncreaser)
            {
                A = "";
                B = "";
                C = "";
                D = "";
                E = "";
                F = "";
                G = "";
                H = "";
                I = "";
                J = "";
                string _sqlReturned = propSQL;
                int _index = intIncreaser;
                A = writeSQL(_sqlReturned, _index, "", 0) + "#";
                _sqlReturned = propSQL.Substring(A.Trim().Length -1);   //must be -1
                B = "#" + writeSQL(_sqlReturned, _index, "", 0) + "#"; ;
                _sqlReturned = _sqlReturned.Substring(B.Length - 2);
                C = "#" + writeSQL(_sqlReturned, _index, "", 0) + "#"; ;
                _sqlReturned = _sqlReturned.Substring(C.Length - 2);
                D = "#" + writeSQL(_sqlReturned, _index, "", 0) + "#"; ;
                _sqlReturned = _sqlReturned.Substring(D.Length - 2);
                E = "#" + writeSQL(_sqlReturned, _index, "", 0) + "#"; ;
                _sqlReturned = _sqlReturned.Substring(E.Length - 2);
                F = "#" + writeSQL(_sqlReturned, _index, "", 0) + "#"; ;
                _sqlReturned = _sqlReturned.Substring(F.Length - 2);
                G = "#" + writeSQL(_sqlReturned, _index, "", 0) + "# "; ;
                _sqlReturned = "#" + _sqlReturned.Substring(G.Length-3);
                H = "#" + _sqlReturned;
            }

            public void divideSQLCustom(int intIncreaser)
            {
                string _sqlReturned = propSQL;
                int _index = intIncreaser;

                A = writeSQL(_sqlReturned, _index, "", 0) + "#"; ;
                _sqlReturned = propSQL.Substring(A.Trim().Length-1);

                B = "#" + writeSQL(_sqlReturned, _index, "", 0) + "#"; ;
                _sqlReturned = _sqlReturned.Substring(B.Trim().Length-2);

                C = "#" +  _sqlReturned;
            }

            public void divideSQL(string strCombiner,int intIncreaser)
            {

                string _sqlReturned = propSQL;
                int _index = _sqlReturned.IndexOf(strCombiner);
                A = writeSQL(_sqlReturned, _index, strCombiner, intIncreaser) + "#";

                if (_index > 0)
                {
                    _sqlReturned = propSQL.Substring(A.Trim().Length-1);
                    _index = _sqlReturned.IndexOf(strCombiner);
                    B = "#" + writeSQL(_sqlReturned, _index, strCombiner, intIncreaser) + "#";


                    if (_index > 0)
                    {
                        _sqlReturned = propSQL.Substring(A.Trim().Length + B.Trim().Length - 2);
                        _index = _sqlReturned.IndexOf(strCombiner);
                        C = "#" + writeSQL(_sqlReturned, _index, strCombiner, intIncreaser) + "#";


                        if (_index > 0)
                        {
                            _sqlReturned = propSQL.Substring(A.Trim().Length + B.Trim().Length -2 + C.Trim().Length -2);
                            _index = _sqlReturned.IndexOf(strCombiner);
                            D = "#" + writeSQL(_sqlReturned, _index, strCombiner, intIncreaser) + "#";


                            if (_index > 0)
                            {
                                _sqlReturned = propSQL.Substring(A.Trim().Length + B.Trim().Length - 2 + C.Trim().Length - 2 + D.Trim().Length - 1);
                                _index = _sqlReturned.IndexOf(strCombiner);
                                E = "#" + writeSQL(_sqlReturned, _index, strCombiner, intIncreaser) + "#";
                            }
                        }
                    }
                }
            }

#endregion


            internal void clearVariablesAtoE()
            {
                A = "X";
                B = "X";
                C = "X";
                D = "X";
                E = "X";
                F = "X";
                G = "X";
                H = "X";
                I = "X";
                J = "X";


            }
    }
    }


using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace Clippy
{
    internal class SyntaxHighlighter
    {
        private bool InLine;

        private List<string> keyWords = new List<string>()
        {
            "if",
            "then",
            "elsif",
            "else",
            "endif",
            "do",
            "for",
            "to",
            "по",
            "each",
            "in",
            "while",
            "enddo",
            "procedure",
            "endprocedure",
            "function",
            "endfunction",
            "var",
            "export",
            "goto",
            "and",
            "or",
            "not",
            "val",
            "break",
            "continue",
            "return",
            "try",
            "except",
            "endtry",
            "raise",
            "false",
            "true",
            "undefined",
            "null",
            "new",
            "execute"
        };

        private string keyWordPattern = "";
        private string operatorPattern = @"[\+\*\/\%\=\>\<]";
        private string notLetterPattern = @"(?:[^ёа-яa-z]|^|$)";

        public SyntaxHighlighter()
        {
            foreach (var keyWord in keyWords)
            {
                if (false
                    || keyWord == "false"
                    || keyWord == "true"
                    || keyWord.Length < 4)
                {
                    continue;
                }
                if (!String.IsNullOrEmpty(keyWordPattern))
                {
                    keyWordPattern += "|";
                }
                keyWordPattern += keyWord;
            }
            keyWordPattern = notLetterPattern + "(?:" + keyWordPattern + ")" + notLetterPattern;
        }

        private string StylePart()
        {
            return
            @"pre
            {
             font - family: Courier;
             color: #0000FF;
             font - size: 9pt;
            }
            .k
            {
             color: red;
            }
            .c
            {
             color: green;
            }
            .s
            {
             color: black;
            }
            .n
            {
             color: black;
            }
            .p
            {
             color: brown;
            }
            ";
        }

        private string GetSymbol(string Line, int Position)
        {
            if (Line.Length > Position)
                return Line.Substring(Position, 1);
            return "";
        }

        private bool IsKeyword(string _Токен)
        {
            string Token = _Токен.ToLower();
            return keyWords.Contains(Token);
        }

        private bool IsSpecialSymbol(string _Symbol)
        {
            string Symbol = _Symbol.ToLower();
            return false ||
                   Symbol == ")" ||
                   Symbol == "(" ||
                   Symbol == "[" ||
                   Symbol == "]" ||
                   Symbol == "." ||
                   Symbol == "," ||
                   Symbol == "=" ||
                   Symbol == "+" ||
                   Symbol == "-" ||
                   Symbol == "<" ||
                   Symbol == ">" ||
                   Symbol == ";" ||
                   Symbol == "?" ||
                   Symbol == "*";
        }

        private void ProcessToken(ref string codeLine, string Token, ref int Position, string Класс)
        {
            int tokenLength = Token.Length;
            codeLine = codeLine.Substring(0, Position - tokenLength + 1) +
                         "<span class=" + Класс + ">" +
                         codeLine.Substring(Position - tokenLength + 1, tokenLength) +
                         "</span>" +
                         codeLine.Remove(0, Position + 1);
            Position = Position + ("<span class=>" + "</span>" + Класс).Length;
        }

        private string ProcessLine(string codeLine)
        {
            int Position = 0;
            int State = 0;
            string Token = "";
            int LineBegin = 0;
             while (Position != codeLine.Length)
            {
                string CurrentSymbol = GetSymbol(codeLine, Position);
                if (CurrentSymbol == "/")
                    State = 1;
                else if (CurrentSymbol == "\t" || CurrentSymbol == " ")
                    State = 2;
                else if (CurrentSymbol == "\"")
                    State = 3;
                else if (CurrentSymbol == "")
                    State = 5;
                else if (IsSpecialSymbol(CurrentSymbol))
                    State = 6;
                else if (CurrentSymbol == "#" || CurrentSymbol == "&")
                    State = 8;
                else
                    State = 4;

                if (State == 1)
                {
                    if (!InLine)
                    {
                        if (GetSymbol(codeLine, Position + 1) == "/")
                        {
                            codeLine = codeLine.Substring(0, Position) +
                                       "<span class=c>" +
                                       System.Net.WebUtility.HtmlEncode(codeLine.Remove(0, Position)) +
                                       "</span>";
                            return codeLine;
                        }
                        else
                        {
                            ProcessToken(ref codeLine, CurrentSymbol, ref Position, "k");
                            Token = "";
                        }
                    }
                }
                else if (State == 2)
                {
                    if (!InLine)
                    {
                        if (!String.IsNullOrWhiteSpace(Token))
                        {
                            Position = Position - 1;
                            if (IsKeyword(Token))
                            {
                                ProcessToken(ref codeLine, Token, ref Position, "k");
                            }
                            else
                            {
                                       bool success = false;
                                try
                                {
                                    int dummy = Convert.ToInt32(Token);
                                    success = true;
                                }
                                catch
                                { }
                                if (success)
                                    ProcessToken(ref codeLine, Token, ref Position, "n");
                            }
                            Position = Position + 1;
                            Token = "";
                        }
                    }
                }

                else if (State == 3)
                {

                    if (InLine)
                    {
                        ProcessStringLiteral(LineBegin, ref codeLine, ref Position);
                        Token = "";
                        InLine = false;
                    }
                    else
                    {
                        LineBegin = Position;
                        InLine = true;
                    }
                }
                else if (State == 6)
                {
                    if (!InLine)
                    {
                        if (!String.IsNullOrWhiteSpace(Token))
                        {
                            Position--;
                            if (IsKeyword(Token) && (CurrentSymbol != "."))
                            {
                                ProcessToken(ref codeLine, Token, ref Position, "k");
                            }
                            else
                            {

                                bool success = false;
                                try
                                {
                                    int dummy = Convert.ToInt32(Token);
                                    success = true;
                                }
                                catch
                                { }
                                if (success)
                                    ProcessToken(ref codeLine, Token, ref Position, "n");
                            }
                            Position++;
                            Token = "";
                        }
                        ProcessToken(ref codeLine, CurrentSymbol, ref Position, "k");
                    }
                }
                else if (State == 8)
                {
                    if (!InLine)
                    {
                        Position = codeLine.Length - 1;
                        ProcessToken(ref codeLine, codeLine, ref Position, "p");
                    }
                }
                else if (State == 4)
                {
                    Token = Token + CurrentSymbol;
                }
                else if (State == 5)
                    break;
                Position = Position + 1;
            }


            if (InLine)
            {
                ProcessStringLiteral(LineBegin, ref codeLine, ref Position);
                Token = "";
            }

            if (!String.IsNullOrWhiteSpace(Token))
            {
                if (IsKeyword(Token))
                {
                    Position--;
                    ProcessToken(ref codeLine, Token, ref Position, "k");
                    Position++;
                }
            }
            return codeLine;
        }

        private static string ProcessStringLiteral(int LineBegin, ref string codeLine, ref int Position)
        {
            string literal = codeLine.Substring(LineBegin, Position - LineBegin);
            string newLiteral = "<span class=s>" + System.Net.WebUtility.HtmlEncode(literal) + "</span>";
            codeLine = codeLine.Substring(0, LineBegin) + newLiteral + codeLine.Remove(0, Position);
            Position = Position + newLiteral.Length - literal.Length;
            return codeLine;
        }

        public string ProcessCode(string code)
        {
            string result = "";
            //result += @"<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN"">" + "\n";
            result += "<html>";
            result += "<head>\n";
            result += "<style type=text/css>";
            result += StylePart();
            result += "</style>\n";
            result += "</head>\n";
            result += "<body>";
            result += "<!--StartFragment-->";
            result += "<pre>\n";

            code = code.Replace("\t", "    ");
            code = Regex.Replace(code, "(\r\n|\n\r|\n|\r)", "\n");
            string[] lines = code.Split("\n"[0]);
            for (int index0 = 0; index0 < lines.Length; index0++)
            {
                string codeLine = lines[index0];
                result += ProcessLine(codeLine) + "\n";
            }
            result += "</pre>";
            result += "<!--EndFragment-->";
            result += "</body>";
            result += "</html>";
            return result;
        }
    }
}
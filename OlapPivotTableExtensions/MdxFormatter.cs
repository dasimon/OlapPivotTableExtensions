using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace OlapPivotTableExtensions
{
    internal static class MdxFormatter
    {
        private static readonly HashSet<string> TopLevelKeywords = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "SELECT", "FROM", "WHERE", "WITH"
        };

        private static readonly HashSet<string> Keywords = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "SELECT", "FROM", "WHERE", "WITH", "MEMBER", "SET", "AS", "ON", "NOT", "AND", "OR",
            "IS", "CASE", "WHEN", "THEN", "ELSE", "END", "NON", "EMPTY", "DIMENSION", "PROPERTIES",
            "CELL", "COLUMNS", "ROWS", "PAGES", "CHAPTERS", "SECTIONS", "AXIS", "NULL", "EXISTING",
            "HAVING", "BY", "MATCHES", "DISTINCT", "DRILLTHROUGH", "MAXROWS", "FIRSTROWSET", "RETURN"
        };

        private static readonly HashSet<string> Functions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "IIF", "FILTER", "ORDER", "TOPCOUNT", "BOTTOMCOUNT", "TOPPERCENT", "BOTTOMPERCENT",
            "TOPSUM", "BOTTOMSUM", "GENERATE", "CROSSJOIN", "UNION", "INTERSECT", "EXCEPT",
            "DESCENDANTS", "ANCESTORS", "ANCESTOR", "ASCENDANTS", "CHILDREN", "SIBLINGS",
            "PARENT", "FIRSTCHILD", "LASTCHILD", "PREVMEMBER", "NEXTMEMBER",
            "PERIODSTODATE", "YTD", "MTD", "QTD", "WTD", "PARALLELPERIOD",
            "OPENINGPERIOD", "CLOSINGPERIOD", "LASTPERIODS", "DATEADD",
            "COUNT", "SUM", "AVG", "MIN", "MAX", "MEDIAN", "VAR", "STDDEV",
            "COALESCEEMPTY", "ISNULL", "ISEMPTY", "ISLEAF", "ISANCESTOR",
            "MEMBERS", "ALLMEMBERS", "ADDCALCULATEDMEMBERS", "NONEMPTY", "NONEMPTYCROSSJOIN",
            "HIERARCHIZE", "DISTINCT", "EXTRACT", "HEAD", "TAIL", "SUBSET", "ITEM",
            "RANK", "STRTOMEMBER", "STRTOSET", "STRTOTUPLE", "MEMBERTOSTR", "SETTOSTR",
            "TUPLETOSTR", "NAMETOSET", "UNIQUENAME", "NAME", "CAPTION", "PROPERTIES",
            "CURRENTMEMBER", "DEFAULTMEMBER", "UNKNOWNMEMBER", "DATAMEMBER",
            "LEVEL", "LEVELS", "HIERARCHY", "DIMENSION", "ROOT", "LEAVES",
            "LINKMEMBER", "WTDAVG", "AGGREGATE", "CALCULATE", "FORMAT_STRING",
            "SOLVE_ORDER", "VISIBLE", "NON_EMPTY_BEHAVIOR", "SCOPE", "THIS",
            "DRILLDOWNLEVEL", "DRILLDOWNMEMBER", "DRILLUPMEMBER", "TOGGLEDRILLSTATE",
            "VALIDMEASURE", "MEASURES"
        };

        /// <summary>
        /// Formats MDX with proper keyword-based indentation and comma placement.
        /// </summary>
        public static string Format(string mdx)
        {
            if (string.IsNullOrWhiteSpace(mdx))
                return mdx;

            mdx = mdx.Trim();
            var tokens = Tokenize(mdx);
            var sb = new StringBuilder();
            int indent = 0;
            bool newLine = false;
            bool firstToken = true;

            for (int i = 0; i < tokens.Count; i++)
            {
                Token tok = tokens[i];

                if (tok.Kind == TokenKind.Whitespace || tok.Kind == TokenKind.Newline)
                    continue;

                string upper = tok.Value.ToUpperInvariant();

                // Top-level keywords always start on a new line at indent 0
                if (tok.Kind == TokenKind.Word && TopLevelKeywords.Contains(upper))
                {
                    if (!firstToken)
                        sb.AppendLine();
                    sb.Append(upper);
                    newLine = false;
                    firstToken = false;

                    // WITH block: next MEMBER/SET gets indented
                    if (upper == "WITH")
                    {
                        indent = 1;
                        newLine = true;
                    }
                    else if (upper == "SELECT" || upper == "FROM" || upper == "WHERE")
                    {
                        indent = 1;
                        newLine = true;
                    }
                    continue;
                }

                // MEMBER and SET under WITH: new line + indent
                if (tok.Kind == TokenKind.Word && (upper == "MEMBER" || upper == "SET") && indent >= 1)
                {
                    sb.AppendLine();
                    sb.Append(IndentString(indent));
                    sb.Append(upper);
                    newLine = false;
                    firstToken = false;
                    continue;
                }

                // ON keyword signals axis spec — new line + indent
                if (tok.Kind == TokenKind.Word && upper == "ON")
                {
                    sb.AppendLine();
                    sb.Append(IndentString(indent));
                    sb.Append("ON");
                    newLine = false;
                    firstToken = false;
                    continue;
                }

                // Opening parenthesis: emit then increase indent
                if (tok.Kind == TokenKind.OpenParen)
                {
                    if (newLine)
                    {
                        sb.AppendLine();
                        sb.Append(IndentString(indent));
                    }
                    sb.Append("(");
                    indent++;
                    newLine = true;
                    firstToken = false;
                    continue;
                }

                // Closing parenthesis: decrease indent then emit on new line
                if (tok.Kind == TokenKind.CloseParen)
                {
                    indent = Math.Max(0, indent - 1);
                    sb.AppendLine();
                    sb.Append(IndentString(indent));
                    sb.Append(")");
                    newLine = false;
                    firstToken = false;
                    continue;
                }

                // Comma: emit at beginning of a new line (per original service settings)
                if (tok.Kind == TokenKind.Comma)
                {
                    sb.AppendLine();
                    sb.Append(IndentString(indent));
                    sb.Append(",");
                    newLine = false;
                    firstToken = false;
                    continue;
                }

                // Everything else
                if (newLine)
                {
                    sb.AppendLine();
                    sb.Append(IndentString(indent));
                    newLine = false;
                }
                else if (!firstToken && sb.Length > 0)
                {
                    char last = sb[sb.Length - 1];
                    if (last != '(' && last != '[' && tok.Value[0] != ')' && tok.Value[0] != ']'
                        && tok.Value[0] != '.' && last != '.')
                    {
                        sb.Append(' ');
                    }
                }

                sb.Append(tok.Value);
                firstToken = false;
            }

            return sb.ToString();
        }

        /// <summary>
        /// Applies syntax highlighting to a RichTextBox using SelectionColor.
        /// Call after setting richTextBox.Text (not Rtf).
        /// </summary>
        public static void ApplySyntaxHighlighting(RichTextBox rtb)
        {
            if (rtb == null || string.IsNullOrEmpty(rtb.Text))
                return;

            string text = rtb.Text;

            // Save scroll/selection state
            int savedStart = rtb.SelectionStart;
            int savedLen = rtb.SelectionLength;
            rtb.SuspendLayout();

            // Reset all to black first
            rtb.SelectAll();
            rtb.SelectionColor = Color.Black;

            // Apply coloring by token
            var tokens = Tokenize(text);
            int pos = 0;
            foreach (Token tok in tokens)
            {
                Color color = GetTokenColor(tok);
                if (color != Color.Black)
                {
                    rtb.Select(pos, tok.Value.Length);
                    rtb.SelectionColor = color;
                }
                pos += tok.Value.Length;
            }

            // Restore
            rtb.Select(savedStart, savedLen);
            rtb.ResumeLayout();
        }

        private static Color GetTokenColor(Token tok)
        {
            switch (tok.Kind)
            {
                case TokenKind.Word:
                    string upper = tok.Value.ToUpperInvariant();
                    if (Keywords.Contains(upper)) return Color.Blue;
                    if (Functions.Contains(upper)) return Color.DarkCyan;
                    return Color.Black;
                case TokenKind.BracketedIdentifier: return Color.DarkRed;
                case TokenKind.StringLiteral:       return Color.Red;
                case TokenKind.Number:              return Color.DarkBlue;
                case TokenKind.LineComment:
                case TokenKind.BlockComment:        return Color.Gray;
                default:                            return Color.Black;
            }
        }

        private static string IndentString(int level) => new string(' ', level * 2);

        // -----------------------------------------------------------------------
        // Tokenizer
        // -----------------------------------------------------------------------

        private enum TokenKind
        {
            Word, Number, StringLiteral, BracketedIdentifier,
            OpenParen, CloseParen, Comma, Dot, Operator,
            LineComment, BlockComment, Whitespace, Newline, Other
        }

        private struct Token
        {
            public string Value;
            public TokenKind Kind;
        }

        private static List<Token> Tokenize(string text)
        {
            var tokens = new List<Token>();
            int i = 0;
            int len = text.Length;

            while (i < len)
            {
                char c = text[i];

                // Newline
                if (c == '\r' || c == '\n')
                {
                    int start = i;
                    while (i < len && (text[i] == '\r' || text[i] == '\n')) i++;
                    tokens.Add(new Token { Value = text.Substring(start, i - start), Kind = TokenKind.Newline });
                    continue;
                }

                // Whitespace
                if (char.IsWhiteSpace(c))
                {
                    int start = i;
                    while (i < len && char.IsWhiteSpace(text[i]) && text[i] != '\r' && text[i] != '\n') i++;
                    tokens.Add(new Token { Value = text.Substring(start, i - start), Kind = TokenKind.Whitespace });
                    continue;
                }

                // Line comment --
                if (c == '-' && i + 1 < len && text[i + 1] == '-')
                {
                    int start = i;
                    while (i < len && text[i] != '\r' && text[i] != '\n') i++;
                    tokens.Add(new Token { Value = text.Substring(start, i - start), Kind = TokenKind.LineComment });
                    continue;
                }

                // Block comment /* ... */
                if (c == '/' && i + 1 < len && text[i + 1] == '*')
                {
                    int start = i;
                    i += 2;
                    while (i < len - 1 && !(text[i] == '*' && text[i + 1] == '/')) i++;
                    i += 2;
                    tokens.Add(new Token { Value = text.Substring(start, i - start), Kind = TokenKind.BlockComment });
                    continue;
                }

                // String literal "..."
                if (c == '"')
                {
                    int start = i++;
                    while (i < len && text[i] != '"') i++;
                    if (i < len) i++;
                    tokens.Add(new Token { Value = text.Substring(start, i - start), Kind = TokenKind.StringLiteral });
                    continue;
                }

                // Bracketed identifier [...]
                if (c == '[')
                {
                    int start = i++;
                    while (i < len && text[i] != ']') i++;
                    if (i < len) i++;
                    tokens.Add(new Token { Value = text.Substring(start, i - start), Kind = TokenKind.BracketedIdentifier });
                    continue;
                }

                // Number
                if (char.IsDigit(c) || (c == '-' && i + 1 < len && char.IsDigit(text[i + 1])))
                {
                    int start = i;
                    if (c == '-') i++;
                    while (i < len && (char.IsDigit(text[i]) || text[i] == '.')) i++;
                    tokens.Add(new Token { Value = text.Substring(start, i - start), Kind = TokenKind.Number });
                    continue;
                }

                // Word / keyword
                if (char.IsLetter(c) || c == '_')
                {
                    int start = i;
                    while (i < len && (char.IsLetterOrDigit(text[i]) || text[i] == '_')) i++;
                    tokens.Add(new Token { Value = text.Substring(start, i - start), Kind = TokenKind.Word });
                    continue;
                }

                // Single-char tokens
                switch (c)
                {
                    case '(': tokens.Add(new Token { Value = "(", Kind = TokenKind.OpenParen });  i++; break;
                    case ')': tokens.Add(new Token { Value = ")", Kind = TokenKind.CloseParen }); i++; break;
                    case ',': tokens.Add(new Token { Value = ",", Kind = TokenKind.Comma });       i++; break;
                    case '.': tokens.Add(new Token { Value = ".", Kind = TokenKind.Dot });         i++; break;
                    default:
                        tokens.Add(new Token { Value = c.ToString(), Kind = TokenKind.Other });
                        i++;
                        break;
                }
            }

            return tokens;
        }
    }
}

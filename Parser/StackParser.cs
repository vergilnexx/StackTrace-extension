using Parser.Enums;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace Parser
{
    public class StackParser : IStackParser
    {
        public List<Tuple<string, InformationType>> ParseStack(string input)
        {
            var output = new List<Tuple<string, InformationType>>();

            var fileExpression = @"[A-Za-z]{1}:[\\A-Za-z0-9\-_.]*.[A-Za-z]+:(строка|line) [0-9]*";
            var methodExpression = @"(at|в)\s[A-Za-z0-9\.]{5,}\(";
            var expressions = new string[]
            {
                fileExpression,
                methodExpression
            };
            var resultExpression = "(" + string.Join(")|(", expressions) + ")";

            Regex regex = new Regex(resultExpression);
            Match match = regex.Match(input);

            int proccessedIndexChar = 0;
            while (match.Success)
            {
                var link = match.Groups[0];
                if (link.Index != proccessedIndexChar)
                {
                    var contentText = input.Substring(proccessedIndexChar, link.Index - proccessedIndexChar);
                    output.Add(new Tuple<string, InformationType>(contentText, InformationType.Text));
                    proccessedIndexChar += contentText.Length;

                }

                var contentLink = link.Value;
                var subregex = new Regex(fileExpression);
                var submatch = subregex.Match(contentLink);
                if (submatch.Success) // it is the path of file
                {
                    output.Add(new Tuple<string, InformationType>(contentLink, InformationType.File));
                }
                else
                {
                    output.Add(new Tuple<string, InformationType>(contentLink, InformationType.Method));
                }
                proccessedIndexChar += contentLink.Length;

                match = match.NextMatch();
            }
            return output;
        }

        public string GetLineNumberString(string input)
        {
            var lineNumberExpression = @"(:line [0-9]+)|(:строка [0-9]+)";

            Regex regex = new Regex(lineNumberExpression);
            Match match = regex.Match(input);
            while (match.Success)
            {
                return match.Groups[0].Value;
            }

            return string.Empty;
        }

        public int GetLineNumber(string input)
        {
            var lineNumberString = GetLineNumberString(input);
            var lastSpaceIndex = lineNumberString.LastIndexOf(" ");
            var stringLineNumber = lineNumberString.Substring(lastSpaceIndex, lineNumberString.Length - lastSpaceIndex);

            var result = 0;
            int.TryParse(stringLineNumber, out result);
            return result;
        }
    }
}

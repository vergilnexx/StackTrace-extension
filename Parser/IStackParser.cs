using Parser.Enums;
using System;
using System.Collections.Generic;

namespace Parser
{
    public interface IStackParser
    {
        List<Tuple<string, InformationType>> ParseStack(string input);

        string GetLineNumberString(string input);

        int GetLineNumber(string input);
    }
}

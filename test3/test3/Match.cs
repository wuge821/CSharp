using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace test3
{
    class Match
    {
        public int Getnum(int idx, char[] str, char[] subStr, ref bool found)
        {
            int idx_org = idx;
            int i;
            while (idx < str.Length)
            {
                //找到第一个字符
                while (idx < str.Length)
                {
                    if (str[idx++] == subStr[0])
                        break;
                }
                //判断第一个字符是否匹配
                if (idx == str.Length || subStr.Length - 1 > str.Length - idx)
                    break;

                //字符相同才能匹配
                for (i = 1; i < subStr.Length; i++, idx++)
                {
                    if (subStr[i] != str[idx])
                    {//不匹配
                        idx_org++;
                        idx = idx_org;
                        break;
                    }
                }
                //整个匹配
                if (i == subStr.Length)
                {
                    found = true;
                    return idx;
                }
            }
            found = false;
            return str.Length;
        }
    }
}

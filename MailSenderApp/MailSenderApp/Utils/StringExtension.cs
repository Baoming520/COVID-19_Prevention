﻿namespace MailSenderApp.Utils
{
    #region Namespaces.
    using System;
    using System.Collections.Generic;
    using System.Linq;
    #endregion

    public static class StringExtension
    {
        public static string GetLastSegment(this string text, char mark, int offset = 0)
        {
            if (string.IsNullOrEmpty(text))
            {
                return text;
            }

            int lastIndex = text.LastIndexOf(mark);

            if (lastIndex < 0 || lastIndex + offset <= 0)
            {
                return text;
            }

            if (lastIndex + offset >= text.Length)
            {
                return string.Empty;
            }

            return text.Substring(lastIndex + 1 + offset);
        }

        public static IEnumerable<string> ParseTextWithSemicolon(this string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return null;
            }

            if (!text.Contains(';'))
            {
                return new List<string>() { text };
            }

            var items = text.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            var list = new List<string>();
            foreach (var address in items)
            {
                string addr = address;
                if (address.Contains(".pdf?"))
                {
                    int lastIndex = address.LastIndexOf('?');
                    addr = address.Remove(lastIndex);
                }

                list.Add(addr);
            }

            return list;
        }
    }
}

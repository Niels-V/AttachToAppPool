using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Management;
using System.ComponentModel;
using System.Runtime.Caching;

namespace NielsV.NielsV_AttachToAppPool
{
    public class LookupAppPools
    {
        private static object locker = new object();
        public static Dictionary<int, string> GetAppPoolProcesses()
        {
            MemoryCache cache = MemoryCache.Default;
            Dictionary<int, string> appPools = cache["NielsVAppPoolList"] as Dictionary<int, string>;
            if (appPools == null)
            {
                lock (locker)
                {
                    appPools = cache["NielsVAppPoolList"] as Dictionary<int, string>;
                    if (appPools == null)
                    {
                        appPools = RetrieveAppPools();
                        cache.Add("NielsVAppPoolList", appPools, new CacheItemPolicy { AbsoluteExpiration = DateTimeOffset.Now.AddSeconds(15) });
                    }
                }
            }


            appPools = RetrieveAppPools();
            return appPools;
        }

        private static Dictionary<int, string> RetrieveAppPools()
        {
            Dictionary<int, string> appPools = new Dictionary<int, string>();
            var processes = Process.GetProcessesByName("w3wp");

            foreach (Process p in processes)
            {
                try
                {
                    using (ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT CommandLine FROM Win32_Process WHERE ProcessId = " + p.Id))
                    {
                        foreach (ManagementObject @object in searcher.Get())
                        {

                            var cmdLine = @object["CommandLine"].ToString();
                            var appPoolName = SplitCommandLine(cmdLine).SkipWhile(x => x != "-ap").Skip(1).FirstOrDefault();
                            appPools.Add(p.Id, appPoolName);
                            break;
                        }
                    }
                }
                catch (Win32Exception ex)
                {
                    if ((uint)ex.ErrorCode != 0x80004005)
                    {
                        throw;
                    }
                }
            }
            return appPools;
        }

        public static IEnumerable<string> SplitCommandLine(string commandLine)
        {
            bool inQuotes = false;

            return commandLine.Split(c =>
            {
                if (c == '\"')
                    inQuotes = !inQuotes;

                return !inQuotes && c == ' ';
            })
                              .Select(arg => arg.Trim().TrimMatchingQuotes('\"'))
                              .Where(arg => !string.IsNullOrEmpty(arg));
        }
    }

    public static class Extension
    {
        public static IEnumerable<string> Split(this string str,
                                            Func<char, bool> controller)
        {
            int nextPiece = 0;

            for (int c = 0; c < str.Length; c++)
            {
                if (controller(str[c]))
                {
                    yield return str.Substring(nextPiece, c - nextPiece);
                    nextPiece = c + 1;
                }
            }

            yield return str.Substring(nextPiece);
        }

        public static string TrimMatchingQuotes(this string input, char quote)
        {
            if ((input.Length >= 2) &&
                (input[0] == quote) && (input[input.Length - 1] == quote))
                return input.Substring(1, input.Length - 2);

            return input;
        }

    }

}

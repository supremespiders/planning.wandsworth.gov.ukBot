using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace planning.wandsworth.gov.ukBot.Models
{
    public static class Utility
    {
        public static string ConnectionString = "Data Source=system.db;Version=3;";
        public static string SimpleDateFormat = "dd/MM/yyyy HH:mm:ss";

        public static EventHandler<string> OnDisplay;
        public static EventHandler<string> OnError;

        public static void Display(string s)
        {
            OnDisplay?.Invoke(null, s);
        }
        public static void Error(string s)
        {
            OnError?.Invoke(null, s);
        }
        public static async Task<List<T>> Work<T, T2>(this List<T2> items, int maxThreads, Func<T2, Task<T>> func)
        {
            var tasks = new List<Task<T>>();
            var results = new List<T>();
            for (var i = 0; i < items.Count; i++)
            {
                var item = items[i];
                Display($"Working on {i + 1} / {items.Count}");
                tasks.Add(func(item));
                if (tasks.Count == maxThreads)
                {
                    try
                    {
                        var t = await Task.WhenAny(tasks);
                        results.Add(t.GetAwaiter().GetResult());
                        tasks.Remove(t);
                    }
                    catch (Exception e)
                    {
                        Error(e.ToString());
                        var t = tasks.FirstOrDefault(x => x.IsFaulted);
                        tasks.Remove(t);
                    }
                }
            }

            while (tasks.Count!=0)
            {
                try
                {
                    var t = await Task.WhenAny(tasks);
                    results.Add(t.GetAwaiter().GetResult());
                    tasks.Remove(t);
                }
                catch (Exception e)
                {
                    Error(e.ToString());
                    var t = tasks.FirstOrDefault(x => x.IsFaulted);
                    tasks.Remove(t);
                }
            }

            Display($"completed {items.Count}");
            return results;
        }
        public static async Task<List<T>> Work<T, T2>(this List<T2> items, int maxThreads, Func<T2, Task<List<T>>> func)
        {
            var tasks = new List<Task<List<T>>>();
            var results = new List<T>();
            for (var i = 0; i < items.Count; i++)
            {
                var item = items[i];
                Display($"Working on {i + 1} / {items.Count}");
                tasks.Add(func(item));
                if (tasks.Count == maxThreads)
                {
                    try
                    {
                        var t = await Task.WhenAny(tasks);
                        results.AddRange(t.GetAwaiter().GetResult());
                        tasks.Remove(t);
                    }
                    catch (Exception e)
                    {
                        Error(e.ToString());
                        var t = tasks.FirstOrDefault(x => x.IsFaulted);
                        tasks.Remove(t);
                    }
                }
            }

            while (tasks.Count!=0)
            {
                try
                {
                    var t = await Task.WhenAny(tasks);
                    results.AddRange(t.GetAwaiter().GetResult());
                    tasks.Remove(t);
                }
                catch (Exception e)
                {
                    Error(e.ToString());
                    var t = tasks.FirstOrDefault(x => x.IsFaulted);
                    tasks.Remove(t);
                }
            }

            Display($"completed {items.Count}");
            return results;
        }
        public static void SaveCookies(CookieContainer cookieContainer, string url)
        {
            try
            {
                var cookies = new List<Cookie>();
                foreach (Cookie cookie in cookieContainer.GetCookies(new Uri(url)))
                    cookies.Add(new Cookie { Name = cookie.Name, Value = cookie.Value });
                File.WriteAllText("ses", JsonConvert.SerializeObject(cookies));
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }

        public static string BetweenStrings(string text, string start, string end)
        {
            var p1 = text.IndexOf(start, StringComparison.Ordinal) + start.Length;
            var p2 = text.IndexOf(end, p1, StringComparison.Ordinal);
            if (end == "") return (text.Substring(p1));
            else return text.Substring(p1, p2 - p1);
        }

        public static CookieContainer LoadCookies(string url)
        {
            var cookieContainer = new CookieContainer();
            try
            {
                var myCookies = JsonConvert.DeserializeObject<List<Cookie>>(File.ReadAllText("ses"));
                foreach (var myCookie in myCookies)
                    cookieContainer.Add(new Uri(url), new Cookie(myCookie.Name, myCookie.Value));
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            return cookieContainer;
        }

    }
}

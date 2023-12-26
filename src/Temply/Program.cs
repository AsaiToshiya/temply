/**
 * Copyright (c) 2020 Asai Toshiya.
 *
 * License: The MIT License (https://opensource.org/licenses/MIT)
 */

using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Temply
{
    static class Program
    {
        private const string ApplicationName = "Temply";

        /// <summary>
        /// アプリケーションのメイン エントリ ポイントです。
        /// </summary>
        [STAThread]
        static void Main()
        {
            using (var mutex = new Mutex(false, ApplicationName))
            {
                if (!mutex.WaitOne(0, false))
                {
                    return;
                }

                var database = Path.Combine(Path.GetDirectoryName(Application.ExecutablePath), "temply.txt");
                if (!File.Exists(database))
                {
                    File.Create(database).Close();
                }

                var texts = ReadTexts(database);

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                using (var notifyIcon = CreateNotifyIcon(database, texts))
                using (var watcher = CreateFileSystemWatcher(database, texts, notifyIcon))
                {
                    Application.Run();
                }
            }
        }

        private static List<string> ReadTexts(string database)
        {
            var exceptions = new List<Exception>();
            for (int attempted = 0; attempted < 5; attempted++)
            {
                if (attempted < 0)
                {
                    Thread.Sleep(50);
                }
                try
                {
                    return File.ReadAllLines(database)
                        .ToList()
                        .SortChain();
                }
                catch (Exception ex)
                {
                    exceptions.Add(ex);
                }
            }
            throw new AggregateException(exceptions);
        }

        private static NotifyIcon CreateNotifyIcon(string database, List<string> texts)
        {
            var notifyIcon = new NotifyIcon
            {
                Icon = SystemIcons.Application,
                Text = ApplicationName,
                ContextMenuStrip = CreateContextMenu(database, texts)
            };
            notifyIcon.Visible = true;
            return notifyIcon;
        }

        private static FileSystemWatcher CreateFileSystemWatcher(string database, List<string> texts, NotifyIcon notifyIcon)
        {
            var watcher = new FileSystemWatcher
            {
                Path = Path.GetDirectoryName(Application.ExecutablePath),
                Filter = "temply.txt",
                NotifyFilter = NotifyFilters.LastWrite
            };

            watcher.Changed += (sender, e) =>
            {
                watcher.EnableRaisingEvents = false;

                texts.ClearChain()
                    .AddRangeChain(ReadTexts(database))
                    .SortChain();
                notifyIcon.ContextMenuStrip = CreateContextMenu(database, texts);

                watcher.EnableRaisingEvents = true;
            };

            watcher.EnableRaisingEvents = true;
            return watcher;
        }

        private static ContextMenuStrip CreateContextMenu(string database, List<string> texts)
        {
            var templateMenuItems = texts.Count <= 0
                     ? new ToolStripMenuItem[] { CreateNothingMenuItem() }
                     : texts.Select(t => CreateTemplateMenuItem(t));
            var addMenuItem = CreateAddMenuItem(database, texts);
            var deleteMenuItem = CreateDeleteMenuItem(database, texts);

            var contextMenu = new ContextMenuStrip();
            contextMenu.Items.AddRange(new List<ToolStripMenuItem>()
                .AddChain(CreatetApplicationNameMenuItem())
                .AddChain(CreateMenuItemSeparator())
                .AddRangeChain(templateMenuItems)
                .AddChain(CreateMenuItemSeparator())
                .AddChain(addMenuItem)
                .AddChain(deleteMenuItem)
                .AddChain(CreateMenuItemSeparator())
                .AddChain(CreateSettingMenuItem())
                .AddChain(CreateMenuItemSeparator())
                .AddChain(CreateExitMenuItem())
                .ToArray()
            );
            contextMenu.Opening += (sender, e) =>
            {
                addMenuItem.Enabled = Clipboard.ContainsText();
                deleteMenuItem.Enabled = 0 < texts.Count;
            };

            return contextMenu;
        }

        private static ToolStripMenuItem CreatetApplicationNameMenuItem()
        {
            var index = Application.ProductVersion.LastIndexOf('.');
            var version = Application.ProductVersion.Substring(0, index);
            return new ToolStripMenuItem(string.Format("{0} {1}", ApplicationName, version));
        }

        private static ToolStripMenuItem CreateMenuItemSeparator()
        {
            return new ToolStripMenuItem("-");
        }

        private static ToolStripMenuItem CreateNothingMenuItem()
        {
            return new ToolStripMenuItem("(何もありません)") { Enabled = false };
        }

        private static ToolStripMenuItem CreateTemplateMenuItem(string text)
        {
            var caption = Ellipsis(text, 60);
            return new ToolStripMenuItem(caption, null, (sender, e) => Clipboard.SetText(Unescape(text)));
        }

        private static ToolStripMenuItem CreateAddMenuItem(string database, List<string> texts)
        {
            return new ToolStripMenuItem("クリップボードから追加(&A)", null, (sender, e) =>
            {
                var text = Clipboard.GetText();
                var escapedText = Escape(text);
                if (!texts.Contains(escapedText))
                {
                    AddText(database, texts, escapedText);
                }
            });
        }

        private static ToolStripMenuItem CreateDeleteMenuItem(string database, List<string> texts)
        {
            var subItems = texts.Select(t => CreateDeleteTemplateMenuItem(database, texts, t));
            return new ToolStripMenuItem("削除(&D)", null, subItems.ToArray());
        }

        private static ToolStripMenuItem CreateSettingMenuItem()
        {
            var subItems = new ToolStripMenuItem[] { CreateAutorunMenuItem() };
            return new ToolStripMenuItem("設定(&S)", null, subItems);
        }

        private static ToolStripMenuItem CreateExitMenuItem()
        {
            return new ToolStripMenuItem("終了(&X)", null, (sender, e) => Application.Exit());
        }

        private static ToolStripMenuItem CreateDeleteTemplateMenuItem(string database, List<string> texts, string text)
        {
            var caption = Ellipsis(text, 60);
            return new ToolStripMenuItem(caption, null, (sender, e) => DeleteText(database, texts, text));
        }

        private static ToolStripMenuItem CreateAutorunMenuItem()
        {
            var autorunMenuItem = new ToolStripMenuItem("自動起動(&A)", null, (sender, e) =>
            {
                var mi = sender as ToolStripMenuItem;
                var isAutorun = IsAutorun();
                var registryKey = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);
                if (isAutorun)
                {
                    registryKey.DeleteValue(ApplicationName, false);
                }
                else
                {
                    registryKey.SetValue(ApplicationName, Application.ExecutablePath);
                }
                mi.Checked = !isAutorun;
            });
            autorunMenuItem.Checked = IsAutorun();
            return autorunMenuItem;
        }

        private static void AddText(string database, List<string> texts, string text)
        {
            texts.Add(text);
            File.WriteAllLines(database, texts.ToArray());
        }

        private static void DeleteText(string database, List<string> texts, string text)
        {
            texts.Remove(text);
            File.WriteAllLines(database, texts.ToArray());
        }

        private static bool IsAutorun()
        {
            var registryKey = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run");
            return registryKey.GetValue(ApplicationName) != null;
        }

        private static string Escape(string s)
        {
            return s.Replace("\\", "\\\\")
                .Replace("\n", "\\n")
                .Replace("\r", "\\r")
                .Replace("\t", "\\t");
        }

        private static string Unescape(string s)
        {
            return s.RegexReplace(@"(?<!\\)\\n", "\n")
                .RegexReplace(@"(?<!\\)\\r", "\r")
                .RegexReplace(@"(?<!\\)\\t", "\t")
                .Replace("\\\\", "\\");
        }

        private static string Ellipsis(string value, int maxChars)
        {
            return value.Length <= maxChars ? value : value.Substring(0, maxChars) + " ...";
        }

        private static string RegexReplace(this string value, string pattern, string replacement)
        {
            return Regex.Replace(value, pattern, replacement);
        }

        private static List<T> AddChain<T>(this List<T> list, T item)
        {
            list.Add(item);
            return list;
        }

        private static List<T> AddRangeChain<T>(this List<T> list, IEnumerable<T> collection)
        {
            list.AddRange(collection);
            return list;
        }

        private static List<T> ClearChain<T>(this List<T> list)
        {
            list.Clear();
            return list;
        }

        private static List<T> SortChain<T>(this List<T> list)
        {
            list.Sort();
            return list;
        }
    }
}

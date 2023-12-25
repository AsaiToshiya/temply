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
            var assembly = Assembly.GetExecutingAssembly();
            var guidAttribute = (GuidAttribute)assembly.GetCustomAttributes(typeof(GuidAttribute), true)[0];
            var applicationGuid = guidAttribute.Value;

            using (var mutex = new Mutex(false, applicationGuid))
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
            // TODO ContextMenu is no longer supported. Use ContextMenuStrip instead. For more details see https://docs.microsoft.com/en-us/dotnet/core/compatibility/winforms#removed-controls
            var notifyIcon = new NotifyIcon
            {
                Icon = SystemIcons.Application,
                Text = ApplicationName,
                ContextMenu = CreateContextMenu(database, texts)
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
                notifyIcon.ContextMenu = CreateContextMenu(database, texts);

                watcher.EnableRaisingEvents = true;
            };

            watcher.EnableRaisingEvents = true;
            return watcher;
        }

        private static // TODO ContextMenu is no longer supported. Use ContextMenuStrip instead. For more details see https://docs.microsoft.com/en-us/dotnet/core/compatibility/winforms#removed-controls
ContextMenu CreateContextMenu(string database, List<string> texts)
        {
            // TODO MenuItem is no longer supported. Use ToolStripMenuItem instead. For more details see https://docs.microsoft.com/en-us/dotnet/core/compatibility/winforms#removed-controls
            var templateMenuItems = texts.Count <= 0
                     ? new MenuItem[] { CreateNothingMenuItem() }
                     : texts.Select(t => CreateTemplateMenuItem(t));
            var addMenuItem = CreateAddMenuItem(database, texts);
            var deleteMenuItem = CreateDeleteMenuItem(database, texts);

            // TODO ContextMenu is no longer supported. Use ContextMenuStrip instead. For more details see https://docs.microsoft.com/en-us/dotnet/core/compatibility/winforms#removed-controls
            var contextMenu = new ContextMenu(new List<MenuItem>()
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
            contextMenu.Popup += (sender, e) =>
            {
                addMenuItem.Enabled = Clipboard.ContainsText();
                deleteMenuItem.Enabled = 0 < texts.Count;
            };

            return contextMenu;
        }

        private static // TODO MenuItem is no longer supported. Use ToolStripMenuItem instead. For more details see https://docs.microsoft.com/en-us/dotnet/core/compatibility/winforms#removed-controls
MenuItem CreatetApplicationNameMenuItem()
        {
            var index = Application.ProductVersion.LastIndexOf('.');
            var version = Application.ProductVersion.Substring(0, index);
            // TODO MenuItem is no longer supported. Use ToolStripMenuItem instead. For more details see https://docs.microsoft.com/en-us/dotnet/core/compatibility/winforms#removed-controls
            return new MenuItem(string.Format("{0} {1}", ApplicationName, version));
        }

        private static // TODO MenuItem is no longer supported. Use ToolStripMenuItem instead. For more details see https://docs.microsoft.com/en-us/dotnet/core/compatibility/winforms#removed-controls
MenuItem CreateMenuItemSeparator()
        {
            // TODO MenuItem is no longer supported. Use ToolStripMenuItem instead. For more details see https://docs.microsoft.com/en-us/dotnet/core/compatibility/winforms#removed-controls
            return new MenuItem("-");
        }

        private static // TODO MenuItem is no longer supported. Use ToolStripMenuItem instead. For more details see https://docs.microsoft.com/en-us/dotnet/core/compatibility/winforms#removed-controls
MenuItem CreateNothingMenuItem()
        {
            // TODO MenuItem is no longer supported. Use ToolStripMenuItem instead. For more details see https://docs.microsoft.com/en-us/dotnet/core/compatibility/winforms#removed-controls
            return new MenuItem("(何もありません)") { Enabled = false };
        }

        private static // TODO MenuItem is no longer supported. Use ToolStripMenuItem instead. For more details see https://docs.microsoft.com/en-us/dotnet/core/compatibility/winforms#removed-controls
MenuItem CreateTemplateMenuItem(string text)
        {
            var caption = Ellipsis(text, 60);
            // TODO MenuItem is no longer supported. Use ToolStripMenuItem instead. For more details see https://docs.microsoft.com/en-us/dotnet/core/compatibility/winforms#removed-controls
            return new MenuItem(caption, (sender, e) => Clipboard.SetText(Unescape(text)));
        }

        private static // TODO MenuItem is no longer supported. Use ToolStripMenuItem instead. For more details see https://docs.microsoft.com/en-us/dotnet/core/compatibility/winforms#removed-controls
MenuItem CreateAddMenuItem(string database, List<string> texts)
        {
            // TODO MenuItem is no longer supported. Use ToolStripMenuItem instead. For more details see https://docs.microsoft.com/en-us/dotnet/core/compatibility/winforms#removed-controls
            return new MenuItem("クリップボードから追加(&A)", (sender, e) =>
            {
                var text = Clipboard.GetText();
                var escapedText = Escape(text);
                if (!texts.Contains(escapedText))
                {
                    AddText(database, texts, escapedText);
                }
            });
        }

        private static // TODO MenuItem is no longer supported. Use ToolStripMenuItem instead. For more details see https://docs.microsoft.com/en-us/dotnet/core/compatibility/winforms#removed-controls
MenuItem CreateDeleteMenuItem(string database, List<string> texts)
        {
            var subItems = texts.Select(t => CreateDeleteTemplateMenuItem(database, texts, t));
            // TODO MenuItem is no longer supported. Use ToolStripMenuItem instead. For more details see https://docs.microsoft.com/en-us/dotnet/core/compatibility/winforms#removed-controls
            return new MenuItem("削除(&D)", subItems.ToArray());
        }

        private static // TODO MenuItem is no longer supported. Use ToolStripMenuItem instead. For more details see https://docs.microsoft.com/en-us/dotnet/core/compatibility/winforms#removed-controls
MenuItem CreateSettingMenuItem()
        {
            // TODO MenuItem is no longer supported. Use ToolStripMenuItem instead. For more details see https://docs.microsoft.com/en-us/dotnet/core/compatibility/winforms#removed-controls
            var subItems = new MenuItem[] { CreateAutorunMenuItem() };
            // TODO MenuItem is no longer supported. Use ToolStripMenuItem instead. For more details see https://docs.microsoft.com/en-us/dotnet/core/compatibility/winforms#removed-controls
            return new MenuItem("設定(&S)", subItems);
        }

        private static // TODO MenuItem is no longer supported. Use ToolStripMenuItem instead. For more details see https://docs.microsoft.com/en-us/dotnet/core/compatibility/winforms#removed-controls
MenuItem CreateExitMenuItem()
        {
            // TODO MenuItem is no longer supported. Use ToolStripMenuItem instead. For more details see https://docs.microsoft.com/en-us/dotnet/core/compatibility/winforms#removed-controls
            return new MenuItem("終了(&X)", (sender, e) => Application.Exit());
        }

        private static // TODO MenuItem is no longer supported. Use ToolStripMenuItem instead. For more details see https://docs.microsoft.com/en-us/dotnet/core/compatibility/winforms#removed-controls
MenuItem CreateDeleteTemplateMenuItem(string database, List<string> texts, string text)
        {
            var caption = Ellipsis(text, 60);
            // TODO MenuItem is no longer supported. Use ToolStripMenuItem instead. For more details see https://docs.microsoft.com/en-us/dotnet/core/compatibility/winforms#removed-controls
            return new MenuItem(caption, (sender, e) => DeleteText(database, texts, text));
        }

        private static // TODO MenuItem is no longer supported. Use ToolStripMenuItem instead. For more details see https://docs.microsoft.com/en-us/dotnet/core/compatibility/winforms#removed-controls
MenuItem CreateAutorunMenuItem()
        {
            // TODO MenuItem is no longer supported. Use ToolStripMenuItem instead. For more details see https://docs.microsoft.com/en-us/dotnet/core/compatibility/winforms#removed-controls
            var autorunMenuItem = new MenuItem("自動起動(&A)", (sender, e) =>
            {
                var mi = sender as MenuItem;
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

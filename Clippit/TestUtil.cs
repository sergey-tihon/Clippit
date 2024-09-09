// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Diagnostics;
using System.IO;

namespace Clippit
{
    public class TestUtil
    {
        private static readonly Lazy<DirectoryInfo> s_tempDir =
            new(() =>
            {
                var now = DateTime.Now;
                var tempDirName =
                    $"Test-{now.Year - 2000:00}-{now.Month:00}-{now.Day:00}-{now.Hour:00}{now.Minute:00}{now.Second:00}";
                var dir = new DirectoryInfo(Path.Combine(".", tempDirName));
                dir.Create();
                return dir;
            });

        private static DirectoryInfo TempDir => s_tempDir.Value;

        public static void NotePad(string str)
        {
            var guidName = Guid.NewGuid().ToString().Replace("-", "") + ".txt";
            var fi = new FileInfo(Path.Combine(TempDir.FullName, guidName));
            File.WriteAllText(fi.FullName, str);
            var notepadExe = new FileInfo(@"C:\Program Files (x86)\Notepad++\notepad++.exe");
            if (!notepadExe.Exists)
                notepadExe = new FileInfo(@"C:\Program Files\Notepad++\notepad++.exe");
            if (!notepadExe.Exists)
                notepadExe = new FileInfo(@"C:\Windows\System32\notepad.exe");
            ExecutableRunner.RunExecutable(notepadExe.FullName, fi.FullName, TempDir.FullName);
        }

        public static void Explorer(DirectoryInfo di)
        {
            Process.Start(di.FullName);
        }
    }
}

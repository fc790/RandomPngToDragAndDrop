// RandomPngDropViaWM_DROPFILES_BatchFemale.cs
// .NET 3.5 + BepInEx 5
// F8: pick png(s) and POST WM_DROPFILES to UnityWndClass so DragAndDrop.Koikatu.dll (WH_GETMESSAGE hook) handles it.
// Extra: optional batch replace all female characters by controlling GuideObjectManager.selectObjectKey.

using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Reflection;

using BepInEx;
using BepInEx.Configuration;
using UnityEngine;

using Manager;
using Studio;

namespace FC790s
{
    [BepInPlugin(GUID, NAME, VER)]
    public class RandomPngDropViaWM_BatchFemale : BaseUnityPlugin
    {
        public const string GUID = "fc790s.dragdrop_random_png_wmdrop";
        public const string NAME = "Random PNG Drop via WM_DROPFILES (F8 + Batch Female)";
        public const string VER  = "1.2.0";

        private ConfigEntry<bool> _enable;
        private ConfigEntry<string> _pngDir;

        private ConfigEntry<bool> _batchReplaceAllFemale;
        private ConfigEntry<bool> _preselectAllFemale;
        private ConfigEntry<bool> _perCharacterRandom;

        private System.Random _rng = new System.Random();

        private bool _busy = false;

        private void Awake()
        {
            _enable = Config.Bind<bool>("General", "Enable", true, "Master switch.");
            _pngDir = Config.Bind<string>("General", "PngDir", "", "Directory to scan for *.png (absolute path recommended).");

            _batchReplaceAllFemale = Config.Bind<bool>("Batch", "BatchReplaceFemale", false, "If true, F8 will iterate all female characters and replace each one by dropping a random PNG while selecting that character.");
            _preselectAllFemale    = Config.Bind<bool>("Batch", "PreselectAllFemale", false, "If true, F8 will first set selection to all female characters (selection highlight / consistency).");
            _perCharacterRandom    = Config.Bind<bool>("Batch", "PerCharacterRandom", true, "If true, each female gets its own random PNG. If false, all females use the same picked PNG.");

            Logger.LogInfo("Loaded. F8 => drop random png via PostMessage(WM_DROPFILES). Batch options available in config.");
        }

        private void Update()
        {
            if (!_enable.Value) return;
            if (_busy) return;

            if (Input.GetKeyDown(KeyCode.F8))
            {
                StartCoroutine(Co_F8Action());
            }
        }
        
        private TreeNodeObject TryGetTreeNodeObjectFromOCI(object ociOrCtrlInfo)
        {
            if (ociOrCtrlInfo == null) return null;
        
            Type t = ociOrCtrlInfo.GetType();
        
            // 1) property treeNodeObject
            try
            {
                PropertyInfo p = t.GetProperty("treeNodeObject", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                if (p != null)
                {
                    object v = p.GetValue(ociOrCtrlInfo, null);
                    TreeNodeObject tn = v as TreeNodeObject;
                    if (tn != null) return tn;
                }
            }
            catch { }
        
            // 2) field treeNodeObject
            try
            {
                FieldInfo f = t.GetField("treeNodeObject", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                if (f != null)
                {
                    object v = f.GetValue(ociOrCtrlInfo);
                    TreeNodeObject tn = v as TreeNodeObject;
                    if (tn != null) return tn;
                }
            }
            catch { }
        
            return null;
        }

        private TreeNodeObject[] CaptureSelectedNodes()
        {
            try
            {
                Studio.Studio st = Singleton<Studio.Studio>.Instance;
                if (!st) return null;
                TreeNodeCtrl tnc = st.treeNodeCtrl;
                if (tnc == null) return null;
        
                // 这是你贴的属性：get => hashSelectNode.ToArray()
                return tnc.selectNodes;
            }
            catch
            {
                return null;
            }
        }
        
        private TreeNodeCtrl GetTreeNodeCtrl()
        {
            try
            {
                Studio.Studio st = Singleton<Studio.Studio>.Instance;
                if (!st) return null;
                return st.treeNodeCtrl;
            }
            catch { return null; }
        }
        
        // cache
        private MethodInfo _miSelectNodesSetter = null;
        private MethodInfo _miGetSelectNodes = null;
        private MethodInfo _miSelectArray = null;   // any Select*(TreeNodeObject[] ...)
        private MethodInfo _miSelectOne = null;     // any Select*(TreeNodeObject ...)
        
        private bool _dumpedTnc = false;
        
        private void ResolveTreeNodeCtrlMethods(TreeNodeCtrl tnc)
        {
            if (tnc == null) return;
            if (_miSelectArray != null || _miSelectOne != null || _miSelectNodesSetter != null || _miGetSelectNodes != null) return;
        
            Type t = tnc.GetType();
        
            // 0) property selectNodes (getter/setter) if exists
            try
            {
                PropertyInfo p = t.GetProperty("selectNodes", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                if (p != null)
                {
                    MethodInfo g = p.GetGetMethod(true);
                    if (g != null) _miGetSelectNodes = g;
        
                    MethodInfo s = p.GetSetMethod(true);
                    if (s != null) _miSelectNodesSetter = s; // KKS likely
                }
            }
            catch { }
        
            // 1) scan methods for selection candidates by signature
            MethodInfo[] ms = null;
            try
            {
                ms = t.GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
            }
            catch { ms = null; }
        
            if (ms != null)
            {
                for (int i = 0; i < ms.Length; i++)
                {
                    MethodInfo m = ms[i];
                    if (m == null) continue;
        
                    string n = m.Name;
                    if (n == null) n = "";
        
                    // prefer methods that look like selection APIs
                    if (n.IndexOf("Select", StringComparison.OrdinalIgnoreCase) < 0 &&
                        n.IndexOf("select", StringComparison.OrdinalIgnoreCase) < 0)
                    {
                        // still allow, but lower priority; skip for now
                        continue;
                    }
        
                    ParameterInfo[] ps = null;
                    try { ps = m.GetParameters(); } catch { ps = null; }
                    if (ps == null || ps.Length == 0) continue;
        
                    Type p0 = ps[0].ParameterType;
                    if (p0 == typeof(TreeNodeObject[]))
                    {
                        // pick first match as array selector
                        if (_miSelectArray == null) _miSelectArray = m;
                    }
                    else if (p0 == typeof(TreeNodeObject))
                    {
                        if (_miSelectOne == null) _miSelectOne = m;
                    }
        
                    if (_miSelectArray != null && _miSelectOne != null) break;
                }
            }
        
            // 2) dump once if we still can't find anything usable
            if (!_dumpedTnc && _miSelectNodesSetter == null && _miSelectArray == null && _miSelectOne == null)
            {
                _dumpedTnc = true;
                try
                {
                    Logger.LogWarning("TreeNodeCtrl selection methods not found by name. Dumping candidates:");
                    if (ms != null)
                    {
                        for (int i = 0; i < ms.Length; i++)
                        {
                            MethodInfo m = ms[i];
                            if (m == null) continue;
                            string n = m.Name;
                            if (n == null) n = "";
                            if (n.IndexOf("Select", StringComparison.OrdinalIgnoreCase) < 0 &&
                                n.IndexOf("select", StringComparison.OrdinalIgnoreCase) < 0)
                                continue;
        
                            ParameterInfo[] ps = null;
                            try { ps = m.GetParameters(); } catch { ps = null; }
                            if (ps == null) continue;
        
                            // log signature
                            StringBuilder sb = new StringBuilder();
                            sb.Append(n);
                            sb.Append("(");
                            for (int k = 0; k < ps.Length; k++)
                            {
                                if (k > 0) sb.Append(", ");
                                sb.Append(ps[k].ParameterType.FullName);
                            }
                            sb.Append(")");
                            Logger.LogWarning("  " + sb.ToString());
                        }
                    }
                }
                catch { }
            }
        }
        
        // build default args for extra params (bool/int/enum)
        private object[] BuildArgsForMethod(MethodInfo mi, TreeNodeObject[] nodes, TreeNodeObject one)
        {
            ParameterInfo[] ps = mi.GetParameters();
            object[] args = new object[ps.Length];
        
            // first arg
            if (ps[0].ParameterType == typeof(TreeNodeObject[]))
                args[0] = nodes;
            else if (ps[0].ParameterType == typeof(TreeNodeObject))
                args[0] = one;
            else
                args[0] = null;
        
            // rest: fill defaults
            for (int i = 1; i < ps.Length; i++)
            {
                Type pt = ps[i].ParameterType;
        
                if (pt == typeof(bool)) args[i] = false;
                else if (pt == typeof(int)) args[i] = 0;
                else if (pt.IsEnum) args[i] = 0;
                else args[i] = null;
            }
        
            return args;
        }
        
        private void SetSelectedNodes_Compat(TreeNodeObject[] nodes)
        {
            try
            {
                TreeNodeCtrl tnc = GetTreeNodeCtrl();
                if (tnc == null) return;
        
                ResolveTreeNodeCtrlMethods(tnc);
        
                if (nodes == null) nodes = new TreeNodeObject[0];
        
                // A) if setter exists (KKS), use it
                if (_miSelectNodesSetter != null)
                {
                    _miSelectNodesSetter.Invoke(tnc, new object[] { nodes });
                    return;
                }
        
                // B) prefer array selection method
                if (_miSelectArray != null)
                {
                    object[] args = BuildArgsForMethod(_miSelectArray, nodes, null);
                    _miSelectArray.Invoke(tnc, args);
                    return;
                }
        
                // C) fallback: select one-by-one using single selector
                if (_miSelectOne != null)
                {
                    for (int i = 0; i < nodes.Length; i++)
                    {
                        TreeNodeObject one = nodes[i];
                        object[] args = BuildArgsForMethod(_miSelectOne, null, one);
                        _miSelectOne.Invoke(tnc, args);
                    }
                    return;
                }
        
                Logger.LogWarning("TreeNodeCtrl: no selectable method found; cannot set selection.");
            }
            catch (Exception ex)
            {
                Logger.LogWarning("SetSelectedNodes_Compat exception: " + ex.GetType().Name + " " + ex.Message);
            }
        }

        private System.Collections.IEnumerator Co_F8Action()
        {
            _busy = true;
            try
            {
                string dir = (_pngDir.Value == null) ? "" : _pngDir.Value.Trim();
                if (dir.Length == 0)
                {
                    Logger.LogWarning("PngDir is empty.");
                    yield break;
                }
                if (!Directory.Exists(dir))
                {
                    Logger.LogWarning("PngDir not exists: " + dir);
                    yield break;
                }

                string[] pngs = null;
                try
                {
                    pngs = Directory.GetFiles(dir, "*.png", SearchOption.AllDirectories);
                }
                catch (Exception ex)
                {
                    Logger.LogWarning("GetFiles failed: " + ex.GetType().Name + " " + ex.Message);
                    yield break;
                }

                if (pngs == null || pngs.Length == 0)
                {
                    Logger.LogWarning("No *.png found under: " + dir);
                    yield break;
                }

                IntPtr hUnity = FindUnityWndClassOnCurrentThread();
                if (hUnity == IntPtr.Zero) hUnity = FindAnyUnityWndClassTopLevel();
                if (hUnity == IntPtr.Zero)
                {
                    Logger.LogWarning("UnityWndClass not found.");
                    yield break;
                }

                // --- flags snapshot ---
                bool B = _batchReplaceAllFemale.Value;
                bool P = _preselectAllFemale.Value;
                bool R = _perCharacterRandom.Value;
                
                // Collect females only if needed by P or R
                List<TreeNodeObject> femaleNodes = null;
                
                // 规则：P=true 且 B=false 且 R=false => 只全选 female
                if (P && !B && !R)
                {
                    femaleNodes = GetAllFemaleCharacterNodes();
                    if (femaleNodes.Count == 0)
                    {
                        Logger.LogWarning("[F8] No female characters found in scene.");
                        yield break;
                    }
                
                    SetSelectedNodes_Compat(femaleNodes.ToArray());
                    yield return null;
                    Logger.LogInfo("[F8] PreselectAllFemale only. females=" + femaleNodes.Count);
                    yield break;
                }
                
                // 规则：R=true 优先级最高 => 逐个 female 单选 + 每个随机 drop
                if (R)
                {
                    femaleNodes = GetAllFemaleCharacterNodes();
                    if (P)
                    {
                        SetSelectedNodes_Compat(femaleNodes.ToArray());
                        yield return null;
                    }

                    TreeNodeObject[] prevSel = CaptureSelectedNodes();
                    if (prevSel.Length == 0)
                    {
                        string path = pngs[_rng.Next(0, pngs.Length)];
                        Logger.LogInfo("[F8] no selected ,but drop one (no preselect). png=" + path);
                        PostOneDrop(hUnity, path);
                        yield break;
                    }
                
                    Logger.LogInfo("[F8] PerCharacterRandom batch. females=" + femaleNodes.Count);
                
                    for (int i = 0; i < prevSel.Length; i++)
                    {
                        TreeNodeObject node = prevSel[i];
                
                        TreeNodeObject[] one = new TreeNodeObject[1];
                        one[0] = node;
                        SetSelectedNodes_Compat(one);
                        yield return null;
                
                        string usePath = pngs[_rng.Next(0, pngs.Length)];
                        Logger.LogInfo("[F8] PerChar drop idx=" + i + " png=" + usePath);
                        PostOneDrop(hUnity, usePath);
                
                        yield return null;
                    }
                
                    SetSelectedNodes_Compat(prevSel);
                    Logger.LogInfo("[F8] PerCharacterRandom done. females=" + femaleNodes.Count);
                    yield break;
                }
                
                // 规则：B=true => 随机 drop 一次（不依赖选中）；若 P=true 则先全选 female
                if (B)
                {
                    if (P)
                    {
                        femaleNodes = GetAllFemaleCharacterNodes();
                        if (femaleNodes.Count == 0)
                        {
                            Logger.LogWarning("[F8] No female characters found in scene.");
                            yield break;
                        }
                
                        TreeNodeObject[] prevSel = CaptureSelectedNodes();
                        SetSelectedNodes_Compat(femaleNodes.ToArray());
                        yield return null;
                
                        string path = pngs[_rng.Next(0, pngs.Length)];
                        Logger.LogInfo("[F8] Batch single drop after preselect. png=" + path);
                        PostOneDrop(hUnity, path);
                
                        // 你如果希望 batch drop 后保持全选，就注释掉下一行
                        SetSelectedNodes_Compat(prevSel);
                
                        yield break;
                    }
                    else
                    {
                        string path = pngs[_rng.Next(0, pngs.Length)];
                        Logger.LogInfo("[F8] Batch single drop (no preselect). png=" + path);
                        PostOneDrop(hUnity, path);
                        yield break;
                    }
                }
                
                // 其它：普通单次随机 drop
                {
                    string path = pngs[_rng.Next(0, pngs.Length)];
                    Logger.LogInfo("[F8] Single drop. png=" + path);
                    PostOneDrop(hUnity, path);
                    yield break;
                }
            }
            finally
            {
                _busy = false;
            }
        }

        // ===== selection control (what DragAndDrop.StudioHandler uses) =====


        private List<TreeNodeObject> GetAllFemaleCharacterNodes()
        {
            List<TreeNodeObject> nodes = new List<TreeNodeObject>();
        
            try
            {
                Studio.Studio st = Singleton<Studio.Studio>.Instance;
                if (!st) return nodes;
        
                Dictionary<int, ObjectCtrlInfo> dic = st.dicObjectCtrl;
                if (dic == null) return nodes;
        
                foreach (KeyValuePair<int, ObjectCtrlInfo> kv in dic)
                {
                    ObjectCtrlInfo oci = kv.Value;
                    OCIChar ch = oci as OCIChar;
                    if (ch == null) continue;
        
                    byte sex = 255;
                    try { sex = ch.charInfo.fileParam.sex; } catch { sex = 255; }
                    if (sex != 1) continue;
        
                    TreeNodeObject tn = TryGetTreeNodeObjectFromOCI(ch);
                    if (tn == null) tn = TryGetTreeNodeObjectFromOCI(oci);
                    if (tn == null) continue;
        
                    nodes.Add(tn);
                }
            }
            catch (Exception ex)
            {
                Logger.LogWarning("GetAllFemaleCharacterNodes exception: " + ex.GetType().Name + " " + ex.Message);
            }
        
            return nodes;
        }

        // ===== drop via WM_DROPFILES (queue message so WH_GETMESSAGE hook sees it) =====

        private void PostOneDrop(IntPtr hUnity, string path)
        {
            IntPtr hDrop = BuildHDropForSingleFile(path, 10, 10);
            if (hDrop == IntPtr.Zero)
            {
                Logger.LogWarning("Build HDROP failed: " + path);
                return;
            }

            bool ok = PostMessage(hUnity, WM_DROPFILES, hDrop, IntPtr.Zero);
            if (!ok)
            {
                Logger.LogWarning("PostMessage WM_DROPFILES failed. hWnd=0x" + hUnity.ToString("X") + " path=" + path);
            }
        }

        // ================= Win32 =================

        private const int WM_DROPFILES = 0x0233;
        private const uint GMEM_MOVEABLE = 0x0002;
        private const uint GMEM_ZEROINIT = 0x0040;

        [StructLayout(LayoutKind.Sequential)]
        private struct DROPFILES
        {
            public int pFiles;
            public int pt_x;
            public int pt_y;
            public int fNC;
            public int fWide;
        }

        private delegate bool EnumThreadWndProc(IntPtr hWnd, IntPtr lParam);
        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumThreadWindows(uint dwThreadId, EnumThreadWndProc lpfn, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpfn, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool PostMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr GlobalAlloc(uint uFlags, UIntPtr dwBytes);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr GlobalLock(IntPtr hMem);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool GlobalUnlock(IntPtr hMem);

        private static bool ClassEquals(IntPtr hWnd, string cls)
        {
            StringBuilder sb = new StringBuilder(256);
            int n = 0;
            try { n = GetClassName(hWnd, sb, sb.Capacity); } catch { n = 0; }
            if (n <= 0) return false;
            return string.Equals(sb.ToString(), cls, StringComparison.Ordinal);
        }

        private static IntPtr FindUnityWndClassOnCurrentThread()
        {
            uint tid = GetCurrentThreadId();
            IntPtr found = IntPtr.Zero;

            EnumThreadWindows(tid, delegate(IntPtr hWnd, IntPtr lp)
            {
                if (found != IntPtr.Zero) return false;
                if (!IsWindowVisible(hWnd)) return true;
                if (ClassEquals(hWnd, "UnityWndClass"))
                {
                    found = hWnd;
                    return false;
                }
                return true;
            }, IntPtr.Zero);

            return found;
        }

        private static IntPtr FindAnyUnityWndClassTopLevel()
        {
            IntPtr found = IntPtr.Zero;

            EnumWindows(delegate(IntPtr hWnd, IntPtr lp)
            {
                if (found != IntPtr.Zero) return false;
                if (!IsWindowVisible(hWnd)) return true;
                if (ClassEquals(hWnd, "UnityWndClass"))
                {
                    found = hWnd;
                    return false;
                }
                return true;
            }, IntPtr.Zero);

            return found;
        }

        private static IntPtr BuildHDropForSingleFile(string filePath, int x, int y)
        {
            string list = filePath + "\0\0";
            byte[] bytes = Encoding.Unicode.GetBytes(list);

            int dfSize = Marshal.SizeOf(typeof(DROPFILES));
            int total = dfSize + bytes.Length;

            IntPtr hGlobal = GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, (UIntPtr)total);
            if (hGlobal == IntPtr.Zero) return IntPtr.Zero;

            IntPtr p = GlobalLock(hGlobal);
            if (p == IntPtr.Zero) return IntPtr.Zero;

            try
            {
                DROPFILES df = new DROPFILES();
                df.pFiles = dfSize;
                df.pt_x = x;
                df.pt_y = y;
                df.fNC = 0;
                df.fWide = 1;

                Marshal.StructureToPtr(df, p, false);

                IntPtr pFiles = new IntPtr(p.ToInt64() + dfSize);
                Marshal.Copy(bytes, 0, pFiles, bytes.Length);
            }
            finally
            {
                GlobalUnlock(hGlobal);
            }

            return hGlobal;
        }
    }
}

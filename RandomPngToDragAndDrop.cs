// RandomPngDropViaWM_DROPFILES_BatchFemale.cs
// .NET 3.5 + BepInEx 5
// F8: pick png(s) and POST WM_DROPFILES to UnityWndClass so DragAndDrop.Koikatu.dll (WH_GETMESSAGE hook) handles it.
// Extra: optional batch replace all female characters by controlling GuideObjectManager.selectObjectKey.

using System;
using System.IO;
using System.Text;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Reflection;

using BepInEx;
using BepInEx.Configuration;
using HarmonyLib;
using UnityEngine;

using Manager;
using Studio;

namespace FC790s
{
    [BepInPlugin(GUID, NAME, VER)]
    public class RandomPngDropViaWM_BatchFemale : BaseUnityPlugin
    {
        public const string GUID = "fc790s.dragdrop_random_png_wmdrop";
        public const string NAME = "Random PNG Drop via WM_DROPFILES (F4/Ctrl+F4 + AutoTriggers)";
        public const string VER  = "1.3.0";

        private ConfigEntry<bool> _enable;

        // Two directory paths:
        // - F4 uses PngDir_F4
        // - Ctrl+F4 uses PngDir_CtrlF4
        private ConfigEntry<string> _pngDir_F4;
        private ConfigEntry<string> _pngDir_CtrlF4;

        private ConfigEntry<bool> _batchReplaceAllFemale;
        private ConfigEntry<bool> _preselectAllFemale;
        private ConfigEntry<bool> _perCharacterRandom;

        // ===== Auto triggers (migrated from AffectedDynamicBonesEnabledPersistForKKS.cs) =====
        private ConfigEntry<bool> _autoOnCharacterReplace;
        private ConfigEntry<bool> _autoOnSceneLoad;
        private ConfigEntry<int>  _autoDelayFrames; // delay frames before auto action
        private ConfigEntry<bool> _patchTriggers;   // master switch for Harmony patches

        // Harmony PatchAny via reflection (avoid overload assumptions)
        private Harmony _harmony;
        private static MethodInfo _miHarmonyPatchAny;

        private static bool _isInternalTrigger = false; // 内部触发锁

        private System.Random _rng = new System.Random();

        private bool _busy = false;
        // reflect
        private static bool S_RefOk = false;
        private static Type T_MainWindow;

        private static RandomPngDropViaWM_BatchFemale _inst;
private void Awake()
        {
            _inst = this;

            _enable = Config.Bind<bool>("General", "Enable", true, "Master switch.");

            _pngDir_F4      = Config.Bind<string>("General", "PngDir_F4", "", "Directory to scan for *.png (F4). Absolute path recommended.");
            _pngDir_CtrlF4  = Config.Bind<string>("General", "PngDir_CtrlF4", "", "Directory to scan for *.png (Ctrl+F4). Absolute path recommended.");

            _batchReplaceAllFemale = Config.Bind<bool>("Batch", "BatchReplaceFemale", false, "If true, hotkey will iterate all female characters and replace each one by dropping a random PNG while selecting that character.");
            _preselectAllFemale    = Config.Bind<bool>("Batch", "PreselectAllFemale", false, "If true, hotkey will first set selection to all female characters (selection highlight / consistency).");
            _perCharacterRandom    = Config.Bind<bool>("Batch", "PerCharacterRandom", true, "If true, each female gets its own random PNG. If false, all females use the same picked PNG.");

            // Auto triggers
            _patchTriggers        = Config.Bind<bool>("AutoTrigger", "PatchTriggers", true, "If true, patch HSPE/Studio triggers (OnCharacterReplace / SceneInfo.Load).");
            _autoOnCharacterReplace = Config.Bind<bool>("AutoTrigger", "AutoOnCharacterReplace", false, "If true, after character replace, auto-run using F4 directory (PngDir_F4).");
            _autoOnSceneLoad        = Config.Bind<bool>("AutoTrigger", "AutoOnSceneLoad", false, "If true, after scene load, auto-run using F4 directory (PngDir_F4).");
            _autoDelayFrames        = Config.Bind<int>("AutoTrigger", "AutoDelayFrames", 30, "Delay frames before auto-run for triggers (0-600).");

            // init harmony patch helper
            if (_patchTriggers.Value)
            {
                TryInitHarmonyPatchAny();
                TryInitReflection();
                Patch_MainWindow_OnCharacterReplace();
                Patch_StudioScene_Load();
            }

            Logger.LogInfo("Loaded. F4 => drop random PNG from PngDir_F4. Ctrl+F4 => from PngDir_CtrlF4.");
        }


        private void Update()
        {
            if (!_enable.Value) return;
            if (_busy) return;

            // F4: use PngDir_F4
            // Ctrl+F4: use PngDir_CtrlF4
            if (Input.GetKeyDown(KeyCode.U))
            {
                bool ctrl = Input.GetKey(KeyCode.LeftAlt) || Input.GetKey(KeyCode.RightAlt);
                string dir = ctrl ? _pngDir_CtrlF4.Value : _pngDir_F4.Value;
                string tag = ctrl ? "CTRL+F4" : "F4";
                StartCoroutine(Co_RunFromDir(dir, tag));
            }
        }

        private static Type FindType(string fullName)
        {
            Type t = Type.GetType(fullName, false);
            if (t != null) return t;

            Assembly[] asms = AppDomain.CurrentDomain.GetAssemblies();
            for (int i = 0; i < asms.Length; i++)
            {
                Assembly a = asms[i];
                if (a == null) continue;
                try
                {
                    t = a.GetType(fullName, false);
                    if (t != null) return t;
                }
                catch { }
            }
            return null;
        }

        private void TryInitReflection()
        {
            if (S_RefOk) return;

            try
            {
                T_MainWindow = FindType("HSPE.MainWindow");

                if (T_MainWindow == null)
                {
                    Logger.LogWarning("[REF] missing members.");
                    return;
                }

                S_RefOk = true;
                Logger.LogInfo("[REF] OK");
            }
            catch (Exception e)
            {
                Logger.LogWarning("[REF] Exception: " + e.Message);
            }
        }

        private void Patch_MainWindow_OnCharacterReplace()
        {
            if (_miHarmonyPatchAny == null) return;

            try
            {
                MethodInfo mi = T_MainWindow.GetMethod("OnCharaLoad", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                if (mi == null)
                {
                    Logger.LogWarning("[PATCH] MainWindow." + "OnCharacterReplace" + " not found (skip).");
                    return;
                }

                MethodInfo post = typeof(RandomPngDropViaWM_BatchFemale).GetMethod("MW_OnCharacterReplace_Postfix", BindingFlags.Static | BindingFlags.NonPublic);
                if (post == null)
                {
                    Logger.LogWarning("[PATCH] Postfix not found: " + "MW_OnCharacterReplace_Postfix");
                    return;
                }

                PatchPostfix(mi, post, Priority.Last);
                Logger.LogInfo("[PATCH] Patched MainWindow." + "OnCharacterReplace" + " postfix.");
            }
            catch (Exception e)
            {
                Logger.LogWarning("[PATCH] Patch_MainWindow_OnCharacterReplace failed: " + e.Message);
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
        
        // ===== TreeNodeCtrl selection (KK authoritative) =====
        private MethodInfo _miAddSelectNode = null;
        private FieldInfo  _fiHashSelectNode = null;
        
        private void ResolveTreeNodeCtrlKKApis(TreeNodeCtrl tnc)
        {
            if (tnc == null) return;
            if (_miAddSelectNode != null && _fiHashSelectNode != null) return;
        
            Type t = tnc.GetType();
        
            // AddSelectNode(TreeNodeObject, bool)
            if (_miAddSelectNode == null)
            {
                try
                {
                    _miAddSelectNode = t.GetMethod(
                        "AddSelectNode",
                        BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance,
                        null,
                        new Type[] { typeof(TreeNodeObject), typeof(bool) },
                        null
                    );
                }
                catch { _miAddSelectNode = null; }
            }
        
            // hashSelectNode field (HashSet<TreeNodeObject> or similar)
            if (_fiHashSelectNode == null)
            {
                try
                {
                    _fiHashSelectNode = t.GetField(
                        "hashSelectNode",
                        BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance
                    );
                }
                catch { _fiHashSelectNode = null; }
            }
        }
        
        private void ClearSelection_KK(TreeNodeCtrl tnc)
        {
            ResolveTreeNodeCtrlKKApis(tnc);
            if (_fiHashSelectNode == null) return;
        
            object hsObj = null;
            try { hsObj = _fiHashSelectNode.GetValue(tnc); } catch { hsObj = null; }
            if (hsObj == null) return;
        
            // Enumerate selected nodes, call OnDeselect, then Clear()
            try
            {
                System.Collections.IEnumerable en = hsObj as System.Collections.IEnumerable;
                if (en != null)
                {
                    foreach (object o in en)
                    {
                        TreeNodeObject n = o as TreeNodeObject;
                        if (n != null)
                        {
                            try { n.OnDeselect(); } catch { }
                        }
                    }
                }
            }
            catch { }
        
            // call Clear()
            try
            {
                MethodInfo miClear = hsObj.GetType().GetMethod("Clear", BindingFlags.Public | BindingFlags.Instance, null, Type.EmptyTypes, null);
                if (miClear != null) miClear.Invoke(hsObj, null);
            }
            catch { }
        }
        
        private TreeNodeObject[] CaptureSelectedNodes_Compat()
        {
            TreeNodeCtrl tnc = GetTreeNodeCtrl();
            if (tnc == null) return null;
        
            ResolveTreeNodeCtrlKKApis(tnc);
            if (_fiHashSelectNode == null) return null;
        
            object hsObj = null;
            try { hsObj = _fiHashSelectNode.GetValue(tnc); } catch { hsObj = null; }
            if (hsObj == null) return null;
        
            // copy into List then ToArray
            List<TreeNodeObject> list = new List<TreeNodeObject>();
            try
            {
                System.Collections.IEnumerable en = hsObj as System.Collections.IEnumerable;
                if (en != null)
                {
                    foreach (object o in en)
                    {
                        TreeNodeObject n = o as TreeNodeObject;
                        if (n != null) list.Add(n);
                    }
                }
            }
            catch { }
        
            return list.ToArray();
        }
        
        private void SetSelectedNodes_Compat(TreeNodeObject[] nodes)
        {
            TreeNodeCtrl tnc = GetTreeNodeCtrl();
            if (tnc == null) return;
        
            ResolveTreeNodeCtrlKKApis(tnc);
        
            if (_miAddSelectNode == null)
            {
                Logger.LogWarning("TreeNodeCtrl: AddSelectNode(TreeNodeObject,bool) not found; cannot set selection.");
                return;
            }
        
            if (nodes == null) nodes = new TreeNodeObject[0];
        
            // EXACTLY mimic KK's SelectMultiple: deselect old + clear hashSelectNode
            ClearSelection_KK(tnc);
        
            // Then build selection set using AddSelectNode
            for (int i = 0; i < nodes.Length; i++)
            {
                TreeNodeObject n = nodes[i];
                if (n == null) continue;
        
                bool append = (i != 0); // first false, rest true
                try
                {
                    _miAddSelectNode.Invoke(tnc, new object[] { n, append });
                }
                catch (Exception ex)
                {
                    Logger.LogWarning("AddSelectNode invoke failed: " + ex.GetType().Name + " " + ex.Message);
                }
            }
        }

        private System.Collections.IEnumerator Co_RunFromDir(string dir, string tag)
        {
            _busy = true;
            try
            {
                dir = (dir == null) ? "" : dir.Trim();
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
                
                    //SetSelectedNodes_Compat(prevSel);
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
            
                        SetSelectedNodes_Compat(femaleNodes.ToArray());
                        yield return null;
                    }
                    string path = pngs[_rng.Next(0, pngs.Length)];
                    Logger.LogInfo("[F8] Batch single drop (no preselect). png=" + path);
                    PostOneDrop(hUnity, path);
                    yield break;
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
        // =========================================================
        // Migrated trigger hooks (from AffectedDynamicBonesEnabledPersistForKKS.cs)
        // - Character replace: HSPE.MainWindow.OnCharacterReplace postfix
        // - Scene load:        Studio.SceneInfo.Load(string, out Version) postfix
        // =========================================================

        private void TryInitHarmonyPatchAny()
        {
            try
            {
                if (_harmony == null) _harmony = new Harmony(GUID);

                // Harmony.Patch(...) overload count differs across Harmony versions.
                // We resolve the instance method "Patch" / "PatchAny" by reflection and call it with a padded args array.
                if (_miHarmonyPatchAny == null)
                {
                    // Prefer method name "Patch", fallback "PatchAny" (some forks)
                    Type th = typeof(Harmony);
                    MethodInfo[] ms = th.GetMethods(BindingFlags.Public | BindingFlags.Instance);
                    for (int i = 0; i < ms.Length; i++)
                    {
                        MethodInfo mi = ms[i];
                        if (mi == null) continue;
                        string n = mi.Name;
                        if (n != "Patch" && n != "PatchAny") continue;

                        ParameterInfo[] ps = mi.GetParameters();
                        if (ps == null || ps.Length < 2) continue;

                        // first param must be MethodBase
                        if (!typeof(MethodBase).IsAssignableFrom(ps[0].ParameterType)) continue;

                        _miHarmonyPatchAny = mi;
                        break;
                    }
                }
            }
            catch (Exception e)
            {
                Logger.LogWarning("[PATCH] TryInitHarmonyPatchAny failed: " + e.Message);
            }
        }

        // ---------------------------------------------------------
        // Reflection helpers (avoid AccessTools.* because old 0Harmony differs)
        // ---------------------------------------------------------

        private static Type ReflectGetType(string fullName)
        {
            if (string.IsNullOrEmpty(fullName)) return null;

            try
            {
                // Try direct Type.GetType first (works if assembly qualified)
                Type t0 = Type.GetType(fullName, false);
                if (t0 != null) return t0;
            }
            catch { }

            try
            {
                Assembly[] asms = AppDomain.CurrentDomain.GetAssemblies();
                int i;
                for (i = 0; i < asms.Length; i++)
                {
                    Assembly a = asms[i];
                    if (a == null) continue;
                    Type t = null;
                    try { t = a.GetType(fullName, false); }
                    catch { t = null; }
                    if (t != null) return t;
                }
            }
            catch { }

            return null;
        }

        private static bool ParamsMatch(ParameterInfo[] ps, Type[] want)
        {
            if (want == null) return true; // any signature
            if (ps == null) return (want.Length == 0);
            if (ps.Length != want.Length) return false;

            int i;
            for (i = 0; i < ps.Length; i++)
            {
                Type pt = ps[i].ParameterType;
                Type wt = want[i];

                if (pt == wt) continue;

                // Handle byref comparison safely
                if (pt != null && wt != null)
                {
                    if (pt.IsByRef && wt.IsByRef)
                    {
                        if (pt.GetElementType() == wt.GetElementType()) continue;
                    }
                }

                return false;
            }
            return true;
        }

        private static MethodInfo ReflectFindMethod(string typeFullName, string methodName, Type[] paramTypes)
        {
            if (string.IsNullOrEmpty(typeFullName) || string.IsNullOrEmpty(methodName)) return null;

            Type t = ReflectGetType(typeFullName);
            if (t == null) return null;

            BindingFlags bf = BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Static;

            try
            {
                MethodInfo[] ms = t.GetMethods(bf);
                int i;
                for (i = 0; i < ms.Length; i++)
                {
                    MethodInfo mi = ms[i];
                    if (mi == null) continue;
                    if (mi.Name != methodName) continue;

                    ParameterInfo[] ps = null;
                    try { ps = mi.GetParameters(); } catch { ps = null; }

                    if (ParamsMatch(ps, paramTypes)) return mi;
                }
            }
            catch { }

            return null;
        }

        private void PatchPostfix(MethodInfo target, MethodInfo postfix, int priority)
        {
            if (_harmony == null || target == null || postfix == null) return;
            if (_miHarmonyPatchAny == null) return;

            HarmonyMethod hmPost = new HarmonyMethod(postfix);
            hmPost.priority = priority;

            ParameterInfo[] ps = _miHarmonyPatchAny.GetParameters();
            object[] args = new object[ps.Length];

            args[0] = target;
            if (args.Length >= 2) args[1] = null;   // prefix
            if (args.Length >= 3) args[2] = hmPost; // postfix
            for (int i = 3; i < args.Length; i++) args[i] = null; // transpiler / finalizer / etc

            try { _miHarmonyPatchAny.Invoke(_harmony, args); }
            catch (Exception e) { Logger.LogWarning("[PATCH] PatchPostfix invoke failed: " + e.Message); }
        }

        private void Patch_StudioScene_Load()
        {
            if (_miHarmonyPatchAny == null) return;
        
            try
            {
                // Studio.SceneInfo.Load(string, out Version) : bool
                MethodInfo miLoad = AccessTools.Method("Studio.SceneInfo:Load",
                    new Type[] { typeof(string), typeof(Version).MakeByRefType() }, null);
        
                if (miLoad == null)
                {
                    Logger.LogWarning("[PATCH] Studio.SceneInfo.Load not found.");
                    return;
                }
        
                MethodInfo post = typeof(RandomPngDropViaWM_BatchFemale).GetMethod(
                    "SceneInfo_Load_Postfix",
                    BindingFlags.Static | BindingFlags.NonPublic);
        
                if (post == null)
                {
                    Logger.LogWarning("[PATCH] SceneInfo_Load_Postfix not found.");
                    return;
                }
        
                PatchPostfix(miLoad, post, Priority.Last);
                Logger.LogInfo("[PATCH] Patched Studio.SceneInfo.Load postfix.");
            }
            catch (Exception e)
            {
                Logger.LogWarning("[PATCH] Patch_StudioScene_Load failed: " + e.Message);
            }
        }

    

        private static void MW_OnCharacterReplace_Postfix()
        {
            if (_inst == null) return;
            if (!_inst._patchTriggers.Value) return;
            if (!_inst._autoOnCharacterReplace.Value) return;
        
            // 如果当前正在处理脚本引发的加载，或者正在执行协程，直接跳过
            if (_isInternalTrigger || _inst._busy) 
            {
                // Logger.LogInfo("[AUTO] Internal trigger or busy, skipping to prevent recursion.");
                return;
            }
        
            // 触发自动执行，理由设为 OnCharaLoad
            _inst.TriggerAutoRun("OnCharaLoad (Limited)");
        }

        private static void SceneInfo_Load_Postfix(string _path, ref Version _dataVersion, bool __result)
        {
            try
            {
                if (!__result) return;
                if (_inst == null) return;
                if (!_inst._patchTriggers.Value) return;
                if (!_inst._autoOnSceneLoad.Value) return;

                _inst.TriggerAutoRun("SceneInfo.Load");
            }
            catch { }
        }

        private void TriggerAutoRun(string reason)
        {
            if (!_enable.Value || _busy || _isInternalTrigger) return;
        
            int frames = Mathf.Clamp(_autoDelayFrames.Value, 0, 600);
            
            string dir;
            if (reason == "SceneInfo.Load")
            {
                // 场景加载时使用的目录（例如使用 F4 目录）
                dir = (_pngDir_F4.Value == null) ? "" : _pngDir_F4.Value.Trim();
                Logger.LogInfo("[AUTO] Scene Load detected. Target: PngDir_F4");
            }
            else
            {
                // 角色更换/OnCharaLoad 时使用的目录（按你要求改为 Ctrl+F4 目录）
                dir = (_pngDir_CtrlF4.Value == null) ? "" : _pngDir_CtrlF4.Value.Trim();
                Logger.LogInfo("[AUTO] Character Load detected. Target: PngDir_CtrlF4");
            }
            if (string.IsNullOrEmpty(dir)) return;
        
            StartCoroutine(Co_DelayedAutoRun(frames, dir, reason));
        }
        private IEnumerator Co_DelayedAutoRun(int frames, string dir, string reason)
        {
            _busy = true;
            _isInternalTrigger = true; // 开启拦截锁
            
            try
            {
                for (int i = 0; i < frames; i++) yield return null;

                if (!_enable.Value) yield break;

                Logger.LogInfo("[AUTO] {reason} triggered. Using Ctrl+F4 Path: "+dir);
                
                // 执行随机 Drop 逻辑
                yield return StartCoroutine(Co_RunFromDir(dir, "AUTO_ONCE"));
                
                // 执行完后额外等待一小段时间，确保 Unity 处理完所有的消息队列
                for (int i = 0; i < 10; i++) yield return null;
            }
            finally
            {
                _busy = false;
                _isInternalTrigger = false; // 解锁，允许下一次手动或合法的自动触发
                Logger.LogInfo("[AUTO] Lock released.");
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

using Microsoft.VisualBasic;
using SolidWorks.Interop.sldworks; // Основное API SolidWorks
using SolidWorks.Interop.swconst; // Добавляем пространство имен для констант
using SolidWorks.Interop.swpublished; // Для ISwAddin
using SwCommands;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using View = SolidWorks.Interop.sldworks.View;

namespace SolidWorksAddIn
{
    [ComVisible(true)]
    [Guid("382FC34A-AE35-42AA-B965-828EC7F10D66")]

    public class MyAddIn : ISwAddin
    {
        #region Переменные
        private ISldWorks swApp; //Основной объект SolidWorks
        private int addInID; //ID плагина 
        private CommandManager cmdMgr;   // Менеджер команд для создания кнопок
        private const int mainCmdGroupID = 5;
        private const int mainCmdID = 1;
        private const string ADDIN_TITLE = "Object Lister"; // Название плагина
        private Dictionary<double, int> cachedHoles;
        private Dictionary<double, (int count, string type)> cachedFeatures = new Dictionary<double, (int, string)>();
        private Dictionary<double, int> cachedSlots = new Dictionary<double, int>();

        HashSet<string> usedCylinderKeys = new HashSet<string>();
        HashSet<string> slotFeatures = new HashSet<string>();

        private int totalSlots = 0; // Общее количество пазов

        public PartDoc pDoc;
        public AssemblyDoc aDoc;
        public DrawingDoc dDoc;
        public ModelDoc2 swModel;
        #endregion

        //Регитсрация COM-интерфейса
        [ComRegisterFunction]
        private static void RegisterFunction(Type t)
        {
            Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry
                .LocalMachine
                .CreateSubKey($@"SOFTWARE\SolidWorks\AddIns\{t.GUID:B}");
            key.SetValue(null, 1);
            key.SetValue("Title", ADDIN_TITLE);
            key.SetValue("Description", "Пример плагина SolidWorks");
        }

        // Удаление COM-регистрации
        [ComUnregisterFunction]
        private static void UnregisterFunction(Type t)
        {
            Microsoft.Win32.Registry.LocalMachine
                .DeleteSubKey($@"SOFTWARE\SolidWorks\AddIns\{t.GUID:B}");
        }

        #region Основные методы плагина
        /// <summary>
        /// Вызывается при подключении плагина
        /// </summary>
        public bool ConnectToSW(object ThisSW, int Cookie)
        {
            swApp = (ISldWorks)ThisSW;
            addInID = Cookie;

            // Регистрация плагина в SolidWorks
            swApp.SetAddinCallbackInfo2(0, this, addInID);

            cmdMgr = swApp.GetCommandManager(addInID);
            if (cmdMgr == null)
            {
                MessageBox.Show("Не удалось получить CommandManager", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false; // Не подключаемся, если не получили CommandManager
            }

            // Создание кнопки в интерфейсе
            AddCommandManager();

            //MessageBox.Show("Плагин запущен!", "MyAddIn", MessageBoxButtons.OK, MessageBoxIcon.Information);

            return true;
        }
        /// <summary>
        /// Вызывается при отключении плагина
        /// </summary>
        public bool DisconnectFromSW()
        {
            // Очистка ресурсов
            Marshal.ReleaseComObject(swApp);
            swApp = null;
            GC.Collect();

            return true;
        }
        #endregion

        #region Создание и удаление кнопки
        private void AddCommandManager()
        {
            int errors = 0;
            var cmdGroup = cmdMgr.CreateCommandGroup2(mainCmdGroupID, "Функции автоматизации", "List all features", "", -1, false, ref errors);
            int cmdIndex = cmdGroup.AddCommandItem2(
                "Список объектов",
                -1,
                "Список вырезов получеых с детали",
                "List Bodies",
                0,
                "ShowObjects",
                "",
                mainCmdID,
                (int)(swCommandItemType_e.swMenuItem | swCommandItemType_e.swToolbarItem));

            cmdGroup.AddCommandItem2(
                "Обновить размеры",
                -1,
                "Подставить количество отверстий",
                "Update Holes",
                1,
                "UpdateDimensionText",
                "",
                mainCmdID + 1,
                (int)(swCommandItemType_e.swMenuItem | swCommandItemType_e.swToolbarItem));

            cmdGroup.AddCommandItem2(
                "Переименовать модель и чертеж",
                -1,
                "Автоматически переименовать файл чертежа при переименовании модели и наоборот",
                "Rename",
                2,
                "Rename",
                "",
                mainCmdID + 2,
                (int)(swCommandItemType_e.swMenuItem | swCommandItemType_e.swToolbarItem));

            cmdGroup.AddCommandItem2(
                "Убрать рамки при смене масшба листа.",
                -1,
                "Автоматически переименовать файл чертежа при переименовании модели и наоборот",
                "Hide Border Keep Title",
                3,
                "HideBorderKeepTitle",
                "",
                mainCmdID + 3,
                (int)(swCommandItemType_e.swMenuItem | swCommandItemType_e.swToolbarItem));
            cmdGroup.HasToolbar = true;
            cmdGroup.HasMenu = true;
            cmdGroup.Activate();
        }
        #endregion

        #region Обработчик нажатия кнопки
        // Этот метод будет вызываться при нажатии на кнопку
        public void ShowObjects()
        {
            IModelDoc2 doc = null;
            try
            {
                doc = swApp.IActiveDoc2;
                if (doc == null)
                {
                    MessageBox.Show("Откройте документ.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (doc.GetType() != (int)swDocumentTypes_e.swDocPART)
                {
                    MessageBox.Show("Откройте деталь.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                cachedFeatures.Clear();
                cachedSlots.Clear();
                totalSlots = 0;

                ProcessFeatures(doc);
                ProcessBodies(doc, usedCylinderKeys, slotFeatures);

                ShowResults();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Критическая ошибка: {ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                ReleaseComObject(doc);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void ProcessFeatures(IModelDoc2 doc)
        {
            IFeature feat = doc.FirstFeature() as IFeature;
            while (feat != null)
            {
                string featType = feat.GetTypeName();
                string featName = feat.Name;

                try
                {
                    if (featType == "HoleWzd")
                    {
                        ProcessHoleFeature(feat, doc);
                    }
                    else if (featType == "CutExtrude" || featType == "Extrude")
                    {
                        ProcessExtrudeFeature(feat);
                    }
                    else if (featType == "LinearPattern" || featType == "CircularPattern" ||
                             featType == "LocalPattern" || featType == "FPattern" || featType == "SketchPattern")
                    {
                        ProcessPatternFeature(feat);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка обработки признака '{featName}': {ex.Message}", "Ошибка",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                IFeature nextFeat = feat.GetNextFeature() as IFeature;
                ReleaseComObject(feat);
                feat = nextFeat;
            }
        }

        private void ProcessHoleFeature(IFeature feat, IModelDoc2 doc)
        {
            IWizardHoleFeatureData2 hole = feat.GetDefinition() as IWizardHoleFeatureData2;
            if (hole == null) return;

            bool accessGranted = false;
            try
            {
                hole.AccessSelections(doc, null);
                accessGranted = true;

                double dia = Math.Round(hole.HoleDiameter * 1000, 2);
                string formattedDia = dia.ToString("0.00", CultureInfo.InvariantCulture);
                double normalizedDia = double.Parse(formattedDia, CultureInfo.InvariantCulture);

                int count = hole.GetSketchPointCount();
                bool isSlot = IsHoleWizardSlot(feat, hole);

                if (normalizedDia > 0 && count > 0)
                {
                    if (isSlot)
                    {
                        totalSlots += count;
                        slotFeatures.Add($"HW_{feat.GetID()}_{normalizedDia}");

                        if (!cachedSlots.ContainsKey(normalizedDia))
                            cachedSlots[normalizedDia] = 0;
                        cachedSlots[normalizedDia] += count;
                    }
                    else
                    {
                        if (!cachedFeatures.ContainsKey(normalizedDia))
                            cachedFeatures[normalizedDia] = (0, "отверстие");
                        cachedFeatures[normalizedDia] = (cachedFeatures[normalizedDia].count + count, "отверстие");
                    }
                }
            }
            finally
            {
                if (accessGranted)
                    hole.ReleaseSelectionAccess();
                ReleaseComObject(hole);
            }
        }

        private void ProcessExtrudeFeature(IFeature feat)
        {
            if (IsSlotFeature(feat))
            {
                totalSlots++;
                slotFeatures.Add($"CE_{feat.GetID()}");

                double slotWidth = GetSlotWidth(feat);
                if (slotWidth > 0)
                {
                    string formattedWidth = slotWidth.ToString("0.00", CultureInfo.InvariantCulture);
                    double normalizedWidth = double.Parse(formattedWidth, CultureInfo.InvariantCulture);

                    if (!cachedSlots.ContainsKey(normalizedWidth))
                        cachedSlots[normalizedWidth] = 0;
                    cachedSlots[normalizedWidth]++;
                }
            }
        }

        private void ProcessPatternFeature(IFeature pattern)
        {
            IFeature subFeat = pattern.GetFirstSubFeature() as IFeature;
            bool containsSlot = false;

            while (subFeat != null)
            {
                if (IsSlotFeature(subFeat))
                {
                    containsSlot = true;
                    ReleaseComObject(subFeat);
                    break;
                }
                IFeature nextSubFeat = subFeat.GetNextSubFeature() as IFeature;
                ReleaseComObject(subFeat);
                subFeat = nextSubFeat;
            }

            if (containsSlot)
            {
                object children = pattern.GetChildren();
                if (children is object[] childArray)
                {
                    totalSlots += childArray.Length;
                }
                else
                {
                    string patternName = pattern.Name;
                    var match = Regex.Match(patternName, @"\((\d+)\)");
                    if (match.Success && int.TryParse(match.Groups[1].Value, out int count))
                    {
                        totalSlots += count;
                    }
                }
                ReleaseComObject(children);
            }
        }

        private void ProcessBodies(IModelDoc2 doc, HashSet<string> usedCylinderKeys, HashSet<string> slotFeatures)
        {
            object bodiesObj = ((PartDoc)doc).GetBodies2((int)swBodyType_e.swSolidBody, false);
            if (bodiesObj == null) return;

            Body2[] bodies = (bodiesObj as object[])?.Cast<Body2>().ToArray();
            if (bodies == null) return;

            foreach (var body in bodies)
            {
                object facesObj = body.GetFaces();
                if (facesObj == null) continue;

                Face2[] faces = (facesObj as object[])?.Cast<Face2>().ToArray();
                if (faces == null) continue;

                foreach (var face in faces)
                {
                    try
                    {
                        Surface surf = face.GetSurface() as Surface;
                        if (surf == null || !surf.IsCylinder()) continue;

                        double[] pars = surf.CylinderParams as double[];
                        if (pars == null || pars.Length < 7) continue;

                        double dia = Math.Round(pars[6] * 2 * 1000, 2);
                        string formattedDia = dia.ToString("0.00", CultureInfo.InvariantCulture);
                        double normalizedDia = double.Parse(formattedDia, CultureInfo.InvariantCulture);

                        string key = $"{pars[0]:F6}_{pars[1]:F6}_{pars[2]:F6}_{pars[6]:F6}";

                        IFeature faceFeature = face.GetFeature() as IFeature;
                        if (faceFeature != null &&
                            !usedCylinderKeys.Contains(key) &&
                            !slotFeatures.Contains($"CE_{faceFeature.GetID()}") &&
                            !slotFeatures.Contains($"HW_{faceFeature.GetID()}_{normalizedDia}"))
                        {
                            usedCylinderKeys.Add(key);
                            if (!cachedFeatures.ContainsKey(normalizedDia))
                                cachedFeatures[normalizedDia] = (0, "отверстие");
                            cachedFeatures[normalizedDia] = (cachedFeatures[normalizedDia].count + 1, "отверстие");
                        }
                        ReleaseComObject(faceFeature);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка обработки поверхности: {ex.Message}", "Ошибка",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    finally
                    {
                        ReleaseComObject(face);
                    }
                }
                ReleaseComObject(facesObj);
                ReleaseComObject(body);
            }
            ReleaseComObject(bodiesObj);
        }

        private void ShowResults()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("Обнаруженные отверстия:");
            foreach (var kv in cachedFeatures.OrderBy(k => k.Key))
                sb.AppendLine($"⌀{kv.Key:F2} мм — {kv.Value.count} шт. ({kv.Value.type})");

            sb.AppendLine("\nОбнаружено пазов:");
            foreach (var kv in cachedSlots.OrderBy(k => k.Key))
                sb.AppendLine($"⌀{kv.Key:F2} мм — {kv.Value} шт.");

            sb.AppendLine($"\nВсего пазов: {totalSlots}");

            if (cachedFeatures.Count == 0)
            {
                MessageBox.Show("Не найдено отверстий в детали!", "Предупреждение",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            if (cachedSlots.Count == 0)
            {
                MessageBox.Show("Не найдено пазов в детали!", "Предупреждение",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            MessageBox.Show(sb.ToString(), "Успешно", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private bool IsHoleWizardSlot(IFeature feature, IWizardHoleFeatureData2 holeData)
        {
            try
            {
                string featName = feature.Name?.ToLower() ?? "";
                if (featName.Contains("паз") || featName.Contains("slot") ||
                    featName.Contains("шестигран") || featName.Contains("hex") ||
                    featName.Contains("гайк") || featName.Contains("nut"))
                {
                    return true;
                }

                IFeature sketchFeat = feature.GetFirstSubFeature() as IFeature;
                while (sketchFeat != null)
                {
                    string subFeatType = sketchFeat.GetTypeName();
                    if (subFeatType == "ProfileFeature" || subFeatType == "Sketch")
                    {
                        ISketch sketch = sketchFeat.GetSpecificFeature2() as ISketch;
                        if (sketch != null)
                        {
                            object segments = sketch.GetSketchSegments();
                            if (segments is object[] segArray && segArray.Length > 0)
                            {
                                foreach (object segObj in segArray)
                                {
                                    ISketchSegment segment = segObj as ISketchSegment;
                                    if (segment != null)
                                    {
                                        int segType = segment.GetType();
                                        if (segType == (int)swSketchSegments_e.swSketchLINE ||
                                            segType == (int)swSketchSegments_e.swSketchARC ||
                                            segType == (int)swSketchSegments_e.swSketchELLIPSE ||
                                            segType == (int)swSketchSegments_e.swSketchPARABOLA ||
                                            segType == (int)swSketchSegments_e.swSketchSPLINE)
                                        {
                                            return true;
                                        }
                                    }
                                }
                            }
                        }
                        break;
                    }
                    IFeature nextSketchFeat = sketchFeat.GetNextSubFeature() as IFeature;
                    ReleaseComObject(sketchFeat);
                    sketchFeat = nextSketchFeat;
                }

                string standard = holeData.Standard?.ToLower() ?? "";
                string fastenerType = holeData.FastenerType?.ToLower() ?? "";

                if (standard.Contains("slot") || standard.Contains("паз") ||
                    fastenerType.Contains("slot") || fastenerType.Contains("паз"))
                {
                    return true;
                }

                if (fastenerType.Contains("hex") || fastenerType.Contains("шестигран") ||
                    fastenerType.Contains("nut") || fastenerType.Contains("гайк"))
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка определения типа отверстия: {ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            return false;
        }

        private bool IsSlotFeature(IFeature feature)
        {
            try
            {
                string featName = feature.Name?.ToLower() ?? "";
                if (featName.Contains("паз") || featName.Contains("slot") ||
                    featName.Contains("шестигран") || featName.Contains("hex") ||
                    featName.Contains("гайк") || featName.Contains("nut"))
                {
                    return true;
                }

                object bodyObjects = feature.GetBody();
                if (bodyObjects is object[] bodies)
                {
                    foreach (Body2 body in bodies)
                    {
                        double[] box = body.GetBodyBox() as double[];
                        if (box != null && box.Length == 6)
                        {
                            double dx = Math.Abs(box[3] - box[0]);
                            double dy = Math.Abs(box[4] - box[1]);
                            double dz = Math.Abs(box[5] - box[2]);

                            double maxDim = Math.Max(dx, Math.Max(dy, dz));
                            double minDim = Math.Min(dx, Math.Min(dy, dz));

                            if (minDim > 0 && maxDim / minDim > 2.0)
                            {
                                return true;
                            }
                        }
                    }
                }
                ReleaseComObject(bodyObjects);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка проверки признака паза: {ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            return false;
        }

        private double GetSlotWidth(IFeature feature)
        {
            try
            {
                object bodyObj = feature.GetBody();
                if (bodyObj is object[] bodies)
                {
                    foreach (Body2 body in bodies)
                    {
                        double[] box = body.GetBodyBox() as double[];
                        if (box != null && box.Length == 6)
                        {
                            double dx = Math.Abs(box[3] - box[0]);
                            double dy = Math.Abs(box[4] - box[1]);
                            double dz = Math.Abs(box[5] - box[2]);

                            return Math.Min(dx, Math.Min(dy, dz));
                        }
                    }
                }
                ReleaseComObject(bodyObj);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка определения ширины паза: {ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            return 0;
        }

        public void UpdateDimensionText()
        {
            IModelDoc2 modelDoc = null;
            DrawingDoc drawing = null;
            Sheet currentSheet = null;
            object viewsObj = null;

            try
            {
                modelDoc = swApp.IActiveDoc2;
                if (modelDoc == null) return;

                // Проверка типа документа
                if (modelDoc.GetType() != (int)swDocumentTypes_e.swDocDRAWING)
                {
                    MessageBox.Show("Откройте чертёж.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Проверка кэша
                if (cachedFeatures.Count == 0 && cachedSlots.Count == 0)
                {
                    MessageBox.Show("Сначала выполните анализ детали (Список объектов).", "Ошибка",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                drawing = (DrawingDoc)modelDoc;
                currentSheet = (Sheet)drawing.GetCurrentSheet();
                viewsObj = currentSheet.GetViews();

                if (viewsObj == null) return;

                object[] views = (object[])viewsObj;
                int updatedHoles = 0;
                int updatedSlots = 0;
                int totalDimensions = 0;
                int processedDimensions = 0;
                HashSet<double> processedHoleDiameters = new HashSet<double>();

                foreach (object viewObj in views)
                {
                    View view = null;
                    object annotationsObj = null;

                    try
                    {
                        view = (View)viewObj;
                        if (view == null) continue;

                        annotationsObj = view.GetAnnotations();
                        if (annotationsObj == null) continue;

                        object[] annotations = (object[])annotationsObj;
                        foreach (object annObj in annotations)
                        {
                            Annotation ann = null;
                            object specific = null;
                            DisplayDimension dispDim = null;
                            Dimension dim = null;

                            try
                            {
                                ann = (Annotation)annObj;
                                if (ann == null) continue;

                                if (ann.GetType() == (int)swAnnotationType_e.swDisplayDimension)
                                {
                                    totalDimensions++;
                                    specific = ann.GetSpecificAnnotation();
                                    dispDim = (DisplayDimension)specific;
                                    if (dispDim == null) continue;

                                    dim = (Dimension)dispDim.GetDimension();
                                    if (dim == null) continue;

                                    double valueInMeters = dim.GetSystemValue2("");
                                    double size = Math.Round(valueInMeters * 1000, 2);
                                    double normalizedSize = double.Parse(size.ToString("0.00", CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);

                                    // Поиск соответствующего диаметра отверстия
                                    double holeKey = cachedFeatures.Keys
                                        .Select(k => new { Key = k, Diff = Math.Abs(k - normalizedSize) })
                                        .Where(x => x.Diff < 0.01)
                                        .OrderBy(x => x.Diff)
                                        .Select(x => x.Key)
                                        .FirstOrDefault();

                                    if (holeKey > 0)
                                    {
                                        if (cachedFeatures.TryGetValue(holeKey, out var holeData) &&
                                            holeData.type == "отверстие" &&
                                            holeData.count > 1 &&
                                            !processedHoleDiameters.Contains(holeKey))
                                        {
                                            string newText = $"<MOD-DIAM>{size:F2} мм\n{holeData.count} отверстий";
                                            dispDim.SetText((int)swDimensionTextParts_e.swDimensionTextAll, newText);
                                            updatedHoles++;
                                            processedHoleDiameters.Add(holeKey);
                                        }
                                    }
                                    else // Обработка пазов
                                    {
                                        double slotKey = cachedSlots.Keys
                                            .Select(k => new { Key = k, Diff = Math.Abs(k - normalizedSize) })
                                            .Where(x => x.Diff < 0.01)
                                            .OrderBy(x => x.Diff)
                                            .Select(x => x.Key)
                                            .FirstOrDefault();

                                        if (slotKey > 0 && cachedSlots.TryGetValue(slotKey, out int slotCount))
                                        {
                                            string newText = slotCount > 1
                                                ? $"<MOD-DIAM>{size:F2} мм\n{slotCount} пазов"
                                                : $"<MOD-DIAM>{size:F2} мм (паз)";

                                            dispDim.SetText((int)swDimensionTextParts_e.swDimensionTextAll, newText);
                                            updatedSlots++;
                                        }
                                    }

                                    processedDimensions++;
                                }
                            }
                            finally
                            {
                                // Упорядоченное освобождение объектов
                                if (dim != null) Marshal.FinalReleaseComObject(dim);
                                if (dispDim != null) Marshal.FinalReleaseComObject(dispDim);
                                if (specific != null) Marshal.FinalReleaseComObject(specific);
                                if (ann != null) Marshal.FinalReleaseComObject(ann);
                            }
                        }
                    }
                    finally
                    {
                        if (annotationsObj != null) Marshal.FinalReleaseComObject(annotationsObj);
                        if (view != null) Marshal.FinalReleaseComObject(view);
                    }
                }

                string message = $"Обновлено: {updatedHoles} отверстий, {updatedSlots} пазов\n" +
                                 $"Обработано размеров: {processedDimensions} из {totalDimensions}";

                MessageBox.Show(message, "Результат обновления", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при обновлении размеров: {ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Освобождение в правильном порядке
                if (viewsObj != null) Marshal.FinalReleaseComObject(viewsObj);
                if (currentSheet != null) Marshal.FinalReleaseComObject(currentSheet);
                if (drawing != null) Marshal.FinalReleaseComObject(drawing);
                if (modelDoc != null) Marshal.FinalReleaseComObject(modelDoc);

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void ReleaseComObject(object comObject)
        {
            if (comObject != null && Marshal.IsComObject(comObject))
            {
                Marshal.ReleaseComObject(comObject);
            }
        }
        #endregion

        ///// <summary>
        ///// Переименовывает файл чертежа при переименовании модели
        ///// </summary>
        public void Rename()
        {
            try
            {
                IModelDoc2 activeDoc = swApp.IActiveDoc2;
                if (activeDoc == null || string.IsNullOrEmpty(activeDoc.GetPathName()))
                {
                    MessageBox.Show("Откройте и сохраните документ.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                string oldBaseName = Path.GetFileNameWithoutExtension(activeDoc.GetPathName());
                string newBaseName = Interaction.InputBox("Введите новое имя (без расширения):", "Переименование", oldBaseName);
                if (string.IsNullOrEmpty(newBaseName) || newBaseName == oldBaseName) return;
                SwRename.Execute(swApp, activeDoc, newBaseName);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при переименовании: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void HideBorderKeepTitle()
        {
            IModelDoc2 model = swApp.IActiveDoc2;
            if (model == null || model.GetType() != (int)swDocumentTypes_e.swDocDRAWING)
            {
                swApp.SendMsgToUser2("Откройте чертёж перед запуском.",
                    (int)swMessageBoxIcon_e.swMbStop,
                    (int)swMessageBoxBtn_e.swMbOk);
                return;
            }

            DrawingDoc drawing = (DrawingDoc)model;

            // Заходим в режим редактирования формата
            drawing.EditTemplate();

            View v = drawing.GetFirstView();
            if (v == null) return;

            Sketch sk = v.GetSketch();
            if (sk == null) return;

            // Ищем блоки в эскизе
            object[] blocks = (object[])sk.GetSketchBlocks();
            if (blocks == null || blocks.Length == 0)
            {
                swApp.SendMsgToUser2("Блоков в формате не найдено.",
                    (int)swMessageBoxIcon_e.swMbInformation,
                    (int)swMessageBoxBtn_e.swMbOk);
            }
            else
            {
                System.Text.StringBuilder sb = new System.Text.StringBuilder();
                sb.AppendLine($"Найдено {blocks.Length} блок(ов).");

                foreach (SketchBlockInstance block in blocks)
                {
                    sb.AppendLine($"Блок: {block.Name}");

                    object[] segs = (object[])block.GetSketchSegments();
                    if (segs != null)
                    {
                        sb.AppendLine($"   Сегментов в блоке: {segs.Length}");
                    }
                }

                swApp.SendMsgToUser2(sb.ToString(),
                    (int)swMessageBoxIcon_e.swMbInformation,
                    (int)swMessageBoxBtn_e.swMbOk);
            }

            // Возврат к редактированию листа
            drawing.EditSheet();
        }
    }
}
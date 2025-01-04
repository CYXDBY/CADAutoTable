using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System.Collections.Generic;
using System.Diagnostics;
using System.Xml.Linq;
using System.Data;
using System;
using System.Text.RegularExpressions;
using System.Linq;
using System.Reflection.Emit;
using System.IO;
using Newtonsoft.Json;
using System.Runtime.Remoting.Metadata.W3cXsd2001;
using System.Collections;
using Newtonsoft.Json.Linq;
using Autodesk.AutoCAD.Colors;
using System.Runtime.InteropServices;



namespace MyCADPlugin
{
    public class TableModifier : IExtensionApplication
    {
        int smallTableSign = 0;

        int textSize = 300;
        int headerSize = 200;
        int headerWidth = 1250;
        int tableWidth = 1400;
        int tableHeight = 660;
        double cellWidth = 3000;
        double cellWidth2 = 600;
        double cellHeight = 600;

        int textSizeBig = 300;
        double cellWidthBig = 2500;
        double cellHeightBig = 700;

        List<string> smallBox = new List<string>();
        List<string> smallUsage = new List<string>();
        List<double[]> smallData = new List<double[]>();

        List<string> bigBox = new List<string>();
        List<string> bigLoop = new List<string>();

        Dictionary<int, string> cableList;
        int[] sortedKeys;
        bool firstRun = true;

        public void WriteSmallTableSign()
        {
            // 获取当前文档和数据库
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database database = doc.Database;

            Dictionary<string, int> ProjectInforArray = new Dictionary<string, int>();
            ProjectInforArray.Add("SmallTableSign", smallTableSign);
            using (Transaction trx = database.TransactionManager.StartTransaction())
            {

                DatabaseSummaryInfoBuilder BuilderSummaryInfo = new DatabaseSummaryInfoBuilder();
                foreach (KeyValuePair<string, int> Project in ProjectInforArray)
                {
                    BuilderSummaryInfo.CustomPropertyTable.Add(Project.Key, Project.Value.ToString());
                }
                database.SummaryInfo = BuilderSummaryInfo.ToDatabaseSummaryInfo();
                trx.Commit();
            }
        }

        public void ReadSmallTableSign()
        {
            // 获取当前文档和数据库
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database database = doc.Database;

            Dictionary<string, string> result = new Dictionary<string, string>();
            IDictionaryEnumerator dictEnum = database.SummaryInfo.CustomProperties;
            while (dictEnum.MoveNext())
            {
                DictionaryEntry entry = dictEnum.Entry;
                result.Add((string)entry.Key, (string)entry.Value);
            }


            foreach (var c in result)
            {
                if (c.Key == "SmallTableSign")
                {
                    smallTableSign = int.Parse(c.Value);
                    break; // 找到后退出循环
                }
            }
        }

        public void Initialize() { }

        public void Terminate() { }

        //—————————————以下为搜索表命令实现————————————————
        // 读取大表小表
        [CommandMethod("GH")]
        public void GetHandle()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            // 提示用户选择对象
            PromptSelectionOptions opts = new PromptSelectionOptions();
            opts.MessageForAdding = "\n选择要读取的对象: ";
            PromptSelectionResult res = ed.GetSelection(opts);

            if (res.Status != PromptStatus.OK)
                return;

            SelectionSet selSet = res.Value;

            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                foreach (SelectedObject selObj in selSet)
                {
                    if (selObj == null) continue;

                    // 尝试获取选中对象的 BlockReference
                    DBObject blockRef = trans.GetObject(selObj.ObjectId, OpenMode.ForRead) as DBObject;
                    if (blockRef != null)
                    {
                        ed.WriteMessage($"\nHandle={blockRef.Handle.ToString()}");
                    }
                }
                trans.Commit();
            }
        }

        // 读取大表小表
        [CommandMethod("RT")]
        public void ReadTables()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            // 提示用户选择对象
            PromptSelectionOptions opts = new PromptSelectionOptions();
            opts.MessageForAdding = "\n选择要读取的对象: ";
            PromptSelectionResult res = ed.GetSelection(opts);

            if (res.Status != PromptStatus.OK)
                return;

            SelectionSet selSet = res.Value;

            List<BlockReference> small2Tables = new List<BlockReference>();
            List<BlockReference> small3Tables = new List<BlockReference>();
            List<BlockReference> bigTables = new List<BlockReference>();
            List<DBText> dbTexts = new List<DBText>();
            List<Table> dbTables = new List<Table>();

            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                foreach (SelectedObject selObj in selSet)
                {
                    if (selObj == null) continue;

                    // 尝试获取选中对象的 BlockReference
                    BlockReference blockRef = trans.GetObject(selObj.ObjectId, OpenMode.ForRead) as BlockReference;
                    if (blockRef != null)
                    {
                        // 获取与 BlockReference 相关联的 BlockTableRecord
                        BlockTableRecord btr = trans.GetObject(blockRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                        if (btr != null)
                        {
                            ed.WriteMessage($"\n找到块: {btr.Handle}");
                            // 检查 BlockTableRecord 的名称是否包含 "Small"
                            if (btr.Name.Contains("Small2"))
                            {
                                small2Tables.Add(blockRef);
                                ed.WriteMessage($"\n找到块: {btr.Name}");
                            }
                            else if (btr.Name.Contains("Small3"))
                            {
                                small3Tables.Add(blockRef);
                                ed.WriteMessage($"\n找到块: {btr.Name}");
                            }
                            else if (btr.Name.Contains("Big"))
                            {
                                bigTables.Add(blockRef);
                                ed.WriteMessage($"\n找到块: {btr.Name}");
                            }
                        }
                    }

                    // 尝试获取选中对象的 DBText
                    DBText dbtext = trans.GetObject(selObj.ObjectId, OpenMode.ForRead) as DBText;
                    if (dbtext != null)
                    {
                        dbTexts.Add(dbtext);
                    }

                    // 尝试获取选中对象的 Table
                    Table dbtable = trans.GetObject(selObj.ObjectId, OpenMode.ForRead) as Table;
                    if (dbtable != null)
                    {
                        dbTables.Add(dbtable);
                    }
                }
                trans.Commit();
            }

            smallBox = new List<string>();
            smallUsage = new List<string>();
            smallData = new List<double[]>();

            bigBox = new List<string>();
            bigLoop = new List<string>();


            // 先根据大小表把DBText进行划分，节省开销
            Dictionary<BlockReference, List<DBText>> classifiedResults = new Dictionary<BlockReference, List<DBText>>();
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                // 初始化每个 BlockReference 的分类列表
                foreach (BlockReference s2 in small2Tables) classifiedResults[s2] = new List<DBText>();
                foreach (BlockReference s3 in small3Tables) classifiedResults[s3] = new List<DBText>();
                foreach (BlockReference bt in bigTables) classifiedResults[bt] = new List<DBText>();

                // 遍历每一个 DBText 对象
                foreach (DBText dbt in dbTexts)
                {
                    // 遍历 BlockReference 列表，判断是否落在该表格内
                    foreach (BlockReference br in small2Tables)
                    {
                        BlockTableRecord btr = trans.GetObject(br.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                        int rowCount = GetRowCount(btr);
                        double maxX = br.Position.X + 3 * cellWidth;
                        double minX = br.Position.X;
                        double maxY = br.Position.Y;
                        double minY = br.Position.Y - rowCount * cellHeight;
                        if (dbt.Position.X >= minX && dbt.Position.X <= maxX && dbt.Position.Y >= minY && dbt.Position.Y <= maxY)
                        {
                            // 将 DBText 添加到当前 DB 对象的分类列表中
                            classifiedResults[br].Add(dbt);
                        }
                    }
                    foreach (BlockReference br in small3Tables)
                    {
                        BlockTableRecord btr = trans.GetObject(br.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                        int rowCount = GetRowCount(btr);
                        double maxX = br.Position.X + 3 * cellWidth;
                        double minX = br.Position.X;
                        double maxY = br.Position.Y;
                        double minY = br.Position.Y - rowCount * cellHeight;
                        if (dbt.Position.X >= minX && dbt.Position.X <= maxX && dbt.Position.Y >= minY && dbt.Position.Y <= maxY)
                        {
                            // 将 DBText 添加到当前 DB 对象的分类列表中
                            classifiedResults[br].Add(dbt);
                        }
                    }
                    foreach (BlockReference br in bigTables)
                    {
                        BlockTableRecord btr = trans.GetObject(br.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                        double maxX = br.Position.X + cellWidthBig;
                        double minX = br.Position.X;
                        double maxY = br.Position.Y;
                        double minY = br.Position.Y - 19 * cellHeightBig;
                        if (dbt.Position.X >= minX && dbt.Position.X <= maxX && dbt.Position.Y >= minY && dbt.Position.Y <= maxY)
                        {
                            // 将 DBText 添加到当前 DB 对象的分类列表中
                            classifiedResults[br].Add(dbt);
                        }
                    }
                }
            }

            // 再在每个表中进行识别
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                foreach (BlockReference br in small2Tables)
                {
                    BlockTableRecord btr = trans.GetObject(br.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;

                    // 读取table
                    double[] data = new double[5];
                    string[] fruits = btr.Name.Split('|'); // 按|分割

                    bool found = false;
                    Table table = new Table();
                    foreach (Table t in dbTables)
                    {
                        // 检查并读取 XData

                        if (t != null)
                        {
                            ResultBuffer xdata = t.GetXDataForApplication("CADAUTOTABLE");

                            if (xdata != null)
                            {
                                foreach (var value in xdata)
                                {
                                    // 检查是否包含特定值
                                    if (value.Value.ToString().Contains("表格编号:"))
                                    {
                                        if (value.Value.ToString().Substring(5) == fruits[1].ToString())
                                        {
                                            found = true;
                                            // 处理找到的XData
                                            ed.WriteMessage($"找到 XData: {value.Value}");
                                            table = t;
                                            break;
                                        }
                                    }
                                }
                            }
                        }

                    }

                    if (found)
                    {
                        data[0] = double.Parse(table.Cells[0, 1].Value.ToString());
                        data[1] = double.Parse(table.Cells[1, 1].Value.ToString());
                        data[2] = data[0] * data[1];
                        data[3] = double.Parse(table.Cells[3, 1].Value.ToString());
                        data[4] = data[2] / 0.38 / data[3] / 1.732;
                    }
                    else
                    {
                        data[0] = 0;
                        data[1] = 0;
                        data[2] = 0;
                        data[3] = 0;
                        data[4] = 0;
                    }

                    int rowCount = GetRowCount(btr);
                    for (int i = 1; i < rowCount; i++)
                    {
                        int count = 0;
                        string[] result = { "无字符串", "无字符串", "无字符串" };
                        double maxX = br.Position.X + 3 * cellWidth;
                        double minX = br.Position.X;
                        double maxY = br.Position.Y - i * cellHeight;
                        double minY = br.Position.Y - (i + 1) * cellHeight;

                        foreach (DBText dbt in classifiedResults[br])
                        {
                            if (dbt.Position.Y >= minY && dbt.Position.Y <= maxY)
                            {
                                if (dbt.Position.X >= minX && dbt.Position.X <= minX + cellWidth && result[0] == "无字符串")
                                {
                                    result[0] = dbt.TextString.ToString();
                                    count++;
                                }
                                else if (dbt.Position.X >= minX + cellWidth && dbt.Position.X <= minX + 2 * cellWidth && result[1] == "无字符串")
                                {
                                    result[1] = dbt.TextString.ToString();
                                    count++;
                                }
                                else if (dbt.Position.X >= minX + 2 * cellWidth && dbt.Position.X <= minX + 3 * cellWidth && result[2] == "无字符串")
                                {
                                    result[2] = dbt.TextString.ToString();
                                    count++;
                                }
                            }
                            if (count == 3) break;
                        }
                        smallUsage.Add(result[0]);
                        smallBox.Add(result[1]);
                        smallData.Add(data);
                    }

                }
                foreach (BlockReference br in small3Tables)
                {
                    BlockTableRecord btr = trans.GetObject(br.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;

                    // 读取table
                    double[] data = new double[5];
                    string[] fruits = btr.Name.Split('|'); // 按|分割

                    bool found = false;
                    Table table = new Table();
                    foreach (Table t in dbTables)
                    {
                        // 检查并读取 XData
                        if (t != null)
                        {
                            ResultBuffer xdata = t.GetXDataForApplication("CADAUTOTABLE");
                            if (xdata != null)
                            {
                                foreach (var value in xdata)
                                {
                                    if (value.Value.ToString().Contains("表格编号:"))
                                    {
                                        if (value.Value.ToString().Substring(5) == fruits[1].ToString())
                                        {
                                            found = true;
                                            // 处理找到的XData
                                            ed.WriteMessage($"找到 XData: {value.Value}");
                                            table = t;
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                    }

                    if (found)
                    {
                        data[0] = double.Parse(table.Cells[1, 2].Value.ToString());
                        data[1] = double.Parse(table.Cells[2, 2].Value.ToString());
                        data[2] = data[0] * data[1];
                        data[3] = double.Parse(table.Cells[4, 2].Value.ToString());
                        data[4] = data[2] / 0.38 / data[3] / 1.732;
                    }
                    else
                    {
                        data[0] = 0;
                        data[1] = 0;
                        data[2] = 0;
                        data[3] = 0;
                        data[4] = 0;
                    }

                    int rowCount = GetRowCount(btr);
                    for (int i = 1; i < rowCount; i++)
                    {
                        int count = 0;
                        string[] result = { "无字符串", "无字符串", "无字符串" };
                        double maxX = br.Position.X + 3 * cellWidth;
                        double minX = br.Position.X;
                        double maxY = br.Position.Y - i * cellHeight;
                        double minY = br.Position.Y - (i + 1) * cellHeight;

                        foreach (DBText dbt in classifiedResults[br])
                        {
                            if (dbt.Position.Y >= minY && dbt.Position.Y <= maxY)
                            {
                                if (dbt.Position.X >= minX && dbt.Position.X <= minX + cellWidth && result[0] == "无字符串")
                                {
                                    result[0] = dbt.TextString.ToString();
                                    count++;
                                }
                                else if (dbt.Position.X >= minX + cellWidth && dbt.Position.X <= minX + 2 * cellWidth && result[1] == "无字符串")
                                {
                                    result[1] = dbt.TextString.ToString();
                                    count++;
                                }
                                else if (dbt.Position.X >= minX + 2 * cellWidth && dbt.Position.X <= minX + 3 * cellWidth && result[2] == "无字符串")
                                {
                                    result[2] = dbt.TextString.ToString();
                                    count++;
                                }
                            }
                            if (count == 3) break;
                        }
                        smallUsage.Add(result[0]);
                        smallBox.Add(result[1]);
                        smallData.Add(data);
                    }
                }
                foreach (BlockReference br in bigTables)
                {
                    int count = 0;
                    string[] result = { "无字符串", "无字符串" };

                    double maxY1 = br.Position.Y - 9 * cellHeightBig;
                    double minY1 = br.Position.Y - 10 * cellHeightBig;
                    double maxY2 = br.Position.Y - 11 * cellHeightBig;
                    double minY2 = br.Position.Y - 12 * cellHeightBig;

                    foreach (DBText dbt in classifiedResults[br])
                    {
                        if (dbt.Position.Y >= minY1 && dbt.Position.Y <= maxY1)
                        {
                            result[0] = dbt.TextString.ToString();
                            count++;
                        }
                        else if (dbt.Position.Y >= minY2 && dbt.Position.Y <= maxY2)
                        {
                            result[1] = dbt.TextString.ToString();
                            count++;
                        }
                        if (count == 2) break;
                    }
                    bigLoop.Add(result[0]);
                    bigBox.Add(result[1]);
                }

                trans.Commit();
            }

            var dataToSave = new
            {
                SmallBox = smallBox,
                SmallData = smallData,
                SmallUsage = smallUsage,
                BigBox = bigBox,
                BigLoop = bigLoop
            };

            // 获取当前用户的 Documents 文件夹路径
            string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            // 目标文件夹路径
            string path = Path.Combine(documentsPath, "CADAUTOTABLE");

            // 检测目录是否存在
            if (!Directory.Exists(path))
            {
                // 如果不存在则创建目录
                Directory.CreateDirectory(path);
            }
            string jsonString = JsonConvert.SerializeObject(dataToSave, Formatting.Indented);
            File.WriteAllText(path + "\\CADAUTOFILL.json", jsonString);
        }

        // 写入大表小表
        [CommandMethod("WT")]
        public void WriteTables()
        {
            // 获取当前用户的 Documents 文件夹路径
            string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            // 目标文件夹路径
            string path = Path.Combine(documentsPath, "CADAUTOTABLE");

            // 检测目录是否存在
            if (!Directory.Exists(path))
            {
                // 如果不存在则创建目录
                Directory.CreateDirectory(path);
            }

            // 读出本地的数据
            {
                string jsonString = File.ReadAllText(path + "\\CADAUTOFILL.json");
                var loadedData = JsonConvert.DeserializeObject<dynamic>(jsonString);

                smallBox = loadedData.SmallBox.ToObject<List<string>>();
                smallUsage = loadedData.SmallUsage.ToObject<List<string>>();
                smallData = loadedData.SmallData.ToObject<List<double[]>>();
                bigBox = loadedData.BigBox.ToObject<List<string>>();
                bigLoop = loadedData.BigLoop.ToObject<List<string>>();
            }

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            // 提示用户选择对象
            PromptSelectionOptions opts = new PromptSelectionOptions();
            opts.MessageForAdding = "\n选择要写入的对象: ";
            PromptSelectionResult res = ed.GetSelection(opts);

            if (res.Status != PromptStatus.OK)
                return;

            SelectionSet selSet = res.Value;

            // 提前保存各类需要写入的表
            List<BlockReference> small2Tables = new List<BlockReference>();
            List<BlockReference> small3Tables = new List<BlockReference>();
            List<BlockReference> small4Tables = new List<BlockReference>();
            List<BlockReference> smallVerticalTables = new List<BlockReference>();
            List<BlockReference> bigTables = new List<BlockReference>();
            List<DBText> dbTexts = new List<DBText>();
            List<Table> dbTables = new List<Table>();

            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                // 预先读取数据保存到list
                {
                    foreach (SelectedObject selObj in selSet)
                    {
                        if (selObj == null) continue;

                        // 尝试获取选中对象的 BlockReference
                        BlockReference blockRef = trans.GetObject(selObj.ObjectId, OpenMode.ForWrite) as BlockReference;
                        if (blockRef != null)
                        {
                            // 获取与 BlockReference 相关联的 BlockTableRecord
                            BlockTableRecord btr = trans.GetObject(blockRef.BlockTableRecord, OpenMode.ForWrite) as BlockTableRecord;
                            if (btr != null)
                            {
                                ed.WriteMessage($"\n找到块: {btr.Handle}");
                                // 检查 BlockTableRecord 的名称是否包含 "Small"
                                if (btr.Name.Contains("Small2"))
                                {
                                    small2Tables.Add(blockRef);
                                    ed.WriteMessage($"\n找到块: {btr.Name}");
                                }
                                else if (btr.Name.Contains("Small3"))
                                {
                                    small3Tables.Add(blockRef);
                                    ed.WriteMessage($"\n找到块: {btr.Name}");
                                }
                                else if (btr.Name.Contains("Big"))
                                {
                                    bigTables.Add(blockRef);
                                    ed.WriteMessage($"\n找到块: {btr.Name}");
                                }
                                else if (btr.Name.Contains("PMT"))
                                {
                                    small4Tables.Add(blockRef);
                                    ed.WriteMessage($"\n找到块: {btr.Name}");
                                }
                                else if (btr.Name.Contains("Vertical"))
                                {
                                    smallVerticalTables.Add(blockRef);
                                    ed.WriteMessage($"\n找到块: {btr.Name}");
                                }
                            }
                        }

                        // 尝试获取选中对象的 DBText
                        DBText dbtext = trans.GetObject(selObj.ObjectId, OpenMode.ForWrite) as DBText;
                        if (dbtext != null)
                        {
                            dbTexts.Add(dbtext);
                        }

                        // 尝试获取选中对象的 Table
                        Table dbtable = trans.GetObject(selObj.ObjectId, OpenMode.ForWrite) as Table;
                        if (dbtable != null)
                        {
                            dbTables.Add(dbtable);
                        }
                    }
                }

                // 根据大小表把DBText进行划分，节省开销
                Dictionary<BlockReference, List<DBText>> classifiedResults = new Dictionary<BlockReference, List<DBText>>();
                Dictionary<BlockReference, Table> tableClassifiedResults = new Dictionary<BlockReference, Table>();
                {
                    // 初始化每个 BlockReference 的分类列表
                    foreach (BlockReference br in small2Tables) classifiedResults[br] = new List<DBText>();
                    foreach (BlockReference br in small3Tables) classifiedResults[br] = new List<DBText>();
                    foreach (BlockReference br in small4Tables) classifiedResults[br] = new List<DBText>();
                    foreach (BlockReference br in smallVerticalTables) classifiedResults[br] = new List<DBText>();
                    foreach (BlockReference br in bigTables) classifiedResults[br] = new List<DBText>();

                    // 遍历每一个 DBText 对象
                    foreach (DBText dbt in dbTexts)
                    {
                        // 遍历 BlockReference 列表，判断是否落在该表格内
                        foreach (BlockReference br in small2Tables)
                        {
                            BlockTableRecord btr = trans.GetObject(br.BlockTableRecord, OpenMode.ForWrite) as BlockTableRecord;
                            int rowCount = GetRowCount(btr);
                            double maxX = br.Position.X + 3 * cellWidth;
                            double minX = br.Position.X;
                            double maxY = br.Position.Y;
                            double minY = br.Position.Y - rowCount * cellHeight;
                            if (dbt.Position.X >= minX && dbt.Position.X <= maxX && dbt.Position.Y >= minY && dbt.Position.Y <= maxY)
                            {
                                // 将 DBText 添加到当前 DB 对象的分类列表中
                                classifiedResults[br].Add(dbt);
                            }
                        }
                        foreach (BlockReference br in small3Tables)
                        {
                            BlockTableRecord btr = trans.GetObject(br.BlockTableRecord, OpenMode.ForWrite) as BlockTableRecord;
                            int rowCount = GetRowCount(btr);
                            double maxX = br.Position.X + 3 * cellWidth;
                            double minX = br.Position.X;
                            double maxY = br.Position.Y;
                            double minY = br.Position.Y - rowCount * cellHeight;
                            if (dbt.Position.X >= minX && dbt.Position.X <= maxX && dbt.Position.Y >= minY && dbt.Position.Y <= maxY)
                            {
                                // 将 DBText 添加到当前 DB 对象的分类列表中
                                classifiedResults[br].Add(dbt);
                            }
                        }
                        foreach (BlockReference br in bigTables)
                        {
                            BlockTableRecord btr = trans.GetObject(br.BlockTableRecord, OpenMode.ForWrite) as BlockTableRecord;
                            double maxX = br.Position.X + cellWidthBig;
                            double minX = br.Position.X;
                            double maxY = br.Position.Y;
                            double minY = br.Position.Y - 19 * cellHeightBig;
                            if (dbt.Position.X >= minX && dbt.Position.X <= maxX && dbt.Position.Y >= minY && dbt.Position.Y <= maxY)
                            {
                                // 将 DBText 添加到当前 DB 对象的分类列表中
                                classifiedResults[br].Add(dbt);
                            }
                        }
                        foreach (BlockReference br in small4Tables)
                        {
                            BlockTableRecord btr = trans.GetObject(br.BlockTableRecord, OpenMode.ForWrite) as BlockTableRecord;
                            int rowCount = GetRowCount(btr);
                            double maxX = br.Position.X + 2 * cellWidth + cellWidth2;
                            double minX = br.Position.X;
                            double maxY = br.Position.Y;
                            double minY = br.Position.Y - rowCount * cellHeight;
                            if (dbt.Position.X >= minX && dbt.Position.X <= maxX && dbt.Position.Y >= minY && dbt.Position.Y <= maxY)
                            {
                                // 将 DBText 添加到当前 DB 对象的分类列表中
                                classifiedResults[br].Add(dbt);
                            }
                        }
                        foreach (BlockReference br in smallVerticalTables)
                        {
                            BlockTableRecord btr = trans.GetObject(br.BlockTableRecord, OpenMode.ForWrite) as BlockTableRecord;
                            double maxX = br.Position.X + cellHeight;
                            double minX = br.Position.X;
                            double maxY = br.Position.Y;
                            double minY = br.Position.Y - 2 * cellWidth;
                            if (dbt.Position.X >= minX && dbt.Position.X <= maxX && dbt.Position.Y >= minY && dbt.Position.Y <= maxY)
                            {
                                // 将 DBText 添加到当前 DB 对象的分类列表中
                                classifiedResults[br].Add(dbt);
                            }
                        }
                    }

                    foreach (BlockReference br in bigTables)
                    {
                        foreach (Table t in dbTables)
                        {
                            if (t != null)
                            {
                                BlockTableRecord btr = trans.GetObject(br.BlockTableRecord, OpenMode.ForWrite) as BlockTableRecord;
                                double maxX = br.Position.X + cellWidthBig * 3 / 4;
                                double minX = br.Position.X - cellWidthBig / 4;  // 向左1/4防止偏移
                                double maxY = br.Position.Y;
                                double minY = br.Position.Y - 19 * cellHeightBig;
                                if (t.Position.X >= minX && t.Position.X <= maxX && t.Position.Y >= minY && t.Position.Y <= maxY)
                                {
                                    // 将 DBText 添加到当前 DB 对象的分类列表中
                                    tableClassifiedResults[br] = t;
                                    break;
                                }
                            }
                        }
                    }
                }

                // 再在每个表中进行识别
                {
                    foreach (BlockReference br in small2Tables)
                    {
                        BlockTableRecord btr = trans.GetObject(br.BlockTableRecord, OpenMode.ForWrite) as BlockTableRecord;

                        int rowCount = GetRowCount(btr);
                        for (int i = 1; i < rowCount; i++)
                        {
                            int count = 0;
                            DBText text1 = null;
                            DBText text2 = null;
                            DBText text3 = null;

                            double minX = br.Position.X;
                            double maxY = br.Position.Y - i * cellHeight;
                            double minY = br.Position.Y - (i + 1) * cellHeight;

                            foreach (DBText dbt in classifiedResults[br])
                            {
                                if (dbt.Position.Y >= minY && dbt.Position.Y <= maxY)
                                {
                                    if (dbt.Position.X >= minX && dbt.Position.X <= minX + cellWidth && text1 == null)
                                    {
                                        text1 = dbt;
                                        count++;
                                    }
                                    else if (dbt.Position.X >= minX + cellWidth && dbt.Position.X <= minX + 2 * cellWidth && text2 == null)
                                    {
                                        text2 = dbt;
                                        count++;
                                    }
                                    else if (dbt.Position.X >= minX + 2 * cellWidth && dbt.Position.X <= minX + 3 * cellWidth && text3 == null)
                                    {
                                        text3 = dbt;
                                        count++;
                                    }
                                }
                                if (count == 3) break;
                            }

                            if (text2 != null && text3 != null)
                            {
                                for (int j = 0; j < bigBox.Count; j++)
                                {
                                    if (bigBox[j] == text2.TextString)
                                    {
                                        text3.TextString = bigLoop[j];
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    foreach (BlockReference br in small3Tables)
                    {
                        BlockTableRecord btr = trans.GetObject(br.BlockTableRecord, OpenMode.ForWrite) as BlockTableRecord;

                        int rowCount = GetRowCount(btr);
                        for (int i = 1; i < rowCount; i++)
                        {
                            int count = 0;
                            DBText text1 = null;
                            DBText text2 = null;
                            DBText text3 = null;

                            double minX = br.Position.X;
                            double maxY = br.Position.Y - i * cellHeight;
                            double minY = br.Position.Y - (i + 1) * cellHeight;

                            foreach (DBText dbt in classifiedResults[br])
                            {
                                if (dbt.Position.Y >= minY && dbt.Position.Y <= maxY)
                                {
                                    if (dbt.Position.X >= minX && dbt.Position.X <= minX + cellWidth && text1 == null)
                                    {
                                        text1 = dbt;
                                        count++;
                                    }
                                    else if (dbt.Position.X >= minX + cellWidth && dbt.Position.X <= minX + 2 * cellWidth && text2 == null)
                                    {
                                        text2 = dbt;
                                        count++;
                                    }
                                    else if (dbt.Position.X >= minX + 2 * cellWidth && dbt.Position.X <= minX + 3 * cellWidth && text3 == null)
                                    {
                                        text3 = dbt;
                                        count++;
                                    }
                                }
                                if (count == 3) break;
                            }

                            if (text2 != null && text3 != null)
                            {
                                for (int j = 0; j < bigBox.Count; j++)
                                {
                                    if (bigBox[j] == text2.TextString)
                                    {
                                        text3.TextString = bigLoop[j];
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    foreach (BlockReference br in small4Tables)
                    {
                        BlockTableRecord btr = trans.GetObject(br.BlockTableRecord, OpenMode.ForWrite) as BlockTableRecord;

                        int rowCount = GetRowCount(btr);
                        for (int i = 0; i < rowCount; i++)
                        {
                            int count = 0;
                            DBText text1 = null;
                            DBText text2 = null;

                            double maxX1 = br.Position.X + cellWidth + cellWidth2;
                            double minX1 = br.Position.X + cellWidth2;
                            double maxX2 = br.Position.X + 2 * cellWidth + cellWidth2;
                            double minX2 = br.Position.X + cellWidth + cellWidth2;
                            double maxY = br.Position.Y - i * cellHeight;
                            double minY = br.Position.Y - (i + 1) * cellHeight;

                            foreach (DBText dbt in classifiedResults[br])
                            {
                                if (dbt.Position.Y >= minY && dbt.Position.Y <= maxY)
                                {
                                    if (dbt.Position.X >= minX1 && dbt.Position.X <= maxX1 && text1 == null)
                                    {
                                        ed.WriteMessage($"\n找到text1");
                                        text1 = dbt;
                                        count++;
                                    }
                                    else if (dbt.Position.X >= minX2 && dbt.Position.X <= maxX2 && text2 == null)
                                    {
                                        ed.WriteMessage($"\n找到text2");
                                        text2 = dbt;
                                        count++;
                                    }
                                }
                                if (count == 2) break;
                            }

                            if (text1 != null && text2 != null)
                            {
                                for (int j = 0; j < bigBox.Count; j++)
                                {

                                    ed.WriteMessage($"\n对比{bigBox[j]}和{text1.TextString}");
                                    if (bigBox[j] == text1.TextString)
                                    {
                                        text2.TextString = bigLoop[j];
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    foreach (BlockReference br in smallVerticalTables)
                    {
                        BlockTableRecord btr = trans.GetObject(br.BlockTableRecord, OpenMode.ForWrite) as BlockTableRecord;

                        DBText text1 = null;
                        DBText text2 = null;

                        int count = 0;
                        double maxY1 = br.Position.Y;
                        double minY1 = br.Position.Y - cellWidth;
                        double maxY2 = br.Position.Y - cellWidth;
                        double minY2 = br.Position.Y - 2 * cellWidth;

                        foreach (DBText dbt in classifiedResults[br])
                        {
                            if (dbt.Position.Y >= minY1 && dbt.Position.Y <= maxY1 && text1 == null)
                            {
                                text1 = dbt;
                                count++;
                            }
                            else if (dbt.Position.Y >= minY2 && dbt.Position.Y <= maxY2 && text2 == null)
                            {
                                text2 = dbt;
                                count++;
                            }
                            if (count == 2) break;
                        }

                        if (text1 != null && text2 != null)
                        {
                            for (int j = 0; j < bigBox.Count; j++)
                            {
                                if (bigBox[j] == text2.TextString)
                                {
                                    text1.TextString = bigLoop[j];
                                    break;
                                }
                            }
                        }

                    }
                    foreach (BlockReference br in bigTables)
                    {
                        int count = 0;
                        DBText[] texts = { null, null, null, null, null, null, null };

                        double[] maxYs = { 0, 1, 3, 9, 10, 11, 18 };

                        foreach (DBText dbt in classifiedResults[br])
                        {
                            for (int j = 0; j < maxYs.Length; j++)
                            {
                                if (dbt.Position.Y >= br.Position.Y - (maxYs[j] + 1) * cellHeightBig && dbt.Position.Y <= br.Position.Y - maxYs[j] * cellHeightBig)
                                {
                                    texts[j] = dbt;
                                    count++;
                                }
                            }
                            if (count == 7) break;
                        }

                        if (texts[5] != null)
                        {
                            for (int j = 0; j < smallBox.Count; j++)
                            {
                                if (smallBox[j] == texts[5].TextString)
                                {
                                    if (texts[4] != null)
                                    {
                                        texts[4].TextString = smallUsage[j];
                                    }

                                    if (tableClassifiedResults.ContainsKey(br))
                                    {
                                        Table table = tableClassifiedResults[br];

                                        if (table.Cells[0, 0].Value.ToString() != smallData[j][0].ToString() || table.Cells[1, 0].Value.ToString() != smallData[j][1].ToString() || table.Cells[3, 0].Value.ToString() != smallData[j][3].ToString())
                                        {
                                            table.Cells[0, 0].Value = smallData[j][0].ToString();
                                            table.Cells[1, 0].Value = smallData[j][1].ToString();
                                            table.Cells[3, 0].Value = smallData[j][3].ToString();

                                            table.Cells[2, 0].DataFormat = "%lu2%pr2";
                                            table.Cells[4, 0].DataFormat = "%lu2%pr2";

                                            double current = smallData[j][0] * smallData[j][1] / 0.38 / smallData[j][3] / 1.732;

                                            string[] currents = GetCurrentAndCable((float)current);

                                            if (texts[0] != null)
                                            {
                                                texts[0].TextString = currents[0];
                                            }

                                            if (texts[6] != null)
                                            {
                                                // 电缆写入
                                                texts[6].TextString = currents[3];
                                            }

                                            if (texts[1] != null)
                                            {
                                                // 保护写入
                                                texts[1].TextString = currents[1];
                                            }

                                            if (texts[2] != null)
                                            {
                                                // 瞬时保护写入
                                                texts[2].TextString = currents[2];
                                            }
                                        }
                                    }
                                    break;
                                }
                            }
                        }
                    }
                }

                trans.Commit();
            }
        }

        // 读取大表小表（老版本）
        [CommandMethod("RT2")]
        public void ReadTables2()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            // 提示用户选择对象
            PromptSelectionOptions opts = new PromptSelectionOptions();
            opts.MessageForAdding = "\n选择要读取的对象: ";
            PromptSelectionResult res = ed.GetSelection(opts);

            if (res.Status != PromptStatus.OK)
                return;

            SelectionSet selSet = res.Value;

            List<string> small2Tables = new List<string>();
            List<string> small3Tables = new List<string>();
            List<string> bigTables = new List<string>();
            List<string> dbTexts = new List<string>();

            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                foreach (SelectedObject selObj in selSet)
                {
                    if (selObj == null) continue;

                    // 尝试获取选中对象的 BlockReference
                    BlockReference blockRef = trans.GetObject(selObj.ObjectId, OpenMode.ForRead) as BlockReference;
                    if (blockRef != null)
                    {
                        // 获取与 BlockReference 相关联的 BlockTableRecord
                        BlockTableRecord btr = trans.GetObject(blockRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                        if (btr != null)
                        {
                            ed.WriteMessage($"\n找到块: {btr.Handle}");
                            // 检查 BlockTableRecord 的名称是否包含 "Small"
                            if (btr.Name.Contains("Small2"))
                            {
                                small2Tables.Add(btr.Name);
                                ed.WriteMessage($"\n找到块: {btr.Name}");
                            }
                            else if (btr.Name.Contains("Small3"))
                            {
                                small3Tables.Add(btr.Name);
                                ed.WriteMessage($"\n找到块: {btr.Name}");
                            }
                            else if (btr.Name.Contains("Big"))
                            {
                                bigTables.Add(btr.Name);
                                ed.WriteMessage($"\n找到块: {btr.Name}");
                            }
                        }
                    }
                }
                Handle handle = new Handle();
                trans.Commit();
            }

            smallBox = new List<string>();
            smallUsage = new List<string>();
            smallData = new List<double[]>();

            bigBox = new List<string>();
            bigLoop = new List<string>();

            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                foreach (string combineId in small2Tables)
                {
                    double[] data = new double[5];
                    string[] fruits = combineId.Split('|'); // 按|分割

                    try
                    {
                        //读取数据
                        Handle handle = new Handle((long)Convert.ToUInt64(fruits[1], 16));
                        ObjectId dataId = db.GetObjectId(false, handle, 0);
                        Table table = trans.GetObject(dataId, OpenMode.ForRead) as Table;
                        data[0] = double.Parse(table.Cells[0, 1].Value.ToString());
                        data[1] = double.Parse(table.Cells[1, 1].Value.ToString());
                        data[2] = data[0] * data[1];
                        data[3] = double.Parse(table.Cells[3, 1].Value.ToString());
                        data[4] = data[2] / 0.38 / data[3] / 1.732;
                    }
                    catch (Autodesk.AutoCAD.Runtime.Exception ex) { }

                    //读取箱号
                    for (int i = 2; i < fruits.Length; i++)
                    {
                        try
                        {
                            Handle handle2 = new Handle((long)Convert.ToUInt64(fruits[i].Split('_')[0], 16));
                            ObjectId boxId = db.GetObjectId(false, handle2, 0);//箱号id

                            Handle handle3 = new Handle((long)Convert.ToUInt64(fruits[i].Split('_')[2], 16));
                            ObjectId usageId = db.GetObjectId(false, handle3, 0);//箱号id

                            DBText text = trans.GetObject(boxId, OpenMode.ForRead) as DBText;
                            DBText usageText = trans.GetObject(usageId, OpenMode.ForRead) as DBText;

                            smallBox.Add(text.TextString);
                            smallUsage.Add(usageText.TextString);
                            smallData.Add(data);
                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception ex) { }
                    }
                }

                foreach (string combineId in small3Tables)
                {
                    double[] data = new double[5];
                    string[] fruits = combineId.Split('|'); // 按|分割

                    try
                    {
                        //读取数据
                        Handle handle = new Handle((long)Convert.ToUInt64(fruits[1], 16));
                        ObjectId dataId = db.GetObjectId(false, handle, 0);
                        Table table = trans.GetObject(dataId, OpenMode.ForRead) as Table;
                        data[0] = double.Parse(table.Cells[1, 2].Value.ToString());
                        data[1] = double.Parse(table.Cells[2, 2].Value.ToString());
                        data[2] = data[0] * data[1];
                        data[3] = double.Parse(table.Cells[4, 2].Value.ToString());
                        data[4] = data[2] / 0.38 / data[3] / 1.732;
                    }
                    catch (Autodesk.AutoCAD.Runtime.Exception ex) { }

                    //读取箱号
                    for (int i = 2; i < fruits.Length; i++)
                    {
                        try
                        {
                            Handle handle2 = new Handle((long)Convert.ToUInt64(fruits[i].Split('_')[0], 16));
                            ObjectId boxId = db.GetObjectId(false, handle2, 0);//箱号id

                            Handle handle3 = new Handle((long)Convert.ToUInt64(fruits[i].Split('_')[2], 16));
                            ObjectId usageId = db.GetObjectId(false, handle3, 0);//箱号id

                            DBText text = trans.GetObject(boxId, OpenMode.ForRead) as DBText;
                            DBText usageText = trans.GetObject(usageId, OpenMode.ForRead) as DBText;

                            smallBox.Add(text.TextString);
                            smallUsage.Add(usageText.TextString);
                            smallData.Add(data);
                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception ex) { }
                    }
                }

                foreach (string combineId in bigTables)
                {
                    string[] fruits = combineId.Split('|'); // 按|分割

                    //读取箱号
                    for (int i = 2; i < fruits.Length; i++)
                    {
                        try
                        {
                            Handle handle = new Handle((long)Convert.ToUInt64(fruits[i].Split('_')[0], 16));
                            ObjectId boxId = db.GetObjectId(false, handle, 0);//箱号id
                            DBText boxText = trans.GetObject(boxId, OpenMode.ForRead) as DBText;
                            bigBox.Add(boxText.TextString);
                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception ex)
                        {
                            bigBox.Add("无");
                        }

                        try
                        {
                            Handle handle = new Handle((long)Convert.ToUInt64(fruits[i].Split('_')[1], 16));
                            ObjectId loopId = db.GetObjectId(false, handle, 0);//回路id
                            DBText loopText = trans.GetObject(loopId, OpenMode.ForRead) as DBText;
                            bigLoop.Add(loopText.TextString);
                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception ex)
                        {
                            bigLoop.Add("无");
                        }
                    }
                }


                trans.Commit();
            }

            var dataToSave = new
            {
                SmallBox = smallBox,
                SmallData = smallData,
                SmallUsage = smallUsage,
                BigBox = bigBox,
                BigLoop = bigLoop
            };

            // 获取当前用户的 Documents 文件夹路径
            string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            // 目标文件夹路径
            string path = Path.Combine(documentsPath, "CADAUTOTABLE");

            // 检测目录是否存在
            if (!Directory.Exists(path))
            {
                // 如果不存在则创建目录
                Directory.CreateDirectory(path);
            }
            string jsonString = JsonConvert.SerializeObject(dataToSave, Formatting.Indented);
            File.WriteAllText(path + "\\CADAUTOFILL.json", jsonString);
        }

        // 写入大表小表（老版本）
        [CommandMethod("WT2")]
        public void WriteTables2()
        {
            // 获取当前用户的 Documents 文件夹路径
            string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            // 目标文件夹路径
            string path = Path.Combine(documentsPath, "CADAUTOTABLE");

            // 检测目录是否存在
            if (!Directory.Exists(path))
            {
                // 如果不存在则创建目录
                Directory.CreateDirectory(path);
            }
            try
            {
                string jsonString = File.ReadAllText(path + "\\CADAUTOFILL.json");
                var loadedData = JsonConvert.DeserializeObject<dynamic>(jsonString);

                smallBox = loadedData.SmallBox.ToObject<List<string>>();
                smallUsage = loadedData.SmallUsage.ToObject<List<string>>();
                smallData = loadedData.SmallData.ToObject<List<double[]>>();
                bigBox = loadedData.BigBox.ToObject<List<string>>();
                bigLoop = loadedData.BigLoop.ToObject<List<string>>();


                Document doc = Application.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;
                Database db = doc.Database;

                // 提示用户选择对象
                PromptSelectionOptions opts = new PromptSelectionOptions();
                opts.MessageForAdding = "\n选择要写入的对象: ";
                PromptSelectionResult res = ed.GetSelection(opts);

                if (res.Status != PromptStatus.OK)
                    return;

                SelectionSet selSet = res.Value;

                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    foreach (SelectedObject selObj in selSet)
                    {
                        if (selObj == null) continue;

                        // 尝试获取选中对象的 BlockReference
                        BlockReference blockRef = trans.GetObject(selObj.ObjectId, OpenMode.ForRead) as BlockReference;
                        if (blockRef != null)
                        {
                            // 获取与 BlockReference 相关联的 BlockTableRecord
                            BlockTableRecord btr = trans.GetObject(blockRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                            if (btr != null)
                            {
                                // 检查 BlockTableRecord 的名称是否包含 "Small"
                                if (btr.Name.Contains("Small"))
                                {
                                    string[] fruits = btr.Name.Split('|');
                                    for (int i = 2; i < fruits.Length; i++)
                                    {
                                        try
                                        {
                                            Handle handle = new Handle((long)Convert.ToUInt64(fruits[i].Split('_')[0], 16));
                                            ObjectId boxId = db.GetObjectId(false, handle, 0);//箱号id
                                            DBText text = trans.GetObject(boxId, OpenMode.ForRead) as DBText;

                                            for (int j = 0; j < bigBox.Count; j++)
                                            {
                                                if (bigBox[j] == text.TextString)
                                                {
                                                    try
                                                    {

                                                        Handle handle2 = new Handle((long)Convert.ToUInt64(fruits[i].Split('_')[1], 16));
                                                        ObjectId loopId = db.GetObjectId(false, handle2, 0);//回路id
                                                        DBText text2 = trans.GetObject(loopId, OpenMode.ForWrite) as DBText;
                                                        text2.TextString = bigLoop[j];
                                                    }
                                                    catch (Autodesk.AutoCAD.Runtime.Exception ex) { }
                                                }
                                            }
                                            ed.WriteMessage($"\n写入块: {btr.Name}");
                                        }
                                        catch (Autodesk.AutoCAD.Runtime.Exception ex) { }
                                    }
                                }
                                else if (btr.Name.Contains("PMT"))
                                {
                                    string[] fruits = btr.Name.Split('|');
                                    for (int i = 2; i < fruits.Length; i++)
                                    {
                                        try
                                        {
                                            Handle handle = new Handle((long)Convert.ToUInt64(fruits[i].Split('_')[0], 16));
                                            ObjectId boxId = db.GetObjectId(false, handle, 0);//箱号id
                                            DBText text = trans.GetObject(boxId, OpenMode.ForRead) as DBText;

                                            for (int j = 0; j < bigBox.Count; j++)
                                            {
                                                if (bigBox[j] == text.TextString)
                                                {
                                                    try
                                                    {
                                                        Handle handle2 = new Handle((long)Convert.ToUInt64(fruits[i].Split('_')[1], 16));
                                                        ObjectId loopId = db.GetObjectId(false, handle2, 0);//回路id
                                                        DBText text2 = trans.GetObject(loopId, OpenMode.ForWrite) as DBText;
                                                        text2.TextString = bigLoop[j];
                                                    }
                                                    catch (Autodesk.AutoCAD.Runtime.Exception ex) { }
                                                }
                                            }
                                            ed.WriteMessage($"\n写入块: {btr.Name}");
                                        }
                                        catch (Autodesk.AutoCAD.Runtime.Exception ex) { }
                                    }
                                }
                                else if (btr.Name.Contains("Big"))
                                {
                                    string[] fruits = btr.Name.Split('|');
                                    try
                                    {
                                        Handle handle = new Handle((long)Convert.ToUInt64(fruits[2].Split('_')[0], 16));
                                        ObjectId boxId = db.GetObjectId(false, handle, 0);//箱号id
                                        DBText text = trans.GetObject(boxId, OpenMode.ForRead) as DBText;
                                        for (int j = 0; j < smallBox.Count; j++)
                                        {
                                            if (smallBox[j] == text.TextString)
                                            {
                                                try
                                                {
                                                    // 用途写入
                                                    Handle handle3 = new Handle((long)Convert.ToUInt64(fruits[2].Split('_')[2], 16));
                                                    ObjectId usageId = db.GetObjectId(false, handle3, 0);//用途id
                                                    DBText usageText = trans.GetObject(usageId, OpenMode.ForWrite) as DBText;
                                                    usageText.TextString = smallUsage[j];
                                                }
                                                catch (Autodesk.AutoCAD.Runtime.Exception ex) { }

                                                try
                                                {
                                                    // 表格写入
                                                    Handle handle2 = new Handle((long)Convert.ToUInt64(fruits[1], 16));
                                                    ObjectId tableId = db.GetObjectId(false, handle2, 0);//大表的表格id
                                                    Table table = trans.GetObject(tableId, OpenMode.ForWrite) as Table;
                                                    table.Cells[0, 0].Value = smallData[j][0].ToString();
                                                    table.Cells[1, 0].Value = smallData[j][1].ToString();
                                                    table.Cells[3, 0].Value = smallData[j][3].ToString();

                                                    table.Cells[2, 0].DataFormat = "%lu2%pr2";
                                                    table.Cells[4, 0].DataFormat = "%lu2%pr2";

                                                    double current = smallData[j][0] * smallData[j][1] / 0.38 / smallData[j][3] / 1.732;

                                                    string[] currents = GetCurrentAndCable((float)current);

                                                    try
                                                    {
                                                        // 电流写入
                                                        Handle handle4 = new Handle((long)Convert.ToUInt64(fruits[2].Split('_')[3], 16));
                                                        ObjectId currentId = db.GetObjectId(false, handle4, 0);//用途id
                                                        DBText currentText = trans.GetObject(currentId, OpenMode.ForWrite) as DBText;
                                                        currentText.TextString = currents[0];
                                                    }
                                                    catch (Autodesk.AutoCAD.Runtime.Exception ex) { }
                                                    try
                                                    {
                                                        // 电缆写入
                                                        Handle handle5 = new Handle((long)Convert.ToUInt64(fruits[2].Split('_')[4], 16));
                                                        ObjectId cableId = db.GetObjectId(false, handle5, 0);//用途id
                                                        DBText cableText = trans.GetObject(cableId, OpenMode.ForWrite) as DBText;
                                                        cableText.TextString = currents[3];
                                                    }
                                                    catch (Autodesk.AutoCAD.Runtime.Exception ex) { }

                                                    try
                                                    {
                                                        // 保护写入
                                                        Handle handle6 = new Handle((long)Convert.ToUInt64(fruits[2].Split('_')[5], 16));
                                                        ObjectId protectId = db.GetObjectId(false, handle6, 0);//用途id
                                                        DBText protectText = trans.GetObject(protectId, OpenMode.ForWrite) as DBText;
                                                        protectText.TextString = currents[1];
                                                    }
                                                    catch (Autodesk.AutoCAD.Runtime.Exception ex) { }

                                                    try
                                                    {
                                                        // 延时保护写入
                                                        Handle handle7 = new Handle((long)Convert.ToUInt64(fruits[2].Split('_')[6], 16));
                                                        ObjectId delayProtectId = db.GetObjectId(false, handle7, 0);//用途id
                                                        DBText delayProtectText = trans.GetObject(delayProtectId, OpenMode.ForWrite) as DBText;
                                                        delayProtectText.TextString = currents[2];
                                                    }
                                                    catch (Autodesk.AutoCAD.Runtime.Exception ex) { }
                                                }
                                                catch (Autodesk.AutoCAD.Runtime.Exception ex) { }
                                            }
                                        }
                                        ed.WriteMessage($"\n写入块: {btr.Name}");
                                    }
                                    catch (Autodesk.AutoCAD.Runtime.Exception ex) { }
                                }
                                else if (btr.Name.Contains("Vertical"))
                                {
                                    string[] fruits = btr.Name.Split('|');
                                    try
                                    {
                                        Handle handle = new Handle((long)Convert.ToUInt64(fruits[1].Split('_')[0], 16));
                                        ObjectId boxId = db.GetObjectId(false, handle, 0);//箱号id
                                        DBText text = trans.GetObject(boxId, OpenMode.ForRead) as DBText;

                                        for (int j = 0; j < bigBox.Count; j++)
                                        {
                                            if (bigBox[j] == text.TextString)
                                            {
                                                Handle handle2 = new Handle((long)Convert.ToUInt64(fruits[1].Split('_')[1], 16));
                                                ObjectId loopId = db.GetObjectId(false, handle2, 0);//回路id
                                                if (loopId != null)
                                                {
                                                    DBText text2 = trans.GetObject(loopId, OpenMode.ForWrite) as DBText;
                                                    text2.TextString = bigLoop[j];
                                                }
                                            }
                                        }
                                        ed.WriteMessage($"\n写入块: {btr.Name}");
                                    }
                                    catch (Autodesk.AutoCAD.Runtime.Exception ex) { }
                                }
                            }
                        }
                    }
                    trans.Commit();
                }
            }
            catch
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;
                ed.WriteMessage("错误：未找到对应表");
            }

        }

        [CommandMethod("RELOADCABLE")]
        public void ReloadCable()
        {
            // 获取当前用户的 Documents 文件夹路径
            string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            // 目标文件夹路径
            string path = Path.Combine(documentsPath, "CADAUTOTABLE");

            // 检测目录是否存在
            if (!Directory.Exists(path))
            {
                // 如果不存在则创建目录
                Directory.CreateDirectory(path);
            }
            try
            {
                string jsonString = File.ReadAllText(path + "\\电缆对应表.json");

                // 将 JSON 字符串反序列化为 Dictionary<string, string>
                cableList = JsonConvert.DeserializeObject<Dictionary<int, string>>(jsonString);
                sortedKeys = cableList.Keys.OrderBy(k => k).ToArray();
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;
                ed.WriteMessage("\n电缆表已重载");
            }
            catch
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;
                ed.WriteMessage("\n错误：未找到电缆表");
            }
        }

        private string[] GetCurrentAndCable(float number5)
        {
            if (firstRun)
            {
                // 获取当前用户的 Documents 文件夹路径
                string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

                // 目标文件夹路径
                string path = Path.Combine(documentsPath, "CADAUTOTABLE");

                // 检测目录是否存在
                if (!Directory.Exists(path))
                {
                    // 如果不存在则创建目录
                    Directory.CreateDirectory(path);
                }
                // 读取 JSON 文件内容
                string jsonString = File.ReadAllText(path + "\\电缆对应表.json");

                // 将 JSON 字符串反序列化为 Dictionary<string, string>
                cableList = JsonConvert.DeserializeObject<Dictionary<int, string>>(jsonString);
                sortedKeys = cableList.Keys.OrderBy(k => k).ToArray();
                firstRun = false;
            }

            float newNumber = number5 * 1.2f;
            int[] currentList = sortedKeys;
            string current1 = "未匹配";
            string current2 = "未匹配";
            string current3 = "未匹配";
            string cable = "未找到合适电缆";
            foreach (int i in currentList)
            {
                if (newNumber < i)
                {
                    current1 = $"{i}A";
                    current2 = $"{i}A";
                    current3 = $"{i * 8}A";

                    // 如果直接访问不存在的键会抛出异常
                    try
                    {
                        cable = cableList[i];
                    }
                    catch (KeyNotFoundException)
                    {
                        cable = "未找到合适电缆";
                    }

                    break;
                }
            }

            string[] returnCurrents = { current1, current2, current3, cable };

            return returnCurrents;
        }

        //—————————————以下为创建表命令实现————————————————

        //创建大表
        [CommandMethod("IT1")]
        public void InsertCustomTable1()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            while (true)
            {
                // 提示用户选择插入表格的位置
                PromptPointOptions ppo = new PromptPointOptions("选择插入表格的位置 (或按 Enter 结束): ");
                ppo.AllowNone = true;

                PromptPointResult ppr = ed.GetPoint(ppo);
                if (ppr.Status == PromptStatus.None)
                {
                    ed.WriteMessage("\n插入操作已结束。");
                    break;  // 用户按 Enter 键时退出循环
                }
                if (ppr.Status != PromptStatus.OK)
                {
                    ed.WriteMessage("\n结束");
                    break;
                }

                Point3d insertPoint = ppr.Value;

                // 插入表格
                string dataTableId = CreateDataTable1(db, new Point3d(insertPoint.X, insertPoint.Y - cellHeightBig * 12, insertPoint.Z));
                string contentId = CreateBigContent(db, new Point3d(insertPoint.X, insertPoint.Y - cellHeightBig * 9, insertPoint.Z));
                CreateBoxLoopTableBig(db, insertPoint, "Big|" + dataTableId + contentId);// 把数据表的objectid作为名字，即可单方面寻找
                ed.WriteMessage("\n表格已插入，继续选择插入点或按 Enter 结束。");
            }
        }

        [CommandMethod("IT2")]
        public void InsertCustomTable2()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            while (true)
            {

                // 提示用户选择插入表格的位置
                PromptPointOptions ppo = new PromptPointOptions("选择插入表格的位置 (或按 Enter 结束): ");
                ppo.AllowNone = true;

                PromptPointResult ppr = ed.GetPoint(ppo);
                if (ppr.Status == PromptStatus.None)
                {
                    ed.WriteMessage("\n插入操作已结束。");
                    break;  // 用户按 Enter 键时退出循环
                }
                if (ppr.Status != PromptStatus.OK)
                {
                    ed.WriteMessage("\n结束");
                    break;
                }

                Point3d insertPoint = ppr.Value;

                ReadSmallTableSign();
                smallTableSign++;
                WriteSmallTableSign();
                // 插入表格
                string dataTableId = CreateDataTable2(db, new Point3d(insertPoint.X, insertPoint.Y - 2000, insertPoint.Z));
                string contentId = CreateBoxLoopContent(db, 1, new Point3d(insertPoint.X, insertPoint.Y - cellHeight, insertPoint.Z));
                CreateBoxLoopTable(db, 2, insertPoint, "Small2|" + smallTableSign.ToString() + "|" + dataTableId + contentId);
                ed.WriteMessage("\n表格已插入，继续选择插入点或按 Enter 结束。");
            }
        }

        [CommandMethod("IT3")]
        public void InsertCustomTable3()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            while (true)
            {
                // 提示用户选择插入表格的位置
                PromptPointOptions ppo = new PromptPointOptions("选择插入表格的位置 (或按 Enter 结束): ");
                ppo.AllowNone = true;

                PromptPointResult ppr = ed.GetPoint(ppo);
                if (ppr.Status == PromptStatus.None)
                {
                    ed.WriteMessage("\n插入操作已结束。");
                    break;  // 用户按 Enter 键时退出循环
                }
                if (ppr.Status != PromptStatus.OK)
                {
                    ed.WriteMessage("\n结束");
                    break;
                }

                Point3d insertPoint = ppr.Value;


                ReadSmallTableSign();
                smallTableSign++;
                WriteSmallTableSign();

                // 插入表格
                string dataTableId = CreateDataTable3(db, new Point3d(insertPoint.X, insertPoint.Y - 2000, insertPoint.Z));
                string contentId = CreateBoxLoopContent(db, 1, new Point3d(insertPoint.X, insertPoint.Y - cellHeight, insertPoint.Z));
                CreateBoxLoopTable(db, 2, insertPoint, "Small3|" + smallTableSign.ToString() + "|" + dataTableId + contentId);// 把数据表的objectid作为名字，即可单方面寻找
                ed.WriteMessage("\n表格已插入，继续选择插入点或按 Enter 结束。");
            }
        }

        [CommandMethod("IT4")]
        public void InsertCustomTable4()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            while (true)
            {
                // 提示用户选择插入表格的位置
                PromptPointOptions ppo = new PromptPointOptions("选择插入表格的位置 (或按 Enter 结束): ");
                ppo.AllowNone = true;

                PromptPointResult ppr = ed.GetPoint(ppo);
                if (ppr.Status == PromptStatus.None)
                {
                    ed.WriteMessage("\n插入操作已结束。");
                    break;  // 用户按 Enter 键时退出循环
                }
                if (ppr.Status != PromptStatus.OK)
                {
                    ed.WriteMessage("\n结束");
                    break;
                }

                Point3d insertPoint = ppr.Value;

                // 插入表格
                string contentId = CreateBoxLoopContent2(db, 2, insertPoint);
                CreateBoxLoopTable2(db, 2, insertPoint, "PMT|" + contentId);// 把数据表的Handle作为名字，即可单方面寻找
                ed.WriteMessage("\n表格已插入，继续选择插入点或按 Enter 结束。");
            }
        }

        [CommandMethod("ITV")]
        public void InsertCustomTableVertical()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            while (true)
            {
                // 提示用户选择插入表格的位置
                PromptPointOptions ppo = new PromptPointOptions("选择插入表格的位置 (或按 Enter 结束): ");
                ppo.AllowNone = true;

                PromptPointResult ppr = ed.GetPoint(ppo);
                if (ppr.Status == PromptStatus.None)
                {
                    ed.WriteMessage("\n插入操作已结束。");
                    break;  // 用户按 Enter 键时退出循环
                }
                if (ppr.Status != PromptStatus.OK)
                {
                    ed.WriteMessage("\n结束");
                    break;
                }

                Point3d insertPoint = ppr.Value;

                // 插入表格
                Point3d rightup_pos = new Point3d(insertPoint.X, insertPoint.Y + cellWidth * 2, insertPoint.Z);
                string contentId = CreateVerticalContent(db, rightup_pos);
                CreateVerticalTable(db, 0, rightup_pos, "Vertical" + contentId);// 把数据表的objectid作为名字，即可单方面寻找
                ed.WriteMessage("\n表格已插入，继续选择插入点或按 Enter 结束。");
            }
        }

        [CommandMethod("MT")]
        public void ModifyTable()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            PromptEntityResult per = ed.GetEntity("选择表格图块: ");
            if (per.Status != PromptStatus.OK) return;

            // 提示用户选择操作
            PromptKeywordOptions pko = new PromptKeywordOptions("\n选择操作 [增加行(Add)/删除行(Remove)]");
            pko.Keywords.Add("Add");
            pko.Keywords.Add("Remove");
            pko.AllowNone = false;
            PromptResult pkr = ed.GetKeywords(pko);

            if (pkr.Status != PromptStatus.OK) return;

            // 提示用户输入要增加或减少的行数
            PromptIntegerOptions poi = new PromptIntegerOptions("\n请输入行数: ");
            poi.AllowNegative = true; // 允许输入负数
            poi.AllowZero = false; // 不允许输入零
            poi.DefaultValue = 1; // 默认值
            PromptIntegerResult pir = ed.GetInteger(poi);

            if (pir.Status != PromptStatus.OK) return;

            int numberOfRows = pir.Value;

            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                BlockReference blockRef = trans.GetObject(per.ObjectId, OpenMode.ForWrite) as BlockReference;
                if (blockRef == null) return;

                BlockTableRecord btr = trans.GetObject(blockRef.BlockTableRecord, OpenMode.ForWrite) as BlockTableRecord;

                // 确认是小表才修改
                if (btr.Name.Contains("Small"))
                {
                    int rowCount = GetRowCount(btr);

                    // 获取表格实际位置
                    Point3d insertPoint = blockRef.Position;
                    string name = blockRef.Name;

                    if (pkr.StringResult == "Add")
                    {
                        // 删除原表格
                        blockRef.Erase();
                        btr.Erase();
                        string contentId = CreateBoxLoopContent(db, numberOfRows, new Point3d(insertPoint.X, insertPoint.Y - cellHeight * rowCount, insertPoint.Z));
                        CreateBoxLoopTable(db, rowCount + numberOfRows, insertPoint, name + contentId);
                    }
                    else if (pkr.StringResult == "Remove")
                    {
                        if (rowCount - numberOfRows >= 2) // 确保行数不小于2
                        {
                            // 删除原表格
                            blockRef.Erase();
                            btr.Erase();
                            Point3d topLeft = new Point3d(insertPoint.X, insertPoint.Y - (rowCount - numberOfRows) * cellHeight, insertPoint.Z);
                            Point3d downRight = new Point3d(insertPoint.X + 3 * cellWidth, insertPoint.Y - rowCount * cellHeight, insertPoint.Z);
                            DeleteRect(topLeft, downRight);
                            CreateBoxLoopTable(db, rowCount - numberOfRows, insertPoint, name);
                        }
                        else
                        {
                            ed.WriteMessage($"\n无法删除：表格至少要有一行数据");
                        }
                    }
                }
                else if (btr.Name.Contains("PMT"))
                {
                    int rowCount = GetRowCount(btr);

                    // 获取表格实际位置
                    Point3d insertPoint = blockRef.Position;
                    string name = blockRef.Name;

                    if (pkr.StringResult == "Add")
                    {
                        // 删除原表格
                        blockRef.Erase();
                        btr.Erase();
                        string contentId = CreateBoxLoopContent2(db, numberOfRows, new Point3d(insertPoint.X, insertPoint.Y - cellHeight * rowCount, insertPoint.Z));
                        CreateBoxLoopTable2(db, rowCount + numberOfRows, insertPoint, name + contentId);
                    }
                    else if (pkr.StringResult == "Remove")
                    {
                        if (rowCount - numberOfRows >= 1) // 确保行数不小于1
                        {
                            // 删除原表格
                            blockRef.Erase();
                            btr.Erase(); 
                            Point3d topLeft = new Point3d(insertPoint.X, insertPoint.Y - (rowCount - numberOfRows) * cellHeight, insertPoint.Z);
                            Point3d downRight = new Point3d(insertPoint.X + 2 * cellWidth + cellWidth2, insertPoint.Y - rowCount * cellHeight, insertPoint.Z);
                            DeleteRect(topLeft, downRight);
                            CreateBoxLoopTable2(db, rowCount - numberOfRows, insertPoint, name);
                        }
                        else
                        {
                            ed.WriteMessage($"\n无法删除：表格至少要有一行数据");
                        }
                    }
                }
                else
                {
                    ed.WriteMessage($"\n操作失败：不是指定格式表格");
                }
                trans.Commit();
            }
        }

        // 删除指定坐标内的东西
        private void DeleteRect(Point3d topLeft, Point3d bottomRight)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            // 使用 SelectCrossingWindow 选择区域内的对象
            PromptSelectionResult selRes = ed.SelectCrossingWindow(topLeft, bottomRight);
            if (selRes.Status != PromptStatus.OK) return;

            SelectionSet selectionSet = selRes.Value;

            using (Transaction trans = doc.TransactionManager.StartTransaction())
            {
                foreach (SelectedObject selObj in selectionSet)
                {
                    if (selObj != null)
                    {
                        Entity entity = trans.GetObject(selObj.ObjectId, OpenMode.ForWrite) as Entity;

                        // 只筛选 DBText 类型的对象
                        if (entity is DBText dbText)
                        {
                            entity.Erase(); // 删除对象
                        }
                    }
                }

                trans.Commit();
            }
        }

        // 创建1列数据表格
        private string CreateDataTable1(Database db, Point3d insertPoint)
        {
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                // 获取模型空间块表记录
                BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord modelSpace = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                // 创建表格对象
                Table table = new Table
                {
                    // 设置表格的行数和列数
                    NumRows = 5,
                    NumColumns = 1,
                    // 设置单元格的宽度和高度
                    Width = cellWidthBig,
                    Height = cellHeightBig * 5,
                    // 设置表格的插入点
                    Position = insertPoint
                };

                table.Cells.TextHeight = textSizeBig;
                table.Cells.Alignment = CellAlignment.MiddleCenter; // 居中
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;


                // 填充表格单元格数据
                table.SetTextString(0, 0, "0");
                table.SetTextString(1, 0, "0");
                table.SetTextString(2, 0, "=A1*A2");
                table.SetTextString(3, 0, "0.8");
                table.SetTextString(4, 0, "=A3/0.38/A4/1.732");

                table.Cells[2, 0].DataFormat = "%lu2%pr2";
                table.Cells[4, 0].DataFormat = "%lu2%pr2";

                // 将表格对象添加到模型空间
                modelSpace.AppendEntity(table);
                trans.AddNewlyCreatedDBObject(table, true);
                trans.Commit();
                return table.Handle.ToString();
            }
        }

        //大表创建箱号回路表格
        private void CreateBoxLoopTableBig(Database db, Point3d insertPoint, string name)
        {
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord btr = new BlockTableRecord
                {
                    Name = name,
                    Origin = insertPoint
                };

                bt.UpgradeOpen();
                ObjectId btrId = bt.Add(btr);
                trans.AddNewlyCreatedDBObject(btr, true);

                DrawTable(btr, insertPoint, 19, 1, cellWidthBig, cellHeightBig);

                BlockTableRecord modelSpace = trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                BlockReference blockRef = new BlockReference(insertPoint, btrId);
                modelSpace.AppendEntity(blockRef);
                trans.AddNewlyCreatedDBObject(blockRef, true);

                trans.Commit();
            }
        }

        //创建大表箱号回路数据 
        private string CreateBigContent(Database db, Point3d insertPoint)
        {
            string idCombine = "";
            Point3d usage_position = new Point3d(insertPoint.X + cellWidthBig / 2, insertPoint.Y - (cellHeightBig + textSizeBig) / 2 - cellHeightBig, insertPoint.Z);
            Point3d box_position = new Point3d(insertPoint.X + cellWidthBig / 2, insertPoint.Y - (cellHeightBig + textSizeBig) / 2 - cellHeightBig * 2, insertPoint.Z);
            Point3d loop_position = new Point3d(insertPoint.X + cellWidthBig / 2, insertPoint.Y - (cellHeightBig + textSizeBig) / 2, insertPoint.Z);
            Point3d current_position = new Point3d(insertPoint.X + cellWidthBig / 2, insertPoint.Y - (cellHeightBig + textSizeBig) / 2 + cellHeightBig * 9, insertPoint.Z);
            Point3d protect_position = new Point3d(insertPoint.X + cellWidthBig / 2, insertPoint.Y - (cellHeightBig + textSizeBig) / 2 + cellHeightBig * 8, insertPoint.Z);
            Point3d delay_protect_position = new Point3d(insertPoint.X + cellWidthBig / 2, insertPoint.Y - (cellHeightBig + textSizeBig) / 2 + cellHeightBig * 6, insertPoint.Z);
            Point3d cable_position = new Point3d(insertPoint.X + cellWidthBig / 2, insertPoint.Y - (cellHeightBig + textSizeBig) / 2 - cellHeightBig * 9, insertPoint.Z);
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                // 获取模型空间块表记录
                BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord modelSpace = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                // 创建文字对象
                DBText box = new DBText
                {
                    Position = box_position,
                    TextString = "电箱编号",
                    Height = textSizeBig,  // 设置文字高度，根据需要调整
                    WidthFactor = 0.7,
                    HorizontalMode = TextHorizontalMode.TextCenter,
                    VerticalMode = TextVerticalMode.TextBase,  // 使用 TextBase 作为对齐模式
                    AlignmentPoint = box_position,
                    Color = Color.FromRgb(255, 255, 127)
                };

                DBText loop = new DBText
                {
                    Position = loop_position,
                    TextString = "回路编号",
                    Height = textSizeBig,  // 设置文字高度，根据需要调整
                    WidthFactor = 0.7,
                    HorizontalMode = TextHorizontalMode.TextCenter,
                    VerticalMode = TextVerticalMode.TextBase,  // 使用 TextBase 作为对齐模式
                    AlignmentPoint = loop_position,
                    Color = Color.FromRgb(255, 255, 127)
                };

                DBText usage = new DBText
                {
                    Position = usage_position,
                    TextString = "电箱用途",
                    Height = textSizeBig,  // 设置文字高度，根据需要调整
                    WidthFactor = 0.7,
                    HorizontalMode = TextHorizontalMode.TextCenter,
                    VerticalMode = TextVerticalMode.TextBase,  // 使用 TextBase 作为对齐模式
                    AlignmentPoint = usage_position,
                    Color = Color.FromRgb(255, 255, 127)
                };

                DBText current = new DBText
                {
                    Position = current_position,
                    TextString = "额定电流",
                    Height = textSizeBig,  // 设置文字高度，根据需要调整
                    WidthFactor = 0.7,
                    HorizontalMode = TextHorizontalMode.TextCenter,
                    VerticalMode = TextVerticalMode.TextBase,  // 使用 TextBase 作为对齐模式
                    AlignmentPoint = current_position,
                    Color = Color.FromRgb(255, 255, 127)
                };

                DBText protect = new DBText
                {
                    Position = protect_position,
                    TextString = "过载保护",
                    Height = textSizeBig,  // 设置文字高度，根据需要调整
                    WidthFactor = 0.7,
                    HorizontalMode = TextHorizontalMode.TextCenter,
                    VerticalMode = TextVerticalMode.TextBase,  // 使用 TextBase 作为对齐模式
                    AlignmentPoint = protect_position,
                    Color = Color.FromRgb(255, 255, 127)
                };

                DBText delay_protect = new DBText
                {
                    Position = delay_protect_position,
                    TextString = "短路瞬时保护",
                    Height = textSizeBig,  // 设置文字高度，根据需要调整
                    WidthFactor = 0.7,
                    HorizontalMode = TextHorizontalMode.TextCenter,
                    VerticalMode = TextVerticalMode.TextBase,  // 使用 TextBase 作为对齐模式
                    AlignmentPoint = delay_protect_position,
                    Color = Color.FromRgb(255, 255, 127)
                };

                DBText cable = new DBText
                {
                    Position = cable_position,
                    TextString = "电缆",
                    Height = textSizeBig,  // 设置文字高度，根据需要调整
                    WidthFactor = 0.7,
                    HorizontalMode = TextHorizontalMode.TextCenter,
                    VerticalMode = TextVerticalMode.TextBase,  // 使用 TextBase 作为对齐模式
                    AlignmentPoint = cable_position,
                    Color = Color.FromRgb(255, 255, 127)
                };

                // 将表格对象添加到模型空间
                modelSpace.AppendEntity(box);
                modelSpace.AppendEntity(loop);
                modelSpace.AppendEntity(usage);
                modelSpace.AppendEntity(current);
                modelSpace.AppendEntity(cable);
                modelSpace.AppendEntity(protect);
                modelSpace.AppendEntity(delay_protect);
                trans.AddNewlyCreatedDBObject(box, true);
                trans.AddNewlyCreatedDBObject(loop, true);
                trans.AddNewlyCreatedDBObject(usage, true);
                trans.AddNewlyCreatedDBObject(cable, true);
                trans.AddNewlyCreatedDBObject(current, true);
                trans.AddNewlyCreatedDBObject(protect, true);
                trans.AddNewlyCreatedDBObject(delay_protect, true);

                trans.Commit();
                idCombine = "|" + box.Handle.ToString() + "_" + loop.Handle.ToString() + "_" + usage.Handle.ToString() + "_" + current.Handle.ToString() + "_" + cable.Handle.ToString() + "_" + protect.Handle.ToString() + "_" + delay_protect.Handle.ToString();
            }
            return idCombine; // 得到结果：|电箱Handle_回路Handle_用途Handle
        }

        // 创建3列数据表格
        private string CreateDataTable3(Database db, Point3d insertPoint)
        {
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                // 获取模型空间块表记录
                BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord modelSpace = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                // 创建表格对象
                Table table = new Table
                {
                    // 设置表格的行数和列数
                    NumRows = 6,
                    NumColumns = 3,
                    // 设置单元格的宽度和高度
                    Width = tableWidth * 3,
                    Height = tableHeight * 6,
                    // 设置表格的插入点
                    Position = insertPoint
                };

                table.Cells.TextHeight = textSize;
                table.Cells[1, 0].TextHeight = headerSize;
                table.Cells[2, 0].TextHeight = headerSize;
                table.Cells[3, 0].TextHeight = headerSize;
                table.Cells[4, 0].TextHeight = headerSize;
                table.Cells[5, 0].TextHeight = headerSize;
                table.Cells.Alignment = CellAlignment.MiddleCenter; // 居中


                // 填充表格单元格数据
                table.SetTextString(0, 0, "      ");
                table.SetTextString(1, 0, "Pe(kW)=");
                table.SetTextString(2, 0, "Kx=");
                table.SetTextString(3, 0, "Pjs(kW)=");
                table.SetTextString(4, 0, "Cosφ=");
                table.SetTextString(5, 0, "Ijs(A)=");

                table.SetTextString(0, 1, "平时");
                table.SetTextString(1, 1, "0");
                table.SetTextString(2, 1, "0");
                table.SetTextString(3, 1, "=B2*B3");
                table.SetTextString(4, 1, "0.8");
                table.SetTextString(5, 1, "=B4/0.38/B5/1.732");

                table.SetTextString(0, 2, "消防");
                table.SetTextString(1, 2, "0");
                table.SetTextString(2, 2, "0");
                table.SetTextString(3, 2, "=C2*C3");
                table.SetTextString(4, 2, "0.8");
                table.SetTextString(5, 2, "=C4/0.38/C5/1.732");

                table.Cells[3, 1].DataFormat = "%lu2%pr2";
                table.Cells[5, 1].DataFormat = "%lu2%pr2";
                table.Cells[3, 2].DataFormat = "%lu2%pr2";
                table.Cells[5, 2].DataFormat = "%lu2%pr2";

                // 注册一个应用程序名称（如果还没有注册）
                RegAppTable regAppTable = trans.GetObject(db.RegAppTableId, OpenMode.ForWrite) as RegAppTable;
                if (!regAppTable.Has("CADAUTOTABLE"))
                {
                    RegAppTableRecord regAppRecord = new RegAppTableRecord();
                    regAppRecord.Name = "CADAUTOTABLE";
                    regAppTable.Add(regAppRecord);
                    trans.AddNewlyCreatedDBObject(regAppRecord, true);
                }

                // 创建扩展数据
                ResultBuffer xdata = new ResultBuffer(
                    new TypedValue((int)DxfCode.ExtendedDataRegAppName, "CADAUTOTABLE"),
                    new TypedValue((int)DxfCode.ExtendedDataAsciiString, $"表格编号:{smallTableSign}")
                );

                // 将扩展数据附加到对象
                table.XData = xdata;


                // 将表格对象添加到模型空间
                modelSpace.AppendEntity(table);
                trans.AddNewlyCreatedDBObject(table, true);
                trans.Commit();
                return table.Handle.ToString();
            }
        }

        // 创建2列数据表格
        private string CreateDataTable2(Database db, Point3d insertPoint)
        {
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                // 获取模型空间块表记录
                BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord modelSpace = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                // 创建表格对象
                Table table = new Table
                {
                    // 设置表格的行数和列数
                    NumRows = 5,
                    NumColumns = 2,
                    // 设置单元格的宽度和高度
                    Width = tableWidth * 2,
                    Height = tableHeight * 5,
                    // 设置表格的插入点
                    Position = insertPoint
                };

                table.Cells.TextHeight = textSize;
                table.Cells[0, 0].TextHeight = headerSize;
                table.Cells[1, 0].TextHeight = headerSize;
                table.Cells[2, 0].TextHeight = headerSize;
                table.Cells[3, 0].TextHeight = headerSize;
                table.Cells[4, 0].TextHeight = headerSize;
                table.Cells.Alignment = CellAlignment.MiddleCenter; // 居中


                // 填充表格单元格数据
                table.SetTextString(0, 0, "Pe(kW)=");
                table.SetTextString(1, 0, "Kx=");
                table.SetTextString(2, 0, "Pjs(kW)=");
                table.SetTextString(3, 0, "Cosφ=");
                table.SetTextString(4, 0, "Ijs(A)=");

                table.SetTextString(0, 1, "0");
                table.SetTextString(1, 1, "0");
                table.SetTextString(2, 1, "=B1*B2");
                table.SetTextString(3, 1, "0.8");
                table.SetTextString(4, 1, "=B3/0.38/B4/1.732");

                table.Cells[2, 1].DataFormat = "%lu2%pr2";
                table.Cells[4, 1].DataFormat = "%lu2%pr2";


                // 注册一个应用程序名称（如果还没有注册）
                RegAppTable regAppTable = trans.GetObject(db.RegAppTableId, OpenMode.ForWrite) as RegAppTable;
                if (!regAppTable.Has("CADAUTOTABLE"))
                {
                    RegAppTableRecord regAppRecord = new RegAppTableRecord();
                    regAppRecord.Name = "CADAUTOTABLE";
                    regAppTable.Add(regAppRecord);
                    trans.AddNewlyCreatedDBObject(regAppRecord, true);
                }

                // 创建扩展数据
                ResultBuffer xdata = new ResultBuffer(
                    new TypedValue((int)DxfCode.ExtendedDataRegAppName, "CADAUTOTABLE"),
                    new TypedValue((int)DxfCode.ExtendedDataAsciiString, $"表格编号:{smallTableSign}")
                );

                // 将扩展数据附加到对象
                table.XData = xdata;


                // 将表格对象添加到模型空间
                modelSpace.AppendEntity(table);
                trans.AddNewlyCreatedDBObject(table, true);
                trans.Commit();
                return table.Handle.ToString();
            }
        }

        //创建箱号回路表格
        private void CreateBoxLoopTable(Database db, int row, Point3d insertPoint, string name)
        {
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord btr = new BlockTableRecord
                {
                    Name = name,
                    Origin = insertPoint
                };

                bt.UpgradeOpen();
                ObjectId btrId = bt.Add(btr);
                trans.AddNewlyCreatedDBObject(btr, true);

                DrawTable(btr, insertPoint, row, 3, cellWidth, cellHeight);
                DrawText(btr, new Point3d(insertPoint.X + cellWidth / 2, insertPoint.Y - (cellHeight + textSize) / 2, insertPoint.Z), "电箱用途", textSize);
                DrawText(btr, new Point3d(insertPoint.X + cellWidth + cellWidth / 2, insertPoint.Y - (cellHeight + textSize) / 2, insertPoint.Z), "电箱编号", textSize);
                DrawText(btr, new Point3d(insertPoint.X + cellWidth + cellWidth + cellWidth / 2, insertPoint.Y - (cellHeight + textSize) / 2, insertPoint.Z), "回路编号", textSize);

                BlockTableRecord modelSpace = trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                BlockReference blockRef = new BlockReference(insertPoint, btrId);
                modelSpace.AppendEntity(blockRef);
                trans.AddNewlyCreatedDBObject(blockRef, true);

                trans.Commit();
            }
        }

        //创建配电平面图箱号回路表格
        private void CreateBoxLoopTable2(Database db, int row, Point3d insertPoint, string name)
        {
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord btr = new BlockTableRecord
                {
                    Name = name,
                    Origin = insertPoint
                };

                bt.UpgradeOpen();
                ObjectId btrId = bt.Add(btr);
                trans.AddNewlyCreatedDBObject(btr, true);

                DrawRow2(btr, insertPoint, row, cellWidth2, cellWidth, cellHeight);

                BlockTableRecord modelSpace = trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                BlockReference blockRef = new BlockReference(insertPoint, btrId);
                modelSpace.AppendEntity(blockRef);
                trans.AddNewlyCreatedDBObject(blockRef, true);

                trans.Commit();
            }
        }
        //创建配电平面图箱号回路数据
        private string CreateBoxLoopContent2(Database db, int row, Point3d insertPoint)
        {
            string all_combine = "";
            for (int i = 0; i < row; i++)
            {
                string idCombine = "";
                Point3d box_position = new Point3d(insertPoint.X + cellWidth2 + cellWidth / 2, insertPoint.Y - (cellHeight + textSize) / 2 - cellHeight * i, insertPoint.Z);
                Point3d loop_position = new Point3d(insertPoint.X + cellWidth2 + cellWidth + cellWidth / 2, insertPoint.Y - (cellHeight + textSize) / 2 - cellHeight * i, insertPoint.Z);
                Point3d number_position = new Point3d(insertPoint.X + cellWidth2 / 2, insertPoint.Y - (cellHeight + textSize) / 2 - cellHeight * i, insertPoint.Z);
                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    // 获取模型空间块表记录
                    BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord modelSpace = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                    // 创建文字对象
                    DBText number = new DBText
                    {
                        Position = number_position,
                        TextString = "编号",
                        Height = textSize,  // 设置文字高度，根据需要调整
                        WidthFactor = 0.7,
                        HorizontalMode = TextHorizontalMode.TextCenter,
                        VerticalMode = TextVerticalMode.TextBase,  // 使用 TextBase 作为对齐模式
                        AlignmentPoint = number_position,
                        Color = Color.FromRgb(0, 255, 0)
                    };

                    DBText box = new DBText
                    {
                        Position = box_position,
                        TextString = "电箱编号",
                        Height = textSize,  // 设置文字高度，根据需要调整
                        WidthFactor = 0.7,
                        HorizontalMode = TextHorizontalMode.TextCenter,
                        VerticalMode = TextVerticalMode.TextBase,  // 使用 TextBase 作为对齐模式
                        AlignmentPoint = box_position,
                        Color = Color.FromRgb(0, 255, 0)
                    };

                    DBText loop = new DBText
                    {
                        Position = loop_position,
                        TextString = "回路编号",
                        Height = textSize,  // 设置文字高度，根据需要调整
                        WidthFactor = 0.7,
                        HorizontalMode = TextHorizontalMode.TextCenter,
                        VerticalMode = TextVerticalMode.TextBase,  // 使用 TextBase 作为对齐模式
                        AlignmentPoint = loop_position,
                        Color = Color.FromRgb(0, 255, 0)
                    };

                    // 将表格对象添加到模型空间
                    modelSpace.AppendEntity(box);
                    modelSpace.AppendEntity(loop);
                    modelSpace.AppendEntity(number);
                    trans.AddNewlyCreatedDBObject(box, true);
                    trans.AddNewlyCreatedDBObject(loop, true);
                    trans.AddNewlyCreatedDBObject(number, true);
                    trans.Commit();
                    idCombine = box.Handle.ToString() + "_" + loop.Handle.ToString() + "_" + number.Handle.ToString();
                }

                all_combine += "|" + idCombine;
            }
            return all_combine; // 得到结果：|电箱Handle_回路Handle|电箱Handle_回路Handle|电箱Handle_回路Handle
        }

        //创建箱号回路数据 
        private string CreateBoxLoopContent(Database db, int row, Point3d insertPoint)
        {
            string all_combine = "";
            for (int i = 0; i < row; i++)
            {
                string idCombine = "";
                Point3d usage_position = new Point3d(insertPoint.X + cellWidth / 2, insertPoint.Y - (cellHeight + textSize) / 2 - cellHeight * i, insertPoint.Z);
                Point3d box_position = new Point3d(insertPoint.X + cellWidth + cellWidth / 2, insertPoint.Y - (cellHeight + textSize) / 2 - cellHeight * i, insertPoint.Z);
                Point3d loop_position = new Point3d(insertPoint.X + cellWidth + cellWidth + cellWidth / 2, insertPoint.Y - (cellHeight + textSize) / 2 - cellHeight * i, insertPoint.Z);
                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    // 获取模型空间块表记录
                    BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord modelSpace = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                    // 创建文字对象
                    DBText box = new DBText
                    {
                        Position = box_position,
                        TextString = "电箱编号",
                        Height = textSize,  // 设置文字高度，根据需要调整
                        WidthFactor = 0.7,
                        HorizontalMode = TextHorizontalMode.TextCenter,
                        VerticalMode = TextVerticalMode.TextBase,  // 使用 TextBase 作为对齐模式
                        AlignmentPoint = box_position,
                        Color = Color.FromRgb(0, 255, 0)
                    };

                    DBText loop = new DBText
                    {
                        Position = loop_position,
                        TextString = "回路编号",
                        Height = textSize,  // 设置文字高度，根据需要调整
                        WidthFactor = 0.7,
                        HorizontalMode = TextHorizontalMode.TextCenter,
                        VerticalMode = TextVerticalMode.TextBase,  // 使用 TextBase 作为对齐模式
                        AlignmentPoint = loop_position,
                        Color = Color.FromRgb(0, 255, 0)
                    };

                    DBText usage = new DBText
                    {
                        Position = usage_position,
                        TextString = "电箱用途",
                        Height = textSize,  // 设置文字高度，根据需要调整
                        WidthFactor = 0.7,
                        HorizontalMode = TextHorizontalMode.TextCenter,
                        VerticalMode = TextVerticalMode.TextBase,  // 使用 TextBase 作为对齐模式
                        AlignmentPoint = usage_position,
                        Color = Color.FromRgb(0, 255, 0)
                    };

                    // 将表格对象添加到模型空间
                    modelSpace.AppendEntity(box);
                    modelSpace.AppendEntity(loop);
                    modelSpace.AppendEntity(usage);
                    trans.AddNewlyCreatedDBObject(box, true);
                    trans.AddNewlyCreatedDBObject(loop, true);
                    trans.AddNewlyCreatedDBObject(usage, true);
                    trans.Commit();
                    idCombine = box.Handle.ToString() + "_" + loop.Handle.ToString() + "_" + usage.Handle.ToString();
                }

                all_combine += "|" + idCombine;
            }
            return all_combine; // 得到结果：|电箱Handle_回路Handle_用途Handle|电箱Handle_回路Handle_用途Handle|电箱Handle_回路Handle_用途Handle
        }

        //创建竖向箱号回路表格
        private void CreateVerticalTable(Database db, int row, Point3d insertPoint, string name)
        {
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord btr = new BlockTableRecord
                {
                    Name = name,
                    Origin = insertPoint
                };

                bt.UpgradeOpen();
                ObjectId btrId = bt.Add(btr);
                trans.AddNewlyCreatedDBObject(btr, true);

                DrawRowVertical(btr, insertPoint, row, 2, cellWidth, cellHeight);

                BlockTableRecord modelSpace = trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                BlockReference blockRef = new BlockReference(insertPoint, btrId);
                modelSpace.AppendEntity(blockRef);
                trans.AddNewlyCreatedDBObject(blockRef, true);

                trans.Commit();
            }
        }

        //创建竖向箱号回路数据 
        private string CreateVerticalContent(Database db, Point3d insertPoint)
        {
            string all_combine = "";
            string idCombine = "";
            Point3d box_position = new Point3d(insertPoint.X + (cellHeight + textSize) / 2, insertPoint.Y - cellWidth - cellWidth / 2, insertPoint.Z);
            Point3d loop_position = new Point3d(insertPoint.X + (cellHeight + textSize) / 2, insertPoint.Y - cellWidth / 2, insertPoint.Z);
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                // 获取模型空间块表记录
                BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord modelSpace = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                // 创建文字对象
                DBText box = new DBText
                {
                    Position = box_position,
                    TextString = "电箱编号",
                    Height = textSize,  // 设置文字高度，根据需要调整
                    Rotation = Math.PI / 2,  // 设置旋转角度为90度（竖向）
                    WidthFactor = 0.7,
                    HorizontalMode = TextHorizontalMode.TextCenter,
                    VerticalMode = TextVerticalMode.TextBase,  // 使用 TextBase 作为对齐模式
                    AlignmentPoint = box_position,
                    Color = Color.FromRgb(0, 255, 0)
                };

                DBText loop = new DBText
                {
                    Position = loop_position,
                    TextString = "回路编号",
                    Height = textSize,  // 设置文字高度，根据需要调整
                    Rotation = Math.PI / 2,  // 设置旋转角度为90度（竖向）
                    WidthFactor = 0.7,
                    HorizontalMode = TextHorizontalMode.TextCenter,
                    VerticalMode = TextVerticalMode.TextBase,  // 使用 TextBase 作为对齐模式
                    AlignmentPoint = loop_position,
                    Color = Color.FromRgb(0, 255, 0)
                };

                // 将表格对象添加到模型空间
                modelSpace.AppendEntity(box);
                modelSpace.AppendEntity(loop);
                trans.AddNewlyCreatedDBObject(box, true);
                trans.AddNewlyCreatedDBObject(loop, true);
                trans.Commit();
                idCombine = box.Handle.ToString() + "_" + loop.Handle.ToString();
            }

            all_combine += "|" + idCombine;

            return all_combine; // 得到结果：|电箱Handle_回路Handle_用途Handle|电箱Handle_回路Handle_用途Handle|电箱Handle_回路Handle_用途Handle
        }

        private int GetRowCount(BlockTableRecord btr)
        {
            int polylineCount = 0;

            foreach (ObjectId entId in btr)
            {
                Entity ent = btr.Database.TransactionManager.GetObject(entId, OpenMode.ForRead) as Entity;
                if (ent is Polyline)
                {
                    polylineCount++;
                }
            }

            // 假设每行有 8 条 Polyline（包括横线和竖线）
            return polylineCount / 3;
        }

        private void DrawRow(BlockTableRecord btr, Point3d insertPoint, int rowIndex, int colCount, double colWidth, double rowHeight)
        {
            using (Transaction trans = btr.Database.TransactionManager.StartTransaction())
            {
                for (int c = 0; c < colCount; c++)
                {
                    Polyline rect = new Polyline();
                    rect.AddVertexAt(0, new Point2d(insertPoint.X + c * colWidth, insertPoint.Y - rowHeight * rowIndex), 0, 0, 0);
                    rect.AddVertexAt(1, new Point2d(insertPoint.X + (c + 1) * colWidth, insertPoint.Y - rowHeight * rowIndex), 0, 0, 0);
                    rect.AddVertexAt(2, new Point2d(insertPoint.X + (c + 1) * colWidth, insertPoint.Y - rowHeight * (rowIndex + 1)), 0, 0, 0);
                    rect.AddVertexAt(3, new Point2d(insertPoint.X + c * colWidth, insertPoint.Y - rowHeight * (rowIndex + 1)), 0, 0, 0);
                    rect.Closed = true;

                    btr.AppendEntity(rect);
                    trans.AddNewlyCreatedDBObject(rect, true);
                }

                trans.Commit();
            }
        }

        // 数字2字母
        static string NumberToLetterString(int number)
        {
            // 确保数字在0到25的范围内
            number = number % 26;

            // 将数字转换为对应的小写字母
            char letter = (char)('a' + number);

            // 返回字母作为字符串
            return letter.ToString();
        }

        // 画配电平面图的表格行
        private void DrawRow2(BlockTableRecord btr, Point3d insertPoint, int rowCount, double colWidth1, double colWidth2, double rowHeight)
        {
            for (int r = 0; r < rowCount; r++)
            {
                using (Transaction trans = btr.Database.TransactionManager.StartTransaction())
                {
                    Polyline rect1 = new Polyline();
                    rect1.AddVertexAt(0, new Point2d(insertPoint.X, insertPoint.Y - rowHeight * r), 0, 0, 0);
                    rect1.AddVertexAt(1, new Point2d(insertPoint.X + colWidth1, insertPoint.Y - rowHeight * r), 0, 0, 0);
                    rect1.AddVertexAt(2, new Point2d(insertPoint.X + colWidth1, insertPoint.Y - rowHeight * (r + 1)), 0, 0, 0);
                    rect1.AddVertexAt(3, new Point2d(insertPoint.X, insertPoint.Y - rowHeight * (r + 1)), 0, 0, 0);
                    rect1.Closed = true;

                    btr.AppendEntity(rect1);
                    trans.AddNewlyCreatedDBObject(rect1, true);

                    for (int c = 1; c < 3; c++)
                    {
                        Polyline rect2 = new Polyline();
                        rect2.AddVertexAt(0, new Point2d(insertPoint.X + colWidth1 + (c - 1) * colWidth2, insertPoint.Y - rowHeight * r), 0, 0, 0);
                        rect2.AddVertexAt(1, new Point2d(insertPoint.X + colWidth1 + c * colWidth2, insertPoint.Y - rowHeight * r), 0, 0, 0);
                        rect2.AddVertexAt(2, new Point2d(insertPoint.X + colWidth1 + c * colWidth2, insertPoint.Y - rowHeight * (r + 1)), 0, 0, 0);
                        rect2.AddVertexAt(3, new Point2d(insertPoint.X + colWidth1 + (c - 1) * colWidth2, insertPoint.Y - rowHeight * (r + 1)), 0, 0, 0);
                        rect2.Closed = true;

                        btr.AppendEntity(rect2);
                        trans.AddNewlyCreatedDBObject(rect2, true);
                    }


                    // DrawText(btr, new Point3d(insertPoint.X + cellWidth2 / 2, insertPoint.Y - (cellHeight + textSize) / 2 - rowHeight * r, insertPoint.Z), NumberToLetterString(r), textSize);

                    trans.Commit();
                }
            }
        }

        private void DrawText(BlockTableRecord btr, Point3d position, string text, int height)
        {
            using (Transaction trans = btr.Database.TransactionManager.StartTransaction())
            {
                // 创建文字对象
                DBText dbText = new DBText
                {
                    Position = position,
                    TextString = text,
                    Height = height,  // 设置文字高度，根据需要调整
                    HorizontalMode = TextHorizontalMode.TextCenter,
                    VerticalMode = TextVerticalMode.TextBase,  // 使用 TextBase 作为对齐模式
                    AlignmentPoint = position
                };

                // 将文字对象添加到 BlockTableRecord
                btr.AppendEntity(dbText);
                trans.AddNewlyCreatedDBObject(dbText, true);

                trans.Commit();
            }
        }

        private void DrawTable(BlockTableRecord btr, Point3d insertPoint, int rowCount, int colCount, double colWidth, double rowHeight)
        {
            for (int r = 0; r < rowCount; r++)
            {
                DrawRow(btr, insertPoint, r, colCount, colWidth, rowHeight);
            }
        }

        private void DrawRowVertical(BlockTableRecord btr, Point3d insertPoint, int rowIndex, int colCount, double colWidth, double rowHeight)
        {
            using (Transaction trans = btr.Database.TransactionManager.StartTransaction())
            {
                for (int c = 0; c < colCount; c++)
                {
                    Polyline rect = new Polyline();
                    rect.AddVertexAt(0, new Point2d(insertPoint.X + rowIndex * rowHeight, insertPoint.Y - colWidth * c), 0, 0, 0);
                    rect.AddVertexAt(1, new Point2d(insertPoint.X + (rowIndex + 1) * rowHeight, insertPoint.Y - colWidth * c), 0, 0, 0);
                    rect.AddVertexAt(2, new Point2d(insertPoint.X + (rowIndex + 1) * rowHeight, insertPoint.Y - colWidth * (c + 1)), 0, 0, 0);
                    rect.AddVertexAt(3, new Point2d(insertPoint.X + rowIndex * rowHeight, insertPoint.Y - colWidth * (c + 1)), 0, 0, 0);
                    rect.Closed = true;

                    btr.AppendEntity(rect);
                    trans.AddNewlyCreatedDBObject(rect, true);
                }

                trans.Commit();
            }
        }
    }
}

using EPDM.Interop.epdm;
using System;
using System.Collections.Generic;
using System.Data;
using System.Runtime.InteropServices;
using System.Threading;
using System.Linq;
using System.Xml;
using System.Text;
using System.IO;
using Patterns.Observer;
using EPDM.Interop.EPDMResultCode;

namespace ExportToXMLLib
{
    public class Export
    {
        const int BoomId = 8;
        private static object mLockObj = new object();
        DBConnection con = DBConnection.DBProp;
        private string filePath;
        private string pathToSave = @"\\pdmsrv\XML\";
        
        List<MyBomShell> AssmblyBom;
        List<MyBomShell> fullDataSpecParts;
        List<MyBomShell> fullDataSpecAsmblAndParts;

        
        public Export(string filePath)
        {
            this.filePath = filePath;

            //////
            List<string> conf = GetConfigurations(filePath);
            AssmblyBom = GetBomShell(filePath, conf, BoomId);
            if (AssmblyBom == null)
            {
                
                return;
            }
            fullDataSpecParts = GetFullSpecification(AssmblyBom);//для каждой детали
        }

        private static IEdmVault5 vault
        {
            get
            {
                IEdmVault5 vault = EdmVaultSingleton.Instance;

                if (!vault.IsLoggedIn)
                {
                    vault.LoginAuto("Vents-PDM", 0);
                }
                return vault;
            }
        }



        private void ExportParts(IEnumerable<MyBomShell> fullDataSpec)
        {
            string tempPath = string.Empty;
            MyBomShell tempBom = null;
            IEnumerable<IGrouping<string, MyBomShell>> grouped = default(IEnumerable<IGrouping<string, MyBomShell>>);
            try
            {
                grouped = fullDataSpec.GroupBy(x => x.FileName);



                foreach (var bomForOneDoc in grouped)
                {
                    if (bomForOneDoc.Key.ToUpper().Contains(".SLDPRT"))
                    {
                        tempBom = bomForOneDoc.ElementAt<MyBomShell>(0);
                        if (tempBom != null)
                        {
                            tempPath = Path.Combine(tempBom.FilePath, tempBom.FileName);
                        }
                        Export ex = new Export(tempPath);

                        foreach (var docWithAllConfigs in ex.fullDataSpecParts)
                        {
                            ex.ExportPartsToXML2(bomForOneDoc);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageObserver.Instance.SetMessage("Failed to ExportToXMLWithSubAsmbl, with error:  " + ex.Message + System.Environment.NewLine +
                    tempPath, MessageType.Error);
            }
            

            //foreach (var bomForOneDoc in fullDataSpec)
            //{
            //    if (bomForOneDoc.FileName.ToUpper().Contains(".SLDPRT"))
            //    {
            //        Export ex = new Export(Path.Combine( bomForOneDoc.FilePath, bomForOneDoc.FileName));

            //        foreach (var docWithAllConfigs in ex.fullDataSpecParts.GroupBy(x => x.FileName))
            //        {
            //            ex.ExportPartsToXML2(docWithAllConfigs);
            //        }
            //    }
            //}
        }
        private void ExportPartsToXML()
        {
            foreach (var part in fullDataSpecParts.GroupBy(x => x.FileName))
            {
                try
                {
                    string name = Path.GetFileNameWithoutExtension(part.Key);
                    var myXml = new XmlTextWriter(pathToSave + name + ".xml", Encoding.UTF8);
                    myXml.WriteStartDocument();
                    myXml.Formatting = Formatting.Indented;

                    myXml.WriteStartElement("xml");
                    myXml.WriteStartElement("transactions");
                    myXml.WriteStartElement("transaction");


                    myXml.WriteAttributeString("vaultName", "Vents-PDM");
                    myXml.WriteAttributeString("type", "");
                    myXml.WriteAttributeString("date", "");

                    // document
                    myXml.WriteStartElement("document");
                    myXml.WriteAttributeString("pdmweid", "");
                    myXml.WriteAttributeString("aliasset", "Export To ERP");

                    foreach (var item in part)
                    {

                        #region XML
                        // Конфигурация
                        myXml.WriteStartElement("configuration");
                        myXml.WriteAttributeString("name", item.Configuration);

                        // Материал
                        myXml.WriteStartElement("attribute");
                        myXml.WriteAttributeString("name", "Материал");
                        myXml.WriteAttributeString("value", item.Material.ToString());
                        myXml.WriteEndElement();

                        // Наименование
                        myXml.WriteStartElement("attribute");
                        myXml.WriteAttributeString("name", "Наименование");
                        myXml.WriteAttributeString("value", item.Description);
                        myXml.WriteEndElement();

                        // Обозначение
                        myXml.WriteStartElement("attribute");
                        myXml.WriteAttributeString("name", "Обозначение");
                        myXml.WriteAttributeString("value", item.PartNumber);
                        myXml.WriteEndElement();

                        // Площадь покрытия
                        myXml.WriteStartElement("attribute");
                        myXml.WriteAttributeString("name", "Площадь покрытия");
                        myXml.WriteAttributeString("value", item.SurfaceArea.ToString());
                        myXml.WriteEndElement();

                        // Код_Материала
                        myXml.WriteStartElement("attribute");
                        myXml.WriteAttributeString("name", "Код_Материала");
                        myXml.WriteAttributeString("value", item.CodeMaterial.ToString());
                        myXml.WriteEndElement();

                        // Длина граничной рамки
                        myXml.WriteStartElement("attribute");
                        myXml.WriteAttributeString("name", "Длина граничной рамки");
                        myXml.WriteAttributeString("value", item.WorkpieceX.ToString());
                        myXml.WriteEndElement();

                        // Ширина граничной рамки
                        myXml.WriteStartElement("attribute");
                        myXml.WriteAttributeString("name", "Ширина граничной рамки");
                        myXml.WriteAttributeString("value", item.WorkpieceY.ToString());
                        myXml.WriteEndElement();

                        // Сгибы
                        myXml.WriteStartElement("attribute");
                        myXml.WriteAttributeString("name", "Сгибы");
                        myXml.WriteAttributeString("value", item.Bend.ToString());
                        myXml.WriteEndElement();

                        // Толщина листового металла
                        myXml.WriteStartElement("attribute");
                        myXml.WriteAttributeString("name", "Толщина листового металла");
                        myXml.WriteAttributeString("value", item.ListThickness.ToString());
                        myXml.WriteEndElement();

                        // PaintX
                        myXml.WriteStartElement("attribute");
                        myXml.WriteAttributeString("name", "PaintX");
                        myXml.WriteAttributeString("value", item.PaintX.ToString());
                        myXml.WriteEndElement();

                        // PaintY
                        myXml.WriteStartElement("attribute");
                        myXml.WriteAttributeString("name", "PaintY");
                        myXml.WriteAttributeString("value", item.PaintY.ToString());
                        myXml.WriteEndElement();

                        // PaintZ
                        myXml.WriteStartElement("attribute");
                        myXml.WriteAttributeString("name", "PaintZ");
                        myXml.WriteAttributeString("value", item.PaintZ.ToString());
                        myXml.WriteEndElement();

                        // Версия
                        myXml.WriteStartElement("attribute");
                        myXml.WriteAttributeString("name", "Версия");
                        myXml.WriteAttributeString("value", item.LastVersion.ToString());
                        myXml.WriteEndElement();
                        myXml.WriteEndElement();// ' элемент Configuration name
                        #endregion
                    }

                    myXml.WriteEndElement(); // ' элемент DOCUMENT
                    myXml.WriteEndElement(); // ' элемент TRANSACTION
                    myXml.WriteEndElement(); // ' элемент TRANSACTIONS
                    myXml.WriteEndElement(); // ' элемент XML

                    myXml.Flush();
                    myXml.Close();
                }
                catch (Exception ex)
                {
                    MessageObserver.Instance.SetMessage("Failed to ExportPartsToXML, with error:  " + ex.Message + Environment.NewLine + part.Key, MessageType.Error);
                }
            }
        }

        /// <summary>
        ///  Exports one doc with all configs into 1 xml file
        /// </summary>
        /// <param name="listBomShell"></param>
        private void ExportPartsToXML2(IGrouping<string, MyBomShell> listBomShell)
        {
            string filePathForLog = string.Empty;
            try
            {
                string name = Path.GetFileNameWithoutExtension(listBomShell.Key);

                var myXml = new XmlTextWriter(pathToSave + name + ".xml", Encoding.UTF8);
                myXml.WriteStartDocument();
                myXml.Formatting = Formatting.Indented;

                myXml.WriteStartElement("xml");
                myXml.WriteStartElement("transactions");
                myXml.WriteStartElement("transaction");

                myXml.WriteAttributeString("vaultName", "Vents-PDM");
                myXml.WriteAttributeString("type", "");
                myXml.WriteAttributeString("date", "");

                // document
                myXml.WriteStartElement("document");
                myXml.WriteAttributeString("pdmweid", "");
                myXml.WriteAttributeString("aliasset", "Export To ERP");

                foreach (var item in listBomShell)
                {
                    filePathForLog = item.FilePath;
                    #region XML
                    // Конфигурация
                    myXml.WriteStartElement("configuration");
                    myXml.WriteAttributeString("name", item.Configuration);

                    // Материал
                    myXml.WriteStartElement("attribute");
                    myXml.WriteAttributeString("name", "Материал");
                    myXml.WriteAttributeString("value", item.Material.ToString());
                    myXml.WriteEndElement();

                    // Наименование
                    myXml.WriteStartElement("attribute");
                    myXml.WriteAttributeString("name", "Наименование");
                    myXml.WriteAttributeString("value", item.Description);
                    myXml.WriteEndElement();

                    // Обозначение
                    myXml.WriteStartElement("attribute");
                    myXml.WriteAttributeString("name", "Обозначение");
                    myXml.WriteAttributeString("value", item.PartNumber);
                    myXml.WriteEndElement();

                    // Площадь покрытия
                    myXml.WriteStartElement("attribute");
                    myXml.WriteAttributeString("name", "Площадь покрытия");
                    myXml.WriteAttributeString("value", item.SurfaceArea.ToString());
                    myXml.WriteEndElement();

                    // Код_Материала
                    myXml.WriteStartElement("attribute");
                    myXml.WriteAttributeString("name", "Код_Материала");
                    myXml.WriteAttributeString("value", item.CodeMaterial.ToString());
                    myXml.WriteEndElement();

                    // Длина граничной рамки
                    myXml.WriteStartElement("attribute");
                    myXml.WriteAttributeString("name", "Длина граничной рамки");
                    myXml.WriteAttributeString("value", item.WorkpieceX.ToString());
                    myXml.WriteEndElement();

                    // Ширина граничной рамки
                    myXml.WriteStartElement("attribute");
                    myXml.WriteAttributeString("name", "Ширина граничной рамки");
                    myXml.WriteAttributeString("value", item.WorkpieceY.ToString());
                    myXml.WriteEndElement();

                    // Сгибы
                    myXml.WriteStartElement("attribute");
                    myXml.WriteAttributeString("name", "Сгибы");
                    myXml.WriteAttributeString("value", item.Bend.ToString());
                    myXml.WriteEndElement();

                    // Толщина листового металла
                    myXml.WriteStartElement("attribute");
                    myXml.WriteAttributeString("name", "Толщина листового металла");
                    myXml.WriteAttributeString("value", item.ListThickness.ToString());
                    myXml.WriteEndElement();

                    // PaintX
                    myXml.WriteStartElement("attribute");
                    myXml.WriteAttributeString("name", "PaintX");
                    myXml.WriteAttributeString("value", item.PaintX.ToString());
                    myXml.WriteEndElement();

                    // PaintY
                    myXml.WriteStartElement("attribute");
                    myXml.WriteAttributeString("name", "PaintY");
                    myXml.WriteAttributeString("value", item.PaintY.ToString());
                    myXml.WriteEndElement();

                    // PaintZ
                    myXml.WriteStartElement("attribute");
                    myXml.WriteAttributeString("name", "PaintZ");
                    myXml.WriteAttributeString("value", item.PaintZ.ToString());
                    myXml.WriteEndElement();

                    // Версия
                    myXml.WriteStartElement("attribute");
                    myXml.WriteAttributeString("name", "Версия");
                    myXml.WriteAttributeString("value", item.LastVersion.ToString());
                    myXml.WriteEndElement();
                    myXml.WriteEndElement();// ' элемент Configuration name
                    #endregion
                }

                myXml.WriteEndElement(); // ' элемент DOCUMENT
                myXml.WriteEndElement(); // ' элемент TRANSACTION
                myXml.WriteEndElement(); // ' элемент TRANSACTIONS
                myXml.WriteEndElement(); // ' элемент XML

                myXml.Flush();
                myXml.Close();
            }
            catch (Exception ex)
            {
                MessageObserver.Instance.SetMessage("Failed to ExportPartsToXML2, with error:  " + ex.Message + Environment.NewLine + filePathForLog, MessageType.Error);
            }
        }

        private void ExportToXMLWithSubAsmbl(List<MyBomShell> list, string nameAddintion, int param)
        {
            string filePathForError = string.Empty;
            try
            {
                int currentTreeLevel;
                int helpCount = 0;
                int previousTreeLevel = 0;
                string type = null;
                bool p = false;
                bool f = false;

                int l = Convert.ToInt32(list[0].FileName.Count()) - 7;
                string fileName = list[0].FileName.Substring(0, l);



                var myXml = new XmlTextWriter(pathToSave + fileName + nameAddintion + ".xml", Encoding.UTF8);
                myXml.WriteStartDocument();
                myXml.Formatting = Formatting.Indented;

                myXml.WriteStartElement("xml");
                myXml.WriteStartElement("transactions");
                myXml.WriteStartElement("transaction");

                myXml.WriteAttributeString("vaultName", "Vents-PDM");
                myXml.WriteAttributeString("type", "");
                myXml.WriteAttributeString("date", "");

                // document
                myXml.WriteStartElement("document");
                myXml.WriteAttributeString("pdmweid", "");
                myXml.WriteAttributeString("aliasset", "Export To ERP");

                foreach (var it in list)
                {
                    filePathForError = it.FilePath;
                    currentTreeLevel = (int)it.TreeLevel + param;

                    if (helpCount != 0)
                    {
                        if (previousTreeLevel > currentTreeLevel && type == "sldasm")//переход на уровень выше
                        {
                            myXml.WriteEndElement(); //configurations 
                            myXml.WriteEndElement(); //configurations 
                            myXml.WriteEndElement(); //references
                            myXml.WriteEndElement(); //document alias 
                            f = true;
                        }
                        /*else if (previousTreeLevel < currentTreeLevel) //следующий элемент вложенный
                        {
                            
                        }*/
                        if (type == "sldasm" && it.FileType == type && previousTreeLevel == currentTreeLevel)// если две сборки подряд одного уровня
                        {
                            if (currentTreeLevel != 0)
                            {
                                // myXml.WriteEndElement();//references
                                myXml.WriteEndElement();//configurations
                                p = true;
                            }
                            else
                            {
                                myXml.WriteEndElement();//document alias
                                myXml.WriteEndElement();//references
                                myXml.WriteEndElement();//configurations
                                p = true;
                            }
                        }
                        if (type == "sldasm" && previousTreeLevel == currentTreeLevel)
                        {
                            if (p == false)
                            {
                                myXml.WriteEndElement();//configurations
                            }
                        }
                        if (currentTreeLevel == 0 && type == "sldprt")
                        {
                            if (p == false)
                            {
                                if (f == false)
                                {
                                    myXml.WriteEndElement();//document alias
                                    myXml.WriteEndElement();//references
                                    myXml.WriteEndElement();//configurations
                                }
                            }
                        }
                        p = false;
                        f = false;
                        helpCount--;
                    }

                    #region XML
                    // Конфигурация
                    myXml.WriteStartElement("configuration");
                    myXml.WriteAttributeString("name", it.Configuration);

                    // Версия
                    myXml.WriteStartElement("attribute");
                    myXml.WriteAttributeString("name", "Версия");
                    myXml.WriteAttributeString("value", it.LastVersion.ToString());
                    myXml.WriteEndElement();

                    // Масса
                    myXml.WriteStartElement("attribute");
                    myXml.WriteAttributeString("name", "Масса");
                    myXml.WriteAttributeString("value", it.Weight.ToString());
                    myXml.WriteEndElement();

                    // Наименование
                    myXml.WriteStartElement("attribute");
                    myXml.WriteAttributeString("name", "Наименование");
                    myXml.WriteAttributeString("value", it.Description);
                    myXml.WriteEndElement();

                    // Обозначение
                    myXml.WriteStartElement("attribute");
                    myXml.WriteAttributeString("name", "Обозначение");
                    myXml.WriteAttributeString("value", it.PartNumber);
                    myXml.WriteEndElement();

                    // Раздел
                    myXml.WriteStartElement("attribute");
                    myXml.WriteAttributeString("name", "Раздел");
                    myXml.WriteAttributeString("value", it.Partition.ToString());
                    myXml.WriteEndElement();

                    // ERP code
                    myXml.WriteStartElement("attribute");
                    myXml.WriteAttributeString("name", "ERP code");
                    myXml.WriteAttributeString("value", it.ErpCode.ToString());
                    myXml.WriteEndElement();

                    // Код_Материала
                    myXml.WriteStartElement("attribute");
                    myXml.WriteAttributeString("name", "Код_Материала");
                    myXml.WriteAttributeString("value", it.CodeMaterial.ToString());
                    myXml.WriteEndElement();

                    // Код Документа
                    myXml.WriteStartElement("attribute");
                    myXml.WriteAttributeString("name", "Код Документа");
                    myXml.WriteAttributeString("value", "");
                    myXml.WriteEndElement();

                    // Кол. Материала
                    myXml.WriteStartElement("attribute");
                    myXml.WriteAttributeString("name", "Кол. Материала");
                    myXml.WriteAttributeString("value", it.SummMaterial.ToString()); //it.Quantity.ToString(
                    myXml.WriteEndElement();

                    // Состояние 
                    myXml.WriteStartElement("attribute");
                    myXml.WriteAttributeString("name", "Состояние");
                    myXml.WriteAttributeString("value", "");
                    myXml.WriteEndElement();

                    // Подсчет ссылок
                    myXml.WriteStartElement("attribute");
                    myXml.WriteAttributeString("name", "Подсчет ссылок");
                    myXml.WriteAttributeString("value", it.Quantity.ToString());
                    myXml.WriteEndElement();

                    // Конфигурация
                    myXml.WriteStartElement("attribute");
                    myXml.WriteAttributeString("name", "Конфигурация");
                    myXml.WriteAttributeString("value", it.Configuration);
                    myXml.WriteEndElement();

                    // Идентификатор
                    myXml.WriteStartElement("attribute");
                    myXml.WriteAttributeString("name", "Идентификатор");
                    myXml.WriteAttributeString("value", "");
                    myXml.WriteEndElement();

                    #endregion

                    if (it.FileType == "sldasm")
                    {
                        if (currentTreeLevel == 0)
                        {
                            myXml.WriteStartElement("references");
                            myXml.WriteStartElement("document");
                            myXml.WriteAttributeString("pdmweid", "");
                            myXml.WriteAttributeString("aliasset", "Export To ERP");
                        }

                        type = "sldasm";
                    }
                    else if (it.FileType == "sldprt")
                    {
                        myXml.WriteEndElement();//configurations
                        type = "sldprt";
                    }
                    helpCount++;
                    previousTreeLevel = currentTreeLevel;

                }

                myXml.WriteEndElement(); // ' элемент DOCUMENT
                myXml.WriteEndElement(); // ' элемент TRANSACTION
                myXml.WriteEndElement(); // ' элемент TRANSACTIONS
                myXml.WriteEndElement(); // ' элемент XML

                myXml.Flush();
                myXml.Close();
            }
            catch(Exception e)
            {
                MessageObserver.Instance.SetMessage("Failed to ExportToXMLWithSubAsmbl, with error:  " + e.Message + System.Environment.NewLine + 
                    filePathForError, MessageType.Error);
            }
        }


        private List<string> GetConfigurations(string filePath)
        {
            IEdmFolder5 oFolder;
            if (filePath == null || filePath == string.Empty)
            {
                MessageObserver.Instance.SetMessage("Failed to get file from path(GetConfigurations): " + filePath?.ToString());
                return new List<string>();
            }
            var edmFile5 = vault.GetFileFromPath(filePath, out oFolder);
            if (edmFile5 == null)
            {
                MessageObserver.Instance.SetMessage("Failed to get file from path(GetConfigurations): " + filePath);
                return new List<string>();
            }

            EdmStrLst5 cfgList = edmFile5.GetConfigurations(0);

            var headPosition = cfgList.GetHeadPosition();
            List<string> configsArr = new List<string>();

            while (!headPosition.IsNull)
            {
                var configName = cfgList.GetNext(headPosition);
                if (configName != "@")
                {
                    configsArr.Add(configName);
                }
            }
            return configsArr;
        }
        private List<MyBomShell> GetBomShell(string filePath, List<string> Configurations, int BoomId)
        {
            try
            {
                List<MyBomShell> BomShellList = new List<MyBomShell>();
                if (Configurations.Count > 0)
                {
                    MyBomShell bom = null;

                
                    IEdmFolder5 oFolder;
                    IEdmFile7 EdmFile7 = (IEdmFile7)vault.GetFileFromPath(filePath, out oFolder);
                    if (EdmFile7 != null)
                    {

                        foreach (var refConfig in Configurations)
                        {
                            EdmBomView bomView = EdmFile7.GetComputedBOM(BoomId, -1, refConfig, 2);
                            if (bomView == null)
                            {
                                MessageObserver.Instance.SetMessage("Failed to get BOM: " + filePath);
                                return null;
                            }
                            object[] bomRows;
                            EdmBomColumn[] bomColumns;
                            bomView.GetRows(out bomRows);
                            bomView.GetColumns(out bomColumns);

                            for (var i = 0; i < bomRows.Length; i++)
                            {
                                List<object> eachItem = new List<object>();
                                IEdmBomCell cell = (IEdmBomCell)bomRows.GetValue(i);
                                int treeLevel = cell.GetTreeLevel();
                                for (var j = 0; j < bomColumns.Length; j++)
                                {
                                    EdmBomColumn column = (EdmBomColumn)bomColumns.GetValue(j);
                                    object value;
                                    object computedValue;
                                    string config;
                                    bool readOnly;
                                    cell.GetVar(column.mlVariableID, column.meType, out value, out computedValue, out config, out readOnly);
                                    eachItem.Add(value);
                                }
                                if (eachItem.Count > 0)
                                {
                                    bom = new MyBomShell()
                                    {
                                        Partition = eachItem[0].ToString(),
                                        PartNumber = eachItem[1].ToString(),
                                        Description = eachItem[2].ToString(),
                                        Material = eachItem[3].ToString(),
                                        CMIMaterial = eachItem[4].ToString(),
                                        ListThickness = eachItem[5].ToString(),
                                        Quantity = (eachItem[6]).ToString(),//?
                                        FileType = eachItem[7].ToString(),
                                        Configuration = refConfig,
                                        LastVersion = Convert.ToInt32(eachItem[9].ToString()),//?
                                        IdPdm = Convert.ToInt32(eachItem[10]),
                                        FileName = eachItem[11].ToString(),
                                        FilePath = eachItem[12].ToString(),
                                        ErpCode = eachItem[13].ToString(),
                                        SummMaterial = eachItem[14].ToString(),
                                        Weight = eachItem[15].ToString(),
                                        CodeMaterial = eachItem[16].ToString(),
                                        Format = eachItem[17].ToString(),
                                        Note = eachItem[18].ToString(),
                                        RefConfig = eachItem[8].ToString(),
                                        TreeLevel = treeLevel
                                    };
                                    BomShellList.Add(bom);
                                }
                                else
                                {
                                    eachItem = new List<object>();
                                }
                            }
                        }
                    }
                    else
                    {
                        MessageObserver.Instance.SetMessage("Failed get bom shell with path: " + filePath);
                    }
                }
                return BomShellList;
            }
            catch (COMException ex)
            {
                MessageObserver.Instance.SetMessage("Failed get bom shell " + (EdmResultErrorCodes_e)ex.ErrorCode + ". Укажите вид PDM или тип спецификации");
                throw ex;
            }
        }
        private List<MyBomShell> GetFullSpecification(List<MyBomShell> lAssmblyBom)
        {
            IEnumerable<MyBomShell> spec = from data in lAssmblyBom
                                         join parts in con.ViewParts
                                         on new { id = data.IdPdm, conf = data.Configuration, version = (int)data.LastVersion }
                                         equals new { id = parts.IDPDM, conf = parts.ConfigurationName, version = parts.Version }
                                         into fullSpec
                                         from f in fullSpec.DefaultIfEmpty()


                                         select new MyBomShell
                                         {
                                             CMIMaterial = data.CMIMaterial,
                                             CodeMaterial = data.CodeMaterial,
                                             Configuration = data.RefConfig,
                                             Description = data.Description,
                                             ErpCode = data.ErpCode,
                                             FileName = data.FileName,
                                             FilePath = data.FilePath,
                                             FileType = data.FileType,
                                             FolderPath = data.FolderPath,
                                             Format = data.Format,
                                             IdPdm = (f == null) ? 0 : f.IDPDM,
                                             LastVersion = (data.LastVersion == null) ? 0 : data.LastVersion,
                                             ListThickness = (f == null) ? string.Empty : f.Thickness.ToString(),
                                             Material = data.Material,
                                             Note = data.Note,
                                             ObjectType = data.ObjectType,
                                             Partition = data.Partition,
                                             PartNumber = data.PartNumber,
                                             Quantity = data.Quantity,
                                             RefConfig = data.Configuration,
                                             SummMaterial = data.SummMaterial,
                                             TreeLevel = (f == null) ? 0 : data.TreeLevel,
                                             Weight = data.Weight,
                                             Bend = (f == null) ? string.Empty : f.Bend.ToString(),
                                             PaintX = (f == null) ? string.Empty : f.PaintX.ToString(),
                                             PaintY = (f == null) ? string.Empty : f.PaintY.ToString(),
                                             PaintZ = (f == null) ? string.Empty : f.PaintZ.ToString(),
                                             DXF = (f == null) ? string.Empty : f.DXF,
                                             SurfaceArea = (f == null) ? string.Empty : f.SurfaceArea.ToString(),
                                             WorkpieceX = (f == null) ? string.Empty : f.WorkpieceX.ToString(),
                                             WorkpieceY = (f == null) ? string.Empty : f.WorkpieceY.ToString()
                                         };
            return spec.ToList();
        }

        private void AssmblAndAll_1_Level(List<MyBomShell> llAssmblyBom)
        {
            int maxAssmblLevel;
            fullDataSpecAsmblAndParts = new List<MyBomShell>();
            List<MyBomShell> listForEveryPartTemp = new List<MyBomShell>();

            GetMaxTreeLevel(out maxAssmblLevel, llAssmblyBom);

            List<List<MyBomShell>> g = new List<List<MyBomShell>>();
            List<string> namesItem = new List<string> { };
            int index = 0;
            foreach (var item in llAssmblyBom.Where(x => x.FileType == "sldasm").GroupBy(x => x.FileName))
            {
                g.Add(new List<MyBomShell>());
                namesItem.Add(item.Key);
            }

            for (int i = 0; i < (maxAssmblLevel + 1); i++)//по каждому уровню
            {
                foreach (var item in llAssmblyBom)
                {
                    if (item.FileType == "sldasm" && (item.TreeLevel == i || item.TreeLevel == (i + 1)))
                    {
                        if (item.TreeLevel == i)
                        {
                            index = namesItem.IndexOf(item.FileName);
                        }
                        g[index].Add(item);
                    }
                    else if (item.FileType == "sldprt" && (item.TreeLevel == (i + 1)))
                    {
                        g[index].Add(item);
                    }
                }
            }

            for (int i = 0; i < g.Count; i++)
            {
                ExportToXMLWithSubAsmbl(g[i], "", (0 - (int)g[i][0].TreeLevel));
            }
            g.Clear();
        }
        private void GetMaxTreeLevel(out int max, List<MyBomShell> llAssmblyBom)
        {
            List<int> list = new List<int>();

            foreach (var item in llAssmblyBom.Where(x => x.FileType == "sldasm"))
            {
                list.Add((int)item.TreeLevel);
            }
            max = list.Max();
        }

        private List<MyBomShell> AssmblAndAllDetails(List<MyBomShell> lAssmblyBom)
        {
            fullDataSpecAsmblAndParts = new List<MyBomShell>();
            foreach (var item in lAssmblyBom)
            {
                if (item.FileType == "sldasm" && item.TreeLevel == 0)
                {
                    fullDataSpecAsmblAndParts.Add(item);
                }
                else if (item.FileType.Equals("sldprt"))
                {
                    fullDataSpecAsmblAndParts.Add(item);
                }
            }
            return fullDataSpecAsmblAndParts;
        }


        public void XML()
        {
            if (filePath.ToUpper().Contains("SLDPRT"))
            {
                ExportPartsToXML();
            }
            else
            {
                AssmblAndAll_1_Level(this.AssmblyBom);

                ExportToXMLWithSubAsmbl(AssmblAndAllDetails(this.AssmblyBom), " Parts", 0);

                ExportParts(this.fullDataSpecParts);
            }
        }



    }

    public class EdmVaultSingleton
    {
        private static EdmVault5 mInstance = null;
        private static object mLockObj = new object();

        public static EdmVault5 Instance
        {
            get
            {
                try
                {
                    if (mInstance == null)
                    {
                        Monitor.Enter(mLockObj);
                        if (mInstance == null)
                        {
                            mInstance = new EdmVault5();
                        }
                        Monitor.Exit(mLockObj);
                    }
                }
                catch (Exception ex)
                {
                    Monitor.Exit(mLockObj);
                }
                return mInstance;
            }
        }
    }    
}
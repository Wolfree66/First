//#define server
//#define test


namespace ArchiveInventoryBook
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using TFlex.DOCs.Model;
    using TFlex.DOCs.Model.Macros;
    using TFlex.DOCs.Model.References;
    using TFlex.DOCs.Model.Stages;

    using System.Windows.Forms;
    using System.Diagnostics;
    using System.Text;
    using TFlex.DOCs.Model.Structure;
    using TFlex.DOCs.Model.References.Files;
    using TFlex.DOCs.Model.Desktop;
    using System.IO;
    using TFlex.DOCs.Model.Signatures;
    using TFlex.DOCs.Model.Search;
    using TFlex.DOCs.Model.References.Users;
    using TFlex.DOCs.Model.Classes;
    using TFlex.DOCs.Model.Macros.ObjectModel;
    using System.Threading;
#if !server

    using TFlex.DOCs.UI.Objects.Managers;
    using Xceed.Words.NET;
    using TFlex.DOCs.UI;
    using TFlex.DOCs.UI.Objects.ReferenceModel;

#endif

    public class Macro : MacroProvider
    {
        public Macro(MacroContext context)
            : base(context)
        {
        }

        public override void Run()
        {
        }

        public void Test()
        {
            НайтиДубликатыОбозначений();
            return;
            ReferenceObject current = DocumentRegistryCards.Reference.Find(1790);
            //ReferenceObject current2 = InventoryBook.Reference.Find(6929);
            //ReferenceObject current = Context.ReferenceObject;
            //bool result = IsFilesIdentical("D:\\Ддокументы\\ТЗ на внедрение T-FLEX - копия.pdf", "D:\\Ддокументы\\Список файлов.docx");
            
            //FileReferenceObject current = Context.ReferenceObject as FileReferenceObject;
            //FolderObject hranFiles = References.FileReference.FindByPath("Хранилище файлов") as FolderObject;
            //MoveFileToFolder(current, hranFiles);
            //CombineFilesInOriginalsAndTemporaryFolder();

            //string stage = current.SystemFields.Stage.Stage.Name;
            DocumentRegistryCard documentRegistryCard = DocumentRegistryCards.GetDocumentRegistryCard(current);// new InventoryBookRecord(current);
            UploadAllFilesAndCreateActOfTransfer(documentRegistryCard);
            //bool res = AllLinkedFilesWithCorrectName(documentRegistryCard);
            //ChangeToSafeKeeping(documentRegistryCard);
            return;
            //CheckUniqueAndCreateInventoryNumber(documentRegistryCard.InventoryBook);
            //DocumentRegistryCards.CreateNewNumberForCode(documentRegistryCard.InventoryBook.Code).ToString();
            //InventoryBookRecord inventoryBookRecord2 = InventoryBook.GetInventoryBookRecord(current2);// new InventoryBookRecord(current);
            //bool equal = inventoryBookRecord.Equals(inventoryBookRecord2);
            //MessageBox.Show("GetInventoryBook");

            //            var book = GetInventoryBook(documentRegistryCard);
            var book = new InventoryBook(documentRegistryCard.ReferenceObject.GetObject(new Guid("480933a3-681c-4fad-8dc0-20c2835f3478")));
            string code = book.Code;
            string proj = book.Name;
            MessageBox.Show(code + "\n" + proj);
            return;

            //bool equald = inventoryBookRecord == inventoryBookRecord2;
            //List<InventoryBookRecord> list = new List<ArchiveInventoryBook.InventoryBookRecord>() { inventoryBookRecord, inventoryBookRecord2 };
            //List<InventoryBookRecord> listd = list.Distinct().ToList();
            //            UploadAllFilesAndCreateActOfTransfer(inventoryBookRecord2);
            //           return;
            var Users = НайтиПользователейПоФамилии(documentRegistryCard.DesignerFullName);
            if (Users.Count == 1)
            {
                if (documentRegistryCard.DesignerFullName != Users.First().ToString())
                {
                    documentRegistryCard.DesignerObject = Users.First();

                }
            }
            string invNum = documentRegistryCard.RegistrationNumber.ToString();
            invNum = RemoveSpaces(invNum);
            documentRegistryCard.SetInventoryNumber(new ArchiveInventoryBook.InventoryNumber(invNum));
            documentRegistryCard.Save();
            //var components = InventoryBook.GetComponentsFor(inventoryBookRecord.RegistrationNumber);
            //FindFilesInFolderSatisfyDesignationTerms(References.FileReference.FindByPath(InventoryBook.FolderName.OriginalsFullPath) as FolderObject, inventoryBookRecord.Designation);
            //int countFiles = inventoryBookRecord.Files.Count();
            //ChangeToSafeKeeping(inventoryBookRecord);

            //List<FileObject> files = new List<FileObject>();
            //FileReference fileReference = new FileReference(Connection);
            //files.Add(fileReference.FindByPath("Архив документации\\Временная\\1.txt") as FileObject);
            //files.Add(fileReference.FindByPath("Архив документации\\1.txt") as FileObject);
            //CombineFilesWithVersions(files);

            //FileObject folderOriginals = fileReference.FindByPath("Архив документации\\Временная\\СКАН_ВГИП.467862.122.pdf") as FileObject;
            //int num = folderOriginals.SystemFields.Version;
            //Desktop.GetObjectChangelists(folderOriginals, false).Where(c=> c.IsAdded || c.IsUpdated);
            //files.Add(folderOriginals);
            //MoveFilesToOriginalsFolder(files);
        }


        public void НайтиДубликатыОбозначений()
        {

            DocumentRegistryCards.Reference.Objects.Load();
            List<ReferenceObject> allObjects = DocumentRegistryCards.Reference.Objects.ToList();
            allObjects.AddRange(DocumentRegistryCards.Reference.GetDeletedObjects());
            MessageBox.Show("allObjects.Count - " + allObjects.Count.ToString());
            Dictionary<string, List<string>> dictDesignToInvNumb = new Dictionary<string, List<string>>();
            List<ReferenceObject> listObjectsToDelete = new List<ReferenceObject>();
            int numDeletedFiles = 0;
            StringBuilder result = new StringBuilder();
            foreach (var item in allObjects)
            {
                DocumentRegistryCard drc = DocumentRegistryCards.GetDocumentRegistryCard(item);
                List<string> listInvNum;
                string designation = drc.Designation;
                if (string.IsNullOrWhiteSpace(designation)) continue;
                designation = RemoveSpaces(designation);
                if (dictDesignToInvNumb.TryGetValue(designation, out listInvNum))
                {
                    if (item.IsInRecycleBin || item.IsDeleted)
                    {
                        //MessageBox.Show("item.IsInRecycleBin - " + item.IsInRecycleBin.ToString() + "\n item.IsDeleted - " + item.IsDeleted.ToString());
                        listObjectsToDelete.Add(item);
                        numDeletedFiles++;
                    }
                    else
                    {
                        //result.AppendLine(designation + " - есть дубликат");
                        //Console.WriteLine(designation + " - есть дубликат");
                        //MessageBox.Show("Обозначение = "+designation + "\nЕсть неудалённый дубликат по Guid - " + item.SystemFields.Guid.ToString());
                        listInvNum.Add(drc.RegistrationNumber);
                        dictDesignToInvNumb[designation] = listInvNum;
                    }
                }
                else
                {
                    listInvNum = new List<string>() { drc.RegistrationNumber };
                    dictDesignToInvNumb.Add(designation, listInvNum);
                }
            }
            MessageBox.Show("dictGuids.Count - " + dictDesignToInvNumb.Count.ToString() + "\nlistObjectsToDelete.Count - " + listObjectsToDelete.Count.ToString());
            //Console.ReadKey();
            foreach (var item in dictDesignToInvNumb.Where(kvp => kvp.Value.Count > 1))
            {
                foreach (var card in item.Value)
                {
                    result.AppendLine(item.Key + "\t" + card);
                }
            }
            Clipboard.Clear();
            Clipboard.SetText(result.ToString());
            //foreach (var item in listObjectsToDelete)
            //{
            //    Desktop.ClearRecycleBin(item); // Удаляем из корзины
            //}

        }

        private InventoryBook GetInventoryBook(DocumentRegistryCard documentRegistryCard)
        {
            return InventoryBooksReference.GetInventoryBookForDocumentCard(documentRegistryCard);

        }

        private string RemoveSpaces(string inputString)
        {
            inputString = inputString.Replace("  ", string.Empty);
            inputString = inputString.Trim().Replace(" ", string.Empty);

            return inputString;
        }

        public List<User> НайтиПользователейПоФамилии(string lastName)
        {
            ParameterInfo lastNameParameter = References.UsersReference.ParameterGroup.OneToOneParameters.Find(User.Fields.LastName);
            List<ReferenceObject> users = References.UsersReference.Find(lastNameParameter, lastName);
            return users.Select(ro => ro as User).ToList();
        }

        public void CombineFilesInOriginalsAndTemporaryFolder()
        {
            //получаем все файлы из папки временная
            string temporaryFolderPath = DocumentRegistryCards.FolderName.TemporaryFullPath;
            FolderObject folderTemporary = References.FileReference.FindByPath(temporaryFolderPath) as FolderObject;
            Dictionary<string, List<FileObject>> dictNameToFiles = new Dictionary<string, List<FileObject>>();

            AddFilesInFolderToDictionaryByName(folderTemporary, dictNameToFiles);
            //получаем все файлы из папки оригиналы
            string originalsFolderPath = DocumentRegistryCards.FolderName.ScansFullPath;
            FolderObject originalsFolder = References.FileReference.FindByPath(originalsFolderPath) as FolderObject;

            //AddFilesInFolderToDictionaryByName(originalsFolder, dictNameToFiles);
            //получаем из списков файлов список уникальных наименований
            //выбираем по наименованию файлы из списков файлов
            //если файлов больше одного то объединяем их в один, с учётом версий
            int i = 0;
            foreach (var kvp in dictNameToFiles)
            {
                i++;
                FileObject resultFile;
                if (kvp.Value.Count > 1)
                {
                    resultFile = CombineFilesWithVersions(kvp.Value);
                    //int num = dictNameToFiles.e

                }
                else
                {
                    resultFile = kvp.Value.First();

                }
                MoveFileToFolder(resultFile, folderTemporary);
            }
        }

        private void MoveFileToFolder(FileObject resultFile, FolderObject folderTemporary)
        {
            if (resultFile != null)
                MoveFileReferenceObjectToFolder(resultFile, folderTemporary);
            else throw new NullReferenceException("Ссылка на файл равна null");
        }

        private void MoveFileReferenceObjectToFolder(FileReferenceObject file, FolderObject folder)
        {
            if (file.Parent == folder) return;
            if (!file.IsCheckedOut) file.CheckOut(false); else Error("Файл на редактировании"); // новый файл
            file.BeginChanges();
            file.SetParent(folder);
            file.EndChanges();
            Desktop.CheckIn(file, "Перемещён в паку - " + folder.Path.ToString(), false);
        }

        /// <summary>
        /// Собирает в самый старый файл все версии остальных файлов из списка
        /// </summary>
        /// <param name="listFiles"></param>
        /// <returns></returns>
        private FileObject CombineFilesWithVersions(List<FileObject> listFiles)
        {
            listFiles.Sort(SortByCreationDateAscending());
            //для каждого файла получаем список изменений
            //собираем список изменений по дате изменения
            List<ChangelistObject> listChanges = new List<ChangelistObject>();
            foreach (var file in listFiles)
            {
                listChanges.AddRange(Desktop.GetObjectChangelists(file, false).Where(c => c.IsAdded || c.IsUpdated).GroupBy(p => p.Version)
  .Select(g => g.First())
  .ToList());
                listChanges = listChanges.OrderBy(c => c.Changelist.Date).ToList();
            }
            //заливаем изменения в самый старый файл в виде версий
            var result = AddChangesToFile(listFiles.First(), listChanges);
            //убираем его из списка
            listFiles.Remove(result);
            //удаляем все оставшиеся в списке файлы
            foreach (var item in listFiles)
            {

                Desktop.CheckOut(item, true); // Удаление
                Desktop.CheckIn(item, "Удаление копии", false);
                Desktop.ClearRecycleBin(item);
            }
            return result;
        }

        private FileObject AddChangesToFile(FileObject fileObject, List<ChangelistObject> listChanges)
        {
            IEnumerable<ChangelistObject> listExistedChanges = Desktop.GetObjectChangelists(fileObject, false).Where(c => c.IsAdded || c.IsUpdated);
            ChangelistObject lastChange = listChanges.Last();
            foreach (var change in listChanges)
            {
                //проверяем на совпадение файла с предыдущей версией

                if (listExistedChanges.Contains(change))
                {
                    bool thisVersionShouldBeActual = (change.Id == lastChange.Id && change.Version == lastChange.Version);
                    if (thisVersionShouldBeActual)
                    {
                        Desktop.CheckOutVersion(fileObject, change.Version);
                        Desktop.CheckIn(fileObject, "Актуализация, версия от - " + change.Changelist.Date.ToString(), false);
                    }
                    else
                        continue;
                }
                else
                {

                    fileObject.GetHeadRevision();
                    string filePath = fileObject.LocalPath;
                    FileObject file = change.ReferenceObject as FileObject;
                    string newVersionFilePath = file.GetFileVersion(change.Version);
                    //проверяем на совпадение файла с текущей версией (обновляется при заливке новой версии)
                    if (!IsFilesIdentical(filePath, newVersionFilePath))
                    {

                        if (!fileObject.IsCheckedOut) fileObject.CheckOut(false); else Error("Файл на редактировании");  // старый файл
                        File.Copy(newVersionFilePath, fileObject.LocalPath, true); // Обновляем старый файл

                        Desktop.CheckIn(fileObject, "Обновление, версия от - " + change.Changelist.Date.ToString(), false); // Применяем новую версию
                    }
                    File.Delete(newVersionFilePath);
                }

            }
            return fileObject;
        }


        /// <summary>
        /// проверяет идентичность файлов по содержимому
        /// </summary>
        /// <param name="xFilePath"></param>
        /// <param name="yFilePath"></param>
        /// <returns></returns>
        private static bool IsFilesIdentical(string xFilePath, string yFilePath)
        {
            bool isIdentical = true;
            FileInfo xFI = new FileInfo(xFilePath);
            FileInfo yFI = new FileInfo(yFilePath);
            //проверяем по размеру файлов
            //if (xFI.Length != yFI.Length) return false;
            //проверяем по содержимому
            using (FileStream xStreamReader = xFI.OpenRead())
            {
                using (FileStream yStreamReader = yFI.OpenRead())
                {
                    long currentPosition = 0;
                    while (currentPosition < xStreamReader.Length)
                    {
                        //Console.ReadKey();
                        currentPosition = xStreamReader.Position;
                        byte[] xB = new byte[1024];
                        byte[] yB = new byte[1024];
                        UTF8Encoding temp = new UTF8Encoding(true);
                        xStreamReader.Read(xB, 0, xB.Length);
                        yStreamReader.Read(yB, 0, yB.Length);

                        for (int i = 0; i < xB.Length; i++)
                        {
                            isIdentical = xB[i] == yB[i];
                            if (!isIdentical) return false;
                        }
                        //Console.WriteLine(isIdentical.ToString());
                        //Console.WriteLine(temp.GetString(xB));
                        //Console.WriteLine(temp.GetString(yB));

                        currentPosition = xStreamReader.Position;
                    }
                    //if (xStreamReader..l)
                    //MD5 md = MD5.Create();
                    //md.ComputeHash(xStreamReader.);
                }
            }
            return isIdentical;
        }

        /// <summary>
        /// Сортирует список объектов справочника по возрастанию даты создания
        /// </summary>
        /// <returns></returns>
        private static Comparison<ReferenceObject> SortByCreationDateAscending()
        {
            return delegate (ReferenceObject x, ReferenceObject y)
            {
                //      	if (x.SystemFields.CreationDate == null && y.SystemFields.CreationDate == null) return 0;
                if (x.SystemFields.CreationDate == null) return -1;
                else if (y.SystemFields.CreationDate == null) return 1;
                else return x.SystemFields.CreationDate.CompareTo(y.SystemFields.CreationDate);
            };
        }

        private void AddFilesInFolderToDictionaryByName(FolderObject folder, Dictionary<string, List<FileObject>> dictNameToFiles)
        {
            var folderTemporaryChildren = GetAllChildFolders(folder);
            folderTemporaryChildren = folderTemporaryChildren.Concat(new List<ReferenceObject> { folder });
            foreach (var item in folderTemporaryChildren)
            {
                var filesInFolder = item.Children.Where(c => c as FileObject != null && c.ToString().ToLower().EndsWith(".pdf"));
                if (filesInFolder == null || filesInFolder.Count() == 0) continue;
                foreach (var file in filesInFolder)
                {
                    List<FileObject> files = null;
                    string fileName = file.ToString().ToLower();
                    if (dictNameToFiles.TryGetValue(fileName, out files))
                    {
                        files.Add(file as FileObject);
                        //dictNameToFiles[fileName] = files;
                    }
                    else
                    {
                        dictNameToFiles.Add(fileName, new List<FileObject> { file as FileObject });
                    }
                }
            }
        }

        #region События справочника
        public void Событие_ПроверитьИНаХранение()
        {
            Thread.Sleep(3000);
            ReferenceObject current = Context.ReferenceObject;
            DocumentRegistryCard inventoryBookRecord = DocumentRegistryCards.GetDocumentRegistryCard(current);// new InventoryBookRecord(current);
            string errors = CanBeKeeping(inventoryBookRecord);
            if (string.IsNullOrWhiteSpace(errors))
            {
                if (ChangeStage(inventoryBookRecord, DocumentRegistryCard.StageNames.CheckDocument)) ;
                Sign(inventoryBookRecord, DocumentRegistryCard.SignNames.Accept);
            }
            else
            {
#if !server
                MessageBox.Show("Проверка не пройдена:\n" + errors, "Проверка не пройдена", MessageBoxButtons.OK, MessageBoxIcon.Warning);
#else

#endif
            }
        }



        public void Событие_ПереводВОбработку()
        {
            ReferenceObject current = Context.ReferenceObject;
            DocumentRegistryCard inventoryBookRecord = DocumentRegistryCards.GetDocumentRegistryCard(current);// new InventoryBookRecord(current);
            ChangeStage(inventoryBookRecord, DocumentRegistryCard.StageNames.Edit);
        }

        public void Событие_ИзменениеСтадииНаПроверкаДокументов()
        {

            ReferenceObject current = Context.ReferenceObject;
            DocumentRegistryCard inventoryBookRecord = DocumentRegistryCards.GetDocumentRegistryCard(current);// new InventoryBookRecord(current);
            ChangeToSafeKeeping(inventoryBookRecord);
        }

        public void Событие_ПроверитьУникальность()
        {
            ReferenceObject current = Context.ReferenceObject;
            DocumentRegistryCard inventoryBookRecord = DocumentRegistryCards.GetDocumentRegistryCard(current);
            IsDesignationAndInventaryNumberUnique(inventoryBookRecord);
        }

        public void Событие_ПрисвоитьИнвентарныйНомерИСохранить()
        {
            ReferenceObject current = Context.ReferenceObject;
            DocumentRegistryCard documentRegistryCard = DocumentRegistryCards.GetDocumentRegistryCard(current);
            CheckUniqueAndCreateInventoryNumber(documentRegistryCard);
        }

#if !server
        public void Событие_ОбъединитьВыделенныеФайлыСУчётомВерсий()
        {
            List<ReferenceObject> listRO = ПолучитьТекущиеОбъекты();
            //MessageBox.Show(listRO.Count.ToString());
            if (listRO == null) return;
            List<FileObject> listFile = listRO.Where(ro => ro as FileObject != null).Select(ro => ro as FileObject).ToList();
            CombineFilesWithVersions(listFile);
        }
#endif
        public void Событие_ПрисоединитьВсеФайлыССовпадающимОбозначениемИзПапкиВременнаяИОригиналы()
        {
            ReferenceObject current = Context.ReferenceObject;
            DocumentRegistryCard inventoryBookRecord = DocumentRegistryCards.GetDocumentRegistryCard(current);// new InventoryBookRecord(current);
            AddFiles(inventoryBookRecord, false);
        }

        public void НажатиеНаКнопку_ПрисоединитьФайлы()
        {
            ReferenceObject current = Context.ReferenceObject;
            DocumentRegistryCard inventoryBookRecord = DocumentRegistryCards.GetDocumentRegistryCard(current);// new InventoryBookRecord(current);
            AddFiles(inventoryBookRecord);
        }

        private void AddFiles(DocumentRegistryCard inventoryBookRecord, bool showMessages = true)
        {
            FolderObject temporaryFolder = References.FileReference.FindByPath(DocumentRegistryCards.FolderName.TemporaryFullPath) as FolderObject;
            FolderObject originalsFolder = References.FileReference.FindByPath(DocumentRegistryCards.FolderName.ScansFullPath) as FolderObject;
            //находим все файлы в папке Временная с подходящим обозначением
            string designation = inventoryBookRecord.Designation;
            if (String.IsNullOrWhiteSpace(designation))
            {
#if !server
                if (showMessages)
                    MessageBox.Show("Обозначение не задано!\nУкажите обозначение или присоедините файлы вручную.", "Обозначение не задано!", MessageBoxButtons.OK, MessageBoxIcon.Error);
#endif
                return;
            }
            if (!IsTheSameDesignationExist(inventoryBookRecord))
            {
#if !server
                if (showMessages)
                    MessageBox.Show("Обозначение не уникально!\nУкажите другое обозначение.", "Обозначение не уникально!", MessageBoxButtons.OK, MessageBoxIcon.Error);
#endif
                return;
            }
            List<FileObject> foundFiles = FindFilesInFolderSatisfyDesignationTerms(temporaryFolder, designation);
            foundFiles.AddRange(FindFilesInFolderSatisfyDesignationTerms(originalsFolder, designation));
            List<FileObject> needToLinkFiles = foundFiles.Where(f => !inventoryBookRecord.Files.Select(af => af.FileObject).Contains(f)).ToList();
            if (needToLinkFiles.Count > 0)
            {
                string result = "";
                foreach (var item in needToLinkFiles)
                {
                    //присоединяем их к инвентарной записи
                    inventoryBookRecord.AddFile(item);
                    result += "\n" + item.Name.ToString();
                }
                inventoryBookRecord.Save();
                //MoveFilesToOriginalsFolder(inventoryBookRecord.Files);
                //ChangeStage(inventoryBookRecord, InventoryBookRecord.StageNames.SafeKeeping);
#if !server
                if (showMessages)
                    MessageBox.Show("Были присоединены следующие файлы:" + result, "Файлы присоединены");
            }
            else
            {
                if (showMessages)
                    MessageBox.Show("Подходящих файлов в папке Временная, не найдено!", "Файлов нет!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
#endif
            }
        }

#if !server

        public void ВыгрузитьВсеФайлыВнизПоДеревуПрименяемостиИСоздатьАктОПередаче()
        {
            ReferenceObject current = Context.ReferenceObject;
            DocumentRegistryCard inventoryBookRecord = DocumentRegistryCards.GetDocumentRegistryCard(current);// new InventoryBookRecord(current);

            UploadAllFilesAndCreateActOfTransfer(inventoryBookRecord);
        }

#endif

        #region Методы  
#if !server
        private List<ReferenceObject> ПолучитьТекущиеОбъекты()
        {
            Guid referenceGuid = References.FileReference.ParameterGroup.ReferenceInfo.Guid; // ТекущийОбъект.Справочник.УникальныйИдентификатор;
            UIMacroContext uiContext = Context as UIMacroContext;
            //MessageBox.Show(referenceGuid.ToString());
            if (uiContext == null)
            {
                //MessageBox.Show("uiContext == null");
                return null;
            }

            var vrs = uiContext.FindReferenceVisualRepresentations(UIMacroContext.FindReferenceVisualRepresentationsType.CurrentWindow);

            if (vrs == null)
            {
                // MessageBox.Show("vrs == null");
                return null;
            }

            var visualRepresentation = vrs.FirstOrDefault(vr => vr.Reference.ParameterGroup.Guid == referenceGuid);

            //foreach (var item in vrs)
            //{
            //    MessageBox.Show(item.ToString());
            //}

            if (visualRepresentation == null)
            {
                //MessageBox.Show("visualRepresentation == null");
                return null;
            }

            return visualRepresentation.GetSelectedObjects()
                .OfType<ReferenceUIObject>()
                .Select(uiObject => uiObject.ReferenceObject).ToList();
        }

#endif

        private string CanBeKeeping(DocumentRegistryCard card)
        {
            StringBuilder errors = new StringBuilder();
            //errors.Append(CanBeKeeping(card));
            if (!IsTheSameDesignationExist(card)) errors.AppendLine("Карточка с таким обозначением документа - " + card.Designation + ", уже существует.");
            //Проверка на редактирование файлов
            foreach (var file in card.Files)
            {
                if (file.FileObject.IsCheckedOut) errors.AppendLine(string.Format("Прикреплённый файл {0}- , взят на редактирование пользователем - {1}.", file.Name, file.FileObject.SystemFields.ClientView.UserName));
            }
            //TODO: проверка на наличие аннулированных документов
            if (!AllLinkedFilesWithCorrectName(card)) errors.AppendLine("Не все подключённые файлы, корректно названы или расположены.");
            return errors.ToString();
        }


        public bool IsTheSameDesignationExist(DocumentRegistryCard card)
        {
            if (card.Designation == "") return false;
            List<ReferenceObject> list = DocumentRegistryCards.GetAllObjectsWithDesignation(card.Designation);
            return list == null || list.Count == 0 || (list.Count == 1 && list.First().SystemFields.Id == card.ReferenceObject.SystemFields.Id);
        }

        /// <summary>
        /// все присоединённые файлы должны соответствовать процессу
        /// </summary>
        /// <param name="inventoryBookRecord"></param>
        /// <returns></returns>
        private bool AllLinkedFilesWithCorrectName(DocumentRegistryCard card)
        {
            if (card.Files.Count() == 0) return false;
            Dictionary<string, List<ArchiveFile>> dictFolderToFiles = new Dictionary<string, List<ArchiveInventoryBook.ArchiveFile>>();
            foreach (var file in card.Files)
            {
                List<ArchiveFile> filesInFolder;
                if (dictFolderToFiles.TryGetValue(file.FolderName, out filesInFolder))
                {
                    filesInFolder.Add(file);
                }
                else
                {
                    filesInFolder = new List<ArchiveInventoryBook.ArchiveFile> { file };
                    dictFolderToFiles.Add(file.FolderName, filesInFolder);
                }
            }

            //все файлы лежат в папках временная, подлинники, технологические файлы
            foreach (string folderName in dictFolderToFiles.Keys)
            {
                if (folderName != DocumentRegistryCards.FolderName.Temporary &&
                    folderName != DocumentRegistryCards.FolderName.Scans &&
                    folderName != DocumentRegistryCards.FolderName.TechnologicalFiles)
                    return false;
            }
            //наименование файлов из технологические файлы не проверяем
            //наименование файлов из подлинники не проверяем (они всё равно будут переименованы)
            //во временной папке должны лежать только .pdf
            List<ArchiveFile> filesPDFInTemporary;
            if (!dictFolderToFiles.TryGetValue(DocumentRegistryCards.FolderName.Temporary, out filesPDFInTemporary)) return true;

            foreach (var item in filesPDFInTemporary)
            {
                if (!item.Name.ToLower().EndsWith(".pdf")) return false;
            }
            //если файл один, то всё нормально
            if (filesPDFInTemporary.Count == 1) return true;
            //если файлов во временной несколько, и это не разные листы именованные согласно новой конвенции
            List<IntervalPageNumbers> intervals = new List<ArchiveInventoryBook.IntervalPageNumbers>();
            foreach (var item in filesPDFInTemporary)
            {
                intervals.Add(ScanFileNameConvention.GetPageNumbersInterval(item.Name));
            }
            return IsIntervalsContinuous(intervals);
        }

        private static bool IsIntervalsContinuous(List<IntervalPageNumbers> intervals)
        {
            intervals.Sort((x, y) =>
                x.StartPage.CompareTo(y.StartPage));
            int currentPage = 1;
            foreach (var interval in intervals)
            {
                if (interval.StartPage != currentPage) return false;
                //присваиваем следующий за интервалом номер
                currentPage = interval.EndPage + 1;
            }
            return true;
        }

        private void CheckUniqueAndCreateInventoryNumber(DocumentRegistryCard documentRegistryCard)
        {
            //MessageBox.Show("documentRegistryCard");
            if (string.IsNullOrWhiteSpace(documentRegistryCard.RegistrationNumber.ToString()))
                if (IsDesignationUnique(documentRegistryCard))
                {
                    //MessageBox.Show("IsDesignationUnique = true");
                    InventoryBook inventoryBook = documentRegistryCard.InventoryBook;// InventoryBooksReference.GetInventoryBookForDocumentCard(documentRegistryCard);
                    documentRegistryCard.SetInventoryNumber(DocumentRegistryCards.CreateNewNumberForCode(inventoryBook.Code));
                    //MessageBox.Show("Save");
                    documentRegistryCard.SaveToDB();
                    MessageBox.Show(string.Format("Инвентарный номер - {0} сохранён", documentRegistryCard.RegistrationNumber.ToString()));
                }
        }

#if !server
        private void UploadAllFilesAndCreateActOfTransfer(DocumentRegistryCard inventoryBookRecord)
        {
            StringBuilder log = new StringBuilder();
#if !test
            string folderPathForUpLoad = ChooseLocalFolder();
#else
            string folderPathForUpLoad = "D:\\проба\\";
#endif
            if (folderPathForUpLoad == null) return;
            if (System.IO.Directory.GetFiles(folderPathForUpLoad).Length > 0)
            {
                DialogResult dialogResult = MessageBox.Show(string.Format("Папка не пуста, удалить все файлы из папки?", folderPathForUpLoad), string.Format("Удалить файлы из папки - {0}?", folderPathForUpLoad), MessageBoxButtons.YesNoCancel);
                if (dialogResult == DialogResult.Yes)
                    ClearFilesFromFolder(folderPathForUpLoad);
                else if (dialogResult == DialogResult.Cancel) return;
            }
            if (folderPathForUpLoad == null) return;
            folderPathForUpLoad += "\\";
            bool isChangeNotificationRecord = false;
            List<DocumentRegistryCard> allInventoryRecords = new List<DocumentRegistryCard>();
            if ((inventoryBookRecord as DesignDocumentRegistryCard) != null ||
                (inventoryBookRecord as TechnologicalDocumentRegistryCard) != null)
            {
                allInventoryRecords = RecursiveGetInventoryRecordComponents(inventoryBookRecord);
            }
            else if (inventoryBookRecord as ChangeNotificationDocumentRegistryCard != null)
            {
                isChangeNotificationRecord = true;
                ChangeNotificationDocumentRegistryCard inventoryBookRecordChangeNotification = inventoryBookRecord as ChangeNotificationDocumentRegistryCard;
                foreach (var item in inventoryBookRecordChangeNotification.ChangedDocuments)
                {
                    allInventoryRecords.AddRange(RecursiveGetInventoryRecordComponents(item));
                }
            }
            allInventoryRecords = allInventoryRecords.Distinct().ToList();
            bool upLoadHaveErrors = false;
            string pathFile = "";
            int j = allInventoryRecords.Count;
            WaitingHelper.Wait(null, "Выгрузка файлов", true, new Action<int>((k) =>
            {
                UploadFiles(log, folderPathForUpLoad, allInventoryRecords, out upLoadHaveErrors, out pathFile);
            }), j);
            if (upLoadHaveErrors && MessageBox.Show("При выгрузке файлов возникли ошибки! Открыть лог ошибок?", "Ошибки при выгрузке файлов!", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.OK)
            {
                Process.Start(pathFile);
                //return;
            }
            string pathActFile = folderPathForUpLoad + "Акт передачи.docx"; ;
            if (isChangeNotificationRecord)
            {
                CreateActOfTransfer(pathActFile, inventoryBookRecord as ChangeNotificationDocumentRegistryCard, allInventoryRecords);
            }
            else
                CreateActOfTransfer(pathActFile, allInventoryRecords);
            Process.Start(pathActFile);
        }

        private string CreateActOfTransfer(string filePath, ChangeNotificationDocumentRegistryCard inventoryBookRecordChangeNotification, List<DocumentRegistryCard> allInventoryRecords)
        {
            var docX = DocX.Create(filePath, Xceed.Words.NET.DocumentTypes.Document);

            //добавляем сегодняшнюю дату
            docX.InsertParagraph(DateTime.Today.Date.ToShortDateString()).Alignment = Alignment.right;
            //добавляем пустые строки
            docX.InsertParagraph();
            docX.InsertParagraph();
            //Добавляем название акта
            Formatting format = new Formatting();
            format.Bold = true;
            format.UnderlineStyle = UnderlineStyle.singleLine;
            format.Size = 16;
            docX.InsertParagraph("Акт передачи документов", false, format).Alignment = Alignment.center;
            docX.InsertParagraph(inventoryBookRecordChangeNotification.ProductIndex, false, format).Alignment = Alignment.center;
            docX.InsertParagraph(string.Format("По извещению {0}", inventoryBookRecordChangeNotification.Designation), false, format).Alignment = Alignment.center;
            //добавляем пустые строки
            docX.InsertParagraph();
            AddTableAndFooterToDocx(allInventoryRecords, docX);
            docX.Save();
            return filePath;
        }

        private void UploadFiles(StringBuilder log, string folderPathForUpLoad, List<DocumentRegistryCard> allInventoryRecords, out bool upLoadHaveErrors, out string pathFile)
        {
            int count = allInventoryRecords.Count;
            int index = 1;
            upLoadHaveErrors = false;
            foreach (var item in allInventoryRecords)
            {
#if test
                Thread.Sleep(5000);
#endif
                WaitingHelper.SetText(string.Format("({0}/{1}) {2}", index++, count, item.Designation));
                if (!WaitingHelper.NextStep())
                {
                    upLoadHaveErrors = true;
                    log.AppendLine("Выгрузка прервана");
                    SaveLog(log, folderPathForUpLoad);
                    break;
                }
                if (item.Stage.Name != DocumentRegistryCard.StageNames.SafeKeeping.Name) upLoadHaveErrors = true;
                bool isNoFileUpload = true;
                
                foreach (var file in item.Files)
                {
                    isNoFileUpload = false;
                    string uploadResult = UploadVersionFileToLocalFolder(folderPathForUpLoad, file);
                    log.AppendLine(item.Stage.Name + "\t" + uploadResult);
                    if (!upLoadHaveErrors) upLoadHaveErrors = uploadResult.StartsWith("Fail");
                }
                if (isNoFileUpload)
                {
                    upLoadHaveErrors = true;
                    log.AppendLine(string.Format("{0} - файл не найден.", item.Designation));
                }
            }
            pathFile = SaveLog(log, folderPathForUpLoad);
        }
#endif
        private string SaveLog(StringBuilder log, string folderPathForUpLoad)
        {
            string pathFile;
            string logFileName = "!Log_UploadFiles.txt";
            //if(upLoadHaveErrors) logFileName = "!Log_Failed.txt";

            pathFile = folderPathForUpLoad + logFileName;
            SaveLog(pathFile, log.ToString());
            return pathFile;
        }

#if !server
        private string CreateActOfTransfer(string filePath, List<DocumentRegistryCard> allInventoryRecords)
        {
            var docX = DocX.Create(filePath, Xceed.Words.NET.DocumentTypes.Document);

            //добавляем сегодняшнюю дату
            docX.InsertParagraph(DateTime.Today.Date.ToShortDateString()).Alignment = Alignment.right;
            //добавляем пустые строки
            docX.InsertParagraph();
            docX.InsertParagraph();
            //Добавляем название акта
            Formatting format = new Formatting();
            format.Bold = true;
            format.UnderlineStyle = UnderlineStyle.singleLine;
            format.Size = 16;
            docX.InsertParagraph("Акт передачи документов АК КТС - 35", false, format).Alignment = Alignment.center;
            docX.InsertParagraph("По извещению ВГИП.436-17", false, format).Alignment = Alignment.center;
            //добавляем пустые строки
            docX.InsertParagraph();
            AddTableAndFooterToDocx(allInventoryRecords, docX);
            docX.Save();
            return filePath;
        }

        private void AddTableAndFooterToDocx(List<DocumentRegistryCard> allInventoryRecords, DocX docX)
        {
            //добавляем таблицу
            var table = docX.InsertTable(1, 6);
            Border border = new Border();
            table.SetBorder(TableBorderType.InsideH, border);
            table.SetBorder(TableBorderType.InsideV, border);
            table.SetBorder(TableBorderType.Left, border);
            table.SetBorder(TableBorderType.Right, border);
            table.SetBorder(TableBorderType.Top, border);
            table.SetBorder(TableBorderType.Bottom, border);
            var tableHeaderRow = table.Rows.First();
            List<string> headers = new List<string> {
                "№ п/п",
                "Формат",
                "Обозначение документа",
                "Наименование документа",
                "Кол-во листов",
                "Инв. номер"};
            FillTheRow(tableHeaderRow, headers);
            int counter = 1;
            foreach (var item in allInventoryRecords)
            {
                var row = table.InsertRow();
                IEnumerable<string> cellValues = GetTheCellValues(counter, item, row);
                FillTheRow(row, cellValues);

                counter++;
            }

            docX.InsertParagraph();
            docX.InsertParagraph("Составлен в двух экземплярах, по одному экземпляру для каждой из сторон.");
            docX.InsertParagraph();

            //убираем границы у таблицы
            border.Tcbs = Xceed.Words.NET.BorderStyle.Tcbs_none;
            //добавляем таблицу с пописантами
            var signs = docX.InsertTable(1, 2);

            List<string> peredal = new List<string>() { "Передал документы:", string.Format("{0} /______________/", Connection.ClientView.GetUser().ShortName) };
            FillTheRow(signs.Rows.First(), peredal);
            signs.InsertRow();
            List<string> prinyal = new List<string>() { "Принял документы:", "Воробьев В.Е. /_______________/" };
            FillTheRow(signs.InsertRow(), prinyal);

            //signs.AutoFit = AutoFit.Contents;
            signs.AutoFit = AutoFit.Window;

            foreach (var row in signs.Rows)
            {

                row.Cells[0].Paragraphs.First().Alignment = Alignment.left;
                row.Cells[1].Paragraphs.First().Alignment = Alignment.right;
            }
        }

        private IEnumerable<string> GetTheCellValues(int counter, DocumentRegistryCard item, Row row)
        {
            List<string> result = new List<string>();

            result.Add(counter.ToString());
            result.Add(item.Format);
            result.Add(item.Designation);
            result.Add(item.DocumentName);
            result.Add(item.PagesCount.ToString());
            result.Add(item.RegistrationNumber.ToString());
            return result;
        }

        private static void FillTheRow(Xceed.Words.NET.Row row, IEnumerable<string> cellValues)
        {
            int index = 0;
            foreach (var item in cellValues)
            {
                row.Cells[index].Paragraphs.First().Append(item);
                index++;
            }

        }
#endif
        private void ClearFilesFromFolder(string folderPathForUpLoad)
        {
            string[] files = System.IO.Directory.GetFiles(folderPathForUpLoad);
            foreach (var item in files)
            {
                File.SetAttributes(item, FileAttributes.Normal);
                File.Delete(item);
            }
        }

        private void SaveLog(string pathFile, string text)
        {
            System.IO.File.WriteAllText(pathFile, text);
        }

        private List<DocumentRegistryCard> RecursiveGetInventoryRecordComponents(DocumentRegistryCard inventoryBookRecord, bool onlySafeKeepingStage = false)
        {
            List<DocumentRegistryCard> result = new List<ArchiveInventoryBook.DocumentRegistryCard> { inventoryBookRecord };
            var components = inventoryBookRecord.Components;
            if (components.Count() > 0)
            {
                foreach (var item in components)
                {
                    if (onlySafeKeepingStage && item.Document.Stage.Name != DocumentRegistryCard.StageNames.SafeKeeping.ToString())
                        continue;
                    result.AddRange(RecursiveGetInventoryRecordComponents(item.Document));
                }
            }
            return result;
        }

        private string UploadVersionFileToLocalFolder(string pathForUpLoad, ArchiveFile file, int version = 0)
        {
            string fileInfo = string.Format("{0}", file.FileObject.Name);
            if (version == 0) fileInfo += ", ver." + file.FileObject.SystemFields.Version.ToString();
            else fileInfo += ", ver." + version.ToString();
            try
            {
                if (version == 0)
                    file.FileObject.GetHeadRevision(pathForUpLoad + file.FileObject.Name);
                else
                {
                    file.FileObject.GetFileVersion(pathForUpLoad + file.FileObject.Name, version);
                }
                return "Success - " + fileInfo;
            }
            catch
            {
                return "Failed - " + fileInfo;
            }
        }

        private string ChooseLocalFolder()
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                return folderBrowserDialog.SelectedPath;
            }
            return null;
        }

        /// <summary>
        /// Переводит инвентарную запись на хранение
        /// </summary>
        /// <param name="inventoryBookRecord"></param>
        private void ChangeToSafeKeeping(DocumentRegistryCard inventoryBookRecord)
        {
            string errors = CanBeKeeping(inventoryBookRecord);
            if (string.IsNullOrWhiteSpace(errors))
            {
                try
                {
                    MovePDFFilesToOriginalsFolder(inventoryBookRecord);
                    if (!ChangeStage(inventoryBookRecord, DocumentRegistryCard.StageNames.SafeKeeping))
                    { throw new Exception("Не удалось изменить стадию, на Хранение"); }
                }
                catch (Exception e)
                {
                    string objectInfo = GetInventoryRecordInfo(inventoryBookRecord);
                    //SendErrorToAdmin(inventoryBookRecord, e.ToString());
                    throw new Exception(objectInfo + e.ToString());
                }
            }
        }

        private string GetInventoryRecordInfo(DocumentRegistryCard inventoryBookRecord)
        {
            string result = "";
            if (inventoryBookRecord != null)
            {
                result += "Reference = " + inventoryBookRecord.ReferenceObject.Reference.Name + Environment.NewLine;
                result += "Id = " + inventoryBookRecord.ReferenceObject.SystemFields.Id.ToString() + Environment.NewLine;
            }
            return result;
        }

        private void SendErrorToAdmin(DocumentRegistryCard inventoryBookRecord, string v)
        {
            throw new NotImplementedException();
        }


        #endregion

        /// <summary>
        /// проверяет уникальность объекта по параметру обозначение
        /// </summary>
        /// <param name="documentRegistryCard"></param>
        private bool IsDesignationAndInventaryNumberUnique(DocumentRegistryCard documentRegistryCard)
        {
            bool isUnique = true;
            List<ReferenceObject> listRecords = DocumentRegistryCards.GetAllObjectsWithDesignation(documentRegistryCard.Designation);
            if (listRecords.Count == 0 || listRecords.First() == documentRegistryCard.ReferenceObject) return isUnique;
            isUnique = false;
            DialogResult dialogResult = MessageBox.Show(string.Format("Запись с обозначением - {0}, уже существует. Открыть запись?", documentRegistryCard.Designation), "Запись уже существует!", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk);
            if (dialogResult == DialogResult.Yes)
            {
                ReferenceObject newObject = listRecords.Where(r => r != documentRegistryCard.ReferenceObject).First();
                CloseDialog(false);
                RefObj newRefObj = RefObj.CreateInstance(newObject, Context);
                ShowPropertyDialog(newRefObj);

            }
            //TODO: объединить проверку обозначения и инвентарного номера в одно действие
            listRecords = DocumentRegistryCards.GetAllObjectsWithRegistrationNumber(documentRegistryCard.RegistrationNumber.ToString());
            if (listRecords.Count == 0 || listRecords.First() == documentRegistryCard.ReferenceObject) return isUnique;
            dialogResult = MessageBox.Show(string.Format("Запись с инвентарным номером - {0}, уже существует. Открыть запись?", documentRegistryCard.RegistrationNumber), "Запись уже существует!", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk);
            if (dialogResult == DialogResult.Yes)
            {
                ReferenceObject newObject = listRecords.Where(r => r != documentRegistryCard.ReferenceObject).First();
                CloseDialog(false);
                RefObj newRefObj = RefObj.CreateInstance(newObject, Context);
                ShowPropertyDialog(newRefObj);
            }
            return isUnique;
        }

        private bool IsDesignationUnique(DocumentRegistryCard documentRegistryCard)
        {
            List<ReferenceObject> listRecords = DocumentRegistryCards.GetAllObjectsWithDesignation(documentRegistryCard.Designation);
            if (listRecords.Count == 0 || listRecords.First() == documentRegistryCard.ReferenceObject) return true;
            else return false;
        }

        /// <summary>
        /// Закрывает диалог свойств если он открыт и сохраняет изменения
        /// </summary>
        public void CloseDialog(bool withSave = true)
        {
#if !server
            var uiContext = Context as UIMacroContext;
            if (!withSave) Context.ReferenceObject.CancelChanges();
            uiContext.CloseDialog(withSave);
#endif
        }


        private List<FileObject> FindFilesInFolderSatisfyDesignationTerms(FolderObject temporaryFolder, string designation)
        {
            List<FileObject> result = new List<FileObject>();
            Filter filter = new Filter(References.FileReference.ParameterGroup.ReferenceInfo);
            // Условие: в названии содержится "чертёж"
            ReferenceObjectTerm term1 = new ReferenceObjectTerm(filter.Terms);
            // устанавливаем параметр
            term1.Path.AddParentObject();
            //term1.Path.AddParameter(FileReferenceInfo.Description.DefaultVisibleParameter);
            // устанавливаем оператор сравнения
            term1.Operator = ComparisonOperator.Equal;
            // устанавливаем значение для оператора сравнения
            //List<ReferenceObject> listOfFolders = new List<ReferenceObject> { temporaryFolder };

            term1.Value = temporaryFolder;

            // Условие: наименование начинается с "name"
            ReferenceObjectTerm term2 = new ReferenceObjectTerm(filter.Terms);
            // устанавливаем параметр
            term2.Path.AddParameter(References.FileReference.ParameterGroup.ReferenceInfo.Description.DefaultVisibleParameter);
            // устанавливаем оператор сравнения
            term2.Operator = ComparisonOperator.ContainsSubstring;
            // устанавливаем значение для оператора сравнения
            term2.Value = designation;

            //// Условие: наименование начинается с "name"
            //ReferenceObjectTerm term3 = new ReferenceObjectTerm(filter.Terms);
            //// устанавливаем параметр
            //term3.Path.AddParameter(References.FileReference.ParameterGroup.ReferenceInfo.Description.DefaultVisibleParameter);
            //// устанавливаем оператор сравнения
            //term3.Operator = ComparisonOperator.EndsWithSubstring;
            //// устанавливаем значение для оператора сравнения
            //term3.Value = ".pdf";

#if test
            //MessageBox.Show(filter.ToString());
#endif
            List<FileObject> searchResult = References.FileReference.Find(filter).Select(r => r as FileObject).ToList();
#if test
            //MessageBox.Show(result.Count.ToString());
#endif
            List<FileObject> partlySatisfiedFiles = new List<FileObject>();
            //проверяем по условиям конвенции
            foreach (var file in searchResult)
            {
                if (FileNameFullSatisfyAnyConventions(file.Name.ToString(), designation))
                {
                    result.Add(file);
                }
#if !server
                else if (FileNamePartlySatisfyAnyConvention(file.Name.ToString(), designation)) partlySatisfiedFiles.Add(file);

            }
            if (partlySatisfiedFiles.Count > 0) result.AddRange(ChooseFilesToAdd(designation, partlySatisfiedFiles));
#else
            }
#endif
            return result;
        }

        private bool FileNamePartlySatisfyAnyConvention(string fileName, string designation)
        {
            return ScanFileNameConvention.PartlySatisfy(fileName.ToString(), designation) || ScanFileNameConvention_Old.PartlySatisfy(fileName.ToString(), designation);
        }

        private bool FileNameFullSatisfyAnyConventions(string fileName, string designation)
        {
            return ScanFileNameConvention.FullSatisfy(fileName.ToString(), designation) || ScanFileNameConvention_Old.FullSatisfy(fileName.ToString(), designation);
        }

        /// <summary>
        /// Возвращает выбранные файлы, если ничего не выбрано то пустой список, 
        /// если нажата отмена то null
        /// </summary>
        /// <param name="designation"></param>
        /// <param name="partlySatisfiedFiles"></param>
        /// <returns></returns>
        private IEnumerable<FileObject> ChooseFilesToAdd(string designation, List<FileObject> partlySatisfiedFiles)
        {
            var dialog = CreateInputDialog("Выберите файлы для - " + designation);
            dialog.AddComment("comment", string.Format("Некоторые файлы частично подходят к инвентарной записи с обозначением - {0}, отметьте те которые должны быть к ней присоединены", designation));
            foreach (var item in partlySatisfiedFiles)
            {
                dialog.AddFlag(item.Name.ToString(), false, true);
            }
            List<FileObject> result = new List<FileObject>();
            if (dialog.Show())
            {
                foreach (var item in partlySatisfiedFiles)
                {
                    if (dialog.GetValue(item.Name.ToString())) result.Add(item);
                }
            }
            else return null;
            return result;
        }



        private void Sign(DocumentRegistryCard inventoryBookRecord, string accept)
        {
            SignatureCollection signCol = new SignatureCollection(inventoryBookRecord.ReferenceObject);
            SignatureType newSign = signCol.Types.Where(st => st.Name == accept).ToList().FirstOrDefault();//находим подпись нужного типа

            inventoryBookRecord.ReferenceObject.SetSignature(Context.Connection.ClientView.GetUser(), newSign, "Принято на хранение");
        }

        private void MovePDFFilesToOriginalsFolder(DocumentRegistryCard inventoryBookRecord)
        {
            IEnumerable<FileObject> files = inventoryBookRecord.Files.Select(af => af.FileObject);
            FileReference fileReference = References.FileReference;
            string PathToOriginalFolder = DocumentRegistryCards.FolderName.ScansFullPath;
            FolderObject folderOriginals = fileReference.FindByPath(PathToOriginalFolder) as FolderObject;

            //делим файлы по нахождению в папках
            List<FileObject> filesPDFInOriginals = new List<FileObject>();
            List<FileObject> filesPDFInTemporary = new List<FileObject>();
            foreach (var file in files)
            {
                switch (file.Parent.Name.ToString())
                {
                    case DocumentRegistryCards.FolderName.Scans:
                        if (file.Name.ToString().ToLower().EndsWith(".pdf"))
                            filesPDFInOriginals.Add(file);
                        break;
                    case DocumentRegistryCards.FolderName.Temporary:
                        if (file.Name.ToString().ToLower().EndsWith(".pdf"))
                            filesPDFInTemporary.Add(file);
                        break;
                    default:
                        break;
                }
            }
            int countFilesTemporary = filesPDFInTemporary.Count;
            int countFilesOriginals = filesPDFInOriginals.Count;
            //если файлов в папке Временная нет, значит ничего не делаем
            if (countFilesTemporary == 0)
            {
                if (countFilesOriginals == 1)
                {
                    string newName = GetNewNameForPDFFile(filesPDFInOriginals.First(), inventoryBookRecord);
                    MoveFileToFolderAndRename(filesPDFInOriginals.First(), folderOriginals, newName);
                }
                return;
            }
            //если в папке Оригиналы файлов нет
            if (countFilesOriginals == 0)
            {
                //если в папке Оригиналы файлов нет, а во Временной один файл то перемещаем его и именуем согласно конвенции
                if (countFilesTemporary == 1)
                {
                    string newName = GetNewNameForPDFFile(filesPDFInTemporary.First(), inventoryBookRecord);
                    MoveFileToFolderAndRename(filesPDFInTemporary.First(), folderOriginals, newName);
                }
                //если в папке Оригиналы файлов нет, а во Временной больше чем один файл то надо придумать что делать                else
                else
                {//TODO: add method Обработка двух файлов во Временной папке
                    throw new NotImplementedException("Не задан метод, для обработки двух файлов в папке Временная и отсутствии файла в Оригиналах!");
                }
            }
            else if (countFilesOriginals == 1)
            {
                if (countFilesTemporary == 1)
                {
                    string newName = GetNewNameForPDFFile(filesPDFInTemporary.First(), inventoryBookRecord);
                    List<FileObject> combineFiles = new List<FileObject>();
                    combineFiles.AddRange(filesPDFInTemporary);
                    combineFiles.AddRange(filesPDFInOriginals);
                    FileObject file = CombineFilesWithVersions(combineFiles);
                    MoveFileToFolderAndRename(file, folderOriginals, newName);
                }
                else
                {//TODO: add method Обработка двух файлов во Временной папке
                    throw new NotImplementedException("Не задан метод, для обработки двух файлов в папке Временная и одного файла в Оригиналах!");
                }
            }
            else
            {//TODO: add method Обработка двух файлов во Временной папке и папке Оригиналы
                throw new NotImplementedException("Не задан метод, для обработки двух файлов в папке Временная и одного файла в Оригиналах!");
            }

        }

        private static string GetNewNameForPDFFile(FileObject file, DocumentRegistryCard inventoryBookRecord)
        {
            string newName = "";
            string oldFileName = file.Name.ToString();
            string oldFileNameWithoutExtension = oldFileName.Remove(oldFileName.IndexOf(".pdf"), ".pdf".Length);
            if (string.IsNullOrWhiteSpace(inventoryBookRecord.Designation))
            {
                newName = "СКАН_" + inventoryBookRecord.RegistrationNumber + "_" + oldFileNameWithoutExtension;
            }
            else newName = "СКАН_" + inventoryBookRecord.Designation;
            //максимальная длина наименования без расширения
            int MaxFileNameLength = 250;
            if (newName.Length > MaxFileNameLength) newName.Remove(MaxFileNameLength, newName.Length - MaxFileNameLength);
            return newName + ".pdf";
        }

        private void MoveFileToFolderAndRename(FileObject file, FolderObject folder, string newName)
        {
            string oldName = file.Name.ToString();
            if (file.Parent == folder && oldName == newName) return;
            if (!file.IsCheckedOut) file.CheckOut(false); else Error("Файл на редактировании"); // новый файл

            file.BeginChanges();
            file.SetParent(folder);
            file.Name.Value = newName;
            file.EndChanges();
            string comment = "Перемещён в папку - " + folder.Path.ToString();
            if (oldName != newName) comment = string.Format("Переименован: {0} на {1}", oldName, newName) + Environment.NewLine + comment;
            Desktop.CheckIn(file, comment, false);
        }

        FileObject FindFileInFolder(FolderObject folder, string fileName)
        {
            FileReference fileReference = new FileReference(Connection);
            FileObject file = fileReference.FindByPath(folder.Path + "\\" + fileName) as FileObject;
            return file;
        }

        /// <summary>
        /// Возвращает рекурсивный список дочерних папок
        /// </summary>
        /// <param name="rootFolder"></param>
        /// <returns></returns>
        private IEnumerable<ReferenceObject> GetAllChildFolders(TFlex.DOCs.Model.References.Files.FolderObject rootFolder)
        {
            List<ReferenceObject> result = new List<ReferenceObject>();
            Filter filter = new Filter(References.FileReference.ParameterGroup.ReferenceInfo);
            // Условие: в названии содержится "чертёж"
            ReferenceObjectTerm term1 = new ReferenceObjectTerm(filter.Terms);
            // устанавливаем параметр
            term1.Path.AddParentObject();
            //term1.Path.AddParameter(FileReferenceInfo.Description.DefaultVisibleParameter);
            // устанавливаем оператор сравнения
            term1.Operator = ComparisonOperator.Equal;
            // устанавливаем значение для оператора сравнения
            term1.Value = rootFolder;

            ReferenceObjectTerm term2 = new ReferenceObjectTerm(filter.Terms);
            //term2.Path.AddChildObjects();
            term2.Path.AddParameter(References.FileReference.ParameterGroup[SystemParameterType.Class]);
            term2.Operator = ComparisonOperator.Equal;
            term2.Value = References.FileReference.Classes.Folder;
            result = References.FileReference.Find(filter);
            List<ReferenceObject> tempResult = new List<ReferenceObject>();
            if (result.Count > 0)
            {

                foreach (var item in result)
                {
                    tempResult.AddRange(GetAllChildFolders(item as TFlex.DOCs.Model.References.Files.FolderObject));
                }

            }
            result.AddRange(tempResult);
            return result;
        }

        /// <summary>
        /// Находит все файлы с расширением .pdf в указанной папке, которые начинаются с name
        /// </summary>
        /// <param name="name"></param>
        /// <param name="vremennayaFolder"></param>
        /// <returns></returns>
        private List<FileObject> GetPDFFilesWithNameFromFolder(string name, TFlex.DOCs.Model.References.Files.FolderObject vremennayaFolder)
        {
            List<FileObject> result = new List<FileObject>();
            Filter filter = new Filter(References.FileReference.ParameterGroup.ReferenceInfo);
            // Условие: в названии содержится "чертёж"
            ReferenceObjectTerm term1 = new ReferenceObjectTerm(filter.Terms);
            // устанавливаем параметр
            term1.Path.AddParentObject();
            //term1.Path.AddParameter(FileReferenceInfo.Description.DefaultVisibleParameter);
            // устанавливаем оператор сравнения
            term1.Operator = ComparisonOperator.IsOneOf;
            // устанавливаем значение для оператора сравнения
            List<ReferenceObject> listOfFolders = new List<ReferenceObject> { vremennayaFolder };
            listOfFolders.AddRange(GetAllChildFolders(vremennayaFolder));
            term1.Value = listOfFolders;

            // Условие: наименование начинается с "name"
            ReferenceObjectTerm term2 = new ReferenceObjectTerm(filter.Terms);
            // устанавливаем параметр
            term2.Path.AddParameter(References.FileReference.ParameterGroup.ReferenceInfo.Description.DefaultVisibleParameter);
            // устанавливаем оператор сравнения
            term2.Operator = ComparisonOperator.StartsWithSubstring;
            // устанавливаем значение для оператора сравнения
            term2.Value = name;

            // Условие: наименование начинается с "name"
            ReferenceObjectTerm term3 = new ReferenceObjectTerm(filter.Terms);
            // устанавливаем параметр
            term3.Path.AddParameter(References.FileReference.ParameterGroup.ReferenceInfo.Description.DefaultVisibleParameter);
            // устанавливаем оператор сравнения
            term3.Operator = ComparisonOperator.EndsWithSubstring;
            // устанавливаем значение для оператора сравнения
            term3.Value = ".pdf";

#if test
            //MessageBox.Show(filter.ToString());
#endif
            result = References.FileReference.Find(filter).Select(r => r as FileObject).ToList();
#if test
            //MessageBox.Show(result.Count.ToString());
#endif
            return result;
        }

        private bool ChangeStage(DocumentRegistryCard inventoryBookRecord, StageName stageName)
        {
            if (inventoryBookRecord.ReferenceObject.SystemFields.Stage.Stage.Name == stageName.ToString()) return true;
            Stage newStage = Stage.GetStages(DocumentRegistryCards.Reference.ParameterGroup.ReferenceInfo).Where(s => s.Name == stageName.ToString()).FirstOrDefault();
            if (newStage == null) throw new ArgumentNullException("Stage named - " + stageName + ", not found!");
            inventoryBookRecord.ReferenceObject.Reload();
            var list = newStage.Change(new List<ReferenceObject> { inventoryBookRecord.ReferenceObject });
            //var list = newStage.AutomaticChange(new List<ReferenceObject> { inventoryBookRecord.ReferenceObject });
            //var list = newStage.Set(new List<ReferenceObject> { inventoryBookRecord.ReferenceObject });
            string stage = inventoryBookRecord.ReferenceObject.SystemFields.Stage.Stage.Name;
            return list.Count > 0;
        }
        #endregion
        static ServerConnection Connection { get { return ServerGateway.Connection; } }
    }

    interface IDocumentRegistryCard
    {
        string Designation { get; }
        string RegistrationNumber { get; }
    }




    /// <summary>
    /// Инвентарный номер (уникальный номер справочника)
    /// </summary>
    public class InventoryNumber
    {
        public InventoryNumber(string code, string stringNumber)
        {
            this.Code = code;
            this.RegistrationNumber = stringNumber;
        }

        public InventoryNumber(string inventoryNumber)
        {
            string code;
            string number;
            ParseInventoryNumberString(inventoryNumber, out code, out number);
            this.Code = code;
            this.RegistrationNumber = number;
        }

        public InventoryNumber(string code, int number)
        {
            this.Code = code;

            int numDigits = 3;
            string newOboznachTP = number.ToString();
            if (number < 100)
                newOboznachTP = number.ToString().PadLeft(numDigits, '0');
            //else 
            this.RegistrationNumber = newOboznachTP;
        }


        private static void ParseInventoryNumberString(string inventoryNumber, out string code, out string number)
        {
            code = "";
            number = "";
            if (string.IsNullOrWhiteSpace(inventoryNumber)) return;
            int indexOfDelimiter = inventoryNumber.LastIndexOf(delimiter);
            if (indexOfDelimiter > 0)
            {
                code = inventoryNumber.Remove(indexOfDelimiter);
                if (indexOfDelimiter < inventoryNumber.Length)
                    number = inventoryNumber.Remove(0, indexOfDelimiter + 1);
            }
        }

        public string RegistrationNumber { get; private set; }

        public const char delimiter = '.';
        /// <summary>
        /// Код плюс разделитель '.'
        /// </summary>
        public string Code { get; private set; }
        public override string ToString()
        {
            return Code + delimiter + RegistrationNumber;
        }
    }

    public class DocumentRegistryCard : IDocumentRegistryCard, IComparable, IEquatable<DocumentRegistryCard>
    {
        #region Статические переменные
        public struct StageNames
        {
            public static StageName Edit = new StageName("Обрабатывается");
            public static StageName CheckDocument = new StageName("Проверка документа");
            public static StageName SafeKeeping = new StageName("Хранение");
        }

        public struct SignNames
        {
            public const string Accept = "Изм. внес";
            //public static StageName Edit = new StageName("Обрабатывается");
        }


        #endregion

        #region
        public static bool operator ==(DocumentRegistryCard obj1, DocumentRegistryCard obj2)
        {
            if (object.ReferenceEquals(null, obj1))
            {
                if (object.ReferenceEquals(null, obj2))
                    return true;
                return false;
            }
            return obj1.Equals(obj2);
        }

        public static bool operator !=(DocumentRegistryCard obj1, DocumentRegistryCard obj2)
        {
            if (object.ReferenceEquals(null, obj1))
            {
                if (object.ReferenceEquals(null, obj2))
                    return false;
                return true;
            }
            return !obj1.Equals(obj2);
        }
        #endregion

        #region Конструкторы
        public DocumentRegistryCard(ReferenceObject referenceObject)
        {
            this.ReferenceObject = referenceObject;
        }
        #endregion

        #region Свойства

        public ReferenceObject ReferenceObject { get; private set; }

        string _Designation;
        public string Designation
        {
            get
            {
                if (_Designation == null)
                {
                    if (this.ReferenceObject != null)
                    {
                        _Designation = this.ReferenceObject[param_Designation_Guid].GetString();
                    }
                    else return "";
                }
                return _Designation;
            }
        }

        string _DocumentName;
        public string DocumentName
        {
            get
            {
                if (_DocumentName == null)
                {
                    if (this.ReferenceObject != null)
                    {
                        _DocumentName = this.ReferenceObject[param_DocumentName_Guid].GetString();
                    }
                    else return "";
                }
                return _DocumentName;
            }
        }

        string _RegistrationNumber;

        /// <summary>
        /// Инвентарный номер
        /// </summary>
        public string RegistrationNumber
        {
            get
            {
                if (_RegistrationNumber == null)
                {
                    string regNum = "";
                    if (this.ReferenceObject != null)
                    {
                        regNum = this.ReferenceObject[param_RegistrationNumber_Guid].GetString();
                    }
                    _RegistrationNumber = regNum;
                }
                return _RegistrationNumber;
            }
            //set
            //{
            //    if (_RegistrationNumber != value)
            //    {
            //        _RegistrationNumber = value;
            //        if (_RegistrationNumber != null) _haveChangesToSave = true;
            //    }
            //}
        }

        internal void SetInventoryNumber(InventoryNumber inventoryNumber)
        {
            _RegistrationNumber = inventoryNumber.ToString();
            _haveChangesToSave = true;
        }

        Status _Status;

        /// <summary>
        /// Инвентарный номер
        /// </summary>
        public Status Status
        {
            get
            {
                if (_Status == null)
                {
                    string status = "";
                    if (this.ReferenceObject != null)
                    {
                        status = this.ReferenceObject[param_Status_Guid].GetString();
                    }
                    _Status = new Status(status);
                }
                return _Status;
            }
        }

        string _DesignerFullName;

        public string DesignerFullName
        {
            get
            {
                if (_DesignerFullName == null)
                {
                    _DesignerFullName = this.ReferenceObject[param_DesignerName_Guid].GetString();
                }
                return _DesignerFullName;
            }
        }

        UserReferenceObject _DesignerObject;

        public UserReferenceObject DesignerObject
        {
            get
            {
                if (_DesignerObject == null)
                {
                    ReferenceObject linkedObject = this.ReferenceObject.GetObject(link_Department_Guid);
                    if (linkedObject != null)
                    {
                        _DesignerObject = linkedObject as UserReferenceObject;
                    }
                }

                return _DesignerObject;
            }
            set
            {
                if (_DesignerObject != value)
                {
                    _DesignerObject = value;
                    if (_DesignerObject == null) _DesignerFullName = "";
                    else _DesignerFullName = _DesignerObject.ToString();
                    _haveChangesToSave = true;
                }
            }
        }

        InventoryBook _InventoryBook;
        internal InventoryBook InventoryBook
        {
            get
            {
                if (_InventoryBook == null)
                    _InventoryBook = new InventoryBook(this.ReferenceObject.GetObject(link_InventoryBook_Guid));
                return _InventoryBook;
            }
        }

        List<ArchiveFile> _Files;

        public IEnumerable<ArchiveFile> Files
        {
            get
            {
                if (_Files == null)
                {
                    //Guid link_Files_Guid = new Guid("017a0e87-63af-4c33-a7db-1b0ac4f9059b");
                    var tempFiles = this.ReferenceObject.GetObjects(link_Files_Guid).Where(r => (r as FileObject) != null).Select(r => (r as FileObject)).ToList();
                    if (tempFiles != null)
                    {
                        _Files = new List<ArchiveFile>();
                        foreach (var item in tempFiles)
                        {
                            ArchiveFile archiveFile = new ArchiveFile(item);
                            _Files.Add(archiveFile);
                        }
                    }
                }
                return _Files;
            }
        }

        List<DocumentComponent> _Components;

        internal IEnumerable<DocumentComponent> Components
        {
            get
            {
                if (_Components == null)
                {
                    _Components = new List<DocumentComponent>();
                    //Guid link_Files_Guid = new Guid("017a0e87-63af-4c33-a7db-1b0ac4f9059b");
                    var tempDocumentComponents = this.ReferenceObject.GetObjects(list_DocumentComponents_Guid);

                    foreach (var item in tempDocumentComponents)
                    {
                        DocumentComponent documentComponent = new DocumentComponent(item);
                        _Components.Add(documentComponent);
                    }
                }
                return _Components;
            }
        }


        public Stage Stage { get { return this.ReferenceObject.SystemFields.Stage.Stage; } }

        /// <summary>
        /// Формат документа
        /// </summary>
        public string Format
        {
            get
            {
                if (this.ReferenceObject != null)
                {
                    return this.ReferenceObject[param_Format_Guid].GetString();
                }
                return "";
            }
        }

        public int PagesCount
        {
            get
            {
                if (this.ReferenceObject != null)
                {
                    return this.ReferenceObject[param_PagesCount_Guid].GetInt16();
                }
                return 0;
            }
        }







        private bool _haveChangesToSave;



        #endregion

        #region Методы
        internal void AddFile(FileObject file)
        {
            this._Files.Add(new ArchiveFile(file));
            _haveChangesToSave = true;
        }

        /// <summary>
        /// Сохраняет изменения
        /// </summary>
        internal void Save()
        {
            if (_haveChangesToSave)
            {
                //MessageBox.Show("BeginChanges ");
                this.ReferenceObject.BeginChanges();
                foreach (var file in this.Files)
                {
                    this.ReferenceObject.AddLinkedObject(link_Files_Guid, file.FileObject);
                }

                this.ReferenceObject.SetLinkedObject(link_Department_Guid, this.DesignerObject);
                this.ReferenceObject[param_DesignerName_Guid].Value = this.DesignerFullName;
                this.ReferenceObject[param_RegistrationNumber_Guid].Value = this.RegistrationNumber.ToString();
                //сохраняем изменения локально
                this.ReferenceObject.EndChanges();
                //MessageBox.Show("saveSet ");
            }
        }

        /// <summary>
        /// Сохраняет изменения в БД
        /// </summary>
        internal void SaveToDB()
        {
            if (_haveChangesToSave)
            {
                this.Save();

                this.ReferenceObject.CreateSaveSet();
                var saveSet = this.ReferenceObject.SaveSet;
                //if (saveSet == null) MessageBox.Show("saveSet = null");
                //else MessageBox.Show("saveSet.EndChanges = " + saveSet.EndChanges().ToString());
                //сохраняем изменения в БД на сервере
                saveSet.EndChanges();
                //Снова берём на редактирование
                this.ReferenceObject.BeginChanges(true);
                this.ReferenceObject.CreateSaveSet();
                //this.ReferenceObject.EndChanges();
            }
        }

        public override string ToString()
        {
            return this.RegistrationNumber.ToString();
        }

        public int CompareTo(object obj)
        {
            DocumentRegistryCard invRec = obj as DocumentRegistryCard;
            if (invRec == null) return 1;
            return CompareTo(invRec);
        }
        public int CompareTo(DocumentRegistryCard inventoryBookRecord)
        {
            int result = this.RegistrationNumber.ToString().CompareTo(inventoryBookRecord.RegistrationNumber.ToString());
            if (result == 0) result = this.Designation.CompareTo(inventoryBookRecord.Designation);
            return result;
        }

        public override bool Equals(object obj)
        {
            DocumentRegistryCard inv = obj as DocumentRegistryCard;
            if (inv != null) return Equals(inv);
            else
                return base.Equals(obj);
        }

        public bool Equals(DocumentRegistryCard other)
        {
            if (other == null) return false;
            return this.RegistrationNumber.ToString().Equals(other.RegistrationNumber.ToString()) &&
                this.Designation.Equals(other.Designation);
        }
        public override int GetHashCode()
        {

            return (this.RegistrationNumber.ToString() + this.Designation).GetHashCode();

        }
        #endregion

        #region параметры БД
        public static Guid param_Designation_Guid = new Guid("743bc2bd-e8ea-4676-ad61-281df141ba61");
        public static Guid param_DocumentName_Guid = new Guid("4f89b0c3-9578-4e75-803b-0494c17f4858");
        public static Guid param_RegistrationNumber_Guid = new Guid("54d74eab-82bb-4199-964c-5a3c22397d6a");
        public static Guid param_Status_Guid = new Guid("4140b838-e909-45b7-b216-69233031fd94");
        public static Guid param_DesignerName_Guid = new Guid("05c4570a-02cf-4d50-98b8-636352977d72");
        public static Guid param_Format_Guid = new Guid("bed14fe8-370c-4fed-93e8-30e8d2c82310");
        public static Guid param_PagesCount_Guid = new Guid("e9b03bc5-5c56-429b-9b24-46fbcf7b142c");

        public static Guid link_Department_Guid = new Guid("b742e4c4-485f-457c-abac-7e2849f1c418");
        public static Guid link_Files_Guid = new Guid("a8432b2a-d9d1-4ab0-9bfe-003d349b6e0d");
        public static Guid link_InventoryBook_Guid = new Guid("480933a3-681c-4fad-8dc0-20c2835f3478");
        public static Guid list_DocumentComponents_Guid = new Guid("16b41b42-e186-4320-8f92-ee7e26aec790");

        #endregion
    }

    class DesignDocumentRegistryCard : DocumentRegistryCard
    {
        public DesignDocumentRegistryCard(ReferenceObject referenceObject) : base(referenceObject)
        {
        }
    }

    class TechnologicalDocumentRegistryCard : DocumentRegistryCard
    {
        public TechnologicalDocumentRegistryCard(ReferenceObject referenceObject) : base(referenceObject)
        {
        }
    }

    class ChangeNotificationDocumentRegistryCard : DocumentRegistryCard
    {
        public ChangeNotificationDocumentRegistryCard(ReferenceObject referenceObject) : base(referenceObject)
        {
        }



        string _ProductIndex;
        public string ProductIndex
        {
            get
            {
                if (_ProductIndex == null)
                    _ProductIndex = this.ReferenceObject[param_ProductIndex_Guid].GetString();
                return _ProductIndex;
            }
        }

        List<DocumentRegistryCard> _ChangedDocuments;
        public IEnumerable<DocumentRegistryCard> ChangedDocuments
        {
            get
            {
                if (_ChangedDocuments == null && this.ReferenceObject != null)
                {
                    _ChangedDocuments = new List<DocumentRegistryCard>();
                    Guid param_designation_Guid = new Guid("ac67607f-1966-46a8-882c-98ff332d410d");
                    foreach (var item in this.ReferenceObject.GetObjects(list_ApplicableDocuments_Guid))
                    {

                        _ChangedDocuments.Add(DocumentRegistryCards.GetDocumentRegistryCard(DocumentRegistryCards.GetAllObjectsWithDesignation(item[param_designation_Guid].GetString()).FirstOrDefault()));
                    }
                }
                return _ChangedDocuments;
            }
        }


        static Guid list_ApplicableDocuments_Guid = new Guid("db4facb4-d234-4519-9bb6-acdf045d23ec");
        public static Guid param_ProductIndex_Guid = new Guid("a621a3a4-a0e5-4ae0-a01f-a476fd3e8d41");
    }

    class DocumentRegistryCards
    {
        public struct FolderName
        {
            public const string Archive = "Архив документации";
            public const string Temporary = "Временная";
            public const string TemporaryFullPath = Archive + "\\" + Temporary;
            public const string Scans = "Подлинники";
            public const string ScansFullPath = Archive + "\\" + Scans;
            public const string MasterFiles = "Оригиналы";
            public const string MasterFilesFullPath = Archive + "\\" + MasterFiles;
            public const string TechnologicalFiles = "Технологические файлы";
            public const string TechnologicalFilesFullPath = MasterFilesFullPath + "\\" + TechnologicalFiles;
        }

        public static dynamic GetDocumentRegistryCard(ReferenceObject referenceObject)
        {
            ClassObject className = referenceObject.Class;
            switch (referenceObject.Class.Name)
            {
                case ("Карточка учёта конструкторских документов"):
                    return new DesignDocumentRegistryCard(referenceObject);

                case ("Карточка учёта технологических документов"):
                    return new TechnologicalDocumentRegistryCard(referenceObject);

                case ("Карточка учёта извещений"):
                    return new ChangeNotificationDocumentRegistryCard(referenceObject);
                default:
                    return new DocumentRegistryCard(referenceObject); ;
            }
        }

        public static List<ReferenceObject> GetAllObjectsWithDesignation(string designation)
        {
            ParameterInfo designationParameter = DocumentRegistryCards.ReferenceInfo.Description[DocumentRegistryCard.param_Designation_Guid];
            var objects = DocumentRegistryCards.Reference.Find(designationParameter, designation);
            return objects;
        }

        internal static List<ReferenceObject> GetAllObjectsWithRegistrationNumber(string registrationNumber)
        {
            ParameterInfo designationParameter = DocumentRegistryCards.ReferenceInfo.Description[DocumentRegistryCard.param_RegistrationNumber_Guid];
            var objects = DocumentRegistryCards.Reference.Find(designationParameter, registrationNumber);
            return objects;
        }

        /*
        public static List<DocumentRegistryCard> GetComponentsFor(string assemblyInventoryNumber)
        {
            //TODO: сделать проверку на соответствие типу КД и ТД
            ReferenceInfo documentReferenceInfo = DocumentRegistryCards.ReferenceInfo;
            ParameterGroup listToApplicability = documentReferenceInfo.Description.OneToManyTables.First(link => link.Guid == new Guid("16b41b42-e186-4320-8f92-ee7e26aec790"));
            ParameterInfo invNumberAppliable = listToApplicability.Parameters.Find(new Guid("6f0387a5-ecce-41fe-a85c-8437eab8b0de"));

            Filter filter = new Filter(documentReferenceInfo);


            // Условие: имя одного из связанных файлов равно "2D чертёж.grb"
            ReferenceObjectTerm term2 = new ReferenceObjectTerm(filter.Terms);
            // добавляем связь в путь к параметру
            term2.Path.AddGroup(listToApplicability);
            term2.Path.AddParameter(invNumberAppliable);
            // список операторов сравнения, которые поддерживает прарамeтр можно получить у него - ParamterInfo.GetComparisonOperators()
            term2.Operator = ComparisonOperator.Equal;
            term2.Value = assemblyInventoryNumber.ToString();

            var tempResult = DocumentRegistryCards.Reference.Find(filter);

            List<DocumentRegistryCard> result = new List<DocumentRegistryCard>();
            foreach (var item in tempResult)
            {
                result.Add(DocumentRegistryCards.GetDocumentRegistryCard(item));
            }
            return result;

        }
        */

        internal static InventoryNumber CreateNewNumberForCode(string code)
        {
            //    public string FindFreeNumberForTP(string tpNumber)
            //{
            //int lastPointIndex = tpNumber.LastIndexOf('.');
            string codeAndDelimiter = code + InventoryNumber.delimiter;
#if test
            //MessageBox.Show(codeOfTP);
#endif

            //Создаем ссылку на справочник
            ReferenceInfo info = DocumentRegistryCards.ReferenceInfo;
            Reference reference = DocumentRegistryCards.Reference;
            //Находим тип «Технологический процесс»
            //ClassObject classObject = info.Classes.Find(new Guid(TechProcClass_ID));
            //Создаем фильтр
            Filter filter = new Filter(info);
            ////Добавляем условие поиска – «Тип = Технологический процесс»
            //filter.Terms.AddTerm(reference.ParameterGroup[SystemParameterType.Class],
            //                     ComparisonOperator.IsInheritFrom, classObject);
            // Условие: в названии содержится "чертёж"
            ReferenceObjectTerm term1 = new ReferenceObjectTerm(filter.Terms);
            //ParameterGroup группаПараметровОсновныеТП = info.Description.OneToOneTables.Find(new Guid(ГруппаПараметровОсновныеТП));
            //term1.Path.AddGroup(группаПараметровОсновныеТП);
            // устанавливаем параметр
            term1.Path.AddParameter(info.Description.Parameters.Find(DocumentRegistryCard.param_RegistrationNumber_Guid));
            // устанавливаем оператор сравнения
            term1.Operator = ComparisonOperator.StartsWithSubstring;
            // устанавливаем значение для оператора сравнения
            term1.Value = codeAndDelimiter;

            //Получаем список объектов, в качестве условия поиска – сформированный фильтр
            List<ReferenceObject> listObj = reference.Find(filter);

            int newNumberOfTP = 0;
            int i = 1;

            List<int> listOfNumbers = new List<int>();

            foreach (var obj in listObj)
            {
                DocumentRegistryCard card = new DocumentRegistryCard(obj);
                string fullNumber = card.RegistrationNumber.ToString();
                if (fullNumber == "") continue;

                i += 1;
                //вычисляем индекс вхождения кода в обозначение ТП
                int numberOfTPStartIndex = fullNumber.LastIndexOf(codeAndDelimiter);

                //если код отсутствует в обозначении переходим к следующему объекту
                if (numberOfTPStartIndex == -1) continue;

                int parsedValue = -1;
                string numForParse = fullNumber.Remove(0, (codeAndDelimiter.Length + numberOfTPStartIndex));
                bool isParsed = Int32.TryParse(numForParse, out parsedValue);

                //если номер удалось отпарсить то добавляем в список
                if (isParsed)
                {
                    listOfNumbers.Add(parsedValue);
                }
            }
            listOfNumbers.Sort();


            //если записей с таким кодом нет то присваиваем номер 1
            if (listOfNumbers.Count == 0) newNumberOfTP = 1;
            else
            {
                int index = 0;
                int currentValueShouldBe = 1 + index;
                while (index < listOfNumbers.Count)
                {
                    //сравниваем текущее значение с предполагаемым если совпадает, то обновляем сравниваемое и переходим
                    //к следующему,
                    if (listOfNumbers[index] == currentValueShouldBe)
                    {
                        index++;
                        currentValueShouldBe++;

                    }
                    else//если не совпадает, то сравниваем предполагаемое со следующим элементом
                    {
                        index++;
                    }
                }
                newNumberOfTP = currentValueShouldBe;
            }
            InventoryNumber invNum = new InventoryNumber(code, newNumberOfTP);

            return invNum;
        }



        // получение описания справочника
        static ReferenceInfo ReferenceInfo { get { return ServerGateway.Connection.ReferenceCatalog.Find(new Guid("6689ec12-ff4b-4301-beb2-ae4dfd11f3b8")); } }

        static Reference _Reference;
        public static Reference Reference
        {
            get
            {
                if (_Reference == null)
                {
                    _Reference = ReferenceInfo.CreateReference();
                }
                else { _Reference.Refresh(); }
                return _Reference;
            }
        }
    }

    class DocumentComponent
    {
        internal DocumentComponent(ReferenceObject referenceObject)
        {
            this.ReferenceObject = referenceObject;
        }

        ReferenceObject ReferenceObject { get; set; }

        DateTime? _IncludeDate;
        internal DateTime IncludeDate
        {
            get
            {
                if (_IncludeDate == null)
                {
                    _IncludeDate = this.ReferenceObject[param_Date_Guid].GetDateTime();
                }
                return (DateTime)_IncludeDate;
            }
        }

        DocumentRegistryCard _Document;
        internal DocumentRegistryCard Document
        {
            get
            {
                if (_Document == null)
                {
                    var ro = this.ReferenceObject.GetObject(link_DocumentRegistryCard_Guid);
                    if (ro != null)
                        _Document = DocumentRegistryCards.GetDocumentRegistryCard(ro);
                }
                return _Document;
            }
        }

        string _Designation;
        internal string Designation
        {
            get
            {
                if (_Designation == null)
                {
                    _Designation = this.ReferenceObject[param_Designation_Guid].GetString();
                }
                return _Designation;
            }
        }

        string _InventoryNumber;
        internal string InventoryNumber
        {
            get
            {
                if (_InventoryNumber == null)
                {
                    _InventoryNumber = this.ReferenceObject[param_InventoryNumber_Guid].GetString();
                }
                return _InventoryNumber;
            }
        }

        static readonly Guid param_Date_Guid = new Guid("6791913e-279d-49df-a9e0-aec3cdc2ac67");
        static readonly Guid param_Designation_Guid = new Guid("3410a791-2e83-4964-8422-5404a971f03b");
        static readonly Guid param_InventoryNumber_Guid = new Guid("6f0387a5-ecce-41fe-a85c-8437eab8b0de");
        static readonly Guid link_DocumentRegistryCard_Guid = new Guid("78e74595-f35e-40e7-b21a-06318fe4b658");
    }

    class InventoryBooksReference
    {

        public static InventoryBook GetInventoryBookForDocumentCard(DocumentRegistryCard card)
        {
            //Создание фильтра
            var filter = CreateFilter(card.ReferenceObject);
            //Поиск объектов с помощью созданного фильтра
            var result = InventoryBooksReference.Reference.Find(filter);
            if (result != null && result.Count > 0) return new ArchiveInventoryBook.InventoryBook(result.First());
            return null;
        }

        private static Filter CreateFilter(ReferenceObject linkedObject)
        {

            ParameterGroup linkToFiles = InventoryBooksReference.ReferenceInfo.Description.OneToManyLinks.First(link => link.Guid == InventoryBooksReference.link_1N_DocumentRegistryCards_Guid);

            Filter filter = new Filter(InventoryBooksReference.ReferenceInfo);


            // Условие: имя одного из связанных файлов равно "2D чертёж.grb"
            ReferenceObjectTerm term2 = new ReferenceObjectTerm(filter.Terms);
            // добавляем связь в путь к параметру
            term2.Path.AddGroup(linkToFiles);
            //term2.Path.AddParameter(linkToFiles.SlaveGroup[SystemParameterType.ObjectId]);
            // список операторов сравнения, которые поддерживает прарамeтр можно получить у него - ParamterInfo.GetComparisonOperators()
            term2.Operator = ComparisonOperator.Equal;
            term2.Value = linkedObject;

            filter.Validate();
            MessageBox.Show(filter.ToString());
            return filter;
        }


        // получение описания справочника
        static ReferenceInfo ReferenceInfo
        { get { return ServerGateway.Connection.ReferenceCatalog.Find(new Guid("5fe8ed4b-30b9-4a36-9437-9d9f3328f4d1")); } }

        static Reference _Reference;
        public static Reference Reference
        {
            get
            {
                if (_Reference == null)
                {
                    _Reference = ReferenceInfo.CreateReference();
                }
                else { _Reference.Refresh(); }
                return _Reference;
            }
        }

        public static readonly Guid link_1N_DocumentRegistryCards_Guid = new Guid("480933a3-681c-4fad-8dc0-20c2835f3478");
    }

    class InventoryBook
    {
        public InventoryBook(ReferenceObject referenceObject)
        {
            this.ReferenceObject = referenceObject;
        }

        ReferenceObject ReferenceObject;
        #region Поля
        string _InventaryCode;
        public string Code
        {
            get
            {
                if (_InventaryCode == null)
                {
                    _InventaryCode = this.ReferenceObject[param_InventaryCode_Guid].GetString();
                }
                return _InventaryCode;
            }
        }

        string _Project;
        public string Name
        {
            get
            {
                if (_Project == null)
                {
                    _Project = this.ReferenceObject[param_Project_Guid].GetString();
                }
                return _Project;
            }
        }

        #endregion
        static readonly Guid param_InventaryCode_Guid = new Guid("af881faa-8784-494c-ae3b-075161ef4ecc");
        static readonly Guid param_Project_Guid = new Guid("b1d97a96-033d-4272-a90d-c12bbde6ee7e");

    }


    public static class ScanFileNameConvention
    {

        public static bool FullSatisfy(string fileName, string designation)
        {
            string extension = fileName.Remove(0, fileName.LastIndexOf('.'));
            bool satisfy = false;
            if (fileName.StartsWith("СКАН_" + designation + "_") && fileName.IndexOf("_УЛ" + extension) < 0) satisfy = true;
            else if (fileName == "СКАН_" + designation + extension) satisfy = true;
            return satisfy;
        }
        public const char PageIntervalDelimiter = '-';
        public static bool PartlySatisfy(string fileName, string designation)
        {
            //для проверки все строки переводим в нижний регистр
            fileName = fileName.ToLower();
            designation = designation.ToLower();
            if (ScanFileNameConvention.FullSatisfy(fileName, designation)) return true;
            return false;
        }

        public static IntervalPageNumbers GetPageNumbersInterval(string fileName)
        {
            IntervalPageNumbers result;
            //СКАН_ВГИП.301431.001СБ_Л2.pdf; СКАН_ВГИП.301431.001СБ_Л5-6.pdf;
            int startIndex = fileName.ToLower().IndexOf("_л");
            if (startIndex < 0)
            {
                result.StartPage = 1;
                result.EndPage = 1;
                return result;
            }
            startIndex = startIndex + 2;
            int endIndex = fileName.ToLower().IndexOf(".pdf");
            string numbersInterval = fileName.Remove(endIndex).Remove(0, startIndex);
            int delimiterIndex = numbersInterval.IndexOf(PageIntervalDelimiter);
            if (delimiterIndex < 0)
            {
                int NumPage = 0;
                if (Int32.TryParse(numbersInterval, out NumPage))
                {
                    result.StartPage = NumPage;
                    result.EndPage = NumPage;
                    return result;
                }
                else throw new InvalidCastException("Не удалось получить номер страницы из имени файла - " + fileName);
            }
            else
            {
                string startPageString = numbersInterval.Remove(delimiterIndex);
                string endPageString = numbersInterval.Remove(0, delimiterIndex + PageIntervalDelimiter.ToString().Length);
                int startPageNumber;
                int endPageNumber;
                if (Int32.TryParse(startPageString, out startPageNumber) && Int32.TryParse(endPageString, out endPageNumber))
                {
                    result.StartPage = startPageNumber;
                    result.EndPage = endPageNumber;
                    return result;
                }
                else throw new InvalidCastException("Не удалось получить номер страницы из имени файла - " + fileName);
            }
        }
    }

    public struct IntervalPageNumbers
    {
        public int StartPage;
        public int EndPage;
    }

    public static class ScanFileNameConvention_Old
    {

        public static bool FullSatisfy(string fileName, string designation)
        {
            bool satisfy = false;
            if (fileName.StartsWith(designation + "_")) satisfy = true;
            else if (fileName == designation + ".pdf") satisfy = true;
            return satisfy;
        }

        public static bool PartlySatisfy(string fileName, string designation)
        {
            //для проверки все строки переводим в нижний регистр
            fileName = fileName.ToLower();
            designation = designation.ToLower();
            if (ScanFileNameConvention.FullSatisfy(fileName, designation)) return true;
            return false;
        }
    }

    public class ArchiveFile
    {
        public ArchiveFile(FileObject file)
        {
            this.FileObject = file;
        }

        public delegate bool Convention(String fileName, string designation);

        Convention _CurrentConvention;
        public void SetConvention(Convention conv)
        {
            _CurrentConvention = conv;
        }


        public FileObject FileObject { get; private set; }

        public string Name { get { return this.FileObject.Name; } }

        public string FolderName
        {
            get
            {
                return this.FileObject.Parent.Name;
            }
        }
        public bool SatisfyConvention(string designation)
        {
            return this._CurrentConvention(this.FileObject.Name, designation);
        }
        public override string ToString()
        {
            return FileObject.Name.ToString();
        }

    }


    /// <summary>
    /// Обёртка над именами стадий
    /// </summary>
    public class StageName
    {
        public StageName(string stageName)
        {
            this.stageName = stageName;
        }

        public string Name {

        get
            {
                return this.stageName;
            } } 

        string stageName;
        public override string ToString()
        {
            return this.stageName;
        }
    }

    /// <summary>
    /// Значения статуса инвентарной записи
    /// </summary>
    public class Status
    {
        public Status(string stringStatus)
        {
            Value = stringStatus;
        }

        public bool IsValid
        { get { return this.AllValues.Contains(this.Value); } }

        public string Value
        { get; private set; }

        static List<string> _AllValues;
        public IEnumerable<string> AllValues
        {
            get
            {
                if (_AllValues == null)
                {
                    _AllValues = new List<string>{"Действителен", "Аннулирован",
                        "При новом конструировании не применять", "Резерв"};
                }
                return _AllValues;
            }
        }
    }




    /// <summary>
    /// Справочники
    /// ver 1.5
    /// </summary>
    public class References
    {

        public static ServerConnection Connection
        {
            get { return ServerGateway.Connection; }
        }
        #region Справочники и классы
        private static Reference GetReference(ref Reference reference, ReferenceInfo referenceInfo)
        {
            if (reference == null)
                reference = referenceInfo.CreateReference();


            return reference;
        }


        static UserReference _UsersReference;
        /// <summary>
        /// Справочник "Группы и пользователи"
        /// </summary>
        public static UserReference UsersReference
        {
            get
            {
                if (_UsersReference == null)
                    _UsersReference = new UserReference(Connection);

                return _UsersReference;
            }
        }

        private static Reference _WorkChangeRequestReference;
        /// <summary>
        /// Справочник "Запросы коррекции работ"
        /// </summary>
        public static Reference WorkChangeRequestReference
        {
            get
            {
                if (_WorkChangeRequestReference == null)
                    return GetReference(ref _WorkChangeRequestReference, WorkChangeRequestReferenceInfo);

                return _WorkChangeRequestReference;
            }
        }

        private static ClassObject _Class_PMWorkChangeRequest;
        /// <summary>
        /// класс Запрос коррекции работ
        /// </summary>
        public static ClassObject Class_WorkChangeRequest
        {
            get
            {
                if (_Class_PMWorkChangeRequest == null)
                {
                    Guid CR_class_PMWorkChangeRequest_Guid = new Guid("f79a353c-00da-4248-bd4f-dab6543f95b0"); // Тип "Запросы коррекции работ"
                    _Class_PMWorkChangeRequest = WorkChangeRequestReference.Classes.Find(CR_class_PMWorkChangeRequest_Guid);
                }

                return _Class_PMWorkChangeRequest;
            }
        }

        private static ClassObject _Class_ProjectManagementWork;
        /// <summary>
        /// тип Работа - справочника Управление проектами
        /// </summary>
        public static ClassObject Class_ProjectManagementWork
        {
            get
            {
                if (_Class_ProjectManagementWork == null)
                {
                    Guid PM_class_Work_Guid = new Guid("c0bef497-cf64-44a7-9839-a704dc3facb2");
                    _Class_ProjectManagementWork = ProjectManagementReference.Classes.Find(PM_class_Work_Guid);
                }
                return _Class_ProjectManagementWork;
            }
        }

        private static Reference _ProjectManagementReference;
        /// <summary>
        /// Справочник "Управление проектами"
        /// </summary>
        public static Reference ProjectManagementReference
        {
            get
            {
                if (_ProjectManagementReference == null)
                    _ProjectManagementReference = GetReference(ref _ProjectManagementReference, ProjectManagementReferenceInfo);

                return _ProjectManagementReference;
            }
        }


        private static ReferenceInfo _ProjectManagementReferenceInfo;
        /// <summary>
        /// Справочник "Управление проектами"
        /// </summary>
        private static ReferenceInfo ProjectManagementReferenceInfo
        {
            get
            {
                Guid ref_ProjectManagement_Guid = new Guid("86ef25d4-f431-4607-be15-f52b551022ff"); // Справочник "Управление проектами"
                return GetReferenceInfo(ref _ProjectManagementReferenceInfo, ref_ProjectManagement_Guid);
            }
        }

        private static ReferenceInfo _WorkChangeRequestReferenceInfo;
        /// <summary>
        /// Справочник "Запросы коррекции работ"
        /// </summary>
        internal static ReferenceInfo WorkChangeRequestReferenceInfo
        {
            get
            {
                Guid WorkChangeRequests_ref_Guid = new Guid("9387cd96-d3cf-4cb6-997d-c4f4e59f8a21"); // Справочник "Запросы коррекции работ"
                return GetReferenceInfo(ref _WorkChangeRequestReferenceInfo, WorkChangeRequests_ref_Guid);
            }
        }

        static FileReference _FileReference;
        /// <summary>
        /// Справочник "Файлы"
        /// </summary>
        public static FileReference FileReference
        {
            get
            {
                if (_FileReference == null)
                    _FileReference = new FileReference(Connection);

                return _FileReference;
            }
        }
        #endregion
        #region "Используемые ресурсы"

        // Справочник "Используемые ресурсы"
        static readonly Guid ref_UsedResources_Guid = new Guid("3459a8fb-6bca-47ca-971a-1572b684e92e");        //Guid справочника - "Используемые ресурсы"
        static readonly Guid UR_class_NonConsumableResources_Guid = new Guid("8473c817-68fd-479c-a4c3-d6b3b405ea5d"); //Guid типа "Нерасходуемые ресурсы"

        private static Reference _UsedResourcesReference;
        /// <summary>
        /// Справочник "Используемые ресурсы"
        /// </summary>
        public static Reference UsedResources
        {
            get
            {
                if (_UsedResourcesReference == null)
                    _UsedResourcesReference = GetReference(ref _UsedResourcesReference, UsedResourcesReferenceInfo);

                return _UsedResourcesReference;
            }
        }

        private static ReferenceInfo _UsedResourcesReferenceInfo;
        /// <summary>
        /// Справочник "Используемые ресурсы"
        /// </summary>
        private static ReferenceInfo UsedResourcesReferenceInfo
        {
            get { return GetReferenceInfo(ref _UsedResourcesReferenceInfo, ref_UsedResources_Guid); }
        }

        private static ClassObject _Class_NonConsumableResources;
        /// <summary>
        /// класс Нерасходуемые ресурсы
        /// </summary>
        public static ClassObject Class_NonConsumableResources
        {
            get
            {
                if (_Class_NonConsumableResources == null)
                    _Class_NonConsumableResources = UsedResources.Classes.Find(UR_class_NonConsumableResources_Guid);

                return _Class_NonConsumableResources;
            }
        }
        #endregion

        #region "Ресурсы"

        // Справочник "Ресурсы"
        static readonly Guid ref_Resources_Guid = new Guid("fe80ab68-01e1-4a95-96cf-602ec877ff19");        //Guid справочника - "Ресурсы"
        private static Reference _ResourcesReference;
        /// <summary>
        /// Справочник "Ресурсы"
        /// </summary>
        public static Reference Resources
        {
            get
            {
                if (_ResourcesReference == null)
                    _ResourcesReference = GetReference(ref _ResourcesReference, ResourcesReferenceInfo);

                return _ResourcesReference;
            }
        }

        private static ReferenceInfo _ResourcesReferenceInfo;
        /// <summary>
        /// Справочник "Ресурсы"
        /// </summary>
        private static ReferenceInfo ResourcesReferenceInfo
        {
            get { return GetReferenceInfo(ref _ResourcesReferenceInfo, ref_Resources_Guid); }
        }
        #endregion

        #region "Регистрационно-Контрольные карточки (Канцелярия)"

        static readonly Guid ref_RegistryControlCards_Guid = new Guid("80831dc7-9740-437c-b461-8151c9f85e59");
        private static Reference _RegistryControlCardsReference;
        /// <summary>
        /// Справочник "Регистрационно-Контрольные карточки"
        /// </summary>
        public static Reference RegistryControlCards
        {
            get
            {
                if (_RegistryControlCardsReference == null)
                    _RegistryControlCardsReference = GetReference(ref _RegistryControlCardsReference, RegistryControlCardsReferenceInfo);

                _RegistryControlCardsReference.Refresh();
                return _RegistryControlCardsReference;
            }
        }

        private static ReferenceInfo _RegistryControlCardsReferenceInfo;
        /// <summary>
        /// Справочник "Регистрационно-Контрольные карточки"
        /// </summary>
        private static ReferenceInfo RegistryControlCardsReferenceInfo
        {
            get { return GetReferenceInfo(ref _RegistryControlCardsReferenceInfo, ref_RegistryControlCards_Guid); }
        }
        #endregion

        #region "Статистика загрузки ресурсов"

        // Справочник "Статистика загрузки ресурсов"
        static readonly Guid ref_StatisticConsumptionResource_Guid = new Guid("fb63e9e9-71b3-4afc-aed5-8312a732a60a");        //Guid справочника - "Статистика загрузки ресурсов"
        private static Reference _StatisticConsumptionResourceReference;
        /// <summary>
        /// Справочник "Статистика загрузки ресурсов"
        /// </summary>
        public static Reference StatisticConsumptionResourceReference
        {
            get
            {
                if (_StatisticConsumptionResourceReference == null)
                    _StatisticConsumptionResourceReference = GetReference(ref _StatisticConsumptionResourceReference, StatisticConsumptionResourceReferenceInfo);

                return _StatisticConsumptionResourceReference;
            }
        }

        private static ReferenceInfo _StatisticConsumptionResourceReferenceInfo;
        /// <summary>
        /// Справочник "Статистика загрузки ресурсов"
        /// </summary>
        private static ReferenceInfo StatisticConsumptionResourceReferenceInfo
        {
            get { return GetReferenceInfo(ref _StatisticConsumptionResourceReferenceInfo, ref_StatisticConsumptionResource_Guid); }
        }


        static readonly Guid SCR_class_StatisticConsumptionResource_Guid = new Guid("41606f1a-16c5-4df4-b507-360590ebfe95"); //Guid типа "Загрузка ресурса"
        static ClassObject _Class_StatisticConsumptionResource;
        /// <summary>
        /// класс Нерасходуемые ресурсы
        /// </summary>
        public static ClassObject Class_StatisticConsumptionResource
        {
            get
            {
                if (_Class_StatisticConsumptionResource == null)
                    _Class_StatisticConsumptionResource = StatisticConsumptionResourceReference.Classes.Find(SCR_class_StatisticConsumptionResource_Guid);

                return _Class_StatisticConsumptionResource;
            }
        }
        #endregion

        #region Регистрационно-контрольные карточки

        static Reference _RegistryControlCardReference;

        /// <summary>
        /// Справочник "Регистрационно-контрольные карточки"
        /// </summary>
        private Reference RegistryControlCardReference
        {
            get
            {
                if (_RegistryControlCardReference == null)

                    return GetReference(ref _RegistryControlCardReference, RegistryControlCardReferenceInfo);
                _RegistryControlCardReference.Refresh();
                return _RegistryControlCardReference;
            }
        }

        private static Guid reference_RCC_Guid = new Guid("80831dc7-9740-437c-b461-8151c9f85e59");    //Guid справочника - "Регистрационно-контрольные карточки"
        static ReferenceInfo _RCCReferenceInfo;
        private ReferenceInfo RegistryControlCardReferenceInfo
        {
            get { return GetReferenceInfo(ref _RCCReferenceInfo, reference_RCC_Guid); }
        }
        #endregion



        private static ReferenceInfo GetReferenceInfo(ref ReferenceInfo referenceInfo, Guid referenceGuid)
        {
            if (referenceInfo == null)
                referenceInfo = Connection.ReferenceCatalog.Find(referenceGuid);

            return referenceInfo;
        }
    }
}

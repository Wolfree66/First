
using Dinamika;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Users;
using TFlex.DOCs.Model.Signatures;
using DinamikaGuids;
using TFlex.DOCs.Model;

namespace WpfApp_DialogueAddSignatories
{
    delegate List<User> DelegateGetUserAtPost(string NamePost);
    delegate List<User> DelegateGetAllActualUsers();
    class Model
    {

        Guids Guids = new Guids();
        public ExtensionDinamika ExtensionDinamika;
        public Guid TechnicalSolution_LinkBaseProject_Guid = new Guid("256e404f-c1e8-45d3-bc48-25868df6d468");
        //public DelegateGetUserAtPost GetUserAtPost = ExtensionDinamika.НайтиПользоватейПоДолжности;
        //public DelegateGetAllActualUsers GetAllActualUsers = ExtensionDinamika.ПолучитьВсехПользователейОрганизационнойСтруктуры;
        public ReferenceObject CurrentTechnicalSolution { get; set; }
        // public User CurrentUser { get; private set; }

        public User ChiefDesignerDirection { get; set; }

        public User ProgramManager { get; set; }

        public IReadOnlyList<User> AllUsers { get; set; }

        public Dictionary<string, List<User>> UsersWithEquallyPost { get; set; }

        public Model(ServerConnection connection, ReferenceObject CurrentTechnicalSolution, string nameServer = "TFLEX")
        {
            if (ExtensionDinamika == null)
                ExtensionDinamika = new ExtensionDinamika(connection, nameServer);

            if (UsersWithEquallyPost == null)
                UsersWithEquallyPost = new Dictionary<string, List<User>>();

            this.CurrentTechnicalSolution = CurrentTechnicalSolution;

            //  CurrentUser = CurrentTechnicalSolution.SystemFields.Author;
        }

        public Model(ServerConnection connection, string nameServer = "TFLEX")
        {

            if (ExtensionDinamika == null)
                ExtensionDinamika = new ExtensionDinamika(connection, nameServer);

            if (UsersWithEquallyPost == null)
                UsersWithEquallyPost = new Dictionary<string, List<User>>();


            this.CurrentTechnicalSolution = ExtensionDinamika.Connection.ReferenceCatalog.Find(Guids.NomenclatureGuids.Ref_Guid)
                .CreateReference().Find(new Guid("a82345ef-2a00-4388-9ff9-b30069904fe2"));

            //   CurrentUser = CurrentTechnicalSolution.SystemFields.Author;

            ReloadDataUsers();
        }

        public void ShowErrorMsg(string text, string caption)
        {
            System.Windows.Forms.MessageBox.Show(text,
                     caption,
                     System.Windows.Forms.MessageBoxButtons.OK,
                     System.Windows.Forms.MessageBoxIcon.Error,
                     System.Windows.Forms.MessageBoxDefaultButton.Button1);
        }

        public void ReloadDataUsers()
        {
            ExtensionDinamika.UserReference.Objects.Reload();
        }

        /// <summary>
        /// Получить руководителя проекта
        /// </summary>
        public List<User> GetProjectManager(string post)
        {
            List<ComplexHierarchyLink> usersLinks = ExtensionDinamika.UserReference.Objects.RecursiveLoadHierarchyLinks();//OfType<User>();

            List<User> Result = new List<User>();
            foreach (var userLink in usersLinks)
            {
                if (userLink[Guids.UserReference.ParametersHierarchy_Post_Guid].GetString().StartsWith(post))
                {
                    User user = userLink.ChildObject as User;
                    if (user != null)
                        Result.Add(user);
                }
            }
            return Result;
        }


        public void add_UsersWithEquallyPost(string post, List<User> Users)
        {
            List<User> users;
            // если категории еще нет в словаре, создаем новый список и добавляем его в словарь
            if (!UsersWithEquallyPost.TryGetValue(post, out users))
            {
                users = new List<User>();
                UsersWithEquallyPost.Add(post, Users);
            }
        }

        /// <summary>
        /// Метод добавления подписей
        /// </summary>
        public bool AddSign()
        {
            if (CurrentTechnicalSolution == null)
            {
                ShowErrorMsg("Отсутствует тех.решение.",
"Отсутствует текущий объект!");
                return false;
            }
            else
            {
                SignatureCollection signCollection = CurrentTechnicalSolution.Signatures;
                SignatureType Согл = signCollection.Types.FirstOrDefault(s => s.Name.StartsWith("Согл."));


                try
                {
                    signCollection.Add(Согл, ChiefDesignerDirection as UserReferenceObject);
                }
                catch
                {
                    MessageBox.Show("Произошла ошибка при добавления подписи для главного конструктора по направлению");
                    return false;
                }

                if (ProgramManager != null)
                    try
                    {
                        signCollection.Add(Согл, ProgramManager as UserReferenceObject);
                    }
                    catch
                    {
                        MessageBox.Show("Произошла ошибка при добавления подписи для Руководитель программы ");
                        return false;
                    }

                return true;
                //SignatureType Утв = signCollection.Types.FirstOrDefault(s => s.Name.StartsWith("Утв."));
                //SignatureType Разраб = signCollection.Types.FirstOrDefault(s => s.Name.StartsWith("Разраб."));
                //if (Разраб == null)
                //    MessageBox.Show("Тип подписи Разраб не найден");
                // else
                //{
                //    signCollection.Add(Разраб, CurrentUser as UserReferenceObject);
                //}

                //if (Согл == null)
                //    MessageBox.Show("Тип подписи согл не найден");
                // else
                //{

                // try
                // {
                //     signCollection.Add(Согл, GetUserAtPost("Начальник опытно-конструкторского бюро").FirstOrDefault() as UserReferenceObject);
                // }
                //catch
                //{
                //    MessageBox.Show("Произошла ошибка при добавления подписи для Начальника опытно-конструкторского бюро ");
                // }
                //try
                //{
                //    signCollection.Add(Согл, GetUserAtPost("Начальник службы качества").FirstOrDefault() as UserReferenceObject);
                //}
                //catch
                // {
                //    MessageBox.Show("Произошла ошибка при добавления подписи для Начальника службы качества ");
                // }
                //                    ReferenceObject LinkedProjectFromBaseProjects = CurrentTechnicalSolution.GetObject(TechnicalSolution_LinkBaseProject_Guid);
                //                    if (LinkedProjectFromBaseProjects == null)
                //                        System.Windows.Forms.MessageBox.Show("У технического решения не указан номер проекта!",
                //"Укажите номер проекта.",
                //System.Windows.Forms.MessageBoxButtons.OK,
                //System.Windows.Forms.MessageBoxIcon.Exclamation,
                //System.Windows.Forms.MessageBoxDefaultButton.Button1);
                //                    else
                //                    {
                //                        try
                //                        {
                //                            signCollection.Add(Согл, ExtensionDinamika.НайтиПользователяПоКодуФизЛица(LinkedProjectFromBaseProjects[Guids.BaseProject.CodUser1C_Guid].GetString()) as UserReferenceObject);
                //                        }
                //                        catch (Exception)
                //                        {

                //                            MessageBox.Show("Произошла ошибка при добавления подписи для Руководителя проекта ");
                //                        }


                //                    }

                //                }
                //                if (Утв == null)
                //                    MessageBox.Show("Тип подписи Утв не найден");
                //                else
                //                {
                //                    try
                //                    {
                //                        signCollection.Add(Утв, GetUserAtPost("Исполнительный директор").FirstOrDefault() as UserReferenceObject);
                //                    }
                //                    catch
                //                    {
                //                        MessageBox.Show("Произошла ошибка при добавления подписи для Исполнительного директор ");
                //                    }
                //}


            }
        }
    }
}

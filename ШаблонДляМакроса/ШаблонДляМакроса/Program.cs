//#define TF_TEST


using System;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.References.ProjectManagement;

namespace ШаблонДляМакроса
{
    class Program
    {
        [STAThread]
        public static void Main()
        {
            ServerConnection sc = ConnectToDocs();
            MacroContext mc = new MacroContext(sc);
            dynamic macro = new Cancelaria.Macro(mc);
            macro.Test();
            return;
            //var project = References.ProjectManagementReference.Find(346161); //План ОПК
            //var dialog = new Report.Views.Report_View(project);
            //dialog.ShowDialog();
            Console.ReadKey();
        }

        private static void TestPM_Report()
        {
            //ProjectElement testProjectElement = References.ProjectManagementReference.Find(new Guid("7e3fac43-19b7-4268-b505-dc4dfcbbe104")) as ProjectElement;//тестовый проект
            //ProjectElement testProjectElement = References.ProjectManagementReference.Find(new Guid("99af9eb2-0c1a-4836-8b82-0ba6ebec3d2d")) as ProjectElement;//Утверждённые проекты из плана ОКБ
            
                    //Thread Messagethread = new Thread(new ThreadStart(delegate ()
                    //{

                    //PM_Report1_Window dial = new PM_Report1_Window(testProjectElement);
                    //dial.ShowDialog();
                    //}));
                    //Messagethread.SetApartmentState(ApartmentState.STA);
                    //Messagethread.Start();
                    //PM_Report_Dialog1 dial = new PM_Report_Dialog1(testProjectElement);
            //dial.ProjectElement = testProjectElement;

            //MessageBox.Show("Start");
            //dial.ShowDialog();
            //Console.ReadKey();
        }

        

        static ServerConnection ConnectToDocs()
        {
#if TF_TEST
                //ServerGateway.Connect("administrator", new MD5HashString("saTflexTest1"), "TF-test");
                ServerGateway.Connect("TF-TEST\\TF_TEST15");
#else

            ServerGateway.Connect("TFLEX");
            //ServerGateway.Connect("administrator", new MD5HashString("MHMr2QqQae"), "TFLEX");
#endif
            if (!ServerGateway.Connect(false))
            {
                Console.WriteLine("Не удалось подключиться к серверу.");
                return null;
            }
            //else { SaveToFile("Подключение к серверу, успешно."); }
#if !TF_TEST
            TFlex.DOCs.Model.Plugins.AssemblyLoader.LoadAssembly("TFlex.DOCs.ProjectManagement.dll");
            TFlex.DOCs.Model.Plugins.AssemblyLoader.LoadAssembly("TFlex.DOCs.UI.Common.dll");

            TFlex.DOCs.Model.Plugins.AssemblyLoader.LoadAssembly("TFlex.DOCs.UI.Objects.dll");
            TFlex.DOCs.Model.Plugins.AssemblyLoader.LoadAssembly("TFlex.DOCs.UI.Types.dll");
            TFlex.DOCs.Model.Plugins.AssemblyLoader.LoadAssembly("TFlex.DOCs.UI.Mail.dll");
#endif
            //ReferenceCatalog.RegisterSpecialReference(ProjectManagement);
            return ServerGateway.Connection;
        }
    }
}

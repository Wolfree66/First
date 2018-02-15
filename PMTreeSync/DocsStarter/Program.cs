using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.References;
using TFLEXDocsClasses;
using WorkSelector.Models;

namespace DocsStarter
{
    class Program
    {
        static void Main(string[] args)
        {
            ServerConnection sc = ConnectToDocs();
            //MacroContext mc = new MacroContext(sc);
            //dynamic macro = new UserRequests.Macro(mc);
            //macro.Test();
            //return;
            Test();
            Console.ReadKey();
        }

        private static void Test()
        {
            ReferenceObject ro = References.ProjectManagementReference.Find(269370);
            ProjectManagementWork work = new ProjectManagementWork(ro);
            ProjectTreeSynchronizer pts = new ProjectTreeSynchronizer(work);
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

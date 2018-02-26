namespace WpfApp_DialogueAddSignatories.Model
{
    static public class ControlDialog
    {
        static ControlDialog()
        {
            IsSyncRes = false;
            IsSyncOnlyPlanRes = false;
        }
        static public bool IsSyncRes { get; set; }



        static public bool IsSyncOnlyPlanRes { get; set; }

    }
}

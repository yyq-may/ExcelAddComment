using System;
using System.Activities;
using System.Activities.XamlIntegration;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            var workflow = ActivityXamlServices.Load("Activity1.xaml", new ActivityXamlServicesSettings() { CompileExpressions = true });
            WorkflowInvoker.Invoke(workflow);
        }
    }
}

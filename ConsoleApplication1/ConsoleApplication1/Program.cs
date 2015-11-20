using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Threading;
using System.Windows.Automation;
using System.Windows.Automation.Text;
using System.Windows.Automation.Provider;

namespace ConsoleApplication1
{
    class Program
    {
        //获取对主窗体对象的引用，该对象实际上就是Form1应用程序
        public static AutomationElement FindWindowByProcessId(int processId)
        {
            AutomationElement targetWindow = null;
            int count = 0;
            try
            {
                Process p = Process.GetProcessById(processId);
                targetWindow = AutomationElement.FromHandle(p.MainWindowHandle);
                return targetWindow;
            }
            catch(Exception ex)
            {
                count++;
                StringBuilder sb = new StringBuilder();
                string message = sb.AppendLine(string.Format("Target window is not existing.try #{0}", count)).ToString();
                if (count > 5)
                {
                    throw new InvalidProgramException(message, ex);
                }
                else
                {
                    return FindWindowByProcessId(processId);
                }
            }
        }

        //根据NameProperty获取元素对象
        public static AutomationElement FindElementById(AutomationElement appForm,string automationId)
        {
            AutomationElement tarFindElement = appForm.FindFirst(TreeScope.Descendants,new PropertyCondition(AutomationElement.AutomationIdProperty, automationId));
            return tarFindElement;
        }

        //根据NameProperty获取元素数组
        public static AutomationElementCollection FindElementCollectionByControlType(AutomationElement appForm, ControlType controlType)
        {
            AutomationElementCollection tarFindElement = appForm.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.ControlTypeProperty, controlType));
            return tarFindElement;
        }

        //获取ComboBox 控件元素
        public static ExpandCollapsePattern GetExpandCollapsePattern(AutomationElement element)
        {
            object currentPattern;
            if (!element.TryGetCurrentPattern(ExpandCollapsePattern.Pattern, out currentPattern))
            {
                throw new Exception(string.Format("Element with AutomationId '{0}' and Name '{1}' does not support the ExpandCollapsePattern.",element.Current.AutomationId, element.Current.Name));
            }
            return currentPattern as ExpandCollapsePattern;
        }

        //点击按钮触发
       public static InvokePattern GetInvokePattern(AutomationElement element)
       {
            object currentPattern;
            if (!element.TryGetCurrentPattern(InvokePattern.Pattern, out currentPattern))
            {
            throw new Exception(string.Format("Element with AutomationId '{0}' and Name '{1}' does not support the InvokePattern.",element.Current.AutomationId, element.Current.Name));
            }
            return currentPattern as InvokePattern;
        }

       //获取文本框元素 ，读取/修改文本框的值
       public static ValuePattern GetValuePattern(AutomationElement element)
       {
           object currentPattern;
           if (!element.TryGetCurrentPattern(ValuePattern.Pattern, out currentPattern))
           {
               throw new Exception(string.Format("Element with AutomationId '{0}' and Name '{1}' does not support the ValuePattern.", element.Current.AutomationId, element.Current.Name));
           }
           return currentPattern as ValuePattern;
       }

       //获取窗口元素 
       public static WindowPattern GetWindowPattern(AutomationElement element)
       {
           object currentPattern;
           if (!element.TryGetCurrentPattern(WindowPattern.Pattern, out currentPattern))
           {
               throw new Exception(string.Format("Element with AutomationId '{0}' and Name '{1}' does not support the WindowPattern.", element.Current.AutomationId, element.Current.Name));
           }
           return currentPattern as WindowPattern;
       }

       //获取选择框元素 
       public static SelectionItemPattern GetSelectionItemPattern(AutomationElement element)
       {
           object currentPattern;
           if (!element.TryGetCurrentPattern(SelectionItemPattern.Pattern, out currentPattern))
           {
               throw new Exception(string.Format("Element with AutomationId '{0}' and Name '{1}' does not support the SelectionItemPattern.", element.Current.AutomationId, element.Current.Name));
           }
           return currentPattern as SelectionItemPattern;
       }

       //获取 CheckBox， TreeView 中的 button 控件元素 
       public static TogglePattern GetTogglePattern(AutomationElement element)
       {
           object currentPattern;
           if (!element.TryGetCurrentPattern(TogglePattern.Pattern, out currentPattern))
           {
               throw new Exception(string.Format("Element with AutomationId '{0}' and Name '{1}' does not support the TogglePattern.", element.Current.AutomationId, element.Current.Name));
           }
           return currentPattern as TogglePattern;
       }


        //获取多选框元素 index数组表示选择的的哪几项，rows是多选列表的各个选项
       public static void MutlSelect(int[] indexes, AutomationElementCollection rows, bool isSelectAll)
       {
           object multiSelect;
           if (isSelectAll)
           {
               for (int i = 1; i < rows.Count - 1; i++)
               {
                   if (rows[i].TryGetCurrentPattern(SelectionItemPattern.Pattern, out multiSelect))
                   {
                       (multiSelect as SelectionItemPattern).AddToSelection();
                   }
               }
           }
           else
           {
               if (indexes.Length > 0)
               {
                   for (int j = 0; j < indexes.Length; j++)
                   {
                       int tempIndex = indexes[j];
                       if (rows[tempIndex].TryGetCurrentPattern(SelectionItemPattern.Pattern, out multiSelect))
                       {
                           (multiSelect as SelectionItemPattern).AddToSelection();
                       }
                   }
               }
           }
       }

        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("\nBegin WinForm UIAutomayion test run\n");
                Console.WriteLine("Launching WinFormTest application");
               
                //启动被测程序
                string appPath = "c:\\users\\geek\\documents\\visual studio 2013\\Projects\\WpfApplication1\\WpfApplication1\\bin\\Debug\\WpfApplication1.exe";
                Process process = Process.Start(@appPath);
                //获取被测程序的主窗口
                AutomationElement appForm = FindWindowByProcessId(process.Id);

                //下拉菜单获取
                AutomationElement combox = FindElementById(appForm, "comboBox1");
                ExpandCollapsePattern currentpattern = GetExpandCollapsePattern(combox);
                currentpattern.Expand();
                Thread.Sleep(1000);
                currentpattern.Collapse();

                //点击按钮
                AutomationElement btn = FindElementById(appForm, "button1");
                InvokePattern ipBtn = GetInvokePattern(btn);
                ipBtn.Invoke();
                Thread.Sleep(1000);


                //文本框操作
                AutomationElement textBox = FindElementById(appForm, "textBox1");
                ValuePattern vpTextBox = GetValuePattern(textBox);
                vpTextBox.SetValue("Kimi @163.com");

                //选择按钮的操作
                SelectionItemPattern sip = GetSelectionItemPattern(FindElementById(appForm, "radioButton1"));
                sip.Select();
                Thread.Sleep(1000);
                sip.Select();

                TogglePattern tp = GetTogglePattern(FindElementById(appForm, "checkBox1"));
                bool isCheck = tp.Current.ToggleState == ToggleState.On;//三种状态 On Off Indeterminate
                if(!isCheck)
                {
                    tp.Toggle();
                }
                Thread.Sleep(1000);
                tp.Toggle();

                //下拉列表操作
                AutomationElement list = FindElementById(appForm, "listView1");
                AutomationElementCollection rows = FindElementCollectionByControlType(list, ControlType.ListItem);
                MutlSelect(new int[] {1}, rows, false);
                Thread.Sleep(3000);



                //对窗口大小/关闭的操作
                WindowPattern wp = GetWindowPattern(appForm);
                //wp.SetWindowVisualState(WindowVisualState.Maximized);           //最大化
                //Thread.Sleep(1000);
                //wp.SetWindowVisualState(WindowVisualState.Minimized);           //最小化
                //Thread.Sleep(1000);
                //wp.SetWindowVisualState(WindowVisualState.Normal);              //正常值
                //Thread.Sleep(1000);
                wp.Close();                                                     //关闭窗口

               
            }
            catch (Exception ex)
            {
                Console.WriteLine("Fatal error:" + ex.Message);
            }
        }
    }
}

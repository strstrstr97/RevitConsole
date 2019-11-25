using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit;
using Autodesk.Revit.DB;

namespace RevitConsole
{
    class Program
    {
        static RevitContext revit;
        static Program()
        {
            RevitContext.AddEnv(@"D:\Program Files\Autodesk\Navisworks Manage 2020\Loaders\Rx\");
        }
        [STAThread]//一定要有
        static void Main(string[] args)
        {
            revit = new RevitContext();
            revit.InitRevitFinished += InitRevitFinished;
            revit.InitRevit();
            Console.ReadKey();
        }

        private static void InitRevitFinished(object sender, Product revitProduct)
        {
            Console.WriteLine("当前使用Revit版本为：" + revitProduct.Application.VersionName);

            Document document = revit.OpenFile(@"E:\test\2019\经典小文件\2020.rvt");

            View3D view = revit.GetView3D(document);
            if (view!=null)
            {
                Console.WriteLine(view.Name);
                var elements =revit.GetElementsWithView(view);
                foreach (var element in elements)
                {
                    Console.WriteLine(element.Name);
                }
            }

        }
    }
    public class RevitContext
    {
        #region private fields

        Product _revitProduct;
        private static bool isLoadEnv = false;//是否已添加过环境变量

        #endregion

        #region public fields

        /// <summary>
        /// revit程序目录
        /// </summary>
        public static string RevitPath;

        #endregion
        #region event

        public event EventHandler<Product> InitRevitFinished;

        #endregion
        #region public properties
        /// <summary>
        /// 打开REVIT文件时的设置
        /// </summary>
        public OpenOptions OpenOptions { get; set; }
        /// <summary>
        /// Revit Application
        /// </summary>
        public Autodesk.Revit.ApplicationServices.Application Application => this._revitProduct?.Application;
        #endregion
        #region constructors
        /// <summary>
        /// 
        /// </summary>
        /// <param name="revitPath">revit安装目录</param>
        public RevitContext(string revitPath)
        {
            RevitPath = revitPath;
            AddEnv();
        }
        /// <summary>
        /// 使用此构造方法前需要调用 RevitContext.AddEnv();
        /// </summary>
        public RevitContext()
        {
           
        }

        #endregion

        #region public methods
        public void InitRevit()
        {
            this.OpenOptions = new OpenOptions
            {
                Audit = true,
                AllowOpeningLocalByWrongUser = false,
                DetachFromCentralOption = DetachFromCentralOption.DetachAndDiscardWorksets //从中心模型分离
            };
            _revitProduct = Product.GetInstalledProduct();
            var clientApplicationId = new ClientApplicationId(Guid.NewGuid(), "RevitContext", "BIM");
            _revitProduct.SetPreferredLanguage(Autodesk.Revit.ApplicationServices.LanguageType.Chinese_Simplified);
            _revitProduct.Init(clientApplicationId, "I am authorized by Autodesk to use this UI-less functionality.");
            OnInitRevitFinished();
        }
        public Document OpenFile(string filename, OpenOptions options = null)
        {
            if (options == null)
            {
                options = this.OpenOptions;
            }
            ModelPath model = new FilePath(filename);
            return this._revitProduct.Application.OpenDocumentFile(model, options);
        }
        /// <summary>
        /// 获取默认三维视图
        /// </summary>
        /// <param name="document">文档</param>
        /// <returns></returns>
        public View3D GetView3D(Document document)
        {
            if (document.ActiveView is View3D view3D && !view3D.IsPerspective && view3D.CanBePrinted)
            {
                return view3D;
            }
            FilteredElementCollector filter=new FilteredElementCollector(document);
            return (View3D) filter.OfClass(typeof(View3D)).FirstElement();
        }

        /// <summary>
        /// 获取指定三维视图
        /// </summary>
        /// <param name="document">文档</param>
        /// <param name="viewName">指定视图名称</param>
        /// <returns></returns>
        public View3D GetView3D(Document document,string viewName)
        {
            FilteredElementCollector filter = new FilteredElementCollector(document);
            return (View3D)filter.OfClass(typeof(View3D)).FirstOrDefault(x => x.Name==viewName);
        }

        public IList<Element> GetElementsWithView(View3D view)
        {
            FilteredElementCollector collector=new FilteredElementCollector(view.Document,view.Id);
            return collector.ToElements();
           
        }

        #endregion

        #region public static methods
        /// <summary>
        /// 添加revit安装路径到环境变量以便加载相应的DLL
        /// </summary>
        /// <param name="revitPath">添加revit安装路径</param>
        public static void AddEnv(string revitPath=null)
        {
            if (isLoadEnv)
            {
                return;
            }

            if (revitPath!=null)
            {
                RevitPath = revitPath;
            }
            AddEnvironmentPaths(RevitPath);
            AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;
        }



        #endregion

        #region private static methods

        /// <summary>
        /// 添加环境变量
        /// </summary>
        /// <param name="paths">revit安装路径</param>
        static void AddEnvironmentPaths(params string[] paths)
        {
            string[] first = {
                Environment.GetEnvironmentVariable("PATH") ?? string.Empty
            };
            string value = string.Join(Path.PathSeparator.ToString(), first.Concat(paths));
            Environment.SetEnvironmentVariable("PATH", value);
        }
        /// <summary>
        /// 动态加载revit相关的dll
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        /// <returns></returns>
        static Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            var assemblyName = new AssemblyName(args.Name);
            var text = $"{Path.Combine(RevitPath, assemblyName.Name)}.dll";
            Assembly result;
            if (File.Exists(text))
            {
                Console.WriteLine($"Load Revit Dll Path:{text}");
                result = Assembly.LoadFrom(text);
            }
            else
            {
                result = null;
            }
            return result;
        }

        #endregion
        #region private methods

        private void OnInitRevitFinished()
        {
            this.InitRevitFinished?.Invoke(this, this._revitProduct);
        }



        #endregion

    }
}

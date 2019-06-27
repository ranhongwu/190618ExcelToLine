using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Aspose.Cells;
using ESRI.ArcGIS.Geometry;
using ESRI.ArcGIS.Display;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.Geodatabase;

namespace _190618ExcelToLine
{
    public partial class Form1 : Form
    {
        Cells cells;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //修改路径
            ReadExcel("C:\\Users\\rhw\\Desktop\\190618ExcelToLine\\190618ExcelToLine\\data\\各区域间出行分布.xls");
            axMapControl1.Refresh();
        }

        //读取excel
        private void ReadExcel(string path)
        {
            Workbook workbook = new Workbook(path);
            cells = workbook.Worksheets[0].Cells;
            for (int i = 1; i < cells.MaxDataRow + 1; i++)
            {
                Data2Line(cells[i, 3].DoubleValue, cells[i, 4].DoubleValue,
                cells[i, 5].DoubleValue, cells[i, 6].DoubleValue);
            }
        }

        //各行数据生成线
        private void Data2Line(double lonF, double latF, double lonT, double latT)
        {
            IPoint point_From = new PointClass();
            IPoint point_To = new PointClass();
            IPolyline pPolyline = new PolylineClass();
            point_From.PutCoords(lonF, latF);
            point_To.PutCoords(lonT, latT);
            pPolyline.FromPoint = point_From;
            pPolyline.ToPoint = point_To;
            ILayer layer =  this.axMapControl1.Map.Layer[0];
            IFeatureLayer featureLayer = (IFeatureLayer)layer;
            IFeature feature = featureLayer.FeatureClass.CreateFeature();
            feature.Shape = pPolyline;
            feature.Store();
        }
    }
}

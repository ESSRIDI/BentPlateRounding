using System;
using System.Windows.Forms;
using System.Collections;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

using Tekla.Structures.Model;
using Tekla.Structures.Model.UI;
using Tekla.Structures.Model.Operations;
using Tekla.Structures.Geometry3d;

using Point = Tekla.Structures.Geometry3d.Point;

namespace BentPlateRounding
{
    public partial class Form1 : Form
    {
      
        
        
        BentPlate bP;
        PolyBeam polyBeam;
        ContourPoint pointC;
        Tekla.Structures.Model.Model model;

        public Form1()
        {
            InitializeComponent();
          
            this.bP = null;
            this.model = new Tekla.Structures.Model.Model();
            this.polyBeam = null;
            this.pointC = null;



        }


        /// <summary>
        /// Open an Excel cfile which included Raduis of each plate thickness
        /// </summary>
        /// <param name="path"></param>
        public void OpenFile(string path)
        {
            Excel excel = new Excel(path, 1);
            MessageBox.Show(excel.ExcelReaderCell(0, 0));
        }

        /// <summary>
        /// Seek Type of plate and change its radius
        /// </summary>
        private void ModifyRoundingPlate()
        {
            // testing raduis
            Dictionary<double, double> rightRidius = new Dictionary<double, double>();
            rightRidius.Add(5, 8);
            rightRidius.Add(6, 10);
            rightRidius.Add(8, 14);
            rightRidius.Add(10, 18);
            rightRidius.Add(12, 22);
            rightRidius.Add(14, 26);
            rightRidius.Add(15, 28);
            rightRidius.Add(16, 32);
            rightRidius.Add(18, 38);
            rightRidius.Add(20, 42);
            rightRidius.Add(25, 42);
            rightRidius.Add(30, 60);

            Picker _picker = new Picker();

            var selectedObj = _picker.PickObjects(Picker.PickObjectsEnum.PICK_N_PARTS, "Please select the PolyBeam plate :");

            foreach (Part obj in selectedObj)
            {
                if ((obj != null) && ((obj is BentPlate) == true))
                {

                   this.bP = obj as BentPlate;
                    var platesEnum = this.bP.Geometry.GetGeometryEnumerator();
                    while (platesEnum.MoveNext())
                    {
                        double plateThk= this.bP.Thickness;

                        if (platesEnum.Current.GeometryNode is CylindricalSurfaceNode)
                        {
                            try
                            {
                                BentPlateGeometrySolver solver = new BentPlateGeometrySolver();
                                var changedGeometry = solver.ModifyRadius(this.bP.Geometry, platesEnum.Current, rightRidius[plateThk] + (plateThk / 2));

                                this.bP.Geometry = changedGeometry;
                                this.bP.Modify();
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("Exception: " + ex.ToString());
                            }
                        }
                        this.model.CommitChanges();

                    }
                }
                else
                    if ((obj != null) && ((obj is PolyBeam) == true))
                {
                    polyBeam = obj as PolyBeam;

                    string polyBeameProfile = polyBeam.Profile.ProfileString;
                    string[] profileSplitedStringArray = Regex.Split(polyBeameProfile, @"\D+");
                    List<int> intProfile = new List<int>();
                    int thkProfile = 0;

                    foreach (string value in profileSplitedStringArray)
                    {
                        if (!string.IsNullOrEmpty(value))
                        {
                            intProfile.Add(int.Parse(value));

                        }

                    }
                    if (intProfile.Count == 2)
                    {
                        thkProfile = intProfile[0];

                        if (thkProfile > intProfile[1])
                        {
                            thkProfile = intProfile[1];
                        }

                        var contour = this.polyBeam.Contour.ContourPoints;
                        ArrayList contourPoints = new ArrayList();
                        foreach (Point p in contour)
                        {
                            this.pointC = new ContourPoint(p, new Chamfer(rightRidius[thkProfile] , 0, Chamfer.ChamferTypeEnum.CHAMFER_ROUNDING));
                            contourPoints.Add(this.pointC);
                        }
                        var firstPoint = contourPoints[0] as Point;
                        var lastPoint = contourPoints[(contour.Count)-1] as Point;

                        contourPoints[0]= new ContourPoint(firstPoint, new Chamfer(0, 0, Chamfer.ChamferTypeEnum.CHAMFER_NONE));
                        contourPoints[(contour.Count) - 1] = new ContourPoint(lastPoint, new Chamfer(0, 0, Chamfer.ChamferTypeEnum.CHAMFER_NONE));

                        this.polyBeam.Contour.ContourPoints = contourPoints;
                        this.polyBeam.Modify();
                        this.model.CommitChanges();


                     
                    }

                   
                }
            }

            Operation.DisplayPrompt("Plate rounding radius operation has been successfully completed");
        }


        #region ! Method excluded to seek type of profile !  
        /* foreach (string str in strPlateProfile)
         {


             if (plateProfile.StartsWith(str) == true)
             {



                 var contour = this.polyBeam.Contour.ContourPoints;
                 ArrayList contourPoints = new ArrayList();
                 foreach (Point p in contour)
                 {
                     this.pointC = new ContourPoint(p, new Chamfer(50, 0, Chamfer.ChamferTypeEnum.CHAMFER_ROUNDING));
                     contourPoints.Add(this.pointC);
                 }
                 this.polyBeam.Contour.ContourPoints = contourPoints;
                 this.polyBeam.Modify();
                 this.model.CommitChanges();


                 MessageBox.Show($"your profile start with ", "Profile");
                 isRightselection = true;

                 break;
             }

         }




    }


} while (!isRightselection);

}

}*/
        #endregion
        private void SelectAllPlate()
        {
            ArrayList plateToSelect = new ArrayList();
            ModelObjectEnumerator ObjectEnum = this.model.GetModelObjectSelector().GetAllObjects();
            while (ObjectEnum.MoveNext())
            {
                if (ObjectEnum.Current is BentPlate)
                {
                    plateToSelect.Add(ObjectEnum.Current);
                }
                
            }

            var MS = new Tekla.Structures.Model.UI.ModelObjectSelector();

            MS.Select(plateToSelect);
            this.model.CommitChanges();
        }


        private void SelectAllPlate2()
        {
            ArrayList plateToSelect = new ArrayList();
            ModelObjectEnumerator ObjectEnum = this.model.GetModelObjectSelector().GetAllObjects();
            while (ObjectEnum.MoveNext())
            {
                if (ObjectEnum.Current is BentPlate)
                {
                    plateToSelect.Add(ObjectEnum.Current);
                }


            }

            var MS = new Tekla.Structures.Model.UI.ModelObjectSelector();

            MS.Select(plateToSelect);
            this.model.CommitChanges();
        }




        private void Form1_Load(object sender, EventArgs e)
        {



        }

        private void mdfRd_btn_Click(object sender, EventArgs e)
        {

            ModifyRoundingPlate();

        }

        private void radiusToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //openFileDialog1.ShowDialog();

        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            OpenFileDialog choofdlog = new OpenFileDialog();
            choofdlog.Filter = "All Files (*.*)|*.*";
            choofdlog.FilterIndex = 1;
            choofdlog.Multiselect = false;
            string sFileName = openFileDialog1.ShowDialog().ToString();
            OpenFile(sFileName);
        }
    }
}


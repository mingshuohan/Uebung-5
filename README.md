# Uebung-5
nicht völlig
using SldWorks;


using SwConst;



using Mathcad;



　


　


namespace Uebung_5


{


    public partial class Form1 : Form


    {


        SldWorks.SldWorks swApp;


        


        public Form1()


        {


            InitializeComponent();


            try


            {


                swApp = (SldWorks.SldWorks)Marshal.GetActiveObject("SolidWorks.Application");



            }


            catch(Exception)


            {


                swApp = new SldWorks.SldWorks();


                swApp.Visible = true;


            }


        }



　


        


        private void button1_Click(object sender, EventArgs e)


        {


            //double a=0, b=0, c;


            //if(double.TryParse(tb_a.Text,out a)==false||


            //    double.TryParse(tb_b.Text,out b)==false)


            //{


            //    MessageBox.Show("fehlerhafte Eingabe");


            //}



            //Mathcad.Application mcApp = (Mathcad.Application)Marshal.GetActiveObject("Mathcad.Application");


            //Mathcad.Worksheet mcWks = mcApp.ActiveWorksheet;



            //mcWks.SetValue("a", a);


            //mcWks.SetValue("b", b);



            //Mathcad.NumericValue mcNum=mcWks.GetValue("c");


            //c=mcNum.Real;



　


　


            Mathcad.Application mcApp = (Mathcad.Application)Marshal.GetActiveObject(


                                        "Mathcad.Application");


            Mathcad.Worksheet mcWks = mcApp.ActiveWorksheet;



            double phi=0,lh=50,a0x=30,a0y=30,rg=20,z=50,x,y;



            mcWks.SetValue("phi", phi);


            mcWks.SetValue("lh",lh);


            mcWks.SetValue("a0x",a0x);


            mcWks.SetValue("a0y",a0y);


            mcWks.SetValue("rg",rg);


            mcWks.SetValue("pcnt",z);


            mcWks.SetValue("pact", 0);


            Mathcad.NumericValue mcNum1 = mcWks.GetValue("x");


            x = mcNum1.Real;


            Mathcad.NumericValue mcNum2 = mcWks.GetValue("y");


            y = mcNum2.Real;


            listBox1.Items.Add(x + ";" + y);



            for (int a = 0; a < 50; a++)


            {


                mcWks.SetValue("pact",a+1);


                mcNum1 = mcWks.GetValue("x");


                x= mcNum1.Real;


                mcNum2 = mcWks.GetValue("y");


                y= mcNum2.Real;


                listBox1.Items.Add(x+";"+y);


            }



　


　


　


            List<double> l1 = new List<double>[z];


            l1.Add(x);


            List<double> l2 = new List<double>[z];


            l2.Add(y);



　


            ModelDoc2 swModel;


            Sketch swSketch;


            Feature swFeat;



            swModel = swApp.NewPart();


            swModel.Extension.SelectByID("Ebene vorne", "PLANE", 0, 0, 0, false, 1, null);


            swModel.InsertSketch2(true);



            swSketch = (Sketch)swModel.GetActiveSketch2();


            swFeat = (Feature)swSketch;


            swFeat.Name = "Grundflaeche";



　


　


            SketchPoint pact = new SketchPoint();



　


　


            for (int i = 0; i < z; i++)


            {


                swModel.CreateLine2(l1[i], l2[i], 0, l2[i + 1], l2[i + 1], 0);


            }


            swModel.CreateLine2(l1[z], l2[z], 0, l1[0], l2[0], 0);



            Feature swFeat2 = swModel.FeatureManager.FeatureExtrusion2(true, false, false,


                0, 0, hoehe, 0, false, false, false, false, 0, 0, false, false, false, false,


                true, true, true, 0, 0, true);



            swFeat2.Name = "Schleife";



            

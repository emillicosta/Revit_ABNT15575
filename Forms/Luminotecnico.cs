using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Lighting;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;

namespace Forms
{
    internal class Luminotecnico : System.Windows.Forms.Form
    {
        private System.Windows.Forms.TextBox textBoxFU;
        private System.Windows.Forms.Label labelFPL;
        private System.Windows.Forms.ComboBox comboBoxFPL;
        private System.Windows.Forms.ComboBox comboBoxAE;
        private System.Windows.Forms.Label labelAE;
        private System.Windows.Forms.Button buttonCalcular;
        private System.Windows.Forms.GroupBox groupBoxCalculo;
        private System.Windows.Forms.GroupBox groupBoxReflexao;
        private System.Windows.Forms.Label labelPiso;
        private System.Windows.Forms.Label labelParede;
        private System.Windows.Forms.Label labelTeto;
        private System.Windows.Forms.TextBox textBoxPiso;
        private System.Windows.Forms.TextBox textBoxParede;
        private System.Windows.Forms.TextBox textBoxTeto;
        private System.Windows.Forms.Label labelFU;

        private readonly double k;
        private Space space;
        private double area;
        private double aei;

        public Luminotecnico(Space space)
        {
            this.space = space;
            k = GetK();
            GetReflectance(space.Document);
            InitializeComponent();
        }

        private void GetReflectance(Document doc)
        {
            List<double> coefficientReflectance = new List<double>();
            SpatialElementGeometryCalculator calculator = new SpatialElementGeometryCalculator(doc);
            SpatialElementGeometryResults results = calculator.CalculateSpatialElementGeometry(space);
            Solid spaceSolid = results.GetGeometry();

            foreach (Face spaceSolidFace in spaceSolid.Faces)
            {
                foreach (SpatialElementBoundarySubface subface in results.GetBoundaryFaceInfo(spaceSolidFace))
                {
                    Face boundingElementface = subface.GetBoundingElementFace();
                    ElementId id = boundingElementface.MaterialElementId;
                    Material material = doc.GetElement(id) as Material;
                    double media = (Convert.ToDouble(material.Color.Red) + Convert.ToDouble(material.Color.Green) + Convert.ToDouble(material.Color.Blue)) / 3;
                    double index = media / 255;

                    coefficientReflectance.Add(index);
                }
            }

            for (int i = 0; i < coefficientReflectance.Count; i++)
            {
                if (coefficientReflectance[i] <= 0.25)
                    coefficientReflectance[i] = 0.1;
                else
                {
                    if (coefficientReflectance[i] > 0.25 && coefficientReflectance[i] <= 0.50)
                        coefficientReflectance[i] = 0.3;
                    else
                    {
                        if (coefficientReflectance[i] > 0.50 && coefficientReflectance[i] <= 0.75)
                            coefficientReflectance[i] = 0.5;
                        else
                            coefficientReflectance[i] = 0.7;
                    }
                }
            }
            using (Transaction transaction = new Transaction(doc))
            {
                if (transaction.Start("Create model curves") == TransactionStatus.Started)
                {
                    space.CeilingReflectance = coefficientReflectance[0];
                    space.FloorReflectance = coefficientReflectance[1];
                    space.WallReflectance = (coefficientReflectance[2] + coefficientReflectance[3] + coefficientReflectance[4] + coefficientReflectance[5]) / 4;
                    transaction.Commit();
                }
            }

        }

        private double GetK()
        {
            area = UnitUtils.ConvertFromInternalUnits(space.Area, DisplayUnitType.DUT_SQUARE_METERS);
            double perimeter = UnitUtils.ConvertFromInternalUnits(space.Perimeter, DisplayUnitType.DUT_METERS);
            double heightWorkPlane = UnitUtils.ConvertFromInternalUnits(space.LightingCalculationWorkplane, DisplayUnitType.DUT_METERS);
            double heightLum = 0.0;
            double ceilingHeight = UnitUtils.ConvertFromInternalUnits(space.LimitOffset - space.BaseOffset, DisplayUnitType.DUT_METERS);

            double heightMontage = ceilingHeight - heightWorkPlane - heightLum;
            return area / (2 * (heightMontage + heightLum) + perimeter / 2);
        }

        private void InitializeComponent()
        {
            this.labelFU = new System.Windows.Forms.Label();
            this.textBoxFU = new System.Windows.Forms.TextBox();
            this.labelFPL = new System.Windows.Forms.Label();
            this.comboBoxFPL = new System.Windows.Forms.ComboBox();
            this.comboBoxAE = new System.Windows.Forms.ComboBox();
            this.labelAE = new System.Windows.Forms.Label();
            this.buttonCalcular = new System.Windows.Forms.Button();
            this.groupBoxCalculo = new System.Windows.Forms.GroupBox();
            this.groupBoxReflexao = new System.Windows.Forms.GroupBox();
            this.labelTeto = new System.Windows.Forms.Label();
            this.labelParede = new System.Windows.Forms.Label();
            this.labelPiso = new System.Windows.Forms.Label();
            this.textBoxTeto = new System.Windows.Forms.TextBox();
            this.textBoxParede = new System.Windows.Forms.TextBox();
            this.textBoxPiso = new System.Windows.Forms.TextBox();
            this.groupBoxCalculo.SuspendLayout();
            this.groupBoxReflexao.SuspendLayout();
            this.SuspendLayout();
            // 
            // labelFU
            // 
            this.labelFU.AutoSize = true;
            this.labelFU.Location = new System.Drawing.Point(8, 23);
            this.labelFU.Name = "labelFU";
            this.labelFU.Size = new System.Drawing.Size(162, 13);
            this.labelFU.TabIndex = 0;
            this.labelFU.Text = "Fator de Utilização com o Índice " + k.ToString("F");
            // 
            // textBoxFU
            // 
            this.textBoxFU.Location = new System.Drawing.Point(206, 23);
            this.textBoxFU.Name = "textBoxFU";
            this.textBoxFU.Size = new System.Drawing.Size(44, 20);
            this.textBoxFU.TabIndex = 1;
            this.textBoxFU.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TextBox1_KeyPress);
            // 
            // labelFPL
            // 
            this.labelFPL.AutoSize = true;
            this.labelFPL.Location = new System.Drawing.Point(8, 61);
            this.labelFPL.Name = "labelFPL";
            this.labelFPL.Size = new System.Drawing.Size(77, 13);
            this.labelFPL.TabIndex = 2;
            this.labelFPL.Text = "Fator de Perda";
            // 
            // comboBoxFPL
            // 
            this.comboBoxFPL.FormattingEnabled = true;
            this.comboBoxFPL.Location = new System.Drawing.Point(129, 61);
            this.comboBoxFPL.Name = "comboBoxFPL";
            this.comboBoxFPL.Size = new System.Drawing.Size(121, 21);
            this.comboBoxFPL.TabIndex = 3;
            this.comboBoxFPL.Items.AddRange(new object[] {
            "Limpo - 2500h",
            "Limpo - 5000h",
            "Limpo - 7500h",
            "Normal - 2500h",
            "Normal - 5000h",
            "Normal - 7500h",
            "Sujo - 2500h",
            "Sujo - 5000h",
            "Sujo - 7500h"
            });
            // 
            // comboBoxAE
            // 
            this.comboBoxAE.FormattingEnabled = true;
            this.comboBoxAE.Location = new System.Drawing.Point(129, 102);
            this.comboBoxAE.Name = "comboBoxAE";
            this.comboBoxAE.Size = new System.Drawing.Size(121, 21);
            this.comboBoxAE.TabIndex = 4;
            this.comboBoxAE.Items.AddRange(new object[] {
            "Recintos para trabalhos não contínuos ou de transição (depósitos, dormitórios, sala de espera..)",
            "Trabalho com tarefas visuais limitadas como salas de aula, arquivo, auditório etc",
            "Trabalhos visuais normais como escritórios, lojas, bancos etc",
            "Recinto para trabalhos que se exige a visualização de detalhes como vitrines, indústrias de roupas etc",
            "Sala de estar, Dormitório, Banheiro, Áera de serviço",
            "Copa/Cozinha",
            "Corredor ou escada interna à unidade, Corredor de uso comum (prédios), Escadaria de uso comum (prédios), Garagens/Estacionamentos internos e cobertos",
            "Garagens/Estacionamentos descobertos",
            "Lojas (vitrines)"
            });
            // 
            // labelAE
            // 
            this.labelAE.AutoSize = true;
            this.labelAE.Location = new System.Drawing.Point(8, 105);
            this.labelAE.Name = "labelAE";
            this.labelAE.Size = new System.Drawing.Size(104, 13);
            this.labelAE.TabIndex = 5;
            this.labelAE.Text = "Atividade do espaço";
            // 
            // buttonCalcular
            // 
            this.buttonCalcular.Location = new System.Drawing.Point(189, 218);
            this.buttonCalcular.Name = "buttonCalcular";
            this.buttonCalcular.Size = new System.Drawing.Size(75, 23);
            this.buttonCalcular.TabIndex = 6;
            this.buttonCalcular.Text = "Calcular";
            this.buttonCalcular.UseVisualStyleBackColor = true;
            this.buttonCalcular.Click += new System.EventHandler(this.ButtonCalcular_Click);
            // 
            // groupBoxCalculo
            // 
            this.groupBoxCalculo.Controls.Add(this.labelAE);
            this.groupBoxCalculo.Controls.Add(this.comboBoxAE);
            this.groupBoxCalculo.Controls.Add(this.comboBoxFPL);
            this.groupBoxCalculo.Controls.Add(this.labelFPL);
            this.groupBoxCalculo.Controls.Add(this.textBoxFU);
            this.groupBoxCalculo.Controls.Add(this.labelFU);
            this.groupBoxCalculo.Location = new System.Drawing.Point(4, 78);
            this.groupBoxCalculo.Name = "groupBoxCalculo";
            this.groupBoxCalculo.Size = new System.Drawing.Size(268, 138);
            this.groupBoxCalculo.TabIndex = 7;
            this.groupBoxCalculo.TabStop = false;
            this.groupBoxCalculo.Text = "Cálculo Luminotécnico";
            // 
            // groupBoxReflexao
            // 
            this.groupBoxReflexao.Controls.Add(this.textBoxPiso);
            this.groupBoxReflexao.Controls.Add(this.textBoxParede);
            this.groupBoxReflexao.Controls.Add(this.textBoxTeto);
            this.groupBoxReflexao.Controls.Add(this.labelPiso);
            this.groupBoxReflexao.Controls.Add(this.labelParede);
            this.groupBoxReflexao.Controls.Add(this.labelTeto);
            this.groupBoxReflexao.Location = new System.Drawing.Point(5, 10);
            this.groupBoxReflexao.Name = "groupBoxReflexao";
            this.groupBoxReflexao.Size = new System.Drawing.Size(266, 68);
            this.groupBoxReflexao.TabIndex = 8;
            this.groupBoxReflexao.TabStop = false;
            this.groupBoxReflexao.Text = "Reflexão das Cores";
            // 
            // labelTeto
            // 
            this.labelTeto.AutoSize = true;
            this.labelTeto.Location = new System.Drawing.Point(7, 33);
            this.labelTeto.Name = "labelTeto";
            this.labelTeto.Size = new System.Drawing.Size(29, 13);
            this.labelTeto.TabIndex = 0;
            this.labelTeto.Text = "Teto";
            // 
            // labelParede
            // 
            this.labelParede.AutoSize = true;
            this.labelParede.Location = new System.Drawing.Point(87, 33);
            this.labelParede.Name = "labelParede";
            this.labelParede.Size = new System.Drawing.Size(46, 13);
            this.labelParede.TabIndex = 1;
            this.labelParede.Text = "Paredes";
            // 
            // labelPiso
            // 
            this.labelPiso.AutoSize = true;
            this.labelPiso.Location = new System.Drawing.Point(183, 33);
            this.labelPiso.Name = "labelPiso";
            this.labelPiso.Size = new System.Drawing.Size(27, 13);
            this.labelPiso.TabIndex = 2;
            this.labelPiso.Text = "Piso";
            // 
            // textBoxTeto
            // 
            this.textBoxTeto.Location = new System.Drawing.Point(39, 30);
            this.textBoxTeto.Name = "textBoxTeto";
            this.textBoxTeto.Size = new System.Drawing.Size(36, 20);
            this.textBoxTeto.TabIndex = 3;
            this.textBoxTeto.Text = space.CeilingReflectance.ToString();
            this.textBoxTeto.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TextBox1_KeyPress);
            // 
            // textBoxParede
            // 
            this.textBoxParede.Location = new System.Drawing.Point(136, 30);
            this.textBoxParede.Name = "textBoxPared";
            this.textBoxParede.Size = new System.Drawing.Size(39, 20);
            this.textBoxParede.TabIndex = 4;
            this.textBoxParede.Text = space.WallReflectance.ToString();
            this.textBoxParede.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TextBox1_KeyPress);
            // 
            // textBoxPiso
            // 
            this.textBoxPiso.Location = new System.Drawing.Point(212, 31);
            this.textBoxPiso.Name = "textBoxPiso";
            this.textBoxPiso.Size = new System.Drawing.Size(37, 20);
            this.textBoxPiso.TabIndex = 5;
            this.textBoxPiso.Text = space.FloorReflectance.ToString();
            this.textBoxPiso.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TextBox1_KeyPress);
            // 
            // Luminotecnico
            // 
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.Controls.Add(this.groupBoxReflexao);
            this.Controls.Add(this.groupBoxCalculo);
            this.Controls.Add(this.buttonCalcular);
            this.Name = "Luminotecnico";
            this.groupBoxCalculo.ResumeLayout(false);
            this.groupBoxCalculo.PerformLayout();
            this.groupBoxReflexao.ResumeLayout(false);
            this.groupBoxReflexao.PerformLayout();
            this.ResumeLayout(false);

        }

        private void ButtonCalcular_Click(object sender, System.EventArgs e)
        {
            if (this.comboBoxFPL.SelectedIndex != -1)
            {
                if (this.textBoxFU.Text != "")
                {
                    if (this.textBoxParede.Text != "")
                    {
                        if (this.textBoxPiso.Text != "")
                        {
                            if (this.textBoxTeto.Text != "")
                            {
                                UpdateReflection();
                                Close();
                                aei= GetAverageEstimatedIllumination();
                                ChecksABNT(aei);
                                //TaskDialog.Show("Revit", "");
                            }
                            else
                            {
                                TaskDialog.Show("Atenção", "Selecione a Reflexão do Teto");
                            }
                        }
                        else
                        {
                            TaskDialog.Show("Atenção", "Selecione a Reflexão do Piso");
                        }
                    }
                    else
                    {
                        TaskDialog.Show("Atenção", "Selecione a Reflexão da parede");
                    }
                }
                else
                {
                    TaskDialog.Show("Atenção", "Selecione o fator de utilização");
                }
            }
            else
            {
                TaskDialog.Show("Atenção", "Selecione o fator de perda");
            }
        }

        private void ChecksABNT(double aei)
        {
            switch (this.comboBoxAE.SelectedIndex)
            {
                case 0:
                    if (aei >= 100 && aei <= 200)
                        TaskDialog.Show("Revit", "O espaço selecionado está de acordo com a norma 5413!\n" +
                            "A ilimunância atual do abiente é aproximadamente " + Math.Round(aei) + " lux");
                    else
                    {
                        double num = NumLightMin(150);
                        TaskDialog.Show("Revit", "O espaço selecionado não está de acordo com a norma 5413!\n" +
                            "A norma recomenda que para o tipo de atividade " + comboBoxAE.Text + " a luminância" +
                            " minima seja 100 lux, a média 150 lux e a máxima 200 lux. \n" +
                            "\nA ilimunância atual do abiente é aproximadamente " + Math.Round(aei) + " lux\n" +
                            "Pelo método dos lúmes é sugerido o numero total de " + Math.Round(num) + " luminárias");
                    }
                    break;
                case 1:
                    if (aei >= 200 && aei <= 500)
                        TaskDialog.Show("Revit", "O espaço selecionado está de acordo com a norma 5413!\n" +
                            "A ilimunância atual do abiente é aproximadamente " + Math.Round(aei) + " lux");
                    else
                    {
                        double num = NumLightMin(300);
                        TaskDialog.Show("Revit", "O espaço selecionado não está de acordo com a norma 5413!\n" +
                            "A norma recomenda que para o tipo de atividade " + comboBoxAE.Text + " a luminância" +
                            " minima seja 200 lux, a média 300 lux e a máxima 500 lux. \n" +
                            "\nA ilimunância atual do abiente é aproximadamente " + Math.Round(aei) + " lux\n" +
                            "Pelo método dos lúmes é sugerido o numero total de " + Math.Round(num) + " luminárias");
                    }
                    break;
                case 2:
                    if (aei >= 300 && aei <= 750)
                        TaskDialog.Show("Revit", "O espaço selecionado está de acordo com a norma 5413!\n" +
                            "A ilimunância atual do abiente é aproximadamente " + Math.Round(aei) + " lux");
                    else
                    {
                        double num = NumLightMin(500);
                        TaskDialog.Show("Revit", "O espaço selecionado não está de acordo com a norma 5413!\n" +
                            "A norma recomenda que para o tipo de atividade " + comboBoxAE.Text + " a luminância" +
                            " minima seja 300 lux, a média 500 lux e a máxima 750 lux. \n" +
                            "\nA ilimunância atual do abiente é aproximadamente " + Math.Round(aei) + " lux\n" +
                            "Pelo método dos lúmes é sugerido o numero total de " + Math.Round(num) + " luminárias");
                    }
                    break;
                case 3:
                    if (aei >= 750 && aei <= 1500)
                        TaskDialog.Show("Revit", "O espaço selecionado está de acordo com a norma 5413!\n" +
                            "A ilimunância atual do abiente é aproximadamente " + Math.Round(aei) + " lux");
                    else
                    {
                        double num = NumLightMin(1000);
                        TaskDialog.Show("Revit", "O espaço selecionado não está de acordo com a norma 5413!\n" +
                            "A norma recomenda que para o tipo de atividade " + comboBoxAE.Text + " a luminância" +
                            " minima seja 750 lux, a média 1000 lux e a máxima 1500 lux. \n" +
                            "\nA ilimunância atual do abiente é aproximadamente " + Math.Round(aei) + " lux\n" +
                            "Pelo método dos lúmes é sugerido o numero total de " + Math.Round(num) + " luminárias");
                    }
                    break;
                case 4:
                    if (aei >= 100 && aei <= 200)
                        TaskDialog.Show("Revit", "O espaço selecionado está de acordo com a norma 5413!\n" +
                            "A ilimunância atual do abiente é aproximadamente " + Math.Round(aei) + " lux");
                    else
                    {
                        double num = NumLightMin(150);
                        TaskDialog.Show("Revit", "O espaço selecionado não está de acordo com a norma 5413!\n" +
                            "A norma recomenda que para o tipo de atividade " + comboBoxAE.Text + " a luminância" +
                            " minima seja 100 lux, a média 150 lux e a máxima 200 lux. \n" +
                            "\nA ilimunância atual do abiente é aproximadamente " + Math.Round(aei) + " lux\n" +
                            "Pelo método dos lúmes é sugerido o numero total de " + Math.Round(num) + " luminárias");
                    }
                    break;
                case 5:
                    if (aei >= 200 && aei <= 500)
                        TaskDialog.Show("Revit", "O espaço selecionado está de acordo com a norma 5413!\n" +
                            "A ilimunância atual do abiente é aproximadamente " + Math.Round(aei) + " lux");
                    else
                    {
                        double num = NumLightMin(300);
                        TaskDialog.Show("Revit", "O espaço selecionado não está de acordo com a norma 5413!\n" +
                            "A norma recomenda que para o tipo de atividade " + comboBoxAE.Text + " a luminância" +
                            " minima seja 200 lux, a média 300 lux e a máxima 500 lux. \n" +
                            "\nA ilimunância atual do abiente é aproximadamente " + Math.Round(aei) + " lux\n" +
                            "Pelo método dos lúmes é sugerido o numero total de " + Math.Round(num) + " luminárias");
                    }
                    break;
                case 6:
                    if (aei >= 75 && aei <= 150)
                        TaskDialog.Show("Revit", "O espaço selecionado está de acordo com a norma 5413!\n" +
                            "A ilimunância atual do abiente é aproximadamente " + Math.Round(aei) + " lux");
                    else
                    {
                        double num = NumLightMin(100);
                        TaskDialog.Show("Revit", "O espaço selecionado não está de acordo com a norma 5413!\n" +
                            "A norma recomenda que para o tipo de atividade " + comboBoxAE.Text + " a luminância" +
                            " minima seja 75 lux, a média 100 lux e a máxima 150 lux. \n" +
                            "\nA ilimunância atual do abiente é aproximadamente " + Math.Round(aei) + " lux\n" +
                            "Pelo método dos lúmes é sugerido o numero total de " + Math.Round(num) + " luminárias");
                    }
                    break;
                case 7:
                    if (aei >= 20 && aei <= 100)
                        TaskDialog.Show("Revit", "O espaço selecionado está de acordo com a norma 5413!\n" +
                            "A ilimunância atual do abiente é aproximadamente " + Math.Round(aei) + " lux");
                    else
                    {
                        double num = NumLightMin(50);
                        TaskDialog.Show("Revit", "O espaço selecionado não está de acordo com a norma 5413!\n" +
                            "A norma recomenda que para o tipo de atividade " + comboBoxAE.Text + " a luminância" +
                            " minima seja 20 lux, a média 50 lux e a máxima 100 lux. \n" +
                            "\nA ilimunância atual do abiente é aproximadamente " + Math.Round(aei) + " lux\n" +
                            "Pelo método dos lúmes é sugerido o numero total de " + Math.Round(num) + " luminárias");
                    }
                    break;
                case 8:
                    if (aei >= 500 && aei <= 1000)
                        TaskDialog.Show("Revit", "O espaço selecionado está de acordo com a norma 5413!\n" +
                            "A ilimunância atual do abiente é aproximadamente " + Math.Round(aei) + " lux");
                    else
                    {
                        double num = NumLightMin(700);
                        TaskDialog.Show("Revit", "O espaço selecionado não está de acordo com a norma 5413!\n" +
                            "A norma recomenda que para o tipo de atividade " + comboBoxAE.Text + " a luminância" +
                            " minima seja 500 lux, a média 700 lux e a máxima 1000 lux. \n" +
                            "\nA ilimunância atual do abiente é aproximadamente " + Math.Round(aei) + " lux\n" +
                            "Pelo método dos lúmes é sugerido o numero total de " + Math.Round(num) + " luminárias");
                    }
                    break;
                default:
                    TaskDialog.Show("Revit", "Erro");
                    break;
            }
        }

        private double NumLightMin(int v)
        {
            List<LightType> lights = GetLight(space);
            List<double> fuLm = new List<double>();
            double luminancia = 0;
            foreach (LightType light in lights)
            {
                double lm = light.GetInitialIntensity().InitialIntensityValue;
                fuLm.Add(lm * GetFU());
                //TaskDialog.Show("Lumem da luz", lm +" lm");
                luminancia += lm * GetFU() * GetFPL();
            }
            return (v * area) / luminancia;
        }

        private void UpdateReflection()
        {
            //this.textBoxParede.Text
            using (Transaction transaction = new Transaction(space.Document))
            {
                if (transaction.Start("Create model curves") == TransactionStatus.Started)
                {
                    space.CeilingReflectance = Convert.ToDouble(this.textBoxTeto.Text);
                    space.FloorReflectance = Convert.ToDouble(this.textBoxPiso.Text);
                    space.WallReflectance = Convert.ToDouble(this.textBoxParede.Text);
                    transaction.Commit();
                }
            }
        }

        private double GetAverageEstimatedIllumination()
        {
            List<LightType> lights = GetLight(space);
            List<double> fuLm = new List<double>();
            double luminancia = 0;
            foreach (LightType light in lights)
            {
                double lm = light.GetInitialIntensity().InitialIntensityValue;
                fuLm.Add(lm * GetFU());
                //TaskDialog.Show("Lumem da luz", lm +" lm");
                luminancia += lm * GetFU() * GetFPL();
            }
            return luminancia/area;
        }

        private double GetFU()
        {
            return Convert.ToDouble(this.textBoxFU.Text);
        }

        public List<LightType> GetLight(Space space)
        {
            BoundingBoxXYZ bb = space.get_BoundingBox(null);

            Outline outline = new Outline(bb.Min, bb.Max);

            BoundingBoxIntersectsFilter filter
              = new BoundingBoxIntersectsFilter(outline);

            Document doc = space.Document;

            // Todo: add category filters and other
            // properties to narrow down the results

            FilteredElementCollector collector
              = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .WhereElementIsViewIndependent()
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_LightingFixtures)
                .WherePasses(filter);

            int spaceid = space.Id.IntegerValue;

            List<LightType> lightType = new List<LightType>();

            foreach (FamilyInstance fi in collector)
            {
                lightType.Add(LightType.GetLightTypeFromInstance(doc, fi.Id));
            }

            return lightType;
        }

        private void TextBox1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
        {
            // Verify that the pressed key isn't CTRL or any non-numeric digit
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != ','))
            {
                e.Handled = true;
            }

            // If you want, you can allow decimal (float) numbers
            System.Windows.Forms.TextBox a;
            a = sender as System.Windows.Forms.TextBox;
            if ((e.KeyChar == ',') && (a.Text.IndexOf(',') > -1))
            {
                e.Handled = true;
            }
        }

        public double GetFPL()
        {
            switch(this.comboBoxFPL.SelectedIndex)
            {
                case 0:
                    return 0.95;
                case 1:
                    return 0.91;
                case 2:
                    return 0.88;
                case 3:
                    return 0.91;
                case 4:
                    return 0.85;
                case 5:
                    return 0.80;
                case 6:
                    return 0.80;
                case 7:
                    return 0.66;
                case 8:
                    return 0.57;
                default:
                    return 0;
            }
        }
    }
}
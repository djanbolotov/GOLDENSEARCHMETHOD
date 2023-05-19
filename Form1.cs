using System;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;
using TextBox = System.Windows.Forms.TextBox;
using MessageBox = System.Windows.Forms.MessageBox;
using aziretParser;
using System.Diagnostics;
using System.Drawing;

namespace POCKETSEARCHMETHOD
{
    public partial class Form1 : Form
    {
        private const string nameOfExcel = @"\Zhanbolot_uulu_Askabek_LookingForOnePoint.xlsm";
        string inputFuncFX = "";
        decimal a = 0;
        decimal b = 0;
        decimal x1 = 0;
        decimal x2 = 0;
        decimal f1;
        decimal f2;
        decimal e_tol = 0;
        int k_max = 0;
        decimal t_max = 0;
        decimal parameterR = 0;
        decimal fplusTol;
        decimal fminusTol;
        Application xls;
        Workbook book = null;
        Worksheet sheet = null;
        public Form1()
        {
            InitializeComponent();
            xls = new Application();
        }

        public int getSign(decimal number)
        {
            if (number < 0)
            {
                return -1;
            }
            else
            {
                return 1;
            }
        }

        public void OpenExcel()
        {
            if (!checkFunction(1)) return;
            string function;
            decimal startPoint, endPoint;

            try
            {
                if (book == null)
                {
                    book = xls.Workbooks.Open(System.IO.Directory.GetCurrentDirectory() + nameOfExcel);
                }
                if (sheet == null)
                {
                    sheet = book.Sheets["Russian"];
                    sheet.Activate();
                }

                xls.Visible = true;
                function = Function.Text;
                if (StartPoint.Text != "" && StartPoint.Text != "-" && StartPoint.Text != "+" && StartPoint.Text != ".")
                {
                    startPoint = Decimal.Parse(StartPoint.Text);
                }
                else
                {
                    startPoint = 1;
                }

                if (EndPoint.Text != "" && EndPoint.Text != "-" && EndPoint.Text != "+" && EndPoint.Text != ".")
                {
                    endPoint = Decimal.Parse(EndPoint.Text);
                }
                else
                {
                    endPoint = startPoint + 2;
                }

                sheet.Cells[4, 9] = startPoint;
                sheet.Cells[4, 10] = endPoint;
                sheet.Cells[2, 1] = "f(x)=" + Function.Text;

                StringBuilder builder = new StringBuilder(function);
                builder.Replace("exp", ":");
                builder.Replace("x", "D4");
                builder.Replace(":", "exp");
                function = builder.ToString();
                sheet.Range["E4:E10003"].Value = "=" + function;
            }
            catch
            {
                book = xls.Workbooks.Open(System.IO.Directory.GetCurrentDirectory() + nameOfExcel);
                sheet = book.Sheets["Russian"];
                sheet.Activate();
                xls.Visible = true;
                function = Function.Text;
                if (StartPoint.Text != "" && StartPoint.Text != "-" && StartPoint.Text != "+" && StartPoint.Text != ".")
                {
                    startPoint = Decimal.Parse(StartPoint.Text);
                }
                else
                {
                    startPoint = 1;
                }

                sheet.Cells[4, 9] = startPoint;
                sheet.Cells[2, 1] = "f(x)=" + Function.Text;

                StringBuilder builder = new StringBuilder(function);
                builder.Replace("exp", ":");
                builder.Replace("x", "D4");
                builder.Replace(":", "exp");
                function = builder.ToString();
                sheet.Range["E4:E10003"].Value = "=" + function;
            }
        }

        private bool parseTry(TextBox t, String type)
        {
            try
            {
                if (type == "Decimal")
                    Decimal.Parse(t.Text, System.Globalization.NumberStyles.Float);
                else if (type == "Integer")
                    int.Parse(t.Text);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private void Clean(Control control)
        {
            foreach (var  element in control.Controls)
            {
                switch (element.GetType().Name)
                {
                    case "TextBox":
                        ((TextBox)element).Text = String.Empty;
                        break;
                    case "RadioButton":
                        ((RadioButton)element).Checked = false;
                        break;
                    case "RichTextBox":
                        ((RichTextBox)element).Text = String.Empty;
                        break;
                    case "GroupBox":
                        Clean((Control)element);
                        break;
                    default:
                        break;
                }
            }
        }


        private bool IsOKForDecimalTextBox(char theCharacter, TextBox theTextBox, bool positive)
        {
            if (!char.IsControl(theCharacter) && !char.IsDigit(theCharacter) && (theCharacter != ',') && (theCharacter != '.')
                && (theCharacter != '-') && (theCharacter != '+') && (theCharacter != 'E') && (theCharacter != 'e'))
            {
                return false;
            }
            if(positive && theCharacter == '-' && (theTextBox.Text.IndexOf('E') == -1 && theTextBox.Text.IndexOf('e') == -1))
            {
                return false;
            }
            if (theCharacter == ',' && (theTextBox.Text.IndexOf(',') > -1 || theTextBox.Text.IndexOf('.') > -1))
            {
                return false;
            }
            if (theCharacter == '.' && (theTextBox.Text.IndexOf('.') > -1 || theTextBox.Text.IndexOf(',') > -1))
            {
                return false;
            }
            if (theCharacter == 'e' && (theTextBox.Text.IndexOf('e') > -1 || theTextBox.Text.IndexOf('E') > -1))
            {
                return false;
            }
            if (theCharacter == 'E' && (theTextBox.Text.IndexOf('E') > -1 || theTextBox.Text.IndexOf('e') > -1))
            {
                return false;
            }
            if (theCharacter == '-' && (theTextBox.Text.IndexOf('-') > -1 || theTextBox.Text.IndexOf('+') > -1))
            {
                return false;
            }
            if (theCharacter == '+' && (theTextBox.Text.IndexOf('+') > -1 || theTextBox.Text.IndexOf('-') > -1))
            {
                return false;
            }
            if (((theCharacter == '-') || (theCharacter == '+')) && (theTextBox.SelectionStart != 0 && (theTextBox.Text.IndexOf('E') == -1 && theTextBox.Text.IndexOf('e') == -1)))
            {
                return false;
            }
            if ((char.IsDigit(theCharacter) || (theCharacter == ',') || (theCharacter == '.')) && ((theTextBox.Text.IndexOf('-') > -1) 
                || (theTextBox.Text.IndexOf('+') > -1)) && theTextBox.SelectionStart == 0)
            {
                return false;
            }
            return true;
        }

        public decimal Fx(decimal x)
        {
            decimal result;
            result = aziretParser.Computer.Compute(inputFuncFX, x);
            return result;
        }
        private void button4_Click(object sender, EventArgs e)
        {
            OpenExcel();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Clean(this);
            progressBar1.Visible = false;
        }

        private void InitialApproximation_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !IsOKForDecimalTextBox(e.KeyChar, StartPoint, false);
            if (e.KeyChar == '.')
            {
                e.KeyChar = ',';
            }
        }

        private void Tolerance_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !IsOKForDecimalTextBox(e.KeyChar, Tolerance, true);
            if (e.KeyChar == '.')
            {
                e.KeyChar = ',';
            }
        }

        private void ParametrR_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !IsOKForDecimalTextBox(e.KeyChar, EndPoint, false);
            if (e.KeyChar == '.')
            {
                e.KeyChar = ',';
            }
        }

        private String checkParse()
        {
            String errorMessage = "";
            if (!parseTry(StartPoint, "Decimal"))
            {
                errorMessage += "Invalid value of the field Start point (a)! Change the input and perform the calculation!\n\n";
            }
            else
            {
                a = Decimal.Parse(StartPoint.Text, System.Globalization.NumberStyles.Float);
            }

            if (parseTry(Tolerance, "Decimal"))
            {
                e_tol = Decimal.Parse(Tolerance.Text, System.Globalization.NumberStyles.Float);
            }
            else
            {
                errorMessage += "Invalid value of the Tolerance(e) field (entered tolerance)! Change the input and perform the calculation!\n\n";
            }

            if (!parseTry(LimitOfIterations, "Integer"))
            {
                errorMessage += "Invalid value of the field limit of iterations! Change the input and perform the calculation!\n\n";
            }
            else
            {
                k_max = Int32.Parse(LimitOfIterations.Text);
            }

            if (!parseTry(EndPoint, "Decimal"))
            {
                errorMessage += "Invalid value of the field End point (b)! Change the input and perform the calculation!\n\n";
            }
            else
            {
                b = Decimal.Parse(EndPoint.Text, System.Globalization.NumberStyles.Float);   
            }

            if (!parseTry(LimitOfTime, "Decimal"))
            {
                errorMessage += "Invalid value of the field limit of time! Change the input and perform the calculation!\n\n";
            }
            else
            {
                t_max = Decimal.Parse(LimitOfTime.Text, System.Globalization.NumberStyles.Float);
            }

            return errorMessage;
        }

        public bool fullCheck()
        {
            bool check = false;
            if (Function.Text == "" || StartPoint.Text == "" ||
                Tolerance.Text == "" || LimitOfIterations.Text == "" ||
                LimitOfTime.Text == "" || EndPoint.Text == "")
            {
                MessageBox.Show("All fields must be filled in! Enter the missing information and make the calculation!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                if (checkParse() != "")
                {
                    MessageBox.Show(checkParse(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    if (isRigth() && checkFunction(a))
                    {
                        check = true;
                    }
                }
            }
            return check;
        }

        public string getComparisonSign(decimal a, decimal b)
        {
            if (a > b)
            {
                return ">";
            }
            else if (a < b)
            {
                return "<";
            }
            else
            {
                return "=";
            }
        }

        private bool isRigth()
        {
            bool isRight = true;
            if(a >= b)
            {
                MessageBox.Show("The value of the Ending point (b) field must be greater than Starting point (a)! Change the input and perform the calculation!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                isRight = false;
            }
            if (e_tol <= 0)
            {
                MessageBox.Show("The value of the tolerance field must be greater than 0! Change the input and perform the calculation!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                isRight = false;
            }
            if (k_max <= 0)
            {
                MessageBox.Show("The value of the limit of iterations field must be greater than 0! Change the input and perform the calculation!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                isRight = false;
            }
            if (t_max <= 0)
            {
                MessageBox.Show("The value of the limit of time field must be greater than 0! Change the input and perform the calculation!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                isRight = false;
            }
            if (!(Maximum.Checked || Minimum.Checked))
            {
                MessageBox.Show("Please select search option maximum or minimum.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                isRight = false;
            }
            if (isRight)
            {
                return true;
            }
            return false;
        }

        private bool checkFunction(decimal x0)
        {
            inputFuncFX = Function.Text;

            if (inputFuncFX == "" || inputFuncFX.IndexOf('x') == -1)
            {
                MessageBox.Show("The function is entered incorrectly! Change the input and perform the calculation!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Clean(this);
                return false;
            }
            try
            {
                if (inputFuncFX.Contains("log") && x0 <= 0 || inputFuncFX.Contains("ln") && x0 <= 0)
                {
                    MessageBox.Show("If you entered function with 'log' or 'ln' value of a must greater than zero!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
                else
                {
                    decimal F1 = Fx(x0);
                    return true;
                }
            }
            catch
            {
                MessageBox.Show("The function or initial approximation is entered incorrectly! Change the input and perform the calculation!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Clean(this);
                return false;
            }
        }

        public bool MaxOrMin(decimal f0, decimal f1)
        {
            if (Maximum.Checked)
            {
                return f0 >= f1;
            }
            return f0 <= f1;
        }

        public void FillResult(string solution, string iterations, string resultTolerance, string fminustol, string fplustol, string fxvalue, string fminusplus, string fminusminus, string searchStep)
        {
            ResultX.Text = solution;
            countofiterations.Text = iterations;
            fxplustolerance.Text = fplustol;
            fxminustolerance.Text = fminustol;
            fxminusplustolerance.Text = fminusplus;
            fxminusminustolerance.Text = fminusminus;
            fx.Text = fxvalue;
        }

        public string getError(TextBox tol, decimal error)
        {
            Console.WriteLine(tol);
            if (tol.Text.Contains("E"))
            {
                return error.ToString("0E0");
            }
            else if (tol.Text.Contains("e"))
            {
                return error.ToString("0e0");
            }
            else
            {
                return error.ToString();
            }
        }

        private void LimitOfTime_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !IsOKForDecimalTextBox(e.KeyChar, LimitOfTime, true);
            if(e.KeyChar == '.')
            {
                e.KeyChar = ',';
            }
        }

        private void LimitOfIterations_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((int)e.KeyChar == (int)48 && LimitOfIterations.Text == "")
            {
                e.Handled = true;
                return;
            }
            e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("1) Choose a function or write your's on field 'Function'\n" +
                     "2) Click on the button 'Show function graph'\n" +
                     "3) In the opened file select the values for a and b,\n" +
                     "then save the document and return to the program\n" +
                     "4) If you need 'a' and b values to insert,\n" +
                     "click the button 'Set 'a' and 'b'' or write your's\n" +
                     "5) Enter tolerance\n" +
                     "6) Enter limit of time in sec\n" +
                     "7) Enter limit of iterations \n" +
                     "8) Select search parameter\n" +
                     "Then click the button 'Run Method'.", "Information",
                     MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (book == null)
                {
                    book = xls.Workbooks.Open(System.IO.Directory.GetCurrentDirectory() + nameOfExcel);
                }
                if (sheet == null)
                {
                    sheet = book.Sheets["Russian"];
                    sheet.Activate();
                }
                book.Save();
                StartPoint.Text = sheet.Cells[4, 9].Value.ToString();
                EndPoint.Text = sheet.Cells[4, 10].Value.ToString();
            }
            catch
            {
                book = xls.Workbooks.Open(System.IO.Directory.GetCurrentDirectory() + nameOfExcel);
                sheet = book.Sheets["Russian"];
                sheet.Activate();
                book.Save();
                StartPoint.Text = sheet.Cells[4, 9].Value.ToString();
                EndPoint.Text = sheet.Cells[4, 10].Value.ToString();
            }
            xls.Visible = false;
            book = null;
            sheet = null;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            a = 0;
            b = 0;
            x1 = 0;
            x2 = 0;
            f1 = 0;
            f2 = 0;
            e_tol = 0;
            k_max = 0;
            t_max = 0;
            parameterR = 0;
            inputFuncFX = "";
            fminusTol = 0;
            fplusTol = 0;

            string extremium;

            try
            {
                if (fullCheck())
                {
                    xls.Visible = false;
                    book = null;
                    sheet = null;
                    if (Maximum.Checked)
                    {
                        extremium = "maximizer";
                    }
                    else
                    {
                        extremium = "minimizer";
                    }
                    progressBar1.Value = 0;
                    Clean(groupBox2);
                    validation.Text = String.Empty;
                    Stopwatch stopwatch = new Stopwatch();
                    stopwatch.Start();

                    //
                    parameterR = Decimal.Parse("0,61803398874989484820", System.Globalization.NumberStyles.Float);
                    x1 = a + (1 - parameterR) * (b - a);
                    f1 = Fx(x1);
                    x2 = a + parameterR * (b - a);
                    f2 = Fx(x2);
                    int k = 0;

                    do
                    {
                        k = k + 1;

                        progressBar1.Visible = true;
                        progressBar1.Maximum = (int)(k + 0.00000001);
                        progressBar1.Value = k;

                        if (k > k_max)
                        {
                            stopwatch.Stop();
                            f2 = Fx(x2);
                            fminusTol = Fx(x2 - e_tol);
                            fplusTol = Fx(x2 + e_tol);
                            DialogResult result = MessageBox.Show("Iteration limit reached. Do you want to add iterations?",
                                "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                            if (result == DialogResult.Yes)
                            {
                                k_max += k_max;
                                LimitOfIterations.Text = k_max.ToString();
                            }
                            else
                            {
                                k--;
                                validation.Text += "Result X* not found because of limit of iterations = " + k_max + "." +
                                    "\nSince the following condition is false, namely:" +
                                    "\nSign(f(X*)-f(X*+Tolerance)) = " + getSign(f2 - fplusTol) + " and Sign(f(X*)-f(X*-Tolerance)) = " + getSign(f2 - fminusTol) + "!" +
                                    "\nResult X* is not " + extremium + " of the function.";
                                validation.ForeColor = Color.Red;

                                FillResult(x2.ToString("F28"), k.ToString(), getError(Tolerance, Math.Abs(x2 - x1)), fminusTol.ToString("F28"), fplusTol.ToString("F28"), f2.ToString("F28"), (f2 - fplusTol).ToString("F28"), (f2 - fminusTol).ToString("F28"), getError(absError, Math.Abs(x2 - x1)));
                                absError.Text = getError(Tolerance, Math.Abs(x2 - x1));

                                DialogResult answer = MessageBox.Show("Result X* not found because of maximum limit of iterations = " + k_max + "." +
                                "\nSince the following condition is false, namely:" +
                                "\nSign(f(X*)-f(X*+Tolerance)) = " + getSign(f2 - fplusTol) + " and Sign(f(X*)-f(X*-Tolerance)) = " + getSign(f2 - fminusTol) + "!" +
                                "\nResult X* is not " + extremium + " of the function." +
                                "\n\nYou probably entered the values of 'a' and 'b' range incorrectly on Ecxel!" +
                                "\nSince the program is looking for an extremum only in the range 'a' and 'b'." +
                                "\nYou need to open the graph and select the correct points [a;b]!" +
                                "\n\nDo you want to open file?", "Error", MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                                if (answer == DialogResult.Yes)
                                {
                                    OpenExcel();
                                }
                                break;
                            }
                            stopwatch.Start();
                        }

                        if (stopwatch.ElapsedMilliseconds >= t_max * 1000)
                        {
                            stopwatch.Stop();
                            f2 = Fx(x2);
                            fminusTol = Fx(x2 - e_tol);
                            fplusTol = Fx(x2 + e_tol);
                            DialogResult result = MessageBox.Show("Time limit reached. Do you want to add time?",
                                "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                            if (result == DialogResult.Yes)
                            {
                                t_max += t_max;
                                LimitOfTime.Text = t_max.ToString();
                            }
                            else
                            {
                                validation.Text += "Result X* not found because of limit of time = " + t_max + " sec." +
                                    "\nSince the following condition is false, namely:" +
                                    "\nSign(f(X*)-f(X*+Tolerance)) = " + getSign(f2 - fplusTol) + " and Sign(f(X*)-f(X*-Tolerance)) = " + getSign(f2 - fminusTol) + "!" +
                                    "\nResult X* is not " + extremium + " of the function.";
                                validation.ForeColor = Color.Red;

                                FillResult(x2.ToString("F28"), k.ToString(), getError(Tolerance, Math.Abs(x2 - x1)), fminusTol.ToString("F28"), fplusTol.ToString("F28"), f2.ToString("F28"), (f2 - fplusTol).ToString("F28"), (f2 - fminusTol).ToString("F28"), getError(absError, Math.Abs(x2 - x1)));
                                absError.Text = getError(Tolerance, Math.Abs(x2 - x1));

                                DialogResult answer = MessageBox.Show("Result X* not found because of maximum time limit = " + t_max + " sec." +
                                "\nSince the following condition is false, namely:" +
                                "\nSign(f(X*)-f(X*+Tolerance)) = " + getSign(f2 - fplusTol) + " and Sign(f(X*)-f(X*-Tolerance)) = " + getSign(f2 - fminusTol) + "!" +
                                "\nResult X* is not " + extremium + " of the function." +
                                "\n\nYou probably entered the values of 'a' and 'b' range incorrectly on Ecxel!" +
                                "\nSince the program is looking for an extremum only in the range 'a' and 'b'." +
                                "\nYou need to open the graph and select the correct points [a;b]!" +
                                "\n\nDo you want to open file?", "Error", MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                                if (answer == DialogResult.Yes)
                                {
                                    OpenExcel();
                                }
                                break;
                            }
                            stopwatch.Start();
                        }
                                                
                        if (MaxOrMin(f1, f2))
                        {
                            b = x2;
                            x2 = x1;
                            f2 = f1;
                            x1 = a + (1 - parameterR) * (b - a);
                            f1 = Fx(x1);
                        }
                        else
                        {
                            a = x1;
                            x1 = x2;
                            f1 = f2;
                            x2 = a + parameterR * (b - a);
                            f2 = Fx(x2);
                        }

                        fminusTol = Fx(x1 - e_tol);
                        fplusTol = Fx(x1 + e_tol);

                        if (Math.Abs(b - a) <= e_tol)
                        {
                            if (extremium == "minimizer")
                            {
                                if (f2 < fminusTol && f2 < fplusTol)
                                {
                                    FillResult(x2.ToString("F28"), k.ToString(), getError(Tolerance, Math.Abs(x2 - x1)), fminusTol.ToString("F28"), fplusTol.ToString("F28"), f2.ToString("F28"), (f2 - fplusTol).ToString("F28"), (f2 - fminusTol).ToString("F28"), getError(absError, Math.Abs(x2 - x1)));

                                    validation.Text += "Since the following condition is true, namely:" +
                                            "\nSign(f(X*)-f(X*+Tolerance)) = " + getSign(f2 - fplusTol) + " and Sign(f(X*)-f(X*-Tolerance)) = " + getSign(f2 - fminusTol) + "!" +
                                            "\nResult X* is minimizer of the function. It has been found with the error = " + Math.Abs(b - a) + ". This is less than or equal to given Tolerance!";
                                    validation.ForeColor = Color.Green;
                                    absError.Text = Convert.ToString(Math.Abs(b - a));
                                    break;
                                }else
                                {
                                    FillResult(x2.ToString("F28"), k.ToString(), getError(Tolerance, Math.Abs(x2 - x1)), fminusTol.ToString("F28"), fplusTol.ToString("F28"), f2.ToString("F28"), (f2 - fplusTol).ToString("F28"), (f2 - fminusTol).ToString("F28"), getError(absError, Math.Abs(x2 - x1)));
                                    validation.Text += "Since the following condition is false, namely:" +
                                            "\nSign(f(X*)-f(X*+Tolerance)) = " + getSign(f2 - fplusTol) + " and Sign(f(X*)-f(X*-Tolerance)) = " + getSign(f2 - fminusTol) + "!" +
                                            "\nResult X* is not minimizer of the function.";
                                    validation.ForeColor = Color.Red;
                                    absError.Text = Convert.ToString(Math.Abs(b - a));
                                    break;
                                }
                            }
                            else
                            {
                                if ((f2 >= fminusTol && f2 >= fplusTol))
                                {
                                    FillResult(x2.ToString("F28"), k.ToString(), getError(Tolerance, Math.Abs(x2 - x1)), fminusTol.ToString("F28"), fplusTol.ToString("F28"), f2.ToString("F28"), (f2 - fplusTol).ToString("F28"), (f2 - fminusTol).ToString("F28"), getError(absError, Math.Abs(x2 - x1)));

                                    validation.Text += "Since the following condition is true, namely:" +
                                            "\nSign(f(X*)-f(X*+Tolerance)) = " + getSign(f2 - fplusTol) + " and Sign(f(X*)-f(X*-Tolerance)) = " + getSign(f2 - fminusTol) + "!" +
                                            "\nResult X* is maximizer of the function. It has been found with the error = " +  Math.Abs(b - a) + ". This is less than or equal to given Tolerance!";
                                    validation.ForeColor = Color.Green;
                                    absError.Text = Convert.ToString(Math.Abs(b - a));
                                    break;
                                }
                            }
                        }
                    } while (true);

                    stopwatch.Stop();
                    elapsedtime.Text = stopwatch.ElapsedMilliseconds / 1000.0 + " sec";
                    timer1.Enabled = true;
                    timer1.Start();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Clean(this);
                progressBar1.Value = 0;
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            progressBar1.Value = 0;
            timer1.Enabled = false;
            timer1.Stop();
        }

        private void Function_TextChanged(object sender, EventArgs e)
        {
            Clean(groupBox2);
            validation.Text = String.Empty;
        }

        private void InitialApproximation_TextChanged(object sender, EventArgs e)
        {
            Clean(groupBox2);
            validation.Text = String.Empty;
        }

        private void Tolerance_TextChanged(object sender, EventArgs e)
        {
            Clean(groupBox2);
            validation.Text = String.Empty;
        }

        private void SearchStep_TextChanged(object sender, EventArgs e)
        {
            Clean(groupBox2);
            validation.Text = String.Empty;
        }

        private void ParametrR_TextChanged(object sender, EventArgs e)
        {
            Clean(groupBox2);
            validation.Text = String.Empty;
        }

        private void LimitOfTime_TextChanged(object sender, EventArgs e)
        {
            Clean(groupBox2);
            validation.Text = String.Empty;
        }

        private void LimitOfIterations_TextChanged(object sender, EventArgs e)
        {
            Clean(groupBox2);
            validation.Text = String.Empty;
        }

        private void Maximum_CheckedChanged(object sender, EventArgs e)
        {
            Clean(groupBox2);
            validation.Text = String.Empty;
        }

        private void Minimum_CheckedChanged(object sender, EventArgs e)
        {
            Clean(groupBox2);
            validation.Text = String.Empty;
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            xls.Quit();
        }
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace 公式生成工具
{
    public partial class Form1 : Form
    {
        private List<Parameter> parameters = new List<Parameter>();
        private List<Formula> formulas = new List<Formula>();
        private bool IsUpdating = false;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var d = new OpenFileDialog();
            d.Filter = "文本文件|*.txt";
            if (d.ShowDialog() == DialogResult.OK)
            {
                var fName = d.FileName;
                ReadTemplate(fName);
                RenderMyControls();
            }
        }

        private void ReadTemplate(string fpath)
        {
            parameters.Clear();
            formulas.Clear();
            var t = File.ReadAllText(fpath);
            var lines = t.Split('\n').Select(l => l.Trim().Replace("，", ",")).Where(l => !l.StartsWith("//"));

            var flag = 1;
            string group = string.Empty;
            foreach (var line in lines)
            {
                if (string.IsNullOrWhiteSpace(line))
                {
                    flag++;
                    continue;
                }
                else if (line.StartsWith("$"))
                {
                    group = line.Substring(1);
                    continue;
                }

                if (flag == 1)
                {
                    var pair = line.Split(',');
                    if (!string.IsNullOrWhiteSpace(pair[0]) && !string.IsNullOrWhiteSpace(pair[1]))
                        parameters.Add(new Parameter()
                        {
                            Group = group,
                            Name = pair[0],
                            Description = pair[1]
                        });
                }
                else if (flag == 2)
                {
                    var equalFlag = line.IndexOf("=");
                    if (equalFlag > -1)
                    {
                        string desc = string.Empty;
                        var variable = line.Substring(0, equalFlag).Trim().Replace("（", "(");
                        var i = variable.IndexOf('(');
                        if (i > -1 && i < variable.Length)
                        {
                            desc = variable.Substring(i + 1, variable.Length - i - 2);
                            variable = variable.Substring(0, i);
                        }
                        var formula = line.Substring(equalFlag + 1, line.Length - equalFlag - 1);
                        formulas.Add(new Formula
                        {
                            LeftVar = variable,
                            Description = desc,
                            Expression = formula
                        });
                    }
                }
            }
        }

        private void RenderMyControls()
        {
            panel1.Controls.Clear();
            panel1.Visible = false;
            if (parameters.Count == 0) return;

            var curGroup = new GroupBox();
            curGroup.Text = parameters[0].Group;

            var seq = 1;
            foreach (var p in parameters)
            {
                if (p.Group != curGroup.Text)
                {
                    panel1.Controls.Add(curGroup);
                    curGroup = new GroupBox();
                    curGroup.Text = p.Group;
                }

                var label = new Label();
                label.Text = p.Description;
                label.AutoSize = false;
                label.TextAlign = ContentAlignment.MiddleRight;
                label.Tag = seq++;
                var textbox = new TextBox();
                textbox.Name = p.Name;
                textbox.Tag = seq++;
                curGroup.Controls.Add(label);
                curGroup.Controls.Add(textbox);
                textbox.BringToFront();
            }

            panel1.Controls.Add(curGroup);
            SetControlPositions();
        }

        private void SetControlPositions()
        {
            panel1.Visible = false;
            IsUpdating = true;
            var groupTop = 20;
            foreach (var c in panel1.Controls)
            {
                var group = c as GroupBox;
                if (group != null)
                {
                    group.Width = panel1.Width - 20;
                    group.Top = groupTop;

                    var children = group.Controls.Cast<Control>().OrderBy(ctrl => ctrl.Tag).ToList();
                    var curLeft = 10;
                    var curTop = 30;
                    for (var i = 0; i < children.Count; i += 2)
                    {
                        children[i].Left = curLeft;
                        children[i].Top = curTop;
                        children[i].Width = 100;
                        children[i + 1].Left = curLeft + 100;
                        children[i + 1].Top = curTop;
                        children[i + 1].Width = 150;
                        if (curLeft >= group.Width - 500)
                        {
                            curLeft = 10;
                            curTop += 30;
                        }
                        else curLeft += 250;
                    }
                    group.Height = curTop + 30;
                    groupTop += group.Height + 5;
                }
            }
            panel1.Visible = true;
            IsUpdating = false;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var variables = new Dictionary<string, string>();
            foreach (var c in panel1.Controls.Cast<Control>())
            {
                foreach (var child in c.Controls.Cast<Control>())
                {
                    if (child is TextBox)
                    {
                        variables.Add(child.Name, child.Text.Trim() == string.Empty ? "0" : child.Text.Trim());
                    }
                }
            }

            var sb = new StringBuilder();
            foreach (var f in formulas)
            {
                List<string> usedExps;
                var fval = GetExpression(variables, f.Expression, out usedExps);
                if (!usedExps.All(v => variables[v] == "0") || checkBox1.Checked)
                    sb.Append(f.LeftVar + "," + f.Description + "," + fval + Environment.NewLine);
            }

            var d = new SaveFileDialog();
            d.Filter = "表格文件（*.csv）|*.csv";
            if (d.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    File.WriteAllText(d.FileName, sb.ToString(), Encoding.UTF8);
                }
                catch (Exception)
                {
                    MessageBox.Show("生成Excel文件失败");
                }
            }
        }

        private string GetExpression(Dictionary<string, string> variables, string f, out List<string> usedVariables)
        {
            string expression = f;
            usedVariables = new List<string>();
            foreach (var v in variables.Keys)
            {
                string replacement = variables[v];
                if (!Regex.IsMatch(replacement, "^-?\\w+\\.?\\w*$")) replacement = "(" + replacement + ")";
                var reg = new Regex("\\b" + v + "\\b");
                var newExp = reg.Replace(expression, replacement);
                if (newExp != expression) usedVariables.Add(v);
                expression = newExp;
            }
            return expression;
        }

        private void Form1_SizeChanged(object sender, EventArgs e)
        {
            SetControlPositions();
        }
    }

    public class Parameter
    {
        public string Group;
        public string Name;
        public string Description;
    }

    public class Formula
    {
        public string LeftVar;
        public string Description;
        public string Expression;
    }
}

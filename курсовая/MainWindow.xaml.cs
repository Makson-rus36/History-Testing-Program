using Microsoft.Win32;
using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Media.Imaging;
using Microsoft.Office.Interop.Excel;
using Window = System.Windows.Window;
using Application = System.Windows.Application;
using GroupBox = System.Windows.Controls.GroupBox;
using Button = System.Windows.Controls.Button;
using Style = System.Windows.Style;
using Label = System.Windows.Controls.Label;
using Point = System.Windows.Point;
using Border = System.Windows.Controls.Border;

namespace kursovay
{




    
    public class Test
    {
        MainWindow mainWindow_style = new MainWindow();
        int[] summa_ballov;
        string[] otvet;
        int[] maks_ball_curr_quest;
        string[] Quest;
         string[] A1;
        string[] A2;
        string[] A3;
        string[] A4;
       string[,] tabl_srav_vopr;
        string[,] tabl_srav_otv;
        DataGrid temp = new DataGrid();


        string[] Right_answ;
        int[] koef;
        RadioButton radioButton_ans_1 = new RadioButton();
        RadioButton radioButton_ans_2 = new RadioButton();
        RadioButton radioButton_ans_3 = new RadioButton();
        RadioButton radioButton_ans_4 = new RadioButton();
        StackPanel stackPanel_in_answ = new StackPanel();
        


        public void Read_file_test(string path_file, Grid name_grid) //Проверка целостности файлов теста
        {
           
            string ext = Path.GetExtension(path_file);
            try
            {
                File.ReadAllLines(path_file);
            }
            catch
            {
                MessageBox.Show("Критическая ошибка! Проверьте наличие всех файлов", "Ошибка загрузки теста");
            }
            finally
            {
                if (ext == ".txt")
                {

                    int s = 0;
                    int size_closed = 0;
                    int size_open = 0;

                    string[] all_lines = File.ReadAllLines(path_file);
                    for (int a = 0; a < all_lines.Length; a++)
                    {
                        if (all_lines[a] == "закрытая")
                        {
                            s++;
                            size_closed++;
                        }
                        else if (all_lines[a] == "сравнение")
                        {
                            s++;
                            size_open++;
                        }
                        else
                        {
                           
                        }

                    }



                    StreamReader read_file = new StreamReader(path_file);
                    string[] type = new string[s];
                    int[] Id_closed = new int[size_closed];
                    int[] Id_open = new int[s];
                    string[] Diff = new string[s];
                    string[] Im_source = new string[s];
                    Quest = new string[s];
                     A1 = new string[size_closed];
                     A2 = new string[size_closed];
                     A3 = new string[size_closed];
                     A4 = new string[size_closed];
                    Right_answ = new string[s];
                    tabl_srav_vopr = new string[s, 4];
                    tabl_srav_otv = new string[s, 4];
                    string[,] tabl_srav = new string[s, 4];
                    summa_ballov = new int[s];
                    otvet = new string[s];
                    koef = new int[s];
                    maks_ball_curr_quest = new int[s];
                    for (int i = 0; i < s;)
                    {
                        type[i] = read_file.ReadLine();
                        if (type[i] == "закрытая")
                        {
                            Id_closed[i] = i;
                            Diff[i] = read_file.ReadLine();
                            Im_source[i] = Path.GetDirectoryName(path_file) + "\\" + read_file.ReadLine();
                            Quest[i] = read_file.ReadLine();
                            A1[i] = read_file.ReadLine();
                            A2[i] = read_file.ReadLine();
                            A3[i] = read_file.ReadLine();
                            A4[i] = read_file.ReadLine();
                            if (A1[i].IndexOf('*') == 0)
                            {
                                Right_answ[i] = "1";
                                A1[i] = A1[i].Replace("*", "");

                            }
                            else if (A2[i].IndexOf('*') == 0)
                            {
                                Right_answ[i] = "2";
                                A2[i] = A2[i].Replace("*", "");


                            }
                            else if (A3[i].IndexOf('*') == 0)
                            {
                                Right_answ[i] = "3";
                                A3[i] = A3[i].Replace("*", "");
                            }
                            else if (A4[i].IndexOf('*') == 0)
                            {
                                Right_answ[i] = "4";
                                A4[i] = A4[i].Replace("*", "");
                            }
                            else
                            {

                            }

                        }
                        else if (type[i] == "сравнение")
                        {
                            Id_open[i] = i;
                            Diff[i] = read_file.ReadLine();
                            Im_source[i] = Path.GetDirectoryName(path_file) + "\\" + read_file.ReadLine();
                            Quest[i] = read_file.ReadLine();
                            for (int a = 0; a < 4; a++)
                            {
                                tabl_srav[i, a] = read_file.ReadLine();
                            }
                            for (int a = 0; a < 4; a++)
                            {
                                string[] words = tabl_srav[i, a].Split(':');
                                tabl_srav_vopr[i, a] = words[0];
                                tabl_srav_otv[i, a] = words[1];
                            }


                        }

                        i++;
                    }
                    for (int i = 0; i < s; i++)
                    {
                        if (Diff[i] == "легко")
                        {
                            koef[i] = 1;
                        }
                        else if (Diff[i] == "средне")
                        {
                            koef[i] = 2;
                        }
                        else if (Diff[i] == "сложно")
                        {
                            koef[i] = 3;
                        }
                        else
                        {

                        }
                    }
                    for (int i = 0; i < s; i++)
                    {
                        if (type[i] == "сравнение")
                        {
                            maks_ball_curr_quest[i] = 4 * koef[i];
                        }
                        else if (type[i] == "закрытая")
                        {
                            maks_ball_curr_quest[i] = 1 * koef[i];
                        }
                    }
                    Create_content_test(s, Im_source, Quest, A1, A2, A3, A4, Right_answ, Diff, type, tabl_srav_vopr, tabl_srav_otv, name_grid);
                }
                else
                {
                    MessageBox.Show("Файл данного расширения не поддерживается", "Ошибка загрузки теста");
                }
            }
            
        }

        public void Create_content_type_closed(Grid grid_work_zone, RowDefinition row_image, RowDefinition row_image_1, int current_answ, string[] Image_source, string[] Quest, string[] Answ1, string[] Answ2, string[] Answ3, string[] Answ4, int a)
        {


            grid_work_zone.Children.Clear();
            Image image_main = new Image();
            GridLength gridLength = new GridLength(3, GridUnitType.Star);
            row_image.Height = gridLength;
            image_main.HorizontalAlignment = HorizontalAlignment.Center;
            image_main.Source = new BitmapImage(new Uri(Image_source[current_answ], UriKind.Absolute));
            grid_work_zone.Children.Add(image_main);
            Grid.SetRow(image_main, 0);

            GroupBox groupBox = new GroupBox();
            grid_work_zone.Children.Add(groupBox);
            Grid.SetRow(groupBox, 1);
            row_image_1.MinHeight = 90;
            Grid.SetRow(groupBox, 1);
            groupBox.VerticalAlignment = VerticalAlignment.Top;
            groupBox.HorizontalAlignment = HorizontalAlignment.Center;
            groupBox.Padding = new Thickness(7);
            groupBox.FontSize = 20;

            ScrollViewer scrollViewer_quest = new ScrollViewer();
            scrollViewer_quest.Content = null;
            scrollViewer_quest.Content = stackPanel_in_answ;
            
            groupBox.Content = scrollViewer_quest;
            scrollViewer_quest.Padding = new Thickness(0, 0, 10, 0);
            scrollViewer_quest.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;

            groupBox.Header = "вопрос №" + Convert.ToString(1 + a) + " " + Quest[current_answ];


            radioButton_ans_1.Content = Answ1[current_answ];
            radioButton_ans_1.GroupName = "answers";


            radioButton_ans_2.Content = Answ2[current_answ];
            radioButton_ans_2.GroupName = "answers";


            radioButton_ans_3.Content = Answ3[current_answ];
            radioButton_ans_3.GroupName = "answers";



            radioButton_ans_4.Content = Answ4[current_answ];
            radioButton_ans_4.GroupName = "answers";
        }
        int i = 0;
        
       
        public void Create_content_type_srav(Grid grid_work_zone, RowDefinition row_image,  string[,] tabl_otvetov, string[,] tabl_voprosov, int a, int current_answ, string[] Quest, string[] Image_source)
        {
            int[] proverka = new int[4];
             RadioButton answ_1 = new RadioButton();
            RadioButton answ_2 = new RadioButton();
            RadioButton answ_3 = new RadioButton();
            RadioButton answ_4 = new RadioButton();
            GroupBox groupBox = new GroupBox();
            Grid.SetRow(groupBox, 0);
            string[] tabl_o_per = new string[4];
           
            tabl_o_per[0] = tabl_voprosov[current_answ, 0];
            tabl_o_per[1] = tabl_voprosov[current_answ, 1];
            tabl_o_per[2] = tabl_voprosov[current_answ, 2];
            tabl_o_per[3] = tabl_voprosov[current_answ, 3];
           
            i++;




            Random rnd = new Random();
            for (int i = tabl_o_per.Length - 1; i > 0; i--)
            {
                int j = rnd.Next(0, i + 1);
                string temp = tabl_o_per[i];
                tabl_o_per[i] = tabl_o_per[j];
                tabl_o_per[j] = temp;
            }
            for (int i = 0; i < 4; i++)
            {
                if (tabl_o_per[0] == tabl_voprosov[current_answ, i])
                {
                    Right_answ[current_answ] += "1";
                    
                }
                else if (tabl_o_per[1] == tabl_voprosov[current_answ, i])
                {
                    Right_answ[current_answ] += "2";
                    
                }
                else if (tabl_o_per[2] == tabl_voprosov[current_answ, i])
                {
                    Right_answ[current_answ] += "3";
                    
                }
                else if (tabl_o_per[3] == tabl_voprosov[current_answ, i])
                {
                    Right_answ[current_answ] += "4";
                    
                }
            }
            



            groupBox.Header = "вопрос №" + Convert.ToString(1 + a) + " " + Quest[current_answ];
            groupBox.VerticalAlignment = VerticalAlignment.Center;
            groupBox.HorizontalAlignment = HorizontalAlignment.Center;
            groupBox.FontSize = 20;
            grid_work_zone.Children.Add(groupBox);
            Grid grid_box = new Grid();
            groupBox.Content = grid_box;
            ColumnDefinition column_1 = new ColumnDefinition();
            ColumnDefinition column_2 = new ColumnDefinition();
            grid_box.ColumnDefinitions.Add(column_1);
            grid_box.ColumnDefinitions.Add(column_2);
            StackPanel stack_vopr = new StackPanel();
            StackPanel stack_otv = new StackPanel();
            grid_box.Children.Add(stack_vopr);
            grid_box.Children.Add(stack_otv);
            Grid.SetColumn(stack_vopr, 0);
            Grid.SetColumn(stack_otv, 1);

            stack_vopr.Children.Add(answ_1);
            answ_1.Content = "А) "+tabl_otvetov[current_answ,0];

            stack_vopr.Children.Add(answ_2);
            answ_2.Content = "Б) " + tabl_otvetov[current_answ, 1];


            stack_vopr.Children.Add(answ_3);
            answ_3.Content = "В) " + tabl_otvetov[current_answ, 2];


            stack_vopr.Children.Add(answ_4);
            answ_4.Content = "Г) " + tabl_otvetov[current_answ, 3];

            RadioButton otvet_1 = new RadioButton();
            stack_otv.Children.Add(otvet_1);
            otvet_1.Content = "1) " + tabl_o_per[0];

            RadioButton otvet_2 = new RadioButton();
            stack_otv.Children.Add(otvet_2);
            otvet_2.Content = "2) " + tabl_o_per[ 1];

            RadioButton otvet_3 = new RadioButton();
            stack_otv.Children.Add(otvet_3);
            otvet_3.Content = "3) " + tabl_o_per[ 2];

            RadioButton otvet_4 = new RadioButton();
            stack_otv.Children.Add(otvet_4);
            otvet_4.Content = "4) " + tabl_o_per[ 3];

            StackPanel stackPanel = new StackPanel();
            grid_work_zone.Children.Add(stackPanel);
            Grid.SetRow(stackPanel, 1);
            
            Button button_check = new Button();
            
            button_check.Content = "Занести ответ в таблицу";
            stackPanel.Children.Add(button_check);
            button_check.HorizontalAlignment = HorizontalAlignment.Center;
            button_check.VerticalAlignment = VerticalAlignment.Top;
            button_check.FontSize = 20;

            DataGrid dataGrid = new DataGrid();
            stackPanel.Children.Add(dataGrid);

            dataGrid.Margin = new Thickness(0, 10, 0, 0);
            dataGrid.AutoGenerateColumns = false;
            dataGrid.HorizontalAlignment = HorizontalAlignment.Center;
            dataGrid.FontSize = 20;
            dataGrid.CanUserSortColumns = false;
            dataGrid.Style = (Style)mainWindow_style.Resources["DataGrid"];
            dataGrid.CanUserReorderColumns = false;
            dataGrid.FontFamily = new FontFamily("Dubai Light");
            dataGrid.CanUserResizeColumns = false;
            dataGrid.CanUserResizeRows = false;
            DataGridTextColumn c1 = new DataGridTextColumn();
            DataGridTextColumn c2 = new DataGridTextColumn();
            DataGridTextColumn c3 = new DataGridTextColumn();
            DataGridTextColumn c4 = new DataGridTextColumn();
            dataGrid.Columns.Add(c1);
            c1.IsReadOnly = true;
            dataGrid.Columns.Add(c2);
            c2.IsReadOnly = true;
            dataGrid.Columns.Add(c3);
            c3.IsReadOnly = true;
            dataGrid.Columns.Add(c4);
            c4.IsReadOnly = true;

            c1.Header = "А";
            c2.Header = "Б";
            c3.Header = "В";
            c4.Header = "Г";
            c1.Binding = new Binding("Answers_1");
            c2.Binding = new Binding("Answers_2");
            c3.Binding = new Binding("Answers_3");
            c4.Binding = new Binding("Answers_4");

            int[] otv = new int[5] ;

             button_check.Click += (s, e) =>
             {
                 if(answ_1.IsChecked == true && otvet_1.IsChecked == true)
                 {
                     otv[1] = 1;
                 }
                 else if (answ_2.IsChecked == true && otvet_1.IsChecked == true)
                 {                     
                     otv[2] = 1;
                 }
                 else if (answ_3.IsChecked == true && otvet_1.IsChecked == true)
                 {
                     
                     otv[3] = 1;
                 }
                 else if (answ_4.IsChecked == true && otvet_1.IsChecked == true)
                 {
                     
                     otv[4] = 1;
                 }

                 else if (answ_1.IsChecked == true && otvet_2.IsChecked == true)
                 {
                    
                     otv[1] = 2;
                 }
                 else if (answ_2.IsChecked == true && otvet_2.IsChecked == true)
                 {
                     
                     otv[2] = 2;
                 }
                 else if (answ_3.IsChecked == true && otvet_2.IsChecked == true)
                 {
                    
                     otv[3] = 2;
                 }
                 else if (answ_4.IsChecked == true && otvet_2.IsChecked == true)
                 {
                     
                     otv[4] = 2;
                 }

                 else if (answ_1.IsChecked == true && otvet_3.IsChecked == true)
                 {

                     otv[1] = 3;
                 }
                 else if (answ_2.IsChecked == true && otvet_3.IsChecked == true)
                 {
                     
                     otv[2] = 3;
                 }
                 else if (answ_3.IsChecked == true && otvet_3.IsChecked == true)
                 {
                     
                     otv[3] = 3;
                 }
                 else if (answ_4.IsChecked == true && otvet_3.IsChecked == true)
                 {
                     
                     otv[4] = 3;
                 }

                 else if (answ_1.IsChecked == true && otvet_4.IsChecked == true)
                 {
                    
                     otv[1] = 4;
                 }
                 else if (answ_2.IsChecked == true && otvet_4.IsChecked == true)
                 {
                    
                     otv[2] = 4;
                 }
                 else if (answ_3.IsChecked == true && otvet_4.IsChecked == true)
                 {
                    
                     otv[3] = 4;
                 }
                 else if (answ_4.IsChecked == true && otvet_4.IsChecked == true)
                 {
                    
                     otv[4] = 4;
                 }

                 for(int i = 0; i < 5; i++)
                 {
                     for(int j = 0; j < 5; j++)
                     {
                         if (i != j)
                         {
                             if (otv[i] == otv[j])
                             {
                                 otv[j] = 0;
                             }
                         }
                     }
                 }

                 dataGrid.Items.Clear();
                 otvet[current_answ] = null;
                 for(int i=0; i < 4; i++)
                 {
                     otvet[current_answ] += Convert.ToString(otv[i + 1]);
                     
                 }





                 dataGrid.Items.Add(new Answer() {Answers_1 = Convert.ToString(otv[1]), Answers_2 = Convert.ToString(otv[2]), Answers_3 = Convert.ToString(otv[3]), Answers_4 = Convert.ToString(otv[4]) });
                 
                 for (int i = 1; i < 5; i++)
                 {
                     if (otv[i] - 1 != -1)
                     {
                         if (tabl_o_per[otv[i] - 1] == tabl_voprosov[current_answ, (i - 1)])
                         {
                             proverka[i - 1] = 1;
                         }
                         else
                         {
                             proverka[i - 1] = 0;
                         }
                     }
                     else
                     {
                         proverka[i - 1] = 0;
                     }
                 }
                 summa_ballov[current_answ] = proverka.Sum()*koef[current_answ];
             };
            
            
        }

        
        public void Create_content_test(int size,  string[] Image_source, string[] Quest,string[] Answ1, string[] Answ2, string[] Answ3, string[] Answ4, string[] Right_answers, 
            string[] Difficulty,string[] type, string[,] tabl_otvetov, string[,] tabl_voprosov , Grid name_grids) //Создание интерфейса теста
        {           
            int current_answ = 0;
            int a = 0;
            int[] mas = Enumerable.Range(0, size).ToArray();

            // Перемешивание
            Random rnd = new Random();
            for (int i = mas.Length - 1; i > 0; i--)
            {
                int j = rnd.Next(0, i + 1);
                int temp = mas[i];
                mas[i] = mas[j];
                mas[j] = temp;
            }
            current_answ = mas[a];
            
           
            
                Grid grid_work_zone = new Grid();
                grid_work_zone.MaxHeight = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height - 130;
                RowDefinition row_image = new RowDefinition();
                RowDefinition row_image_1 = new RowDefinition();
                RowDefinition row_image_2 = new RowDefinition();
                grid_work_zone.RowDefinitions.Add(row_image);
                grid_work_zone.RowDefinitions.Add(row_image_1);
                grid_work_zone.RowDefinitions.Add(row_image_2);
                name_grids.Children.Add(grid_work_zone);

            stackPanel_in_answ.Children.Add(radioButton_ans_1);
            stackPanel_in_answ.Children.Add(radioButton_ans_2);
            stackPanel_in_answ.Children.Add(radioButton_ans_3);
            stackPanel_in_answ.Children.Add(radioButton_ans_4);

            if (type[current_answ] == "закрытая")
            {
               
            
                Create_content_type_closed(grid_work_zone,row_image, row_image_1,  current_answ, Image_source, Quest, Answ1, Answ2, Answ3, Answ4, a);



            }
            else if(type[current_answ] == "сравнение")
            {

                Create_content_type_srav(grid_work_zone, row_image, tabl_otvetov, tabl_voprosov,a, current_answ, Quest, Image_source);

            }


                Button button_left = new Button();
                name_grids.Children.Add(button_left);
                GridLength gridLength_1 = new GridLength(70);
                row_image_2.Height = gridLength_1;
                Grid.SetRow(button_left, 2);
                button_left.Background = Brushes.Transparent;
                button_left.BorderBrush = Brushes.Transparent;
                button_left.MaxHeight = 40;
                button_left.MaxWidth = 40;
                button_left.HorizontalAlignment = HorizontalAlignment.Left;
                button_left.VerticalAlignment = VerticalAlignment.Bottom;
                button_left.Style = (Style)mainWindow_style.Resources["DefaultButtonStyle"];
                button_left.Content = Application.Current.FindResource("arrow_left");
                button_left.Click += (s, e) =>
                {
                    if (a > 0)
                    {

                        a--;
                        current_answ = mas[a];
                        if (type[current_answ] == "закрытая")
                        {
                            grid_work_zone.Children.Clear();
                            Create_content_type_closed(grid_work_zone, row_image, row_image_1, current_answ, Image_source, Quest, Answ1, Answ2, Answ3, Answ4, a);
                        }
                        else if (type[current_answ] == "сравнение")
                        {
                            grid_work_zone.Children.Clear();
                            Create_content_type_srav(grid_work_zone, row_image, tabl_otvetov, tabl_voprosov, a, current_answ, Quest,  Image_source);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Вы в начале теста");
                    }
                };

                Button button_right = new Button();
                name_grids.Children.Add(button_right);
                Grid.SetRow(button_right, 2);
                button_right.Background = Brushes.Transparent;
                button_right.BorderBrush = Brushes.Transparent;
                button_right.MaxHeight = 40;
                button_right.MaxWidth = 40;
                button_right.Margin = new Thickness(0, 0, 10, 0);
                button_right.HorizontalAlignment = HorizontalAlignment.Right;
                button_right.VerticalAlignment = VerticalAlignment.Bottom;
                button_right.Style = (Style)mainWindow_style.Resources["DefaultButtonStyle"];
                button_right.Content = Application.Current.FindResource("arrow_right");
                button_right.Click += (s, e) =>
                 {
                     if (a<size-1)
                     {
                         if (type[current_answ] == "закрытая")
                         {
                             if (radioButton_ans_1.IsChecked == true)
                             {
                                 if ("1" == Right_answ[current_answ])
                                 {
                                     summa_ballov[current_answ] = 1 * koef[current_answ];
                                 }
                                 else
                                 {
                                     summa_ballov[current_answ] = 0;
                                 }
                                 otvet[current_answ] = "1";

                             }
                             else if (radioButton_ans_2.IsChecked == true)
                             {
                                 if ("2" == Right_answ[current_answ])
                                 {
                                     summa_ballov[current_answ] = 1 * koef[current_answ];
                                 }
                                 else
                                 {
                                     summa_ballov[current_answ] = 0;
                                 }
                                 otvet[current_answ] = "2";
                             }
                             else if (radioButton_ans_3.IsChecked == true)
                             {
                                 if ("3" == Right_answ[current_answ])
                                 {
                                     summa_ballov[current_answ] = 1 * koef[current_answ];
                                 }
                                 else
                                 {
                                     summa_ballov[current_answ] = 0;
                                 }
                                 otvet[current_answ] = "3";
                             }
                             else if (radioButton_ans_4.IsChecked == true)
                             {
                                 if ("4" == Right_answ[current_answ])
                                 {
                                     summa_ballov[current_answ] = 1 * koef[current_answ];
                                 }
                                 else
                                 {
                                     summa_ballov[current_answ] = 0;
                                 }
                                 otvet[current_answ] = "4";
                             }
                             else { }
                         }
                            a++;
                         current_answ = mas[a];
                         if (type[current_answ] == "закрытая")
                             {
                             
                            
                             grid_work_zone.Children.Clear();
                             Create_content_type_closed(grid_work_zone,row_image, row_image_1, current_answ, Image_source, Quest, Answ1, Answ2, Answ3, Answ4, a);
                         }
                         else if (type[current_answ] == "сравнение")
                         {
                             grid_work_zone.Children.Clear();
                             Create_content_type_srav(grid_work_zone, row_image, tabl_otvetov, tabl_voprosov, a, current_answ, Quest, Image_source);
                            

                         }
                         
                     }

                     else if(a==size-1)
                     {
                         if (type[current_answ] == "закрытая")
                         {
                             if (radioButton_ans_1.IsChecked == true)
                             {
                                 if ("1" == Right_answ[current_answ])
                                 {
                                     summa_ballov[current_answ] = 1 * koef[current_answ];
                                 }
                                 else
                                 {
                                     summa_ballov[current_answ] = 0;
                                 }
                                 otvet[current_answ] = "1";

                             }
                             else if (radioButton_ans_2.IsChecked == true)
                             {
                                 if ("2" == Right_answ[current_answ])
                                 {
                                     summa_ballov[current_answ] = 1 * koef[current_answ];
                                 }
                                 else
                                 {
                                     summa_ballov[current_answ] = 0;
                                 }
                                 otvet[current_answ] = "2";
                             }
                             else if (radioButton_ans_3.IsChecked == true)
                             {
                                 if ("3" == Right_answ[current_answ])
                                 {
                                     summa_ballov[current_answ] = 1 * koef[current_answ];
                                 }
                                 else
                                 {
                                     summa_ballov[current_answ] = 0;
                                 }
                                 otvet[current_answ] = "3";
                             }
                             else if (radioButton_ans_4.IsChecked == true)
                             {
                                 if ("4" == Right_answ[current_answ])
                                 {
                                     summa_ballov[current_answ] = 1 * koef[current_answ];
                                 }
                                 else
                                 {
                                     summa_ballov[current_answ] = 0;
                                 }
                                 otvet[current_answ] = "4";
                             }
                             else { }
                         }
                         System.Windows.Forms.DialogResult result = System.Windows.Forms.MessageBox.Show("Конец теста! Готовы проверить ответы?", "Проверка",
                                                                                                         System.Windows.Forms.MessageBoxButtons.YesNo,
                                                                                                         System.Windows.Forms.MessageBoxIcon.Question);
                         if (result == System.Windows.Forms.DialogResult.Yes)
                         {

                             grid_work_zone.Children.Clear();
                             grid_work_zone.RowDefinitions.Clear();

                             RowDefinition row_win = new RowDefinition();
                             RowDefinition row_win_1 = new RowDefinition();
                             
                             grid_work_zone.RowDefinitions.Add(row_win);
                             grid_work_zone.RowDefinitions.Add(row_win_1);
                             GridLength gridLength = new GridLength(2, GridUnitType.Star);
                             row_win_1.Height = gridLength;



                             Label label_win = new Label();
                             grid_work_zone.Children.Add(label_win);
                             Grid.SetRow(label_win, 0);
                             label_win.FontSize = 20;
                             label_win.FontFamily = new FontFamily("Dubai Light");
                             label_win.Content = "Поздравляем! Вы прошли тест!\n\tВаши результаты";
                             label_win.VerticalAlignment = VerticalAlignment.Bottom;
                             label_win.HorizontalAlignment = HorizontalAlignment.Center;

                             Button Export = new Button();
                             grid_work_zone.Children.Add(Export);
                             Grid.SetRow(Export, 0);
                             Export.Background = Brushes.Transparent;
                             Export.BorderBrush = Brushes.Transparent;
                             Export.MaxHeight = 40;
                             Export.MaxWidth = 140;
                             Export.Margin = new Thickness(0, 0, 10, 0);
                             Export.HorizontalAlignment = HorizontalAlignment.Right;
                             Export.FontSize = 20;
                             Export.VerticalAlignment = VerticalAlignment.Bottom;
                             Export.Style = (Style)mainWindow_style.Resources["DefaultButtonStyle"];
                             Export.Content = "Export";
                             Export.Click += Export_Click;


                             int[] maks_ball = new int[size];

                             for(int i = 0; i < size; i++)
                             {
                                 if (type[i] == "закрытая")
                                 {
                                     if (koef[i] == 1)
                                     {
                                         maks_ball[i] = 1;
                                     }
                                     else if (koef[i] == 2)
                                     {
                                         maks_ball[i] = 2;
                                     }
                                     else if (koef[i] == 3)
                                     {
                                         maks_ball[i] = 3;
                                     }
                                     else
                                     {

                                     }
                                 }
                                 else if (type[i] == "сравнение")
                                 {
                                     if (koef[i] == 1)
                                     {
                                         maks_ball[i] = 4* 1;
                                     }
                                     else if (koef[i] == 2)
                                     {
                                         maks_ball[i] = 4*2;
                                     }
                                     else if (koef[i] == 3)
                                     {
                                         maks_ball[i] = 4*3;
                                     }
                                     else
                                     {

                                     }
                                 }
                                 }

                             DataGrid dataGrid_res = new DataGrid();
                             grid_work_zone.Children.Add(dataGrid_res);
                             Grid.SetRow(dataGrid_res, 1);
                             dataGrid_res.AutoGenerateColumns = false;
                             dataGrid_res.HorizontalAlignment = HorizontalAlignment.Center;
                             dataGrid_res.VerticalAlignment = VerticalAlignment.Top;
                             dataGrid_res.FontSize = 20;
                             dataGrid_res.CanUserSortColumns = false;
                             dataGrid_res.Style = (Style)mainWindow_style.Resources["DataGrid"];
                             dataGrid_res.CanUserReorderColumns = false;
                             dataGrid_res.FontFamily = new FontFamily("Dubai Light");
                             dataGrid_res.CanUserResizeColumns = false;
                             dataGrid_res.CanUserResizeRows = false;
                             DataGridTextColumn c1 = new DataGridTextColumn();
                             DataGridTextColumn c2 = new DataGridTextColumn();
                             DataGridTextColumn c3 = new DataGridTextColumn();
                             DataGridTextColumn c4 = new DataGridTextColumn();
                             DataGridTextColumn c5 = new DataGridTextColumn();
                             DataGridTextColumn c6 = new DataGridTextColumn();
                             DataGridTextColumn c7 = new DataGridTextColumn();
                             dataGrid_res.Columns.Add(c1);
                             c1.IsReadOnly = true;
                             dataGrid_res.Columns.Add(c2);
                             c2.IsReadOnly = true;
                             dataGrid_res.Columns.Add(c3);
                             c3.IsReadOnly = true;
                             dataGrid_res.Columns.Add(c4);
                             c4.IsReadOnly = true;
                             dataGrid_res.Columns.Add(c5);
                             c5.IsReadOnly = true;
                             dataGrid_res.Columns.Add(c6);
                             c6.IsReadOnly = true;
                             dataGrid_res.Columns.Add(c7);
                             c7.IsReadOnly = true;

                             c1.Header = "№";
                             c2.Header = "Ваш ответ";
                             c3.Header = "Правильный ответ";
                             c4.Header = "Балл";
                             c5.Header = "Макс. Балл";
                             c6.Header = "Итоговый балл";
                             c7.Header = "Итоговый макс. Балл";
                             c1.Binding = new Binding("Number");
                             c2.Binding = new Binding("User_answ");
                             c3.Binding = new Binding("Answ");
                             c4.Binding = new Binding("ball");
                             c5.Binding = new Binding("maks_ball_cur");
                             c6.Binding = new Binding("user_ball");
                             c7.Binding = new Binding("maks_Ball");

                             for (int i = 0; i < size; i++)
                             {
                                 dataGrid_res.Items.Add(new Item() { Number = Convert.ToString(i + 1),
                                     User_answ = otvet[mas[i]], Answ = Right_answ[mas[i]], ball = summa_ballov[mas[i]], maks_ball_cur = maks_ball_curr_quest[mas[i]] 
                                 });
                             }
                             dataGrid_res.Items.Add(new Itog_ball() { user_ball = summa_ballov.Sum(), maks_Ball = maks_ball_curr_quest.Sum() });
                             button_left.Visibility = Visibility.Hidden;
                             button_right.Visibility = Visibility.Hidden;
                             temp = dataGrid_res;
                         }

                     }
                 };
            }

        private void Export_Click(object sender, RoutedEventArgs e)
        {
            string path = "D:\\test.docx";


            System.Windows.Forms.SaveFileDialog saveFileDialog = new System.Windows.Forms.SaveFileDialog();

            if (saveFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                path = saveFileDialog.FileName;
            }

            Word._Application word_app = new Word.Application();





            // Создаем документ Word.
            object missing = Type.Missing;
            Word._Document word_doc = word_app.Documents.Add(
                ref missing, ref missing, ref missing, ref missing);

            // Создаем абзац заголовка.
            Word.Paragraph para = word_doc.Paragraphs.Add(ref missing);
            para.Range.Text = "Тест по истории";
            object style_name = "Заголовок 1";
            para.Range.set_Style(ref style_name);
            para.Range.InsertParagraphAfter();

            // Добавить текст.


            for (int i = 0; i < A1.Length; i++)
            {
                para.Range.Text = Quest[i];
                para.Range.InsertParagraphAfter();
                para.Range.Text = "1) " + A1[i];
                para.Range.InsertParagraphAfter();
                para.Range.Text = "2) " + A2[i];
                para.Range.InsertParagraphAfter();
                para.Range.Text = "3) " + A3[i];
                para.Range.InsertParagraphAfter();
                para.Range.Text = "4) " + A4[i];
                para.Range.InsertParagraphAfter();
                para.Range.InsertParagraphAfter();
            }
            for(int i = 0; i <tabl_srav_otv.GetLength(0);i++)
            {
                if (tabl_srav_vopr[i, 0] != null)
                {
                    para.Range.InsertParagraphAfter();
                    para.Range.Text = "Установите соответствие";
                    para.Range.InsertParagraphAfter();
                    para.Range.Text = "A) " + tabl_srav_otv[i, 0] + "    " + "1) " + tabl_srav_vopr[i, 1];
                    para.Range.InsertParagraphAfter();
                    para.Range.Text = "Б) " + tabl_srav_otv[i, 1] + "    " + "2) " + tabl_srav_vopr[i, 2];
                    para.Range.InsertParagraphAfter();
                    para.Range.Text = "В) " + tabl_srav_otv[i, 2] + "    " + "3) " + tabl_srav_vopr[i, 3];
                    para.Range.InsertParagraphAfter();
                    para.Range.Text = "Г) " + tabl_srav_otv[i, 3] + "    " + "4) " + tabl_srav_vopr[i, 0];
                    para.Range.InsertParagraphAfter();
                    para.Range.Text = "А:   Б:   В:   Г:";
                    para.Range.InsertParagraphAfter();
                }
                
            }
           

            // Сохраним текущий шрифт и начнем с использования Courier New.
            string old_font = para.Range.Font.Name;
            para.Range.Font.Name = "Courier New";



            // Сохраним документ.
            object filename = path + ".docx";

            word_doc.SaveAs(ref filename, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing,
                ref missing);

            // Закрыть.
            object save_changes = false;
            word_doc.Close(ref save_changes, ref missing, ref missing);
            word_app.Quit(ref save_changes, ref missing, ref missing);


            Excel.Application excel = new Excel.Application();
            
            Workbook workbook = excel.Workbooks.Add(System.Reflection.Missing.Value);
            Worksheet sheet1 = (Worksheet)workbook.Sheets[1];

            for (int j = 0; j < temp.Columns.Count; j++)
            {
                Range myRange = (Range)sheet1.Cells[1, j + 1];
                sheet1.Cells[1, j + 1].Font.Bold = true;
                sheet1.Columns[j + 1].ColumnWidth = 15;
                myRange.Value2 = temp.Columns[j].Header;
                myRange.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlRgbColor.rgbBlack);
                myRange.EntireColumn.AutoFit();
            }
            for (int i = 0; i < temp.Columns.Count; i++)
            {
                for (int j = 0; j < temp.Items.Count; j++)
                {
                    TextBlock b = temp.Columns[i].GetCellContent(temp.Items[j]) as TextBlock;
                    Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[j + 2, i + 1];
                    myRange.Value2 = b.Text;
                    myRange.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlRgbColor.rgbBlack);
                }
            }

            excel.Application.ActiveWorkbook.SaveAs(path + ".xlsx", Type.Missing,
  Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
  Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            excel.Quit();
            MessageBox.Show("Экспорт выполнен успешно");




        }
    }
    
    class Item //Класс для работы финальной таблицы результатов
    {
        public string Number { get; set; }
        public string User_answ { get; set; }
        public string Answ { get; set; }
        public int ball { get; set; }
        public int maks_ball_cur { get; set; }
        

    }
    class Itog_ball
    {
        public int user_ball { get; set; }
        public int maks_Ball { get; set; }
    }
    class Answer
    {
        public string Answers_1 { get; set; }
        public string Answers_2 { get; set; }
        public string Answers_3 { get; set; }
        public string Answers_4 { get; set; }
    }

    public partial class MainWindow : Window
    {
        double font_size = 20;
        public MainWindow()
        {
            InitializeComponent();
        }

        // Settings
        private void But_settings_Click(object sender, RoutedEventArgs n)
        {
            //<===Point for gradient===>
            System.Windows.Point start_point = new System.Windows.Point(0.0, 0.0);
            Point end_point = new Point(0.0, 0.9);
            //<!===Point for gradient===>
            
            //<===Brushes for gradient===>
            LinearGradientBrush brush_fon_1_gr = new LinearGradientBrush(Colors.White, Colors.LightGray, start_point, end_point);
            LinearGradientBrush brush_fon_2_gr = new LinearGradientBrush(Colors.White, Colors.Gray, start_point, end_point);
            //<!===Brushes for gradient===>
            grid.Children.Clear();

            //Label Color of fon
            Label label_fon_color = new Label();
            grid.Children.Add(label_fon_color);
            string label_fon_content = "Цвет фона:";
            label_fon_color.FontFamily = new FontFamily("Dubai Light");
            label_fon_color.FontSize = font_size;
            label_fon_color.Content = label_fon_content;
            double label_width = label_fon_content.Length * 1 / 2 * font_size;
            //End Label Color of fon
            //<===Button 1 change fon===>
            Button button = new Button();
            grid.Children.Add(button);
            //<===Button Property
            button.Style = (Style)Resources["DefaultButtonStyle"];
            button.MaxHeight = font_size * 2;
            button.MaxWidth = font_size * 2;
            button.MinHeight = font_size * 2;
            button.MinWidth = font_size * 2;
            button.VerticalAlignment = VerticalAlignment.Top;
            button.HorizontalAlignment = HorizontalAlignment.Left;
            button.Margin = new Thickness(label_width + button.MinWidth / 2, font_size / 4, 0, 0);
            //Button Property===!>
            //<===Button Action
            button.Background = brush_fon_1_gr;
            button.MouseMove += (s, e) =>
            {
                button.BorderBrush = Brushes.Red;
                button.BorderThickness = new Thickness(3);
            };
            button.MouseLeave += (s, e) =>
            {
                button.BorderBrush = Brushes.Black;
                button.BorderThickness = new Thickness(1);
            };
            button.Click += (s, e) =>
            {
                window_main.Background = brush_fon_1_gr;
            };
            //Button Action===!>
            //<!===Button 1 change fon===>
            //<===Button 2 change fon===>
            //<===Button Property
            Button button1 = new Button();
            grid.Children.Add(button1);
            button1.Style = (Style)Resources["DefaultButtonStyle"];
            button1.MaxHeight = font_size * 2;
            button1.MaxWidth = font_size * 2;
            button1.MinHeight = font_size * 2;
            button1.MinWidth = font_size * 2;
            button1.VerticalAlignment = VerticalAlignment.Top;
            button1.HorizontalAlignment = HorizontalAlignment.Left;
            button1.Margin = new Thickness(label_width + button.MinWidth + button1.MinWidth / 2 + 10, font_size / 4, 0, 0);
            button1.Background = brush_fon_2_gr;
            //Button Property===!>
            //<===Button Action
            button1.MouseMove += (s, e) =>
            {
                button1.BorderBrush = Brushes.Red;
                button1.BorderThickness = new Thickness(3);
            };
            button1.MouseLeave += (s, e) =>
            {
                button1.BorderBrush = Brushes.Black;
                button1.BorderThickness = new Thickness(1);
            };
            button1.Click += (s, e) =>
            {
                window_main.Background = brush_fon_2_gr;
            };
            //Button Action===!>
            //<!===Button 2 change fon===>
        }

        
        //End Settings
        //<===Action but_settings
        private void But_settings_move(object sender, MouseEventArgs e)
        {
            but_settings.Background = window_main.Background;
            but_settings.BorderBrush = window_main.Background;   
        }
        private void But_settings_move_over(object sender, MouseEventArgs e)
        {
            but_settings.Background = Brushes.Transparent;
            but_settings.BorderBrush = Brushes.Transparent;
        }

        //===Action but_settings===!>

        //Function for work test
        private void Animation_completed(StackPanel sp)
        {
            DoubleAnimation opacityAnimation = new DoubleAnimation();
            opacityAnimation.From = 1;
            opacityAnimation.To = 0.0;
            opacityAnimation.Duration = TimeSpan.FromSeconds(0.8);
            opacityAnimation.Completed += (se, ev) =>
            {
                string name_file = "Test_first" + "\\" + "testfirst.txt";
                grid.Children.Clear();
                
                Loading_Test(name_file);
            };
            sp.BeginAnimation(OpacityProperty, opacityAnimation);
        }

        private void Begin_tests(int number_test)
        {
            StackPanel stackPanel_begin = new StackPanel();
            grid.Children.Add(stackPanel_begin);
            stackPanel_begin.VerticalAlignment = VerticalAlignment.Center;
            stackPanel_begin.HorizontalAlignment = HorizontalAlignment.Center;
            Label new_test = new Label();
            stackPanel_begin.Children.Add(new_test);
            new_test.Content = "Привет пользователь! Ты октрыл: " + number_test + " тест. Удачи в выполнении!";
            new_test.FontFamily = new FontFamily("Dubai Light");
            new_test.FontSize = font_size;
            new_test.HorizontalAlignment = HorizontalAlignment.Center;
            new_test.VerticalAlignment = VerticalAlignment.Center;

            Button button_begin = new Button();
            string begin = "Начать!";
            stackPanel_begin.Children.Add(button_begin);
            button_begin.FontSize = font_size+5;
            button_begin.MaxWidth = stackPanel_begin.ViewportWidth / 2;
            button_begin.Style = (Style)Resources["ButtonStyle1"];
            button_begin.Content = new CornerRadius(15);
            button_begin.Content = new Border();
            button_begin.FontFamily = new FontFamily("Dubai Light");
            button_begin.Content = begin;
            button_begin.MaxWidth = begin.Length*font_size + 15; 
            button_begin.Click += (s, e) =>
            {
                Animation_completed(stackPanel_begin);
            };
            
        }
        private void Loading_Test(string name_file)
        {
            grid.Children.Clear();
            string path_file = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\"+"Test_History" +"\\"+name_file;
            Test test = new Test();
            test.Read_file_test(path_file, grid);
        }
        //End function for work test


        //Unit: Tests
        //Settings Expander in main menu
        private void but_test_spisok_move(object sender, MouseEventArgs e)
        {
          
            but_test_spisok.BorderBrush = window_main.Background;
            but_test_spisok.BorderThickness = new Thickness(2);
        }

        private void but_test_spisok_leave(object sender, MouseEventArgs e)
        {
            
            but_test_spisok.BorderBrush = Brushes.Transparent;
            but_test_spisok.BorderThickness = new Thickness(0);
        }
        //End settings Expander

        //_______test_1____
        private void But_test1_Click(object sender, RoutedEventArgs e)
        {
            grid.Children.Clear();
            Begin_tests(1);
        }
        //button test_1 Action
        private void test_1_move(object sender, MouseEventArgs e)
        {
            but_test1.Background = window_main.Background;
        }

        private void test_1_move_over(object sender, MouseEventArgs e)
        {
            but_test1.Background = Brushes.Transparent;
        }
        //End button test_1 Action
        //!_______test_1____

        

        //Button test Open
        private void But_test_open_click(object sender, RoutedEventArgs e)
        {
            Test test = new Test();
            grid.Children.Clear();
            string FilePath;
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Test Files(*.txt)|*.txt";
            if (openFileDialog.ShowDialog() == true)
            {
                FilePath = openFileDialog.FileName;
                test.Read_file_test(FilePath, grid);
            }
            else
            {
                MessageBox.Show("Файл не выбран");
            }
        }

        private void But__open_test_leave(object sender, MouseEventArgs e)
        {
            but_test_open.Background = Brushes.Transparent;
            but_test_open.BorderBrush = Brushes.Transparent;
        }

        private void But__open_test_move(object sender, MouseEventArgs e)
        {
            but_test_open.Background = window_main.Background;
            but_test_open.BorderBrush = window_main.Background;
        }

        private void close_mainwindow(object sender, EventArgs e)
        {
            Application.Current.Shutdown();
        }
        //! Button test Open
    }
    //End unit: Tests
}

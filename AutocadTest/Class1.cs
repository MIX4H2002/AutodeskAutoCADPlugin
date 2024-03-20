using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System.Collections.Generic;
using System.Data.SQLite;

namespace YourNamespace
{
    public class BlockInformationClass
    {
        [CommandMethod("YourBlockInformation")]
        public void ProvidingBlockInformation()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // Запрос выбора вхождений блоков
            PromptSelectionResult selRes = ed.GetSelection();
            if (selRes.Status != PromptStatus.OK)
            {
                ed.WriteMessage("Ошибка при выборе вхождений блоков.");
                return;
            }

            string connectionString = @"C:\Users\m-3k\OneDrive\Рабочий стол\стажировка\PartsDataBase.sqlite";
            using (SQLiteConnection connection = new SQLiteConnection($"Data Source={connectionString}"))
            {
                connection.Open();

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    PromptPointOptions ppo = new PromptPointOptions("\nУкажите точку для размещения таблицы: ");
                    PromptPointResult ppr = ed.GetPoint(ppo);
                    if (ppr.Status != PromptStatus.OK)
                    {
                        ed.WriteMessage("\nОтменено.");
                        return;
                    }

                    // Создание таблицы для хранения информации
                    Table table = new Table();
                    table.SetDatabaseDefaults();
                    table.TableStyle = db.Tablestyle;
                    table.Position = ppr.Value;
                    table.NumRows = selRes.Value.Count + 2;
                    table.NumColumns = 4;

                    double rowHeight = 100.0;
                    double[] columnWidths = new double[] { 200.0, 1500.0, 300.0, 300.0 };
                    double textSize = 40;
                    for (int i = 0; i < table.NumRows; i++)
                    {
                        table.SetRowHeight(i, rowHeight);
                    }
                    for (int j = 0; j < table.NumColumns; j++)
                    {
                        table.SetColumnWidth(j, columnWidths[j]);
                    }

                    // Установка размера текста для каждой ячейки таблицы
                    for (int row = 0; row < table.NumRows; row++)
                    {
                        for (int col = 0; col < table.NumColumns; col++)
                        {
                            table.SetTextHeight(row, col, textSize);
                        }
                    }

                    // Заголовки таблицы
                    table.SetTextString(1, 0, "№");
                    table.SetTextString(1, 1, "Наименование");
                    table.SetTextString(1, 2, "Кол-во");
                    table.SetTextString(1, 3, "Масса ед., кг");

                    int rowIndex = 2;

                    // Словарь для отслеживания числа экземпляров каждого GUID блока
                    Dictionary<string, int> blockCount = new Dictionary<string, int>();

                    Point3d basePoint = ppr.Value;

                    foreach (SelectedObject selObj in selRes.Value)
                    {
                        if (selObj != null)
                        {
                            BlockReference blkRef = tr.GetObject(selObj.ObjectId, OpenMode.ForRead) as BlockReference;
                            if (blkRef != null)
                            {

                                // Получение GUID блока из XData
                                string guid = "";

                                ResultBuffer rb = blkRef.XData;
                                if (rb != null)
                                {
                                    TypedValue[] values = rb.AsArray();
                                    for (int i = 0; i < values.Length; i += 1)
                                    {
                                        if (values[i].TypeCode == 1000)
                                        {
                                            guid = values[i].Value.ToString();
                                            break;
                                        }
                                    }
                                }

                                // Увеличение счетчика экземпляров блока
                                if (blockCount.ContainsKey(guid))
                                {
                                    blockCount[guid]++;
                                }
                                else
                                {
                                    blockCount.Add(guid, 1);
                                }

                                // Получение информации о блоке из базы данных по GUID
                                if (!string.IsNullOrEmpty(guid))
                                {
                                    string fullName = GetFullNameFromDatabase(guid, connection);
                                    double weight = GetWeightFromDatabase(guid, connection);

                                    // Заполнение таблицы
                                    table.SetTextString(rowIndex, 0, (rowIndex - 1).ToString());
                                    table.SetTextString(rowIndex, 1, fullName);
                                    table.SetTextString(rowIndex, 2, blockCount[guid].ToString());
                                    table.SetTextString(rowIndex, 3, weight.ToString());

                                    rowIndex++;
                                }
                            }

                        }
                    }
                    // Добавление таблицы в пространство модели
                    BlockTableRecord btr = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                    btr.AppendEntity(table);
                    tr.AddNewlyCreatedDBObject(table, true);

                    tr.Commit();
                }
            }
        }

        private string GetFullNameFromDatabase(string guid, SQLiteConnection connection)
        {
            string fullName = "";
            string query = "SELECT FullNameTemplate FROM Parts WHERE ID = @ID;";

            using (SQLiteCommand cmd = new SQLiteCommand(query, connection))
            {
                cmd.Parameters.AddWithValue("@ID", guid);
                using (SQLiteDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        fullName = reader["FullNameTemplate"].ToString();
                    }
                }
            }

            // Замена <$D$> на значение диаметра
            if (fullName.Contains("<$D1$>"))
            {
                fullName = fullName.Replace("<$D$>", GetDiameterFromDatabase(guid, connection)).Replace("<$D1$>", GetDiameterFromDatabase(guid, connection));
            }
            else
            {
                fullName = fullName.Replace("<$D$>", GetDiameterFromDatabase(guid, connection));
            }

            return fullName;
        }

        private string GetDiameterFromDatabase(string guid, SQLiteConnection connection)
        {
            string diameter = "";
            string query = "SELECT Diameter FROM Parts WHERE ID = @ID;";

            using (SQLiteCommand cmd = new SQLiteCommand(query, connection))
            {
                cmd.Parameters.AddWithValue("@ID", guid);
                using (SQLiteDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        diameter = reader["Diameter"].ToString();
                    }
                }
            }

            return diameter;
        }

        private double GetWeightFromDatabase(string guid, SQLiteConnection connection)
        {
            double weight = 0.0;
            string query = "SELECT Weight FROM Parts WHERE ID = @ID;";

            using (SQLiteCommand cmd = new SQLiteCommand(query, connection))
            {
                cmd.Parameters.AddWithValue("@ID", guid);
                using (SQLiteDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        double.TryParse(reader["Weight"].ToString(), out weight);
                    }
                }
            }

            return weight;
        }
    }
}
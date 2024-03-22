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
                    // Счетчик для нумерации мультивыносок
                    int number = 1;
                    int leaderNumber = 1;
                    Dictionary<ObjectId, int> leaderNumbers = new Dictionary<ObjectId, int>();
                    List<Point3d> firstPoints = new List<Point3d>(); // Список для хранения первых точек выносок
                    List<int> leaderCount = new List<int>(); // Список для хранения номеров выносок для каждого объекта


                    Dictionary<string, int> blockCount = new Dictionary<string, int>();
                    Dictionary<string, int> blockRowIndices = new Dictionary<string, int>();


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

                                    // Проверяем, содержится ли GUID в blockRowIndices
                                    if (blockRowIndices.ContainsKey(guid))
                                    {
                                        // Обновляем количество в существующей строке
                                        int existingRowIndex = blockRowIndices[guid];
                                        int newCount = blockCount[guid];
                                        table.SetTextString(existingRowIndex, 2, newCount.ToString());

                                        int rowIndex1 = blockRowIndices[guid];
                                        string mleaderText = table.GetTextString(rowIndex1, 0, 0); // Используем номер из таблицы

                                        Point3d blockPosition = blkRef.Position;
                                        CreateMLeader(db, tr, blockPosition, mleaderText);
                                    }
                                    else
                                    {
                                        // Заполнение таблицы
                                        table.SetTextString(rowIndex, 0, (rowIndex - 1).ToString());
                                        table.SetTextString(rowIndex, 1, fullName);
                                        table.SetTextString(rowIndex, 2, blockCount[guid].ToString());
                                        table.SetTextString(rowIndex, 3, weight.ToString());

                                        blockRowIndices[guid] = rowIndex;

                                        rowIndex++;

                                        // Текущие координаты блока
                                        Point3d blockPosition = blkRef.Position;
                                        CreateMLeader(db, tr, blockPosition, number.ToString());

                                        number++;
                                    }
                                }
                            }

                        }
                    }

                    rowIndex = 2;
                    while (rowIndex < table.NumRows)
                    {
                        bool isEmpty = true;

                        // Проверяем каждую ячейку в строке на наличие текста
                        for (int colIndex = 0; colIndex < table.NumColumns; colIndex++)
                        {
                            string cellText = table.GetTextString(rowIndex, 0, colIndex);
                            if (!string.IsNullOrEmpty(cellText))
                            {
                                isEmpty = false;
                                break;
                            }
                        }

                        if (isEmpty)
                        {
                            table.DeleteRows(rowIndex, 1);
                        }
                        else
                        {
                            rowIndex++;
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

        public void CreateMLeader(Database db, Transaction tr, Point3d blockPosition, string mleaderText)
        {
            double xOffset = 500.0;
            double yOffset = 500.0;
            double zOffset = 500.0;
            Point3d endPoint = new Point3d(blockPosition.X + xOffset, blockPosition.Y + yOffset, blockPosition.Z + zOffset);

            MLeader leader = new MLeader();
            leader.SetDatabaseDefaults();
            leader.ContentType = ContentType.MTextContent;

            MText text = new MText();
            text.SetDatabaseDefaults();
            text.Contents = mleaderText;
            text.Location = endPoint;

            leader.MText = text;

            leader.TextStyleId = db.Textstyle;
            leader.TextHeight = 45;
            leader.ArrowSize = 40;

            int leaderIndex = leader.AddLeaderLine(endPoint);
            leader.SetFirstVertex(leaderIndex, blockPosition);
            leader.LeaderLineType = LeaderType.StraightLeader;

            BlockTableRecord btr1 = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
            btr1.AppendEntity(leader);
            tr.AddNewlyCreatedDBObject(leader, true);
        }
    }
}
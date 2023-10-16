using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using Autodesk.AutoCAD.Geometry;
using AddIn;
using Autodesk.AutoCAD.Colors;

namespace CadProject
{
    public class ImportAndDrawGrid : ICadCommand
    {
        public override void Execute()
        {
            ImportFromExcelAndDrawGrid();
        }

        public void ImportFromExcelAndDrawGrid()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor editor = doc.Editor;

            string excelFilePath = GetExcelFilePath(editor);

            if (string.IsNullOrEmpty(excelFilePath))
            {
                editor.WriteMessage("No valid Excel file path provided.");
                return;
            }

            try
            {
                List<Tuple<double, double>> gridSpacingList = ReadGridSpacingFromExcel(excelFilePath);
                List<string> axisNumbersList = new List<string>
                {
                    "1","2","3","4","5","6","7","8","9","10","11","12","13","14","15","16"
                };
                List<string> axisLettersList = new List<string>
                {
                    "A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q"
                };

                if (gridSpacingList == null || gridSpacingList.Count == 0)
                {
                    editor.WriteMessage("Failed to read grid spacing from Excel.");
                    return;
                }

                // Lock the document
                using (doc.LockDocument())
                {
                    // Start a new transaction for drawing grid lines
                    using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
                    {
                        BlockTable bt = tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead) as BlockTable;
                        BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                        // Get the ObjectId of the Centerline linetype
                        string centerlineLtypeName = "Centerline";
                        ObjectId centerlineLtypeId = GetCenterlineLinetypeId(doc, centerlineLtypeName);
                       
                        double xGridSpacing = 0;
                        double yGridSpacing = 0;
                        for (var index = 0; index < gridSpacingList.Count; index++)
                        {
                            var gridSpacing = gridSpacingList[index];
                            xGridSpacing += gridSpacing.Item1;
                            yGridSpacing += gridSpacing.Item2;

                            // Draw vertical grid lines along the X-axis
                            Line verticalLine =
                                new Line(new Point3d(xGridSpacing, 0, 0),
                                    new Point3d(xGridSpacing, 340, 0)); // Adjust the limits as needed

                            verticalLine.LinetypeId = centerlineLtypeId; // Set the linetype
                            btr.AppendEntity(verticalLine);
                            tr.AddNewlyCreatedDBObject(verticalLine, true);
                            
                            if (index == gridSpacingList.Count - 1)
                            {
                                verticalLine.Erase();
                            }

                            // Draw horizontal grid lines along the Y-axis
                            Line horizontalLine =
                                new Line(new Point3d(0, yGridSpacing, 0),
                                    new Point3d(340, yGridSpacing, 0)); // Adjust the limits as needed

                            horizontalLine.LinetypeId = centerlineLtypeId; // Set the linetype

                            btr.AppendEntity(horizontalLine);
                            tr.AddNewlyCreatedDBObject(horizontalLine, true);

                            LinetypeTable lt = tr.GetObject(HostApplicationServices.WorkingDatabase.LinetypeTableId, OpenMode.ForRead) as LinetypeTable;
                            
                            horizontalLine.Color = Color.FromRgb(0x86, 0x8e, 0x96);
                            verticalLine.Color = Color.FromRgb(0x86, 0x8e, 0x96);

                            // Draw a circle to put axis number
                            Circle axisCircleStart = new Circle();
                            axisCircleStart.Center = new Point3d(0-4, yGridSpacing, 0); // Adjust position
                            axisCircleStart.Radius = 4.0; // Adjust radius as needed

                            if (lt.Has("CONTINUOUS"))
                            {
                                axisCircleStart.LinetypeId = lt["CONTINUOUS"];
                            }
                            axisCircleStart.Color = Color.FromRgb(0x86, 0x8e, 0x96);

                            Circle axisCircleEnd = new Circle();
                            axisCircleEnd.Center = new Point3d(340+4, yGridSpacing, 0); // Adjust position
                            axisCircleEnd.Radius = 4.0; // Adjust radius as needed


                            if (lt.Has("CONTINUOUS"))
                            {
                                axisCircleEnd.LinetypeId = lt["CONTINUOUS"];
                            }
                            axisCircleEnd.Color = Color.FromRgb(0x86, 0x8e, 0x96);

                            btr.AppendEntity(axisCircleStart);
                            tr.AddNewlyCreatedDBObject(axisCircleStart, true);
                            btr.AppendEntity(axisCircleEnd);
                            tr.AddNewlyCreatedDBObject(axisCircleEnd, true);

                            // Add text to the circle for axis number
                            string axisNumber = axisNumbersList[index];
                            DBText axisTextStart = new DBText();
                            axisTextStart.TextString = axisNumber;
                            axisTextStart.Position = new Point3d(axisCircleStart.Center.X - 1.5, yGridSpacing - 1.5, 0); // Adjust position
                            axisTextStart.Height = 3; // Adjust text height

                            DBText axisTextEnd = new DBText();
                            axisTextEnd.TextString = axisNumber;
                            axisTextEnd.Position = new Point3d(axisCircleEnd.Center.X - 1.5, yGridSpacing - 1.5, 0); // Adjust position
                            axisTextEnd.Height = 3; // Adjust text height

                            btr.AppendEntity(axisTextStart);
                            tr.AddNewlyCreatedDBObject(axisTextStart, true);
                            btr.AppendEntity(axisTextEnd);
                            tr.AddNewlyCreatedDBObject(axisTextEnd, true);

                            // Draw a circle to put axis number
                            Circle axisCircleLetterStart = new Circle();
                            axisCircleLetterStart.Center = new Point3d( xGridSpacing, 0 - 4, 0); // Adjust position
                            axisCircleLetterStart.Radius = 4.0; // Adjust radius as needed

                            if (lt.Has("CONTINUOUS"))
                            {
                                axisCircleLetterStart.LinetypeId = lt["CONTINUOUS"];
                            }
                            axisCircleLetterStart.Color = Color.FromRgb(0x86, 0x8e, 0x96);

                            Circle axisCircleLetterEnd = new Circle();
                            axisCircleLetterEnd.Center = new Point3d( xGridSpacing, 340 + 4, 0); // Adjust position
                            axisCircleLetterEnd.Radius = 4.0; // Adjust radius as needed

                            if (lt.Has("CONTINUOUS"))
                            {
                                axisCircleLetterEnd.LinetypeId = lt["CONTINUOUS"];
                            }
                            axisCircleLetterEnd.Color = Color.FromRgb(0x86, 0x8e, 0x96);

                            btr.AppendEntity(axisCircleLetterStart);
                            tr.AddNewlyCreatedDBObject(axisCircleLetterStart, true);
                            btr.AppendEntity(axisCircleLetterEnd);
                            tr.AddNewlyCreatedDBObject(axisCircleLetterEnd, true);

                            // Add text to the circle for axis number
                            string axisLetter = axisLettersList[index];
                            DBText axisTextLetterStart = new DBText();
                            axisTextLetterStart.TextString = axisLetter;
                            axisTextLetterStart.Position = new Point3d( xGridSpacing - 1.5, axisCircleLetterStart.Center.Y - 1.5, 0); // Adjust position
                            axisTextLetterStart.Height = 3; // Adjust text height

                            DBText axisTextLetterEnd = new DBText();
                            axisTextLetterEnd.TextString = axisLetter;
                            axisTextLetterEnd.Position = new Point3d( xGridSpacing - 1.5, axisCircleLetterEnd.Center.Y - 1.5, 0); // Adjust position
                            axisTextLetterEnd.Height = 3; // Adjust text height

                            btr.AppendEntity(axisTextLetterStart);
                            tr.AddNewlyCreatedDBObject(axisTextLetterStart, true);
                            btr.AppendEntity(axisTextLetterEnd);
                            tr.AddNewlyCreatedDBObject(axisTextLetterEnd, true);

                            if (index == gridSpacingList.Count - 1)
                            {
                                horizontalLine.Erase();
                                axisCircleStart.Erase();
                                axisCircleEnd.Erase();
                                axisTextStart.Erase();
                                axisTextEnd.Erase();
                                axisCircleLetterStart.Erase();
                                axisCircleLetterEnd.Erase();
                                axisTextLetterStart.Erase();
                                axisTextLetterEnd.Erase();
                            }
                        }

                        tr.Commit();
                    }
                }

                editor.WriteMessage("Grid lines drawn successfully.");
            }
            catch (Exception ex)
            {
                editor.WriteMessage($"Error: {ex.Message}\n");
            }
        }

        private string GetExcelFilePath(Editor editor)
        {
            PromptResult excelFilePrompt = editor.GetString("\nEnter the path to the Excel file: ");
            return excelFilePrompt.Status == PromptStatus.OK ? excelFilePrompt.StringResult : null;
        }

        private List<Tuple<double, double>> ReadGridSpacingFromExcel(string filePath)
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // or LicenseContext.Commercial

                FileInfo fileInfo = new FileInfo(filePath);
                using (ExcelPackage package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Assuming the data is in the first worksheet

                    // Define the starting row and columns for horizontal and vertical spacing
                    int startRow = 2; // Starting from row 2
                    int horizontalColumn = 1; // Column A for horizontal spacing
                    int verticalColumn = 2; // Column B for vertical spacing

                    List<Tuple<double, double>> gridSpacingList = new List<Tuple<double, double>>();

                    for (int row = startRow; row <= worksheet.Dimension.End.Row; row++)
                    {
                        double xGridSpacing = double.Parse(worksheet.Cells[row, horizontalColumn].Value.ToString());
                        double yGridSpacing = double.Parse(worksheet.Cells[row, verticalColumn].Value.ToString());

                        gridSpacingList.Add(new Tuple<double, double>(xGridSpacing, yGridSpacing));
                    }

                    return gridSpacingList;
                }
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        // Helper function to get the ObjectId of the Centerline linetype
        private ObjectId GetCenterlineLinetypeId(Document doc, string lineTypeName)
        {
            // Get the current database
            Database db = doc.Database;

            // Open the LinetypeTable for read
            using (Transaction tr = db.TransactionManager.StartOpenCloseTransaction())
            {
                LinetypeTable lt = tr.GetObject(db.LinetypeTableId, OpenMode.ForRead) as LinetypeTable;

                // Check if the Centerline linetype exists
                if (lt.Has(lineTypeName))
                {
                    // Get the ObjectId of the Centerline linetype
                    ObjectId centerlineLtypeId = lt[lineTypeName];
                    tr.Commit();
                    return centerlineLtypeId;
                }
            }

            // If the Centerline linetype does not exist, return ObjectId.Null
            return ObjectId.Null;
        }
    }
}

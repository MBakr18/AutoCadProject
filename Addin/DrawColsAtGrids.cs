using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using AddIn;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Exception = System.Exception;
using Line = Autodesk.AutoCAD.DatabaseServices.Line;


namespace CadProject
{
    public class DrawColsAtGrids : ICadCommand
    {
        public override void Execute()
        {
            FindIntersection();
        }

        private void FindIntersection()
        {

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor editor = doc.Editor;

            try
            {
                using (doc.LockDocument())
                {
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        // Create a selection filter for lines
                        TypedValue[] filterList = new TypedValue[]
                        {
                            new TypedValue((int)DxfCode.Operator, "<AND"),
                            new TypedValue((int)DxfCode.Start, "LINE"),
                            new TypedValue((int)DxfCode.Operator, "AND>"),
                        };

                        SelectionFilter filter = new SelectionFilter(filterList);

                        // Prompt the user to select all lines
                        PromptSelectionResult selectionResult = editor.SelectAll(filter);

                        if (selectionResult.Status == PromptStatus.OK)
                        {
                            // Prompt the user to specify the column size
                            PromptDoubleOptions options = new PromptDoubleOptions("\nEnter the column width: ");
                            options.AllowNegative = false;
                            options.AllowZero = false;
                            PromptDoubleResult widthResult = editor.GetDouble(options);

                            if (widthResult.Status != PromptStatus.OK)
                                return; // User canceled

                            // Prompt the user to specify the column size
                            PromptDoubleOptions options2 = new PromptDoubleOptions("\nEnter the column height: ");
                            options.AllowNegative = false;
                            options.AllowZero = false;
                            PromptDoubleResult heightResult = editor.GetDouble(options2);

                            if (heightResult.Status != PromptStatus.OK)
                                return; // User canceled

                            // Define the width and height of the column
                            double width = widthResult.Value; // Adjust as needed
                            double height = heightResult.Value; // Adjust as needed

                            SelectionSet selectionSet = selectionResult.Value;

                            foreach (var lineId in selectionSet.GetObjectIds())
                            {
                                Line line = tr.GetObject(lineId, OpenMode.ForRead) as Line;

                                // Check for intersections with other lines
                                foreach (var otherLineId in selectionSet.GetObjectIds())
                                {
                                    if (otherLineId != lineId)
                                    {
                                        Line otherLine = tr.GetObject(otherLineId, OpenMode.ForRead) as Line;

                                        Point3dCollection intersectionPoints = new Point3dCollection();
                                        line.IntersectWith(otherLine, Intersect.OnBothOperands, intersectionPoints,
                                            IntPtr.Zero, IntPtr.Zero);

                                        if (intersectionPoints != null)
                                        {
                                            BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                                            BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                                            foreach (Point3d point in intersectionPoints)
                                            {
                                                
                                                // Calculate the column's corner points
                                                Point3d topLeft = new Point3d(point.X - (width / 2), point.Y + (height / 2), 0);
                                                Point3d topRight = new Point3d(point.X + (width / 2), point.Y + (height / 2), 0);
                                                Point3d bottomRight = new Point3d(point.X + (width / 2), point.Y - (height / 2), 0);
                                                Point3d bottomLeft = new Point3d(point.X - (width / 2), point.Y - (height / 2), 0);

                                                // Create a rectangle by adding a closed Polyline
                                                using (Polyline rect = new Polyline())
                                                {
                                                    rect.AddVertexAt(0, new Point2d(topLeft.X, topLeft.Y), 0, 0, 0);
                                                    rect.AddVertexAt(1, new Point2d(topRight.X, topRight.Y), 0, 0, 0);
                                                    rect.AddVertexAt(2, new Point2d(bottomRight.X, bottomRight.Y), 0, 0, 0);
                                                    rect.AddVertexAt(3, new Point2d(bottomLeft.X, bottomLeft.Y), 0, 0, 0);
                                                    rect.Closed = true;

                                                    // Set the linetype to Solid
                                                    // rect.LinetypeId = GetSolidLinetypeId();

                                                    LinetypeTable lt = tr.GetObject(HostApplicationServices.WorkingDatabase.LinetypeTableId, OpenMode.ForRead) as LinetypeTable;

                                                    if (lt.Has("Continuous"))
                                                    {
                                                        rect.LinetypeId = lt["Continuous"];
                                                    }
                                                    // Set the color to Cyan
                                                    rect.Color = Color.FromColorIndex(ColorMethod.ByAci, 4); // Cyan is color index 4

                                                    // Add the rectangle to the block table record
                                                    btr.AppendEntity(rect);
                                                    tr.AddNewlyCreatedDBObject(rect, true);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        tr.Commit();
                    }
                }
            }
            catch (Exception ex)
            {
                // Handle any exceptions
                editor.WriteMessage($"Error finding intersections: {ex.Message}\n");
            }
        }

        private ObjectId GetSolidLinetypeId()
        {
            ObjectId solidLinetypeId = ObjectId.Null;

            using (Transaction tr = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction())
            {
                LinetypeTable lt = tr.GetObject(HostApplicationServices.WorkingDatabase.LinetypeTableId, OpenMode.ForRead) as LinetypeTable;

                if (lt.Has("SOLID"))
                {
                    solidLinetypeId = lt["SOLID"];
                }

                tr.Commit();
            }

            return solidLinetypeId;
        }
    }
}

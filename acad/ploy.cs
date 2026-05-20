using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

[assembly: CommandClass(typeof(PolylineEnclosure.PolylineEnclosureCommands))]

namespace PolylineEnclosure
{
    /// <summary>
    /// AutoCAD 2024 Plugin:
    /// Prüft ob eine gewählte geschlossene Polyline andere geschlossene Polylines einschließt.
    /// Falls ja, wird eine kreuzungsfreie Hüll-Polyline erstellt, die alle eingeschlossenen
    /// Polylines umfasst (Convex Hull über alle Punkte).
    /// </summary>
    public class PolylineEnclosureCommands
    {
        [CommandMethod("POLYEINSCHLUSS")]
        public void PolylineEnclosure()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            ed.WriteMessage("\n=== POLYLINE EINSCHLUSS-ANALYSE ===\n");

            // 1. Benutzer wählt die Haupt-Polyline
            PromptEntityOptions peo = new PromptEntityOptions(
                "\nHaupt-Polyline auswählen (muss geschlossen sein): ");
            peo.SetRejectMessage("\nNur Polylines erlaubt.");
            peo.AddAllowedClass(typeof(Polyline), true);

            PromptEntityResult per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK) return;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    // 2. Haupt-Polyline laden
                    Polyline mainPoly = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Polyline;
                    if (mainPoly == null || !mainPoly.Closed)
                    {
                        ed.WriteMessage("\nFehler: Die gewählte Polyline ist nicht geschlossen!");
                        tr.Abort();
                        return;
                    }

                    ed.WriteMessage($"\nHaupt-Polyline geladen ({mainPoly.NumberOfVertices} Eckpunkte).");

                    // 3. Alle anderen geschlossenen Polylines im Modellbereich finden
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord btr = tr.GetObject(
                        bt[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;

                    List<Polyline> candidatePolys = new List<Polyline>();

                    foreach (ObjectId objId in btr)
                    {
                        if (objId == per.ObjectId) continue;

                        DBObject obj = tr.GetObject(objId, OpenMode.ForRead);
                        if (obj is Polyline poly && poly.Closed)
                        {
                            candidatePolys.Add(poly);
                        }
                    }

                    ed.WriteMessage($"\n{candidatePolys.Count} weitere geschlossene Polyline(s) gefunden.");

                    // 4. Prüfen welche Polylines innerhalb der Haupt-Polyline liegen
                    List<Polyline> enclosedPolys = new List<Polyline>();

                    foreach (Polyline candidate in candidatePolys)
                    {
                        if (IsPolylineInsidePolyline(candidate, mainPoly))
                        {
                            enclosedPolys.Add(candidate);
                        }
                    }

                    if (enclosedPolys.Count == 0)
                    {
                        ed.WriteMessage("\nKeine eingeschlossenen Polylines gefunden. Keine Aktion nötig.");
                        tr.Commit();
                        return;
                    }

                    ed.WriteMessage($"\n{enclosedPolys.Count} eingeschlossene Polyline(s) gefunden!");

                    // 5. Alle Punkte der eingeschlossenen Polylines + Hauptpolyline sammeln
                    List<Point2d> allPoints = new List<Point2d>();

                    // Punkte der Hauptpolyline
                    for (int i = 0; i < mainPoly.NumberOfVertices; i++)
                        allPoints.Add(new Point2d(mainPoly.GetPoint2dAt(i).X, mainPoly.GetPoint2dAt(i).Y));

                    // Punkte aller eingeschlossenen Polylines
                    foreach (Polyline poly in enclosedPolys)
                    {
                        for (int i = 0; i < poly.NumberOfVertices; i++)
                            allPoints.Add(new Point2d(poly.GetPoint2dAt(i).X, poly.GetPoint2dAt(i).Y));
                    }

                    ed.WriteMessage($"\nInsgesamt {allPoints.Count} Punkte für Hüll-Berechnung.");

                    // 6. Convex Hull berechnen (Graham Scan)
                    List<Point2d> hullPoints = ComputeConvexHull(allPoints);

                    ed.WriteMessage($"\nConvex Hull: {hullPoints.Count} Eckpunkte.");

                    // 7. Hüll-Polyline erstellen
                    BlockTableRecord btrWrite = tr.GetObject(
                        bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    Polyline hullPoly = new Polyline();
                    hullPoly.SetDatabaseDefaults();
                    hullPoly.Closed = true;

                    for (int i = 0; i < hullPoints.Count; i++)
                        hullPoly.AddVertexAt(i, hullPoints[i], 0, 0, 0);

                    // Optional: Farbe der Hüll-Polyline auf Rot setzen
                    hullPoly.ColorIndex = 1;

                    btrWrite.AppendEntity(hullPoly);
                    tr.AddNewlyCreatedDBObject(hullPoly, true);

                    ed.WriteMessage("\n✓ Hüll-Polyline erfolgreich erstellt (rot, geschlossen, kreuzungsfrei).");
                    ed.WriteMessage($"\n  Layer: {hullPoly.Layer}");
                    ed.WriteMessage($"\n  Eckpunkte: {hullPoly.NumberOfVertices}");

                    tr.Commit();
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nFehler: {ex.Message}");
                    tr.Abort();
                }
            }
        }

        // ─────────────────────────────────────────────────────────────────────────
        // HILFSMETHODEN
        // ─────────────────────────────────────────────────────────────────────────

        /// <summary>
        /// Prüft ob eine Polyline vollständig innerhalb einer anderen liegt.
        /// Verwendet Point-in-Polygon (Ray Casting) für alle Eckpunkte.
        /// </summary>
        private bool IsPolylineInsidePolyline(Polyline inner, Polyline outer)
        {
            // Alle Eckpunkte der inneren Polyline müssen im Inneren der äußeren liegen
            for (int i = 0; i < inner.NumberOfVertices; i++)
            {
                Point2d pt = inner.GetPoint2dAt(i);
                if (!IsPointInPolygon(pt, outer))
                    return false;
            }
            return true;
        }

        /// <summary>
        /// Ray-Casting Algorithmus: Punkt-in-Polygon Test.
        /// Gibt true zurück wenn der Punkt innerhalb der Polyline liegt.
        /// </summary>
        private bool IsPointInPolygon(Point2d point, Polyline polygon)
        {
            int n = polygon.NumberOfVertices;
            bool inside = false;
            double px = point.X, py = point.Y;

            int j = n - 1;
            for (int i = 0; i < n; i++)
            {
                Point2d vi = polygon.GetPoint2dAt(i);
                Point2d vj = polygon.GetPoint2dAt(j);

                double xi = vi.X, yi = vi.Y;
                double xj = vj.X, yj = vj.Y;

                if (((yi > py) != (yj > py)) &&
                    (px < (xj - xi) * (py - yi) / (yj - yi) + xi))
                {
                    inside = !inside;
                }
                j = i;
            }
            return inside;
        }

        /// <summary>
        /// Graham Scan Algorithmus für die konvexe Hülle.
        /// Liefert die Punkte in Gegenuhrzeigersinn-Reihenfolge.
        /// </summary>
        private List<Point2d> ComputeConvexHull(List<Point2d> points)
        {
            if (points.Count < 3)
                return new List<Point2d>(points);

            // Duplikate entfernen
            points = points.Distinct(new Point2dComparer()).ToList();

            if (points.Count < 3)
                return new List<Point2d>(points);

            // Startpunkt: niedrigster Y-Wert (bei Gleichheit: kleinster X-Wert)
            Point2d pivot = points.OrderBy(p => p.Y).ThenBy(p => p.X).First();

            // Punkte nach Polarwinkel zum Pivot sortieren
            List<Point2d> sorted = points
                .Where(p => p != pivot)
                .OrderBy(p => Math.Atan2(p.Y - pivot.Y, p.X - pivot.X))
                .ThenBy(p => Distance(p, pivot))
                .ToList();

            // Graham Scan
            Stack<Point2d> stack = new Stack<Point2d>();
            stack.Push(pivot);
            stack.Push(sorted[0]);

            for (int i = 1; i < sorted.Count; i++)
            {
                while (stack.Count >= 2)
                {
                    Point2d top = stack.Pop();
                    Point2d nextToTop = stack.Peek();

                    if (CrossProduct(nextToTop, top, sorted[i]) > 0)
                    {
                        stack.Push(top);
                        break;
                    }
                    // top wird verworfen (nicht Teil der konvexen Hülle)
                }
                stack.Push(sorted[i]);
            }

            return stack.ToList();
        }

        /// <summary>
        /// Kreuzprodukt für drei Punkte (Orientierungstest).
        /// Positiv = Linkskurve, Negativ = Rechtskurve, 0 = kollinear.
        /// </summary>
        private double CrossProduct(Point2d o, Point2d a, Point2d b)
        {
            return (a.X - o.X) * (b.Y - o.Y) - (a.Y - o.Y) * (b.X - o.X);
        }

        private double Distance(Point2d a, Point2d b)
        {
            double dx = a.X - b.X, dy = a.Y - b.Y;
            return Math.Sqrt(dx * dx + dy * dy);
        }

        // ─────────────────────────────────────────────────────────────────────────
        // Zweiter Befehl: Interaktive Auswahl mehrerer Polylines
        // ─────────────────────────────────────────────────────────────────────────

        [CommandMethod("POLYHUELLE")]
        public void PolylineHull()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            ed.WriteMessage("\n=== POLYLINE HÜLLE ERSTELLEN ===\n");

            // Benutzer wählt mehrere Polylines
            PromptSelectionOptions pso = new PromptSelectionOptions();
            pso.MessageForAdding = "\nPolylines auswählen die eingeschlossen werden sollen: ";

            SelectionFilter filter = new SelectionFilter(new TypedValue[]
            {
                new TypedValue((int)DxfCode.Start, "LWPOLYLINE")
            });

            PromptSelectionResult psr = ed.GetSelection(pso, filter);
            if (psr.Status != PromptStatus.OK) return;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    List<Point2d> allPoints = new List<Point2d>();
                    int closedCount = 0;

                    foreach (SelectedObject so in psr.Value)
                    {
                        Polyline poly = tr.GetObject(so.ObjectId, OpenMode.ForRead) as Polyline;
                        if (poly == null) continue;
                        if (!poly.Closed)
                        {
                            ed.WriteMessage($"\nHinweis: Eine Polyline ist nicht geschlossen - wird trotzdem einbezogen.");
                        }
                        else closedCount++;

                        for (int i = 0; i < poly.NumberOfVertices; i++)
                            allPoints.Add(poly.GetPoint2dAt(i));
                    }

                    ed.WriteMessage($"\n{closedCount} geschlossene Polyline(s) verarbeitet.");
                    ed.WriteMessage($"\n{allPoints.Count} Punkte gesammelt.");

                    if (allPoints.Count < 3)
                    {
                        ed.WriteMessage("\nZu wenige Punkte für eine Hülle.");
                        tr.Abort();
                        return;
                    }

                    // Convex Hull berechnen
                    List<Point2d> hullPoints = ComputeConvexHull(allPoints);

                    // Hüll-Polyline erstellen
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord btr = tr.GetObject(
                        bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    Polyline hullPoly = new Polyline();
                    hullPoly.SetDatabaseDefaults();
                    hullPoly.Closed = true;

                    for (int i = 0; i < hullPoints.Count; i++)
                        hullPoly.AddVertexAt(i, hullPoints[i], 0, 0, 0);

                    hullPoly.ColorIndex = 3; // Grün

                    btr.AppendEntity(hullPoly);
                    tr.AddNewlyCreatedDBObject(hullPoly, true);

                    ed.WriteMessage($"\n✓ Hüll-Polyline erstellt (grün).");
                    ed.WriteMessage($"\n  Eckpunkte der Hülle: {hullPoly.NumberOfVertices}");

                    tr.Commit();
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nFehler: {ex.Message}");
                    tr.Abort();
                }
            }
        }

        // ─────────────────────────────────────────────────────────────────────────
        // Hilfsklasse für Point2d Vergleich
        // ─────────────────────────────────────────────────────────────────────────
        private class Point2dComparer : IEqualityComparer<Point2d>
        {
            private const double Tolerance = 1e-10;
            public bool Equals(Point2d a, Point2d b) =>
                Math.Abs(a.X - b.X) < Tolerance && Math.Abs(a.Y - b.Y) < Tolerance;
            public int GetHashCode(Point2d p) =>
                Math.Round(p.X, 8).GetHashCode() ^ Math.Round(p.Y, 8).GetHashCode();
        }
    }
}
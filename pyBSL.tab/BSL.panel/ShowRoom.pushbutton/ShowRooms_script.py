# -*- coding: utf-8 -*-
"""3D Room  Highlighter"""

__title__ = "3D Room\nHighlight"
__author__ = "Holger Rasch"

from pyrevit import script
from pyrevit.forms import WPFWindow

class RoomHighlightWindow(WPFWindow):
    def __init__(self, xaml_file_name):
        WPFWindow.__init__(self, xaml_file_name)
        
        from pyrevit import revit, DB, script
        self.doc = revit.doc
        self.uidoc = revit.uidoc
        self.DB = DB
        self.revit = revit
        self.logger = script.get_logger()
        self.preview_server = None
        """important to catch the closing by the cross in the right corner of the form"""
        self.Closed += self.on_window_closed
        
        """Initialize DC3D Server"""
        try:
            self.preview_server = self.revit.dc3dserver.Server(
                uidoc=self.uidoc,
                name="Room Preview",
                description="Preview for rooms in 3D",
            )
        except Exception:
            self.logger.warning("Could not initialize DC3D server")

    def build_big_mesh(self, mesh_or_meshes):
        new_triangles = []
        new_edges = []
        for m in mesh_or_meshes:
            new_triangles.extend(m.triangles)
            new_edges.extend(m.edges)

        bigmesh = self.revit.dc3dserver.Mesh ( new_edges, new_triangles)

        return bigmesh

    def show_preview_mesh(self, mesh):
        """Show the Preview Mesh"""
        if mesh:
            self.preview_server.meshes = [mesh]
            self.uidoc.RefreshActiveView()

    def highlight_button_click(self, sender, e):
        """Event Handler Highlight Button (unterstützt einzelne und kommagetrennte Raumnummern)."""
        try:
            raw_input = (self.room_number_input.Text or "").strip()
            if not raw_input:
                self.status_text.Text = "Bitte mindestens eine Raumnummer eingeben."
                return

            # Input: “101, 102, 103” -> {“101”, ‘102’, “103”}
            # Optional: Also allow semicolons
            raw_input = raw_input.replace(";", ",")
            requested_numbers = [
                part.strip()
                for part in raw_input.split(",")
                if part.strip()
            ]

            if not requested_numbers:
                self.status_text.Text = "Bitte gültige Raumnummern eingeben."
                return

            requested_set = set(requested_numbers)
            
            # Collect all rooms once, but only remember the relevant ones.
            collector = (
                self.DB.FilteredElementCollector(self.doc)
                .OfCategory(self.DB.BuiltInCategory.OST_Rooms)
                .WhereElementIsNotElementType()
            )
            print (requested_set)
            rooms_to_highlight = []
            found_numbers = set()

            for room in collector:
                room_number = getattr(room, "Number", None)
                if not room_number:
                    continue

                len_requested = len(requested_set)

                if room_number in requested_set:
                    rooms_to_highlight.append(room)
                    found_numbers.add(room_number)
                    requested_set.remove(room_number)
                    # Performance: Once we have found all requested numbers, do not continue running through the model.
                    if len(found_numbers) == len_requested:
                        break

            if not rooms_to_highlight:
                self.status_text.Text = "Keiner der angegebenen Räume wurde im Modell gefunden."
                return
            print (rooms_to_highlight)
            # Create mesh for all found rooms
            meshes = []
            for room in rooms_to_highlight:
                room_solid = self.get_room_solid(room)
                if not room_solid:
                    continue

                room_mesh = self.create_preview_mesh_from_solid(room_solid)
                colored_mesh = self.recolor_mesh(room_mesh)
                meshes.append(colored_mesh)

            if not meshes:
                self.status_text.Text = "Für die gefundenen Räume konnte keine Geometrie erzeugt werden."
                return
            
            bigmesh = self.build_big_mesh(meshes)

            self.show_preview_mesh(bigmesh)

            missing = requested_set - found_numbers
            if missing:
                self.status_text.Text = (
                    "Es wurden {found} Räume hervorgehoben. "
                    "Nicht gefunden: {missing}"
                ).format(
                    found=len(found_numbers),
                    missing=", ".join(sorted(missing)),
                )
            else:
                self.status_text.Text = (
                    "Es wurden {found} Räume hervorgehoben."
                ).format(found=len(found_numbers))

        except Exception as ex:
            self.status_text.Text = "Error: {}".format(str(ex))
            self.logger.error("Error when highlighting: {}".format(ex))

    def highlight_button_click_old(self, sender, e):
        """Event Handler Highlight Button"""
        try:
            """room number from input"""
            room_number_to_find = self.room_number_input.Text.strip()
            
            if not room_number_to_find:
                self.status_text.Text = "Please enter a room number!"
                return
            
            """all rooms"""
            collector = self.DB.FilteredElementCollector(self.doc)\
                          .OfCategory(self.DB.BuiltInCategory.OST_Rooms)\
                          .WhereElementIsNotElementType()
            
            room_found = False
            room_to_highlight = None
            
            """looking for room with number"""
            for room in collector:
                room_number = room.Number
                if room_number == room_number_to_find:
                    room_to_highlight = room
                    room_found = True
                    break
            
            if room_found:
                room_solid = self.get_room_solid(room_to_highlight)
                room_mesh = self.create_preview_mesh_from_solid(room_solid)
                """The mesh has a useless color; we need a transparent color."""
                colored_mesh = self.recolor_mesh(room_mesh)
                self.show_preview_mesh(colored_mesh)
                self.status_text.Text = "Room ‘{}’ has been highlighted!".format(room_number_to_find)
            else:
                self.status_text.Text = "Raum '{}' wurde nicht gefunden!".format(room_number_to_find)
        
        except Exception as ex:
            self.status_text.Text = "Error: {}".format(str(ex))
            self.logger.error("Error when highlighting: {}".format(ex))
    
    def clear_button_click(self, sender, e):
        """Event Handler Clear Button"""
        self.clear_preview()
        self.status_text.Text = "Preview was removed."
    
    def close_button_click(self, sender, e):
        """Event Handler Close Button"""
        self.clear_preview()
        self.Close()

    def get_room_solid(self, room):
        """Calculate a room's geometry and find its boundary faces"""
        from Autodesk.Revit.DB import SpatialElementGeometryCalculator
        calculator = SpatialElementGeometryCalculator(self.doc)
        results = calculator.CalculateSpatialElementGeometry(room)
        roomSolid = results.GetGeometry()
        return roomSolid
    
    def clear_preview(self):
        """remove all geometry from Preview Server"""
        if self.preview_server:
            self.preview_server.meshes = []
            self.uidoc.RefreshActiveView()

    def on_window_closed(self, sender, e):
        """the close event from the closing cross."""
        self.clear_preview()
        
    def create_preview_mesh_from_solid(self, solid):
        """Create a mesh for DC3D Preview."""
        try:
            mesh = self.revit.dc3dserver.Mesh.from_solid(
                self.doc,
                solid
            )
            return mesh
        except Exception as e:
            self.logger.error("Error creating preview mesh: {}".format(e))
            return None
        
    def recolor_edges(self,edges, color):
        """Replace the color of Edges  with a new color."""
        new_edges = []
        for ed in edges:
            new_edge = type(ed)(
                ed.a,
                ed.b,
                color   
            )
            new_edges.append(new_edge)

        return new_edges    

    def recolor_triangles(self, triangles, color):
        """give all triangles a new color"""
        new_triangles = []
        for tri in triangles:
            new_tri = type(tri)(
                tri.a,
                tri.b,
                tri.c,
                tri.normal,
                color
            )
            new_triangles.append(new_tri)

        return new_triangles

    def get_SelectionColor(self, tranparency=200):
            """Get the revit option color for the selection as."""
            color_options = self.DB.ColorOptions.GetColorOptions()
            selection_color = color_options.SelectionColor

            SelectionSemitransparent = color_options.SelectionSemitransparent

            if SelectionSemitransparent:
                alpha = tranparency
            else:
                alpha = 0

            new_color = self.DB.ColorWithTransparency(
                selection_color.Red, 
                selection_color.Green, 
                selection_color.Blue, 
                alpha)
            
            return new_color

    def recolor_mesh(self, mesh, new_color=None):
        """Uniformly recolors the mesh."""
        if new_color is None:
            """Uses the same color that Revit uses."""
            #new_color = self.DB.ColorWithTransparency(100, 150, 255, 150)
            new_color = self.get_SelectionColor()

        new_triangles = self.recolor_triangles(mesh.triangles, new_color)
            
        new_edges = self.recolor_edges(mesh.edges, new_color)

        return type(mesh)(new_edges, new_triangles)

if __name__ == '__main__':

    xaml_file = script.get_bundle_file('ShowRoom.xaml')
    window = RoomHighlightWindow(xaml_file)

    window.Show()
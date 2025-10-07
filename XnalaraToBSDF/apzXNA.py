bl_info = {
    "name": "apzTools Material CSV",
    "author": "Homar",
    "version": (1, 0),
    "blender": (4, 4, 0),
    "location": "View3D > Sidebar > apzTools",
    "description": "Export and rebuild materials from CSV, with global sliders",
    "category": "Material",
}

import bpy
import csv
import os

# ---------------------- Export Script ----------------------
def export_material_csv():
    output_path = bpy.path.abspath("//lista.csv")

    def get_texture_name_from_input(mat, input_name):
        if not mat.use_nodes:
            return ""
        for node in mat.node_tree.nodes:
            if node.type == 'GROUP' and input_name in node.inputs:
                socket = node.inputs[input_name]
                if socket.is_linked:
                    from_node = socket.links[0].from_node
                    if from_node.type == 'TEX_IMAGE' and from_node.image:
                        return from_node.image.name
        return ""

    used_materials = set()
    for obj in bpy.context.scene.objects:
        for slot in obj.material_slots:
            if slot.material:
                used_materials.add(slot.material)

    header = ["Material","Diffuse","Lightmap","Specular","Emission","Bump Map","Environment","Alpha"]
    rows = [header]

    for mat in used_materials:
        row = [
            mat.name,
            get_texture_name_from_input(mat, "Diffuse"),
            get_texture_name_from_input(mat, "Lightmap"),
            get_texture_name_from_input(mat, "Specular"),
            get_texture_name_from_input(mat, "Emission"),
            get_texture_name_from_input(mat, "Bump Map"),
            get_texture_name_from_input(mat, "Environment"),
            get_texture_name_from_input(mat, "Alpha")
        ]
        rows.append(row)

    with open(output_path, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerows(rows)

    print(f"CSV saved to: {output_path}")

# ---------------------- Import Script ----------------------
def import_material_csv():
    blend_dir = bpy.path.abspath("//")
    csv_path = os.path.join(blend_dir, "listaOUT.csv")

    with open(csv_path, newline='', encoding='utf-8') as csvfile:
        reader = csv.reader(csvfile)
        next(reader)
        for row in reader:
            mat_name = row[0].strip()
            diffuse_name   = row[1].strip() if len(row) > 1 else ""
            roughness_name = row[3].strip() if len(row) > 3 else ""
            metallic_name  = row[4].strip() if len(row) > 4 else ""
            normal_name    = row[5].strip() if len(row) > 5 else ""
            emission_name  = row[6].strip() if len(row) > 6 else ""
            alpha_name     = row[7].strip() if len(row) > 7 else ""

            mat = bpy.data.materials.get(mat_name)
            if not mat:
                print(f"Material '{mat_name}' not found.")
                continue
            if not mat.use_nodes:
                mat.use_nodes = True

            nodes = mat.node_tree.nodes
            links = mat.node_tree.links

            for node in nodes:
                if node.type == 'BSDF_PRINCIPLED':
                    nodes.remove(node)

            bsdf_node = nodes.new(type='ShaderNodeBsdfPrincipled')
            bsdf_node.location = (400, 0)

            output_node = next((n for n in nodes if n.type == 'OUTPUT_MATERIAL'), None)
            if output_node:
                links.new(bsdf_node.outputs['BSDF'], output_node.inputs['Surface'])

            def create_image_node(name, location, non_color=False):
                tex_node = nodes.new(type='ShaderNodeTexImage')
                tex_node.location = location
                if name:
                    tex_path = os.path.join(blend_dir, name)
                    if os.path.exists(tex_path):
                        image = bpy.data.images.load(tex_path)
                        tex_node.image = image
                        if non_color:
                            tex_node.image.colorspace_settings.name = 'Non-Color'
                    else:
                        print(f"Texture file '{name}' not found for material '{mat_name}'.")
                return tex_node

            diffuse_node = create_image_node(diffuse_name, (-60, 700))
            if diffuse_node.image:
                links.new(diffuse_node.outputs['Color'], bsdf_node.inputs['Base Color'])

            rough_node = create_image_node(roughness_name, (-60, 100), non_color=True)
            if rough_node.image:
                links.new(rough_node.outputs['Color'], bsdf_node.inputs['Roughness'])

            metallic_node = create_image_node(metallic_name, (-60, 400), non_color=True)
            if metallic_node.image:
                links.new(metallic_node.outputs['Color'], bsdf_node.inputs['Metallic'])

            normal_tex_node = create_image_node(normal_name, (-350, -500), non_color=True)
            normal_map_node = nodes.new(type='ShaderNodeNormalMap')
            normal_map_node.location = (-60, -500)
            links.new(normal_tex_node.outputs['Color'], normal_map_node.inputs['Color'])
            if normal_tex_node.image:
                links.new(normal_map_node.outputs['Normal'], bsdf_node.inputs['Normal'])

            emission_node = create_image_node(emission_name, (-60, -800), non_color=False)
            if emission_node.image:
                links.new(emission_node.outputs['Color'], bsdf_node.inputs['Emission'])

            alpha_node = create_image_node(alpha_name, (-60, -200), non_color=True)
            if alpha_node.image:
                links.new(alpha_node.outputs['Color'], bsdf_node.inputs['Alpha'])

    print("All materials and texture nodes created.")

# ---------------------- Global Slider Logic ----------------------
def update_all_materials(self, context):
    for mat in bpy.data.materials:
        if mat.use_nodes:
            for node in mat.node_tree.nodes:
                if node.type == 'BSDF_PRINCIPLED':
                    node.inputs['Roughness'].default_value = context.scene.apztools_roughness
                    node.inputs['Metallic'].default_value = context.scene.apztools_metallic

def register_properties():
    bpy.types.Scene.apztools_roughness = bpy.props.FloatProperty(
        name="Roughness",
        description="Global Roughness for all materials",
        min=0.0, max=1.0,
        default=0.5,
        update=update_all_materials
    )
    bpy.types.Scene.apztools_metallic = bpy.props.FloatProperty(
        name="Metallic",
        description="Global Metallic for all materials",
        min=0.0, max=1.0,
        default=0.0,
        update=update_all_materials
    )

def unregister_properties():
    del bpy.types.Scene.apztools_roughness
    del bpy.types.Scene.apztools_metallic

# ---------------------- UI Panel ----------------------
class APZTOOLS_PT_panel(bpy.types.Panel):
    bl_label = "Material CSV Tools"
    bl_idname = "APZTOOLS_PT_panel"
    bl_space_type = 'VIEW_3D'
    bl_region_type = 'UI'
    bl_category = 'apzTools'

    def draw(self, context):
        layout = self.layout
        layout.label(text="Export / Import Materials")
        layout.operator("apztools.export_csv", icon='EXPORT')
        layout.operator("apztools.import_csv", icon='IMPORT')
        layout.separator()
        layout.label(text="Global Material Controls")
        layout.prop(context.scene, "apztools_roughness", slider=True)
        layout.prop(context.scene, "apztools_metallic", slider=True)

# ---------------------- Operators ----------------------
class APZTOOLS_OT_export_csv(bpy.types.Operator):
    bl_idname = "apztools.export_csv"
    bl_label = "Export Material CSV"

    def execute(self, context):
        export_material_csv()
        self.report({'INFO'}, "Material CSV exported.")
        return {'FINISHED'}

class APZTOOLS_OT_import_csv(bpy.types.Operator):
    bl_idname = "apztools.import_csv"
    bl_label = "Import Material CSV"

    def execute(self, context):
        import_material_csv()
        self.report({'INFO'}, "Material CSV imported.")
        return {'FINISHED'}

# ---------------------- Registration ----------------------
classes = [
    APZTOOLS_PT_panel,
    APZTOOLS_OT_export_csv,
    APZTOOLS_OT_import_csv,
]

def register():
    for cls in classes:
        bpy.utils.register_class(cls)
    register_properties()

def unregister():
    for cls in reversed(classes):
        bpy.utils.unregister_class(cls)
    unregister_properties()

if __name__ == "__main__":
    register()